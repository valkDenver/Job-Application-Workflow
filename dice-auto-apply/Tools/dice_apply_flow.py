"""The Playwright page flow for one Dice job (design section 8). Imported only by commands that open Dice.

Session opens the installed Google Chrome with its own profile (never the user's everyday profile), watches every dice.com response
for block signals, and runs apply_job() as a state machine: job page -> Apply -> resume step (File options -> Replace) -> questions ->
review -> submit. The script never guesses: any unexpected state ends that job with a result code and a screenshot.
"""
import math
import random
import time
from pathlib import Path

from playwright.sync_api import Error as PwError
from playwright.sync_api import TimeoutError as PwTimeout
from playwright.sync_api import sync_playwright

from dice_apply_common import label_key
import dice_apply_guard as G
import dice_apply_selectors as S

CHROME_ARGS = ["--disable-background-timer-throttling", "--disable-backgrounding-occluded-windows", "--disable-renderer-backgrounding"]


class Stop(Exception):
    """Ends the whole round: kind is BLOCKED, LOGIN_REQUIRED, DAILY_LIMIT_DICE or BROWSER_CLOSED."""

    def __init__(self, kind, detail=""):
        super().__init__(f"{kind}: {detail}")
        self.kind, self.detail = kind, detail
        self.unconfirmed = None  # the Job ID when the Stop came after the Submit click


class JobEnd(Exception):
    """Ends one job with a result code (STUCK, UI_CHANGED, JOB_GONE, ...)."""

    def __init__(self, code, detail="", step=""):
        super().__init__(f"{code}: {detail}")
        self.code, self.detail, self.step = code, detail, step


class Outcome:
    def __init__(self, code, detail="", shot="", wizard_opened=False, submitted=False, url=""):
        self.code, self.detail, self.shot, self.wizard_opened, self.submitted, self.url = code, detail, shot, wizard_opened, submitted, url


def _off_dice(url):
    from urllib.parse import urlparse
    host = urlparse(url or "").hostname or ""
    return bool(host) and not (host == "dice.com" or host.endswith(".dice.com"))


def _closed(e):
    return "closed" in str(e).lower() and ("target" in str(e).lower() or "browser" in str(e).lower() or "context" in str(e).lower())


class Session:
    def __init__(self, cfg, rt, run_dir=None, signals=None, log=print):
        self.cfg, self.rt, self.run_dir, self.log = cfg, Path(rt), Path(run_dir) if run_dir else None, log
        self.signals = signals or G.SoftSignals()
        self.hard = ""
        self.soft_stop = False
        self.pw = self.ctx = self.page = None
        self.stuck_ms = int(cfg.get("stuck_timeout_sec", 15) * 1000)

    # ---------- browser ----------
    def open(self):
        self.pw = sync_playwright().start()
        profile = self.rt / "browser-profile"
        profile.mkdir(parents=True, exist_ok=True)
        try:
            self.ctx = self.pw.chromium.launch_persistent_context(str(profile), channel="chrome", headless=False, no_viewport=True, args=CHROME_ARGS)
        except PwError as e:
            self.pw.stop()
            raise Stop("CHROME_MISSING", f"Google Chrome could not be started ({str(e).splitlines()[0]}). Install Chrome, then run doctor.")
        self.page = self.ctx.pages[0] if self.ctx.pages else self.ctx.new_page()
        self.page.set_default_timeout(self.stuck_ms)
        self.ctx.on("response", self._on_response)
        return self

    def close(self):
        for f in (lambda: self.ctx and self.ctx.close(), lambda: self.pw and self.pw.stop()):
            try:
                f()
            except Exception:
                pass

    def _on_response(self, resp):
        try:
            kind = G.classify_response(resp.url, resp.status, resp.request.resource_type)
        except Exception:
            return
        if kind == "hard" and not self.hard:
            self.hard = f"HTTP {resp.status} from {resp.url[:120]}"
        elif kind == "soft":
            if self.signals.add():
                self.soft_stop = True

    # ---------- pacing ----------
    def pause(self, scale=1.0):
        d = self.cfg.get("delay_between_actions_ms") or {"median": 900, "min": 400, "max": 3000}
        ms = random.lognormvariate(math.log(d["median"]), 0.45)
        ms = min(max(ms, d["min"]), d["max"]) * scale * self.signals.factor()
        time.sleep(ms / 1000)

    def type_text(self, loc, value):
        lo, hi = self.cfg.get("typing_ms") or [60, 180]
        loc.click()
        loc.press_sequentially(str(value), delay=random.randint(lo, hi))

    # ---------- checks ----------
    def signal_text(self):
        p = self.page
        try:
            title = p.title()
            heads = " ".join(p.locator("h1, h2, h3").all_inner_texts()[:15])
            body = p.locator("body").inner_text(timeout=3000)
        except PwError as e:
            if _closed(e):
                raise Stop("BROWSER_CLOSED", "the Chrome window was closed")
            return "", ""
        return f"{title} {heads}", body

    def visible_frames(self):
        """(url, width, height) of each child frame as shown on the page; hidden frames report 0 x 0."""
        out = []
        for f in self.page.frames[1:]:
            try:
                el = f.frame_element()
                box = el.bounding_box() if el.is_visible() else None
            except PwError:
                box = None
            out.append((f.url, box["width"] if box else 0, box["height"] if box else 0))
        return out

    def check_block(self):
        if self.hard:
            raise Stop("BLOCKED", self.hard)
        if self.soft_stop:
            raise Stop("SOFT_STOP", "three slow or failing Dice responses within an hour")
        head, body = self.signal_text()
        c = G.classify_text(head, body) or G.classify_frames(self.visible_frames())
        if c.startswith("hard:"):
            raise Stop("BLOCKED", c[5:])
        if c == "limit":
            raise Stop("DAILY_LIMIT_DICE", "Dice says a daily or application limit was reached")
        return body

    def visible(self, loc, timeout=1500):
        try:
            loc.wait_for(state="visible", timeout=timeout)
            return True
        except PwTimeout:
            return False

    def logged_in(self):
        return not self.visible(S.login_link(self.page), 2500)

    def shot(self, job_id, step):
        if not self.run_dir:
            return ""
        path = self.run_dir / f"{job_id}_{step}.png"
        for full in (True, False):  # the review's resume line is below the fold; a page still loading can refuse a full-page shot
            try:
                self.page.screenshot(path=str(path), full_page=full)
                return path.name
            except Exception:
                time.sleep(1)
        return ""

    def goto(self, url):
        t = time.time()
        try:
            self.page.goto(url, wait_until="domcontentloaded", timeout=self.stuck_ms * 2)
        except PwTimeout:
            raise JobEnd("STUCK", f"page did not load: {url}", "open")
        if time.time() - t > (self.cfg.get("slow_load_sec") or 10):
            if self.signals.add():
                self.soft_stop = True

    # ---------- login ----------
    def wait_for_login(self, minutes=10):
        self.goto(S.HOME_URL)
        end = time.time() + minutes * 60
        while time.time() < end:
            if self.logged_in():
                return True
            time.sleep(3)
        return False

    # ---------- one job ----------
    def apply_job(self, job, resume, mode, answers, before_submit, confirm=None):
        """Returns an Outcome; raises Stop for round-ending signals."""
        jid, t0 = job["Job ID"], time.time()
        state = {"wizard": False, "submitted": False}

        def budget(step):
            if time.time() - t0 > (self.cfg.get("job_timeout_sec") or 180):
                raise JobEnd("STUCK", "the job took longer than its time limit", step)

        def after_submit(detail):
            # once Submit was clicked, every ending except a seen success page is SUBMIT_UNCONFIRMED: ledgered, never retried
            return Outcome("SUBMIT_UNCONFIRMED", detail, self.shot(jid, "submit"), True, True)

        try:
            return self._apply(job, Path(resume), mode, answers, before_submit, confirm, state, budget)
        except Stop as e:
            if state["submitted"]:
                e.unconfirmed = jid  # the round records this job as SUBMIT_UNCONFIRMED before it stops
            raise
        except JobEnd as e:
            if state["submitted"]:
                return after_submit(f"{e.code} after Submit: {e.detail}")
            return Outcome(e.code, e.detail + (f" (step {e.step})" if e.step else ""), self.shot(jid, (e.step or "end").lower()),
                           state["wizard"], state["submitted"])
        except PwTimeout as e:
            if state["submitted"]:
                return after_submit("timed out after Submit")
            return Outcome("STUCK", str(e).splitlines()[0][:160], self.shot(jid, "stuck"), state["wizard"], state["submitted"])
        except PwError as e:
            if _closed(e):
                s = Stop("BROWSER_CLOSED", "the Chrome window was closed")
                if state["submitted"]:
                    s.unconfirmed = jid
                raise s
            if state["submitted"]:
                return after_submit("page error after Submit: " + str(e).splitlines()[0][:120])
            return Outcome("UI_CHANGED", str(e).splitlines()[0][:160], self.shot(jid, "error"), state["wizard"], state["submitted"])
        except Exception as e:
            if state["submitted"]:
                return after_submit(f"{type(e).__name__} after Submit")
            raise

    def _apply(self, job, resume, mode, answers, before_submit, confirm, state, budget):
        p, jid = self.page, job["Job ID"]
        self.goto(job["Job URL"])
        body = self.check_block()
        if not self.logged_in():
            # seen live 2026-11-01: after a while idle, Dice first renders the page logged out while it refreshes the session from its
            # stored token, then logs back in by itself. Give it a moment and one reload before calling the session lost.
            time.sleep(6)
            self.goto(job["Job URL"])
            body = self.check_block()
            if not self.logged_in():
                raise Stop("LOGIN_REQUIRED", "Dice shows Log in: the session is not logged in")
        low = body.lower()
        if any(t in low for t in S.GONE_TEXTS):
            raise JobEnd("JOB_GONE", "Dice says the job is no longer available")
        apply_btn = S.apply_button(p)
        has_apply = self.visible(apply_btn, self.stuck_ms)
        if not has_apply and self.visible(S.applied_badge(p), 1500):
            return Outcome("ALREADY_APPLIED_ON_DICE", "the job page shows Applied", self.shot(jid, "applied"))
        if not has_apply:
            raise JobEnd("UI_CHANGED", "no Apply button on the job page", "OPEN_JOB")

        # read the page like a person: dwell and scroll once
        lo, hi = self.cfg.get("job_page_dwell_sec") or [6, 20]
        time.sleep(random.uniform(lo, hi) * self.signals.factor())
        try:
            p.mouse.wheel(0, random.randint(400, 900))
        except PwError:
            pass
        self.pause()
        budget("OPEN_JOB")

        before_pages = set(self.ctx.pages)
        apply_btn.click()
        state["wizard"] = True
        end = time.time() + self.stuck_ms / 1000
        while time.time() < end:
            new = [x for x in self.ctx.pages if x not in before_pages]
            if new:
                url = new[0].url
                try:
                    stop_at = time.time() + 8
                    while (not url or url == "about:blank") and time.time() < stop_at:  # a new tab starts blank, then navigates
                        new[0].wait_for_load_state("domcontentloaded", timeout=5000)
                        url = new[0].url
                        time.sleep(0.2)
                    new[0].close()
                except PwError:
                    pass
                return Outcome("EXTERNAL_APPLY", f"url={url}", "", True, url=url)
            if _off_dice(p.url):
                return Outcome("EXTERNAL_APPLY", f"url={p.url}", "", True, url=p.url)
            # the wizard may open in place on the job page (seen 2026-11-01) or at its own /wizard address (seen 2026-10-30)
            if S.WIZARD_URL.search(p.url or "") or self.visible(p.get_by_text(S.STEP_RESUME_TEXT).first, 250):
                break
            time.sleep(0.25)
        else:
            raise JobEnd("STUCK", "the Apply click did not open the wizard", "APPLY")
        self.check_block()

        # WIZARD_STEP_RESUME
        if not self.visible(p.get_by_text(S.STEP_RESUME_TEXT).first, self.stuck_ms):
            raise JobEnd("UI_CHANGED", f'no "{S.STEP_RESUME_TEXT}" step', "WIZARD_STEP_RESUME")
        self.pause()
        # Module 03 names files <Candidate>_<Title>.docx, so another job (or a saved draft) can already show this exact name.
        # Then only the "New file" marker proves the upload happened; otherwise the visible name must appear after it.
        same_name_before = resume.name.lower() in p.locator("body").inner_text().lower()
        opts = S.file_options_button(p)
        if self.visible(opts, 3000):
            opts.click()
            self.pause(0.6)
            with p.expect_file_chooser(timeout=self.stuck_ms) as fc:
                S.replace_menuitem(p).click()
            fc.value.set_files(str(resume))
        else:
            zone = S.resume_upload_zone(p)
            if not self.visible(zone, 3000):
                raise JobEnd("UI_CHANGED", "neither File options nor a resume upload area", "WIZARD_STEP_RESUME")
            with p.expect_file_chooser(timeout=self.stuck_ms) as fc:
                zone.click()
            fc.value.set_files(str(resume))
        if not self.visible(p.get_by_text(resume.name, exact=False).first, self.stuck_ms):
            raise JobEnd("RESUME_NOT_ATTACHED", f"the page does not show {resume.name} after upload", "WIZARD_STEP_RESUME")
        new_marker = self.visible(p.get_by_text(S.NEW_FILE_TEXT).first, self.stuck_ms if same_name_before else 5000)
        if same_name_before and not new_marker:
            raise JobEnd("RESUME_NOT_ATTACHED", f'{resume.name} was already shown before the upload and Dice did not mark it "New file"; '
                         "the upload cannot be confirmed", "WIZARD_STEP_RESUME")
        self.check_block()
        budget("WIZARD_STEP_RESUME")
        self.pause()
        S.next_button(p).click()

        # QUESTIONS steps until REVIEW
        for _ in range(6):
            end = time.time() + self.stuck_ms / 1000
            while time.time() < end and not self.visible(p.get_by_text(S.STEP_REVIEW_TEXT).first, 500):
                if self.visible(S.next_button(p), 300) and not self.visible(p.get_by_text(S.STEP_RESUME_TEXT).first, 200):
                    break
            self.check_block()
            budget("QUESTIONS")
            if self.visible(p.get_by_text(S.STEP_REVIEW_TEXT).first, 500):
                break
            missing = self.answer_questions(answers)
            if missing:
                return Outcome("NEEDS_INPUT", "questions=" + "; ".join(f'"{m}"' for m in missing), self.shot(jid, "questions"), True)
            self.pause()
            S.next_button(p).click()
        else:
            raise JobEnd("UI_CHANGED", "the review step never appeared", "QUESTIONS")

        # REVIEW
        review_text = p.locator("body").inner_text()
        if resume.name.lower() not in review_text.lower():
            raise JobEnd("RESUME_NOT_ATTACHED", f"the review page does not show {resume.name}", "REVIEW")
        shot = self.shot(jid, "review")
        gate = before_submit()
        if gate:
            return Outcome(gate, "not submitted", shot, True)
        if mode == "dry-run":
            return Outcome("DRY_RUN_OK", "reached review; not submitted", shot, True)
        if mode == "confirm" and not (confirm and confirm(job)):
            return Outcome("DRY_RUN_OK", "not confirmed at the terminal; not submitted", shot, True)
        self.check_block()
        self.pause()
        S.submit_button(p).click()
        state["submitted"] = True
        end = time.time() + self.stuck_ms * 2 / 1000
        while time.time() < end:
            if self.visible(p.get_by_text(S.SUCCESS_TEXT, exact=False).first, 1000):  # the heading is required; the URL may not change
                return Outcome("APPLIED", "", self.shot(jid, "success"), True, True)
            self.check_block()  # a Stop here is marked SUBMIT_UNCONFIRMED by apply_job
            time.sleep(0.5)
        return Outcome("SUBMIT_UNCONFIRMED", "no success page after Submit", self.shot(jid, "submit"), True, True)

    # ---------- screening questions ----------
    def answer_questions(self, answers):
        """Fill only questions that have a saved answer. Returns the labels of required questions left unanswered."""
        p = self.page
        fields = p.locator("input:visible, select:visible, textarea:visible")
        missing, radios_done = [], set()
        for i in range(fields.count()):
            el = fields.nth(i)
            info = el.evaluate("""e => {
                const lab = (e.labels && e.labels[0] && e.labels[0].innerText) || e.getAttribute('aria-label')
                    || (e.getAttribute('aria-labelledby') && (e.getRootNode().getElementById(e.getAttribute('aria-labelledby')) || {}).innerText)
                    || (e.closest('fieldset') && e.closest('fieldset').querySelector('legend') && e.closest('fieldset').querySelector('legend').innerText)
                    || e.placeholder || '';
                const legend = e.closest('fieldset') && e.closest('fieldset').querySelector('legend');
                return {tag: e.tagName.toLowerCase(), type: (e.type || '').toLowerCase(), name: e.name || '', label: lab.trim(),
                        group: legend ? legend.innerText.trim() : '', value: e.value || '', checked: !!e.checked,
                        required: !!e.required || e.getAttribute('aria-required') === 'true'};
            }""")
            if info["type"] in ("hidden", "file", "submit", "button", "search"):
                continue
            if info["type"] == "radio":
                group = info["group"] or info["name"]
                if group in radios_done:
                    continue
                radios_done.add(group)
                radios = p.locator(f"input[type=radio][name='{info['name']}']") if info["name"] else el
                if any(radios.nth(k).is_checked() for k in range(radios.count())):
                    continue
                ans = answers.get(label_key(info["group"] or info["label"]))
                picked = False
                if ans:
                    for k in range(radios.count()):
                        r = radios.nth(k)
                        lab = r.evaluate("e => (e.labels && e.labels[0] && e.labels[0].innerText) || e.value || ''")
                        if label_key(lab) == label_key(ans):
                            self.pause(0.5)
                            r.check()
                            picked = True
                            break
                if not picked and info["required"]:
                    missing.append(info["group"] or info["label"] or info["name"])
                continue
            key = label_key(info["label"])
            ans = answers.get(key)
            if info["type"] == "checkbox":
                if ans and ans.lower() in ("yes", "true", "y", "1", "agree") and not info["checked"]:
                    self.pause(0.5)
                    el.check()
                elif info["required"] and not info["checked"]:
                    missing.append(info["label"] or info["name"])
                continue
            if info["value"]:
                continue  # already filled (from the Dice profile)
            if ans is None:
                if info["required"]:
                    missing.append(info["label"] or info["name"])
                continue
            self.pause(0.6)
            if info["tag"] == "select":
                try:
                    el.select_option(label=ans)
                except PwError:
                    el.select_option(value=ans)
            else:
                self.type_text(el, ans)
        return missing
