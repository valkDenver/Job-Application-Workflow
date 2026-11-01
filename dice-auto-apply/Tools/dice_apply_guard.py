"""Account-safety and run-state guards (design section 10 and 12).

Run lock, cooldown, stop request, live status, ledger, daily counters, the working window, and the rules that turn a page or an HTTP
response into a block signal. Everything here is plain Python with no browser, so it is unit-tested directly.
"""
import datetime as dt
import os
import re
import socket
import time
from pathlib import Path

from dice_apply_common import Refused, append_jsonl, now_iso, read_json, read_jsonl, today, write_json
from dice_apply_platform import pid_alive
import dice_apply_selectors as S


def _p(rt, name):
    return Path(rt) / name


# ---------- one round at a time ----------
def lock_info(rt):
    info = read_json(_p(rt, "run.lock"))
    if info and pid_alive(info.get("pid")):
        return info
    return None


def acquire_lock(rt):
    live = lock_info(rt)
    if live and live.get("pid") != os.getpid():
        raise Refused(f"ALREADY_RUNNING: a round started at {live.get('started')} is still running. Use status, or stop.")
    write_json(_p(rt, "run.lock"), {"pid": os.getpid(), "host": socket.gethostname(), "started": now_iso()})


def release_lock(rt):
    info = read_json(_p(rt, "run.lock"))
    if info and info.get("pid") == os.getpid():
        try:
            os.remove(_p(rt, "run.lock"))
        except OSError:
            pass


# ---------- cooldown ----------
def read_cooldown(rt):
    return read_json(_p(rt, "cooldown.json"))


def check_cooldown(rt):
    c = read_cooldown(rt)
    if c:
        raise Refused(f"COOLDOWN until {c.get('until')}: {c.get('reason')}. Dice is not touched until the user clears it after that time.")


def write_cooldown(rt, reason, screenshot=None, hours=48):
    until = (dt.datetime.now() + dt.timedelta(hours=hours)).replace(microsecond=0)
    write_json(_p(rt, "cooldown.json"), {"until": until.isoformat(sep=" "), "reason": reason, "screenshot": str(screenshot or ""),
                                         "written": now_iso()})
    return until


def clear_cooldown(rt, now=None):
    c = read_cooldown(rt)
    if not c:
        return "No cooldown is active."
    until = dt.datetime.fromisoformat(c["until"])
    if (now or dt.datetime.now()) < until:
        raise Refused(f"The cooldown runs until {c['until']} ({c.get('reason')}). It can only be cleared after that time.")
    os.remove(_p(rt, "cooldown.json"))
    return f"Cooldown cleared (it was: {c.get('reason')}). Check Dice by hand first; consider a small first round."


# ---------- stop request ----------
def request_stop(rt):
    _p(rt, "stop.request").write_text(now_iso(), encoding="utf-8")


def stop_requested(rt):
    return _p(rt, "stop.request").exists()


def clear_stop(rt):
    try:
        os.remove(_p(rt, "stop.request"))
    except OSError:
        pass


# ---------- live status ----------
def write_status(rt, **kw):
    write_json(_p(rt, "status.json"), dict(kw, updated=now_iso(), pid=os.getpid()))


def read_status(rt):
    return read_json(_p(rt, "status.json"), {}) or {}


# ---------- ledger and daily counters ----------
def ledger(rt):
    return read_jsonl(_p(rt, "ledger.jsonl"))


def ledger_add(rt, job_id, guid, result, resume=""):
    append_jsonl(_p(rt, "ledger.jsonl"), {"job": job_id, "guid": guid, "result": result, "resume": str(resume), "at": now_iso()})


def ledger_guids(rt):
    return {e.get("guid") for e in ledger(rt) if e.get("guid")}


class Daily:
    """Today's counters. Rebuilt from the ledger and the run results when daily.json is missing or from another day."""

    KEYS = ("wizard_opens", "submits", "unconfirmed", "soft_signals")

    def __init__(self, rt):
        self.rt = rt
        d = read_json(_p(rt, "daily.json"))
        if not d or d.get("date") != today():
            d = self._rebuild()
        self.d = d

    def _rebuild(self):
        d = {"date": today(), **{k: 0 for k in self.KEYS}}
        for e in ledger(self.rt):
            if str(e.get("at", "")).startswith(today()):
                if e.get("result") == "unconfirmed":
                    d["unconfirmed"] += 1
                d["submits"] += 1
        runs = _p(self.rt, "runs")
        if runs.is_dir():
            for r in runs.glob(today() + "_*/results.jsonl"):
                d["wizard_opens"] += sum(1 for x in read_jsonl(r) if x.get("wizard_opened"))
        return d

    def add(self, k, n=1):
        self.d[k] = self.d.get(k, 0) + n
        write_json(_p(self.rt, "daily.json"), self.d)

    def get(self, k):
        return self.d.get(k, 0)


def limits(cfg):
    daily = cfg.get("max_submits_per_day")
    opens = cfg.get("max_wizard_opens_per_day") or (int(daily * 1.5) if daily else None)
    return daily, opens


# ---------- the working window ----------
DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def window_state(cfg, now=None):
    """-> (ok, reason). reason is '' when inside the working window."""
    now = now or dt.datetime.now()
    days = [d[:3].title() for d in cfg.get("active_days") or DAYS]
    if DAYS[now.weekday()] not in days:
        return False, f"window closed (today is {DAYS[now.weekday()]}; working days are {', '.join(days)})"
    start, end = (cfg.get("active_hours") or ["00:00", "23:59"])[:2]
    hm = now.strftime("%H:%M")
    if not (start <= hm < end):
        return False, f"window closed (working hours are {start}-{end})"
    return True, ""


def window_end(cfg, now=None):
    now = now or dt.datetime.now()
    end = (cfg.get("active_hours") or ["00:00", "23:59"])[1]
    h, m = map(int, end.split(":"))
    return now.replace(hour=h, minute=m, second=0, microsecond=0)


# ---------- block signals ----------
def classify_text(title_and_headings, body=""):
    """-> 'hard:<reason>', 'limit', or ''. Generic block phrases count only in the title and headings; the body of a short page counts
    only for phrases seen on bot-check pages alone (a job description may say "unusual activity"). Texts: dice_apply_selectors.py."""
    head = (title_and_headings or "").lower()
    low = (body or "").lower()
    for t in S.BLOCK_TEXTS:
        if t in head:
            return "hard:" + t
    if len(low) < S.SHORT_PAGE_CHARS:
        for t in S.BLOCK_BODY_TEXTS:
            if t in low:
                return "hard:" + t
    short = len(low) < S.SHORT_PAGE_CHARS  # a job description can say "daily limit" (API roles); only short pages count
    for t in S.DICE_LIMIT_TEXTS:
        if t in head or (short and t in low):
            return "limit"
    return ""


def classify_frames(frames):
    """frames: (url, width, height) of each frame as shown on the page. Only a CAPTCHA frame the user could actually see counts.
    Seen live 2026-11-01: Dice's wizard always loads Google's invisible reCAPTCHA helper (recaptcha/api2/aframe, 0x0, display:none);
    that is not a challenge. A real one is a visible frame of real size (the checkbox, or the image challenge)."""
    for f in frames or ():
        u, w, h = (f, S.MIN_CHALLENGE_PX, S.MIN_CHALLENGE_PX) if isinstance(f, str) else f
        if (w or 0) < S.MIN_CHALLENGE_PX or (h or 0) < S.MIN_CHALLENGE_PX or re.search(r"recaptcha/[^/]+/aframe", u or "", re.I):
            continue
        for pat in S.CAPTCHA_FRAME_PATTERNS:
            if re.search(pat, u or "", re.I):
                return "hard:captcha frame " + pat
    return ""


# Ads and analytics requests that Dice routes through its own domain. A 403 on one of these is not a block: seen live 2026-11-01,
# "https://www.dice.com/vc7eq8/a?id=AW-...&gtm=..." (a Google Ads beacon) answered 403 on a normal job page.
TRACKING = re.compile(r"[?&](gtm|tid|id=AW-|id=G-|cid)=|/collect\b|/g/collect|/pixel|/beacon|analytics|doubleclick|/gtag/|/gtm\.js|/vc7eq8/", re.I)


def classify_response(url, status, resource_type="document"):
    """A dice.com response -> 'hard', 'soft' or ''. Only the page itself and Dice's data calls count; images, scripts, fonts,
    beacons and ads/analytics requests are ignored, so a blocked tracking pixel never stops a round."""
    if not re.match(r"^https?://([a-z0-9-]+\.)*dice\.com(/|$)", url or "", re.I) or TRACKING.search(url or ""):
        return ""
    rt = (resource_type or "").lower()
    if rt == "document":
        if status in (403, 429):
            return "hard"
    elif rt in ("xhr", "fetch"):
        if status == 429:
            return "hard"
        if status == 403:
            return "soft"
    else:
        return ""
    if 500 <= (status or 0) <= 599:
        return "soft"
    return ""


class SoftSignals:
    """One soft signal doubles every delay for 60 minutes; three within 60 minutes stop the round."""

    def __init__(self, window_sec=3600, stop_at=3):
        self.times, self.window, self.stop_at = [], window_sec, stop_at

    def add(self, t=None):
        t = t or time.time()
        self.times.append(t)
        self.times = [x for x in self.times if t - x < self.window]
        return len(self.times) >= self.stop_at

    def factor(self, t=None):
        t = t or time.time()
        return 2.0 if any(t - x < self.window for x in self.times) else 1.0
