"""Dice Auto Apply: applies to Dice jobs from the Master Job Tracker with the job's own resume, at a natural pace, in rounds the
candidate approved. Driven by an AI assistant through the Module 05 prompt, or by the launcher (design: designs/Dice Auto Apply - Design.md).

Usage:  python Tools/dice_apply.py <command> [options]        (on a Mac: python3)
Run from the dice-auto-apply folder. When a venv exists at ~/.job-workflow/venv, the script re-runs itself with it automatically.

Setup:
  doctor [--fix]                 checks Python, the venv, Playwright, Chrome, the tracker, the Resumes folder, login and settings;
                                 --fix installs Python 3.12 if needed, creates the venv, installs requirements, copies tracker_path.txt
  login                          opens Chrome (its own profile) so the candidate logs in to Dice by hand; never types credentials
  config show | config set key=value | config answer "<question>" "<answer>"
                                 settings and saved screening answers. `config set tracker=<path>` sets the tracker file.
Rounds:
  plan [filters] [--limit N] [--pace P] [--json] [--all]
                                 which jobs a round would take and why the others are skipped; reads only the tracker, never opens Dice.
                                 Ends with an IDS: line, the exact list to pass to run --jobs
  run --jobs IDS --mode dry-run|submit|confirm [--pace P] [--until HH:MM] [--break-every N --break-minutes M] [--detach]
                                 a round over exactly these jobs (filters work too). dry-run is the default and never submits
  round [filters] [--limit N] [--pace P]
                                 for the launcher: plan, ask [a]pply / [t]est only / [n]o, start detached, show progress, then report
  status [--wait SECONDS]        RUNNING / STOPPING / IDLE, progress and today's counts; never opens Dice
  stop                           ends the round after the current step (never between Submit and its success check)
  report [--run FOLDER]          the RUN: line and only the jobs that need attention, with the next step for each
  sync-tracker                   writes tracker updates that are waiting (for example because Excel had the file open)
  clear-cooldown                 removes a block cooldown after its end time (only when the candidate asks)
  bridge --line-file FILE        the launcher's paste bridge: runs one pasted line after checking it (never through a shell)

Filters: --jobs JOB-1,JOB-2  --status S1,S2  --run-date today|yesterday|YYYY-MM-DD  --run-latest  --found today|last:Nd|YYYY-MM-DD
         --priority 1,2  --title-contains TEXT (repeatable)  --allow JOB-1,...  --allow-duplicate  --resume FILE (one job only)

Result codes per job: APPLIED, DRY_RUN_OK, ALREADY_APPLIED_ON_DICE, EXTERNAL_APPLY, NEEDS_INPUT, JOB_GONE, RESUME_NOT_ATTACHED,
  SUBMIT_UNCONFIRMED, STUCK, UI_CHANGED, WINDOW_CLOSED, DAILY_LIMIT, STOPPED, BLOCKED, LOGIN_REQUIRED.
Skip codes in plan: NO_RESUME, RESUME_NOT_FOUND, RESUME_AMBIGUOUS, RESUME_MISMATCH, RESUME_TOO_LARGE, STATUS_<status>, ALREADY_APPLIED,
  EMAILED_RESUME, IN_LEDGER, DUPLICATE_OF <JOB>, SAME_RECRUITER <JOB>, NOT_DICE.
Exit codes: 0 done, 1 some jobs need attention, 2 refused or the round stopped (blocked, cooldown, already running, login, stuck twice,
  Dice limit), 3 bad arguments or settings.
"""
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dice_apply_platform import FOLDER, HERE, in_venv, venv_python  # noqa: E402


def _bootstrap():
    """Re-run with the venv's Python when it exists and we are not in it (design section 5)."""
    vp = venv_python()
    if not in_venv() and vp.exists() and os.environ.get("DICE_APPLY_NO_VENV") != "1":
        env = dict(os.environ, DICE_APPLY_NO_VENV="1", PYTHONIOENCODING="utf-8")
        return subprocess.call([str(vp), os.path.abspath(__file__)] + sys.argv[1:], env=env)
    return None


if __name__ == "__main__":
    _rc = _bootstrap()
    if _rc is not None:
        sys.exit(_rc)

import argparse  # noqa: E402
import datetime as dt  # noqa: E402
import json  # noqa: E402
import math  # noqa: E402
import random  # noqa: E402
import shlex  # noqa: E402
import shutil  # noqa: E402
import time  # noqa: E402
from pathlib import Path  # noqa: E402

import dice_apply_config as C  # noqa: E402
import dice_apply_guard as G  # noqa: E402
import dice_apply_platform as P  # noqa: E402
import dice_apply_select as SEL  # noqa: E402
import resume_find as RF  # noqa: E402
import tracker_io as T  # noqa: E402
from dice_apply_common import Refused, Usage, append_jsonl, now_iso, out, read_jsonl, today, write_json  # noqa: E402

TRACKER_TXT = FOLDER / "tracker_path.txt"
OK_CODES = {"APPLIED", "DRY_RUN_OK", "ALREADY_APPLIED_ON_DICE"}
ERROR_CODES = {"STUCK", "UI_CHANGED"}
NEXT = {
    "EXTERNAL_APPLY": "apply on the company site (link shown)",
    "NEEDS_INPUT": "answer these, save with config answer, then re-run the job",
    "JOB_GONE": "closed on Dice; skip it in the tracker",
    "SUBMIT_UNCONFIRMED": "check Dice > Applied Jobs; it may have gone through",
    "RESUME_NOT_ATTACHED": "re-run once; if again, check the resume file",
    "STUCK": "re-run once; if again, the Dice page changed",
    "UI_CHANGED": "the Dice page changed; the selectors need a fix",
    "WINDOW_CLOSED": "not submitted (outside working hours); next round",
    "DAILY_LIMIT": "not submitted (daily limit reached); next round",
    "STOPPED": "not submitted (stopped); next round",
    "BLOCKED": "do not touch Dice; cooldown started",
}


# ---------- shared helpers ----------
def cfg_rt():
    cfg = C.load()
    rt = P.runtime_dir(cfg)
    rt.mkdir(parents=True, exist_ok=True)
    return cfg, rt


def tracker_path(a):
    if getattr(a, "tracker", None):
        p = Path(a.tracker).expanduser()
    elif TRACKER_TXT.exists():
        p = Path(TRACKER_TXT.read_text(encoding="utf-8").strip().strip('"')).expanduser()
    else:
        raise Usage("No tracker is set. Run doctor --fix (it copies application-workflow's tracker_path.txt) or config set tracker=<path>.")
    if not p.exists():
        raise Usage(f"Tracker not found: {p}")
    return p


def index_for(cfg):
    return RF.ResumeIndex(cfg.get("resumes_root") or None)


def filters_of(a):
    return {"jobs": a.jobs, "status": a.status, "run_date": a.run_date, "run_latest": a.run_latest, "found": a.found,
            "priority": a.priority, "title_contains": a.title_contains or []}


def remaining_today(cfg, rt):
    daily, _ = G.limits(cfg)
    if not daily:
        return None
    return max(0, daily - G.Daily(rt).get("submits"))


def make_plan(a, cfg, rt, jobs, for_submit):
    allow = [x.strip() for x in (a.allow or "").split(",") if x.strip()]
    return SEL.build_plan(jobs, filters_of(a), cfg, G.ledger_guids(rt), T.queued_applied_jobs(rt), index_for(cfg),
                          limit=a.limit, allow=allow, allow_duplicate=a.allow_duplicate,
                          remaining_today=remaining_today(cfg, rt) if for_submit else None, override_resume=a.resume)


def pace_of(a, cfg):
    pace = (a.pace or cfg.get("default_pace") or "normal").lower()
    if pace not in (cfg.get("paces") or {}):
        raise Usage("pace must be brisk, normal or relaxed")
    return pace


def fmt_minutes(m):
    return f"{m // 60} h {m % 60:02d} min" if m >= 60 else f"{m} min"


# ---------- plan ----------
def print_plan(p, cfg, rt, pace, show_all=False):
    n = len(p.picked)
    mins = SEL.estimate_minutes(n, cfg, pace)
    finish = (dt.datetime.now() + dt.timedelta(minutes=mins)).strftime("%H:%M")
    daily, _ = G.limits(cfg)
    used = G.Daily(rt).get("submits")
    left = f"; {len(p.left)} more match and can go in the next round" if p.left else ""
    out(f"PLAN: {n} Dice job(s) for this round{left}.")
    if p.limit_note:
        out(f"Limit: {p.limit_note}.")
    out(f"Pace: {pace}, about {fmt_minutes(mins)} (done around {finish}). Today: {used}/{daily or 'not set'} submits used.")
    shown = p.picked if show_all else p.picked[:10]
    for j, r, note in shown:
        extra = []
        if r.via == "name search":
            extra.append("resume via name search")
        if note:
            extra.append(note)
        out(f"  {j['Job ID']:<9} {j['Job Title'][:60]} @ {j['Company'][:30]}" + (f"  [{'; '.join(extra)}]" if extra else ""))
    if len(shown) < n:
        out(f"  ... and {n - len(shown)} more (plan --all lists them)")
    if p.skipped:
        groups = {}
        for _, code, _ in p.skipped:
            groups[code.split(" ")[0]] = groups.get(code.split(" ")[0], 0) + 1
        out("Skipped: " + ", ".join(f"{v} {k}" for k, v in sorted(groups.items(), key=lambda x: -x[1])))
        act = [s for s in p.skipped if s[1].split(" ")[0] in ("RESUME_NOT_FOUND", "RESUME_AMBIGUOUS", "RESUME_MISMATCH", "RESUME_TOO_LARGE",
                                                             "SAME_RECRUITER", "DUPLICATE_OF", "NOT_DICE")]
        for j, code, detail in act[: (len(act) if show_all else 10)]:
            out(f"  {j['Job ID']:<9} {code}: {detail}")
        if not show_all and len(act) > 10:
            out(f"  ... {len(act) - 10} more to fix (plan --all lists them)")
    for note in p.notes:
        out("Note: " + note)
    out("IDS: " + ",".join(j["Job ID"] for j, _, _ in p.picked))


def cmd_plan(a):
    cfg, rt = cfg_rt()
    tr = tracker_path(a)
    msg = T.try_flush(tr, rt) if not G.lock_info(rt) else ""
    pace = pace_of(a, cfg)
    p = make_plan(a, cfg, rt, T.read_jobs(tr), for_submit=True)
    if a.json:
        d = p.as_dict()
        d.update(pace=pace, minutes=SEL.estimate_minutes(len(p.picked), cfg, pace), today_submits=G.Daily(rt).get("submits"),
                 max_submits_per_day=cfg.get("max_submits_per_day"), ids=[j["Job ID"] for j, _, _ in p.picked])
        if not a.all:
            d["picked"] = d["picked"][:10] + ([{"more": len(d["picked"]) - 10}] if len(d["picked"]) > 10 else [])
            d["skipped"] = [s for s in d["skipped"] if not s["code"].startswith(("NO_RESUME", "STATUS_", "ALREADY_APPLIED", "IN_LEDGER", "EMAILED"))][:15]
        out(json.dumps(d, ensure_ascii=False))
    else:
        if msg:
            out(msg)
        print_plan(p, cfg, rt, pace, a.all)
    return 0


# ---------- run ----------
class Round:
    def __init__(self, a, cfg, rt, tracker, run_dir):
        self.a, self.cfg, self.rt, self.tracker, self.run_dir = a, cfg, rt, tracker, run_dir
        self.mode, self.pace = a.mode, pace_of(a, cfg)
        self.pc = dict(cfg["paces"][self.pace])
        if a.break_every:
            self.pc["break_every"] = [a.break_every, a.break_every]
        if a.break_minutes:
            self.pc["break_min"] = [a.break_minutes, a.break_minutes]
        self.daily = G.Daily(rt)
        self.max_submits, self.max_opens = G.limits(cfg)
        self.signals = G.SoftSignals()
        self.counts = {"applied": 0, "dry": 0, "attention": 0, "errors": 0}
        self.done = self.total = 0
        self.current = ""
        self.started = now_iso()
        self.next_break = None
        self.until = None
        if a.until:
            h, m = map(int, a.until.split(":"))
            self.until = dt.datetime.now().replace(hour=h, minute=m, second=0, microsecond=0)

    def log(self, s):
        out(f"[{dt.datetime.now().strftime('%H:%M:%S')}] {s}")

    def status(self, state="RUNNING", **kw):
        G.write_status(self.rt, state=state, started=self.started, mode=self.mode, pace=self.pace, done=self.done, total=self.total,
                       current=self.current, run_dir=str(self.run_dir), next_break_after=self.next_break, eta=self.eta(),
                       today_submits=self.daily.get("submits"), max_submits=self.max_submits, window_end=self.cfg["active_hours"][1],
                       **self.counts, **kw)

    def eta(self):
        left = max(0, self.total - self.done)
        return (dt.datetime.now() + dt.timedelta(minutes=SEL.estimate_minutes(left, self.cfg, self.pace))).strftime("%H:%M")

    def gate(self, at_submit=False):
        """'' to carry on, or why not. At Submit the answer becomes the job's result code."""
        if G.stop_requested(self.rt):
            return "STOPPED" if at_submit else "stopped by user"
        if G.read_cooldown(self.rt):
            return "STOPPED" if at_submit else "cooldown active"
        if self.mode != "dry-run":
            ok, why = G.window_state(self.cfg)
            if not ok:
                return "WINDOW_CLOSED" if at_submit else why
            if self.max_submits and self.daily.get("submits") >= self.max_submits:
                return "DAILY_LIMIT" if at_submit else "daily limit"
        if self.until and dt.datetime.now() >= self.until:
            return "WINDOW_CLOSED" if at_submit else "until time"
        if not at_submit and self.max_opens and self.daily.get("wizard_opens") >= self.max_opens:
            return "daily limit (wizard opens)"
        return ""

    def sleep(self, seconds, label):
        """Pacing waits in small steps so stop requests and the window are honoured during long gaps."""
        end = time.time() + seconds * self.signals.factor()
        self.status(waiting=label, wait_until=dt.datetime.fromtimestamp(end).strftime("%H:%M:%S"))
        while time.time() < end:
            if self.gate():
                return
            time.sleep(min(2.0, max(0.0, end - time.time())))

    def gap(self):
        pc = self.pc
        s = random.lognormvariate(math.log(pc["gap_median"]), 0.5)
        return min(max(s, pc["gap_min"]), pc["gap_max"])

    def record(self, job, res, o):
        jid, guid = job["Job ID"], SEL.guid_of(job["Job URL"])
        if o.submitted and o.code != "APPLIED":  # belt and braces: anything after a Submit click is ledgered and never retried
            o.code = "SUBMIT_UNCONFIRMED"
        code = o.code
        if o.wizard_opened:
            self.daily.add("wizard_opens")
        resume_name = res.file.name if res and res.file else ""
        if code == "APPLIED":
            G.ledger_add(self.rt, jid, guid, "applied", res.file)
            self.daily.add("submits")
            T.queue_add(self.rt, jid, "applied", f"Applied on Dice (Apply wizard) by dice-auto-apply with {resume_name}; no cover letter; "
                        "work auth and location from Dice profile.", resume_path=str(res.file), resume_filename=resume_name, guid=guid, applied_date=today())
        elif code == "ALREADY_APPLIED_ON_DICE" and self.mode != "dry-run":  # a test run reports it only; it never writes anything
            G.ledger_add(self.rt, jid, guid, "found_applied", "")
            T.queue_add(self.rt, jid, "applied", "Found already applied on Dice by dice-auto-apply (the job page shows Applied).", guid=guid, applied_date=today())
        elif code == "SUBMIT_UNCONFIRMED":
            G.ledger_add(self.rt, jid, guid, "unconfirmed", res.file if res else "")
            self.daily.add("submits")
            self.daily.add("unconfirmed")
            T.queue_add(self.rt, jid, "note", "dice-auto-apply clicked Submit on Dice but saw no success page; check Dice > Applied Jobs.")
        elif self.mode != "dry-run" and code in ("EXTERNAL_APPLY", "NEEDS_INPUT", "JOB_GONE", "STUCK", "RESUME_NOT_ATTACHED"):
            T.queue_add(self.rt, jid, "note", f"dice-auto-apply: {code} {o.detail}".strip())
        if code == "APPLIED" or code == "ALREADY_APPLIED_ON_DICE":
            self.counts["applied"] += 1
        elif code == "DRY_RUN_OK":
            self.counts["dry"] += 1
        else:
            self.counts["attention"] += 1
            if code in ERROR_CODES:
                self.counts["errors"] += 1
        row = {"job": jid, "code": code, "title": job["Job Title"], "company": job["Company"], "detail": o.detail, "resume": resume_name,
               "via": res.via if res else "", "wizard_opened": o.wizard_opened, "shot": o.shot, "url": o.url, "at": now_iso()}
        append_jsonl(self.run_dir / "results.jsonl", row)
        extra = f"resume={resume_name}" if code in OK_CODES and resume_name else o.detail
        line = f"{jid:<9} {code:<23} {job['Job Title'][:50]} @ {job['Company'][:30]}  {extra}".rstrip()
        with open(self.run_dir / "summary.txt", "a", encoding="utf-8") as f:
            f.write(line + "\n")
        self.log(line)

    def flush(self, final=False):
        waits = [30, 120, 300] if final else []
        for i in range(len(waits) + 1):
            try:
                n, rep = T.flush(self.tracker, self.rt)
                for r in rep:
                    self.log("tracker: " + r)
                if n:
                    self.log(f"tracker: wrote {n} update(s)")
                return True
            except Refused as e:
                self.log(f"tracker waiting: {e}")
                if i < len(waits):
                    self.status(state="RUNNING" if not final else "FINISHING", waiting="tracker waiting: close Excel")
                    time.sleep(waits[i])
            except Exception as e:  # never lose the round over a tracker problem; the queue keeps the updates
                self.log(f"tracker waiting: {type(e).__name__}: {e}")
                break
        return False

    def run(self, picked):
        from dice_apply_flow import Session, Stop
        self.total = len(picked)
        answers = C.answers(self.cfg)
        reason, exit_code = "done", 0
        session = None
        stuck_row = 0
        self.next_break = self.done + random.randint(*self.pc["break_every"])
        self.status()
        try:
            session = Session(self.cfg, self.rt, self.run_dir, self.signals, self.log).open()
            with P.KeepAwake():
                for job, res, note in picked:
                    why = self.gate()
                    if why:
                        reason = why
                        break
                    if SEL.guid_of(job["Job URL"]) in G.ledger_guids(self.rt):
                        continue
                    self.current = job["Job ID"]
                    self.status()
                    o = session.apply_job(job, res.file, self.mode, answers, lambda: self.gate(at_submit=True), confirm=self.confirm)
                    self.record(job, res, o)
                    self.done += 1
                    stuck_row = stuck_row + 1 if o.code in ERROR_CODES else 0
                    if stuck_row >= 2:
                        reason, exit_code = "stuck/UI changed twice", 2
                        break
                    if o.code in ("WINDOW_CLOSED", "DAILY_LIMIT", "STOPPED"):
                        reason = {"WINDOW_CLOSED": "window closed", "DAILY_LIMIT": "daily limit", "STOPPED": "stopped by user"}[o.code]
                        break
                    if self.done % int(self.cfg.get("tracker_batch_size") or 5) == 0:
                        self.flush()
                    if self.done >= self.total:
                        break
                    if self.done >= self.next_break:
                        self.flush()
                        self.sleep(random.uniform(*self.pc["break_min"]) * 60, "break")
                        self.next_break = self.done + random.randint(*self.pc["break_every"])
                    else:
                        self.sleep(self.gap(), "gap")
        except Stop as s:
            exit_code = 2
            if getattr(s, "unconfirmed", None):  # the Stop came after the Submit click: record it before anything else
                job, res = next(((j, r) for j, r, _ in picked if j["Job ID"] == s.unconfirmed), (None, None))
                if job:
                    from dice_apply_flow import Outcome
                    self.record(job, res, Outcome("SUBMIT_UNCONFIRMED", f"{s.kind} right after Submit", "", True, True))
                    self.done += 1
            kind = s.kind
            if kind == "LOGIN_REQUIRED" and self.done > 0:  # logged in for an earlier job of this round, then logged out
                kind = "BLOCKED"  # logged out partway through a round: the session was revoked (design section 8)
                s.detail = "logged out partway through the round (session revoked)"
            if kind == "BLOCKED":
                shot = session.shot(self.current or "round", "blocked") if session else ""
                until = G.write_cooldown(self.rt, s.detail, self.run_dir / shot if shot else "", self.cfg.get("cooldown_hours") or 48)
                reason = f"BLOCKED {s.detail}; COOLDOWN until {until}"
                try:
                    (self.run_dir / f"{self.current or 'round'}_blocked.html").write_text(session.page.content(), encoding="utf-8")
                except Exception:
                    pass
            elif kind == "BROWSER_CLOSED":
                reason, exit_code = "stopped: browser closed", 0
            elif kind == "SOFT_STOP":
                reason = "stopped: Dice was slow or failing (3 times within an hour)"
            else:
                reason = kind + (f" ({s.detail})" if s.detail else "")
        except Exception as e:  # never leave without the final flush and RUN line
            exit_code = 2
            reason = f"error: {type(e).__name__}: {str(e)[:160]}"
        finally:
            if session:
                session.close()
            self.current = ""
            self.status(state="FINISHING", waiting="writing the tracker")
            self.flush(final=True)
            waiting = len(T.queue_items(self.rt))
            daily_txt = f"Today {self.daily.get('submits')}/{self.max_submits or 'not set'} submits, " \
                        f"{self.daily.get('wizard_opens')}/{self.max_opens or 'not set'} wizard opens."
            run_line = (f"RUN: {self.counts['applied']} applied, {self.counts['dry']} dry-run, {self.counts['attention']} need attention, "
                        f"{self.counts['errors']} error(s). Round {self.done}/{self.total}. {daily_txt} Ended: {reason}."
                        + (f" {waiting} tracker update(s) waiting (Excel open?)." if waiting else ""))
            with open(self.run_dir / "summary.txt", "a", encoding="utf-8") as f:
                f.write(run_line + "\n")
            self.log(run_line)
            G.write_status(self.rt, state="IDLE", last_run=str(self.run_dir), last_line=run_line)
        if exit_code == 0 and self.counts["attention"]:
            exit_code = 1
        return exit_code

    def confirm(self, job):
        try:
            return input(f"Submit {job['Job ID']} {job['Job Title']} @ {job['Company']}? [y/N] ").strip().lower() in ("y", "yes")
        except EOFError:
            return False


def new_run_dir(rt):
    d = Path(rt) / "runs" / dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    d.mkdir(parents=True, exist_ok=True)
    return d


def preflight(a, cfg, rt):
    G.check_cooldown(rt)
    live = G.lock_info(rt)
    if live:
        raise Refused(f"ALREADY_RUNNING: a round started at {live.get('started')} is still running. Use status, or stop.")
    if a.mode in ("submit", "confirm") and not cfg.get("max_submits_per_day"):
        raise Usage("Set the daily limit first: config set max_submits_per_day=<number>.")
    if a.mode == "confirm" and a.detach:
        raise Usage("--mode confirm asks at the terminal, so it cannot run detached.")
    if a.resume and (not a.jobs or "," in a.jobs):
        raise Usage("--resume is only for a single job (--jobs JOB-####).")


def cmd_run(a, argv=None):
    cfg, rt = cfg_rt()
    tr = tracker_path(a)
    if not a.child:
        preflight(a, cfg, rt)
    if a.detach and not a.child:
        run_dir = new_run_dir(rt)
        args = [x for x in (argv or []) if x != "--detach"] + ["--child", "--run-dir", str(run_dir)]
        py = str(P.venv_python()) if P.venv_python().exists() else sys.executable
        proc = P.start_detached([py, str(HERE / "dice_apply.py")] + args, run_dir / "run.log")
        for _ in range(20):
            time.sleep(0.5)
            info = G.lock_info(rt)
            if info and info.get("pid") == proc.pid:
                break
            if proc.poll() is not None:
                tail = (run_dir / "run.log").read_text(encoding="utf-8", errors="ignore").strip().splitlines()[-5:]
                if proc.returncode in (0, 1) and any(x.startswith("RUN:") or "RUN:" in x for x in tail):
                    out(next(x for x in reversed(tail) if "RUN:" in x))  # a very short round can finish before we look
                    return proc.returncode
                raise Refused("The round did not start:\n  " + "\n  ".join(tail))
        out(f"Round started in the background ({a.mode}). Folder: {run_dir}")
        out("Use status to see progress, stop to end it after the current step.")
        return 0
    G.acquire_lock(rt)
    G.clear_stop(rt)
    try:
        run_dir = Path(a.run_dir) if a.run_dir else new_run_dir(rt)
        run_dir.mkdir(parents=True, exist_ok=True)
        T.try_flush(tr, rt)
        p = make_plan(a, cfg, rt, T.read_jobs(tr), for_submit=a.mode != "dry-run")
        write_json(run_dir / "request.json", {"argv": argv, "mode": a.mode, "pace": pace_of(a, cfg), "ids": [j["Job ID"] for j, _, _ in p.picked],
                                              "skipped": [[j["Job ID"], c] for j, c, _ in p.skipped], "started": now_iso()})
        if not p.picked:
            line = "RUN: 0 applied. Nothing to do: no job passed the checks (see plan). Ended: queue empty."
            (run_dir / "summary.txt").write_text(line + "\n", encoding="utf-8")
            G.write_status(rt, state="IDLE", last_run=str(run_dir), last_line=line)
            out(line)
            return 0
        return Round(a, cfg, rt, tr, run_dir).run(p.picked)
    finally:
        G.release_lock(rt)
        G.clear_stop(rt)


# ---------- status / stop / report ----------
def latest_run(rt):
    runs = sorted((Path(rt) / "runs").glob("*/summary.txt"))
    return runs[-1].parent if runs else None


def cmd_status(a):
    cfg, rt = cfg_rt()
    end = time.time() + (a.wait or 0)
    while True:
        live = G.lock_info(rt)
        st = G.read_status(rt)
        if not live or time.time() >= end or G.read_cooldown(rt):
            break
        time.sleep(5)
    cd = G.read_cooldown(rt)
    daily = G.Daily(rt)
    mx, _ = G.limits(cfg)
    if live:
        state = "STOPPING" if G.stop_requested(rt) else st.get("state", "RUNNING")
        cur = f"job {st.get('done', 0) + (1 if st.get('current') else 0)}/{st.get('total')}" + (f" ({st['current']})" if st.get("current") else "")
        out(f"{state}  started {str(st.get('started', ''))[11:16]}  {st.get('mode')}  pace {st.get('pace')}  {cur}  applied {st.get('applied', 0)}  "
            f"attention {st.get('attention', 0)}  errors {st.get('errors', 0)}")
        wait = f"{st.get('waiting')} until {st.get('wait_until')} · " if st.get("waiting") and st.get("wait_until") else (f"{st.get('waiting')} · " if st.get("waiting") else "")
        out(f"{wait}next break after job {st.get('next_break_after')} · estimated finish {st.get('eta')} · today {daily.get('submits')}/{mx or 'not set'} "
            f"submits · window ends {st.get('window_end')}")
        return 0
    out("IDLE" + (f"  COOLDOWN until {cd.get('until')}: {cd.get('reason')}" if cd else ""))
    last = st.get("last_line") or ""
    if not last and latest_run(rt):
        lines = (latest_run(rt) / "summary.txt").read_text(encoding="utf-8").strip().splitlines()
        last = next((x for x in reversed(lines) if x.startswith("RUN:")), "")
    out(f"last round: {last or 'none yet'}")
    out(f"today {daily.get('submits')}/{mx or 'not set'} submits, {daily.get('wizard_opens')} wizard opens")
    try:
        msg = T.try_flush(tracker_path(a), rt)
        if msg:
            out(msg)
    except Refused:
        pass
    return 2 if cd else 0


def cmd_stop(a):
    cfg, rt = cfg_rt()
    if not G.lock_info(rt):
        out("No round is running.")
        return 0
    G.request_stop(rt)
    out("Stop requested. The current step finishes first (never between Submit and its success check), then the round ends.")
    return 0


def cmd_report(a):
    cfg, rt = cfg_rt()
    run = Path(a.run) if a.run else latest_run(rt)
    if not run or not (run / "summary.txt").exists():
        out("No rounds yet.")
        return 0
    lines = (run / "summary.txt").read_text(encoding="utf-8").strip().splitlines()
    run_line = next((x for x in reversed(lines) if x.startswith("RUN:")), "RUN: still running (status shows progress).")
    out(run_line)
    for r in read_jsonl(run / "results.jsonl"):
        if r["code"] in OK_CODES:
            continue
        det = r.get("detail") or ""
        out(f"{r['job']:<9} {r['code']:<20} {r['title'][:40]} @ {r['company'][:24]}  next: {NEXT.get(r['code'], 'see the screenshot')}"
            + (f"  ({det[:160]})" if det else "") + (f"  shot={r['shot']}" if r.get("shot") and r["code"] in ERROR_CODES else ""))
    try:
        msg = T.try_flush(tracker_path(a), rt) if not G.lock_info(rt) else ""
        if msg:
            out(msg)
    except Refused:
        pass
    return 0


# ---------- round (launcher) ----------
def cmd_round(a, argv=None):
    cfg, rt = cfg_rt()
    tr = tracker_path(a)
    G.check_cooldown(rt)
    if G.lock_info(rt):
        out("A round is already running:")
        return cmd_status(argparse.Namespace(wait=0, tracker=a.tracker))
    if a.limit is None and not cfg.get("max_per_round"):
        try:
            a.limit = int(input("How many jobs in this round? ").strip())
        except (ValueError, EOFError):
            raise Usage("Give a number of jobs (for example --limit 10).")
    T.try_flush(tr, rt)
    pace = pace_of(a, cfg)
    p = make_plan(a, cfg, rt, T.read_jobs(tr), for_submit=True)
    print_plan(p, cfg, rt, pace)
    if not p.picked:
        return 0
    ans = input(f"Apply to these {len(p.picked)} jobs now? [a]pply / [t]est only / [n]o: ").strip().lower()
    if ans not in ("a", "apply", "t", "test"):
        out("Nothing started.")
        return 0
    mode = "submit" if ans.startswith("a") else "dry-run"
    if mode == "submit" and not cfg.get("max_submits_per_day"):
        raise Usage("Set the daily limit first: config set max_submits_per_day=<number>.")
    ids = ",".join(j["Job ID"] for j, _, _ in p.picked)
    run_argv = ["run", "--jobs", ids, "--mode", mode, "--pace", pace, "--detach"]
    for flag, v in (("--until", a.until), ("--break-every", a.break_every), ("--break-minutes", a.break_minutes), ("--tracker", a.tracker)):
        if v:
            run_argv += [flag, str(v)]
    if a.allow:  # the overrides behind the approved plan go to the run too, so no approved job is dropped silently
        run_argv += ["--allow", a.allow]
    if a.allow_duplicate:
        run_argv += ["--allow-duplicate"]
    if a.resume:
        run_argv += ["--resume", a.resume]
    rc = main(run_argv)
    if rc != 0:
        return rc
    out("Progress (close this window any time; the round keeps running):")
    last = ""
    try:
        while G.lock_info(rt):
            st = G.read_status(rt)
            line = f"  {st.get('done', 0)}/{st.get('total')} done, applied {st.get('applied', 0)}, attention {st.get('attention', 0)}" \
                   + (f", {st.get('waiting')}" if st.get("waiting") else "")
            if line != last:
                out(line)
                last = line
            time.sleep(10)
    except KeyboardInterrupt:
        out("The round keeps running in the background. Use stop to end it.")
        return 0
    out("")
    return cmd_report(argparse.Namespace(run=None, tracker=a.tracker))


# ---------- bridge (launcher paste) ----------
BRIDGE_OK = {"round", "status", "stop", "report", "config", "plan", "sync-tracker", "doctor", "login", "clear-cooldown"}


def bridge_tokens(line):
    try:
        toks = shlex.split(line.strip(), posix=True)
    except ValueError as e:
        raise Usage(f"Could not read that line ({e}). Paste exactly the line your assistant gave.")
    while toks and (toks[0].lower() in ("python", "python3", "py") or toks[0].replace("\\", "/").lower().endswith("dice_apply.py")):
        toks.pop(0)
    if not toks:
        return ["round"]
    if toks[0].startswith("--"):
        toks = ["round"] + toks
    if toks[0] not in BRIDGE_OK:
        raise Usage(f'"{toks[0]}" is not allowed here. Allowed: {", ".join(sorted(BRIDGE_OK))}.')
    return toks


def cmd_bridge(a):
    line = Path(a.line_file).read_text(encoding="utf-8-sig", errors="ignore").strip() if a.line_file else ""
    toks = bridge_tokens(line)
    out("> " + " ".join(shlex.quote(t) for t in toks))
    return main(toks)


# ---------- config / sync / cooldown ----------
def cmd_config(a):
    if a.action == "set":
        if not a.args or "=" not in a.args[0]:
            raise Usage('Use: config set key=value')
        k, v = a.args[0].split("=", 1)
        if k.strip() == "tracker":
            p = Path(v.strip().strip('"')).expanduser()
            if not p.exists():
                raise Usage(f"Tracker not found: {p}")
            TRACKER_TXT.write_text(str(p) + "\n", encoding="utf-8")
            out(f"tracker = {p}")
            return 0
        out(f"{k.strip()} = {C.set_value(k, v)}")
        return 0
    if a.action == "answer":
        if len(a.args) != 2:
            raise Usage('Use: config answer "<question>" "<answer>" (an empty answer removes it)')
        q = C.set_answer(a.args[0], a.args[1])
        out(f'saved answer for "{q}"' if a.args[1].strip() else f'removed the answer for "{q}"')
        return 0
    cfg, mine = C.load(), C.personal()
    keys = ["max_submits_per_day", "max_per_round", "max_wizard_opens_per_day", "active_days", "active_hours", "default_pace",
            "cooldown_hours", "stuck_timeout_sec", "job_timeout_sec", "tracker_batch_size", "resumes_root"]
    for k in keys:
        out(f"{k} = {cfg.get(k)}" + ("   (your setting)" if k in mine else ""))
    out(f"tracker = {TRACKER_TXT.read_text(encoding='utf-8').strip() if TRACKER_TXT.exists() else 'not set'}")
    ans = C.answers(cfg)
    out(f"screening answers ({len(ans)}): " + ("; ".join(f"{q} = {v}" for q, v in ans.items()) or "none"))
    return 0


def cmd_sync(a):
    cfg, rt = cfg_rt()
    if G.lock_info(rt):
        out("A round is running; it writes the tracker itself.")
        return 0
    n, rep = T.flush(tracker_path(a), rt)
    for r in rep:
        out(r)
    out(f"Wrote {n} tracker update(s)." if n else "Nothing was waiting.")
    return 0


def cmd_clear_cooldown(a):
    cfg, rt = cfg_rt()
    out(G.clear_cooldown(rt))
    return 0


# ---------- login / doctor ----------
def cmd_login(a):
    cfg, rt = cfg_rt()
    G.check_cooldown(rt)
    if G.lock_info(rt):
        raise Refused("A round is running; log in after it ends.")
    from dice_apply_flow import Session
    s = Session(cfg, rt).open()
    try:
        out("Chrome is opening dice.com. Log in there by hand (Log in, top right). This window waits up to 10 minutes.")
        ok = s.wait_for_login(10)
    finally:
        s.close()
    out("Logged in. The session is remembered for later rounds." if ok else "Not logged in yet. Run login again when ready.")
    return 0 if ok else 2


PY_VERSION = "3.12.10"


def install_python():
    """Install Python 3.12 for the venv (design section 13, setup step 2)."""
    import tempfile
    import urllib.request
    if P.IS_WIN:
        if shutil.which("winget"):
            r = subprocess.run(["winget", "install", "-e", "--id", "Python.Python.3.12", "--scope", "user",
                                "--accept-package-agreements", "--accept-source-agreements"])
            if r.returncode == 0:
                return True
        url = f"https://www.python.org/ftp/python/{PY_VERSION}/python-{PY_VERSION}-amd64.exe"
        dest = Path(tempfile.gettempdir()) / f"python-{PY_VERSION}-amd64.exe"
        out(f"Downloading {url}")
        urllib.request.urlretrieve(url, dest)
        return subprocess.run([str(dest), "/quiet", "InstallAllUsers=0", "Include_launcher=1", "PrependPath=0"]).returncode == 0
    if P.IS_MAC:
        url = f"https://www.python.org/ftp/python/{PY_VERSION}/python-{PY_VERSION}-macos11.pkg"
        dest = Path(tempfile.gettempdir()) / f"python-{PY_VERSION}-macos11.pkg"
        out(f"Downloading {url}")
        urllib.request.urlretrieve(url, dest)
        subprocess.run(["open", "-W", str(dest)])
        out("Finish the Python installer window, then run doctor --fix again if the venv is not created.")
        return True
    out("Install Python 3.12 with your system's package manager.")
    return False


def cmd_doctor(a):
    cfg, rt = cfg_rt()
    problems = 0

    def line(ok, what, fix=""):
        nonlocal problems
        if not ok:
            problems += 1
        out(f"{'OK  ' if ok else 'FIX '} {what}" + (f"  ->  {fix}" if fix and not ok else ""))

    vp = P.venv_python()
    if a.fix and not vp.exists():
        py = P.find_supported_python()
        if not py:
            out("No Python 3.12/3.13 found; installing Python 3.12 ...")
            install_python()
            py = P.find_supported_python()
        if py:
            out(f"Creating the venv at {P.VENV} ...")
            subprocess.run(py + ["-m", "venv", str(P.VENV)], check=False)
    if a.fix and vp.exists():
        out("Installing requirements into the venv ...")
        subprocess.run([str(vp), "-m", "pip", "install", "--disable-pip-version-check", "-q", "-r", str(FOLDER / "requirements.txt")], check=False)
    if a.fix and not TRACKER_TXT.exists():
        src = FOLDER.parent / "application-workflow" / "tracker_path.txt"
        if src.exists():
            shutil.copyfile(src, TRACKER_TXT)
            out("Copied tracker_path.txt from application-workflow.")

    line(sys.version_info[:2] >= (3, 9), f"Python {sys.version.split()[0]} runs this script")
    venv_ok = vp.exists()
    line(venv_ok, f"venv at {P.VENV}", "doctor --fix (creates it with Python 3.12/3.13)")
    if venv_ok:
        r = subprocess.run([str(vp), "-c", "import playwright, openpyxl, sys; print(sys.version.split()[0])"], capture_output=True, text=True)
        line(r.returncode == 0, f"Playwright and openpyxl in the venv (Python {r.stdout.strip() or '?'})", "doctor --fix")
    line(P.chrome_installed(), "Google Chrome installed", "install Google Chrome from google.com/chrome")
    try:
        tr = tracker_path(a)
        line(True, f"tracker: {tr}")
    except Usage as e:
        line(False, "tracker", str(e))
    root = Path(cfg.get("resumes_root") or RF.resumes_root())
    line(root.is_dir(), f"Resumes folder: {root}", "create resumes with Module 03, or config set resumes_root=<folder>")
    prof = rt / "browser-profile"
    line(prof.is_dir() and any(prof.iterdir()), "Dice browser profile (login is checked when a round starts)", "run login once")
    line(bool(cfg.get("max_submits_per_day")), f"daily limit: {cfg.get('max_submits_per_day') or 'not set'}", "config set max_submits_per_day=<number>")
    line(len(C.answers(cfg)) > 0, f"screening answers saved: {len(C.answers(cfg))}", 'config answer "<question>" "<answer>" for the common ones')
    cd = G.read_cooldown(rt)
    if cd:
        out(f"NOTE COOLDOWN until {cd.get('until')}: {cd.get('reason')}")
    if G.lock_info(rt):
        out("NOTE a round is running now")
    out("All good." if not problems else f"{problems} thing(s) to fix.")
    return 0 if not problems else 1


# ---------- argument parsing ----------
def add_filters(s):
    s.add_argument("--tracker")
    s.add_argument("--jobs")
    s.add_argument("--status")
    s.add_argument("--run-date", dest="run_date")
    s.add_argument("--run-latest", dest="run_latest", action="store_true")
    s.add_argument("--found")
    s.add_argument("--priority")
    s.add_argument("--title-contains", dest="title_contains", action="append")
    s.add_argument("--allow")
    s.add_argument("--allow-duplicate", dest="allow_duplicate", action="store_true")
    s.add_argument("--resume")
    s.add_argument("--limit", type=int)
    s.add_argument("--pace", choices=["brisk", "normal", "relaxed"])


def hhmm(v):
    try:
        dt.datetime.strptime(v, "%H:%M")
    except ValueError:
        raise argparse.ArgumentTypeError("use HH:MM, for example 14:00")
    return v


def main(argv):
    p = argparse.ArgumentParser(prog="dice_apply.py", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("doctor"); s.add_argument("--fix", action="store_true"); s.add_argument("--tracker"); s.set_defaults(fn=cmd_doctor)
    s = sub.add_parser("login"); s.set_defaults(fn=cmd_login)
    s = sub.add_parser("config"); s.add_argument("action", nargs="?", default="show", choices=["show", "set", "answer"])
    s.add_argument("args", nargs="*"); s.set_defaults(fn=cmd_config)
    s = sub.add_parser("plan"); add_filters(s); s.add_argument("--json", action="store_true"); s.add_argument("--all", action="store_true")
    s.set_defaults(fn=cmd_plan)
    for name, fn in (("run", cmd_run), ("round", cmd_round)):
        s = sub.add_parser(name)
        add_filters(s)
        s.add_argument("--until", type=hhmm)
        s.add_argument("--break-every", dest="break_every", type=int)
        s.add_argument("--break-minutes", dest="break_minutes", type=int)
        if name == "run":
            s.add_argument("--mode", choices=["dry-run", "submit", "confirm"], default="dry-run")
            s.add_argument("--detach", action="store_true")
            s.add_argument("--child", action="store_true", help=argparse.SUPPRESS)
            s.add_argument("--run-dir", dest="run_dir", help=argparse.SUPPRESS)
        s.set_defaults(fn=fn, wants_argv=True)
    s = sub.add_parser("status"); s.add_argument("--wait", type=int, default=0); s.add_argument("--tracker"); s.set_defaults(fn=cmd_status)
    s = sub.add_parser("stop"); s.set_defaults(fn=cmd_stop)
    s = sub.add_parser("report"); s.add_argument("--run"); s.add_argument("--tracker"); s.set_defaults(fn=cmd_report)
    s = sub.add_parser("sync-tracker"); s.add_argument("--tracker"); s.set_defaults(fn=cmd_sync)
    s = sub.add_parser("clear-cooldown"); s.set_defaults(fn=cmd_clear_cooldown)
    s = sub.add_parser("bridge"); s.add_argument("--line-file", dest="line_file"); s.set_defaults(fn=cmd_bridge)
    try:
        a = p.parse_args(argv)
    except SystemExit as e:
        return 3 if e.code else 0
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    try:
        return a.fn(a, argv) if getattr(a, "wants_argv", False) else a.fn(a)
    except Usage as e:
        out("REFUSED: " + str(e))
        return 3
    except Refused as e:
        out("REFUSED: " + str(e))
        return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
