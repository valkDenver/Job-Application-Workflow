"""The CLI and the round loop end to end, with a fake browser session (no Playwright, no Dice): plan, run, status, report, the
paste bridge, stuck-twice, BLOCKED -> cooldown, a closed browser, and the tracker batch."""
import io
import json
import os
import sys
import types
import unittest
from contextlib import redirect_stdout
from unittest import mock

import openpyxl

from helpers import TempDir, job, make_resume, make_tracker

# a fake dice_apply_flow so the round loop runs without a browser
fake = types.ModuleType("dice_apply_flow")


class Stop(Exception):
    def __init__(self, kind, detail="", unconfirmed=None):
        super().__init__(kind)
        self.kind, self.detail, self.unconfirmed = kind, detail, unconfirmed


class Outcome:
    def __init__(self, code, detail="", shot="", wizard_opened=False, submitted=False, url=""):
        self.code, self.detail, self.shot, self.wizard_opened, self.submitted, self.url = code, detail, shot, wizard_opened, submitted, url


SCRIPT = []  # what the fake session returns per job: a code, or a Stop


class Session:
    def __init__(self, *a, **k):
        self.page = mock.Mock()

    def open(self):
        return self

    def close(self):
        pass

    def shot(self, *a):
        return ""

    def apply_job(self, job, resume, mode, answers, before_submit, confirm=None):
        step = SCRIPT.pop(0) if SCRIPT else ("APPLIED" if mode == "submit" else "DRY_RUN_OK")
        if isinstance(step, Stop):
            raise step
        if isinstance(step, tuple):                                                # (code, submitted): an ending after Submit
            return Outcome(step[0], "after submit", "", True, step[1])
        if step in ("APPLIED", "DRY_RUN_OK"):
            gate = before_submit()
            if gate:
                return Outcome(gate, "not submitted", "", True)
        return Outcome(step, "detail" if step not in ("APPLIED", "DRY_RUN_OK") else "", "", True, step == "APPLIED")


fake.Stop, fake.Outcome, fake.Session = Stop, Outcome, Session

import dice_apply as D  # noqa: E402  (imports dice_apply_flow only inside a round, where the fake is patched in)
import dice_apply_config as C  # noqa: E402
import dice_apply_guard as G  # noqa: E402
import shutil  # noqa: E402

SHARED_CONFIG = C.CONFIG


def run(argv):
    buf = io.StringIO()
    with redirect_stdout(buf):
        rc = D.main(argv)
    return rc, buf.getvalue()


class CLI(unittest.TestCase):
    def setUp(self):
        self.tmp = TempDir()
        d = self.tmp.__enter__()
        self.rt = d / "rt"
        os.environ["DICE_APPLY_RUNTIME"] = str(self.rt)
        self.rt.mkdir(parents=True, exist_ok=True)
        self.cfgfile = d / "cfg.json"                                              # a copy: tests never change the shared file
        shutil.copyfile(SHARED_CONFIG, self.cfgfile)
        self.real_config, C.CONFIG = C.CONFIG, self.cfgfile
        root = d / "Resumes"
        rows = []
        for n in range(1, 8):
            name = f"Pat_{n}.docx"
            make_resume(root, "Workday", "HCM", f"JOB-{n}", name)
            rows.append(job(n, **{"Resume Filename": name, "Run Timestamp": "2026-11-02 07:00 MST"}))
        rows.append(job(8))                                                        # no resume
        self.tracker = make_tracker(d / "tracker.xlsx", rows)
        C.set_value("resumes_root", str(root))
        C.set_value("max_submits_per_day", "50")
        C.set_value("active_days", "Mon,Tue,Wed,Thu,Fri,Sat,Sun")
        C.set_value("active_hours", "00:00-23:59")
        SCRIPT.clear()
        self.sleep = mock.patch.object(D.Round, "sleep", lambda *a: None)
        self.sleep.start()
        self.flow = mock.patch.dict(sys.modules, {"dice_apply_flow": fake})
        self.flow.start()

    def tearDown(self):
        C.CONFIG = self.real_config
        self.flow.stop()
        self.sleep.stop()
        self.tmp.__exit__()
        os.environ.pop("DICE_APPLY_RUNTIME", None)

    def status_of(self, jid):
        wb = openpyxl.load_workbook(self.tracker, read_only=True)
        try:
            for r in wb["Master Tracker"].iter_rows(values_only=True):
                if r[0] == jid:
                    return r[16]
        finally:
            wb.close()

    def test_plan_json(self):
        rc, o = run(["plan", "--tracker", str(self.tracker), "--limit", "3", "--json"])
        d = json.loads(o)
        self.assertEqual((rc, len(d["ids"]), d["left_for_next_round"]), (0, 3, 4))

    def test_plan_text_has_ids_line(self):
        rc, o = run(["plan", "--tracker", str(self.tracker), "--limit", "2", "--pace", "relaxed"])
        self.assertIn("PLAN: 2 Dice job(s)", o)
        self.assertIn("IDS: JOB-", o)
        self.assertIn("1 NO_RESUME", o)

    def test_submit_round_batches_tracker_and_reports(self):
        SCRIPT[:] = ["APPLIED", "EXTERNAL_APPLY", "APPLIED", "APPLIED", "NEEDS_INPUT", "APPLIED"]
        rc, o = run(["run", "--tracker", str(self.tracker), "--jobs", "JOB-1,JOB-2,JOB-3,JOB-4,JOB-5,JOB-6", "--mode", "submit"])
        self.assertEqual(rc, 1)                                                    # some need attention
        self.assertIn("RUN: 4 applied", o)
        self.assertEqual(self.status_of("JOB-1"), "APPLIED")
        self.assertEqual(self.status_of("JOB-2"), "REVIEW")
        self.assertEqual(G.Daily(self.rt).get("submits"), 4)
        rc, o = run(["report", "--tracker", str(self.tracker)])
        lines = o.strip().splitlines()
        self.assertTrue(lines[0].startswith("RUN:"))
        self.assertEqual(len(lines), 3)                                            # only the two attention jobs
        rc, o = run(["plan", "--tracker", str(self.tracker), "--json"])
        self.assertNotIn("JOB-1", json.loads(o)["ids"])                            # never picked again

    def test_dry_run_writes_nothing(self):
        before = os.path.getmtime(self.tracker)
        rc, o = run(["run", "--tracker", str(self.tracker), "--jobs", "JOB-1,JOB-2"])
        self.assertIn("RUN: 0 applied, 2 dry-run", o)
        self.assertEqual(os.path.getmtime(self.tracker), before)

    def test_dry_run_already_applied_writes_nothing(self):
        SCRIPT[:] = ["ALREADY_APPLIED_ON_DICE"]
        before = os.path.getmtime(self.tracker)
        rc, o = run(["run", "--tracker", str(self.tracker), "--jobs", "JOB-1"])
        self.assertIn("ALREADY_APPLIED_ON_DICE", o)
        self.assertEqual(os.path.getmtime(self.tracker), before)
        self.assertEqual(G.ledger_guids(self.rt), set())

    def test_stuck_twice_stops(self):
        SCRIPT[:] = ["STUCK", "UI_CHANGED", "APPLIED"]
        rc, o = run(["run", "--tracker", str(self.tracker), "--jobs", "JOB-1,JOB-2,JOB-3", "--mode", "submit"])
        self.assertEqual(rc, 2)
        self.assertIn("stuck/UI changed twice", o)

    def test_blocked_starts_cooldown(self):
        SCRIPT[:] = ["APPLIED", Stop("BLOCKED", "HTTP 429 from https://www.dice.com/x")]
        rc, o = run(["run", "--tracker", str(self.tracker), "--jobs", "JOB-1,JOB-2,JOB-3", "--mode", "submit"])
        self.assertEqual(rc, 2)
        self.assertIsNotNone(G.read_cooldown(self.rt))
        self.assertEqual(self.status_of("JOB-1"), "APPLIED")                       # the batch is still written
        rc, o = run(["run", "--tracker", str(self.tracker), "--jobs", "JOB-3", "--mode", "submit"])
        self.assertEqual(rc, 2)
        self.assertIn("COOLDOWN", o)

    def test_any_stop_after_submit_is_ledgered_unconfirmed(self):
        for kind in ("BLOCKED", "SOFT_STOP", "DAILY_LIMIT_DICE"):
            with self.subTest(kind=kind):
                for f in ("ledger.jsonl", "daily.json", "cooldown.json"):
                    (self.rt / f).unlink(missing_ok=True)
                SCRIPT[:] = [Stop(kind, "x", unconfirmed="JOB-1")]
                rc, o = run(["run", "--tracker", str(self.tracker), "--jobs", "JOB-1,JOB-2", "--mode", "submit"])
                self.assertIn("SUBMIT_UNCONFIRMED", o)
                self.assertIn(job(1)["Job URL"].rsplit("/", 1)[1], G.ledger_guids(self.rt))   # never retried
                self.assertEqual(G.Daily(self.rt).get("submits"), 1)

    def test_submitted_error_outcome_becomes_unconfirmed(self):
        SCRIPT[:] = [("UI_CHANGED", True)]
        rc, o = run(["run", "--tracker", str(self.tracker), "--jobs", "JOB-1", "--mode", "submit"])
        self.assertIn("SUBMIT_UNCONFIRMED", o)
        self.assertEqual(G.Daily(self.rt).get("unconfirmed"), 1)

    def test_logout_partway_starts_cooldown_but_not_at_start(self):
        SCRIPT[:] = [Stop("LOGIN_REQUIRED", "Log in shown")]
        run(["run", "--tracker", str(self.tracker), "--jobs", "JOB-1,JOB-2", "--mode", "submit"])
        self.assertIsNone(G.read_cooldown(self.rt))                                # not logged in yet: just log in
        SCRIPT[:] = ["APPLIED", Stop("LOGIN_REQUIRED", "Log in shown")]
        rc, o = run(["run", "--tracker", str(self.tracker), "--jobs", "JOB-1,JOB-2,JOB-3", "--mode", "submit"])
        self.assertIsNotNone(G.read_cooldown(self.rt))                             # revoked mid-round: treated as BLOCKED
        self.assertIn("session revoked", o)

    def test_config_set_never_changes_the_shared_file(self):
        before = self.cfgfile.read_text(encoding="utf-8")
        run(["config", "set", "max_submits_per_day=77"])
        self.assertEqual(self.cfgfile.read_text(encoding="utf-8"), before)
        self.assertEqual(C.load()["max_submits_per_day"], 77)
        rc, o = run(["config", "show"])
        self.assertIn("max_submits_per_day = 77   (your setting)", o)
        rc, o = run(["config", "set", "runtime_dir=C:/elsewhere"])
        self.assertEqual(rc, 3)                                                    # cannot move the cooldown/ledger/lock aside

    def test_browser_closed_after_submit(self):
        SCRIPT[:] = [Stop("BROWSER_CLOSED", "the Chrome window was closed", unconfirmed="JOB-1")]
        rc, o = run(["run", "--tracker", str(self.tracker), "--jobs", "JOB-1,JOB-2", "--mode", "submit"])
        self.assertIn("stopped: browser closed", o)
        self.assertIn("SUBMIT_UNCONFIRMED", o)
        self.assertEqual(G.Daily(self.rt).get("unconfirmed"), 1)

    def test_stop_request_before_submit(self):
        self.rt.mkdir(parents=True, exist_ok=True)
        with mock.patch.object(D.G, "clear_stop", lambda rt: None):               # keep the request a running round would see
            G.request_stop(self.rt)
            rc, o = run(["run", "--tracker", str(self.tracker), "--jobs", "JOB-1,JOB-2", "--mode", "submit"])
        self.assertIn("Ended: stopped by user", o)
        G.clear_stop(self.rt)

    def test_daily_limit_not_set_refuses(self):
        C.set_value("max_submits_per_day", "none")
        rc, o = run(["run", "--tracker", str(self.tracker), "--jobs", "JOB-1", "--mode", "submit"])
        self.assertEqual(rc, 3)

    def test_answers_live_outside_the_repo(self):
        rc, o = run(["config", "answer", "Willing to relocate?", "No"])
        self.assertEqual(rc, 0)
        self.assertEqual(C.answers(), {"willing to relocate": "No"})
        self.assertTrue(str(C.answers_path()).startswith(str(self.rt)))
        self.assertNotIn("screening_answers", self.cfgfile.read_text(encoding="utf-8"))

    def test_status_idle(self):
        rc, o = run(["status", "--tracker", str(self.tracker)])
        self.assertTrue(o.startswith("IDLE"))

    def test_bridge_tokens(self):
        self.assertEqual(D.bridge_tokens("--run-date today --limit 5"), ["round", "--run-date", "today", "--limit", "5"])
        self.assertEqual(D.bridge_tokens("python Tools/dice_apply.py status"), ["status"])
        self.assertEqual(D.bridge_tokens('config answer "Willing to relocate?" "No"'), ["config", "answer", "Willing to relocate?", "No"])
        self.assertEqual(D.bridge_tokens(""), ["round"])
        for bad in ("run --mode submit", "& del C:\\x", "; rm -rf ~"):
            with self.assertRaises(D.Usage):
                D.bridge_tokens(bad)
        rc, _ = run(D.bridge_tokens("status && calc"))                             # extra words are rejected by argparse, never run
        self.assertEqual(rc, 3)


if __name__ == "__main__":
    unittest.main()
