"""Guards (design section 10): lock, cooldown, block signals from hand-written fixtures, soft signals, window, daily counters."""
import datetime as dt
import json
import os
import unittest
from pathlib import Path

from helpers import TempDir
from dice_apply_common import Refused, write_json
import dice_apply_guard as G

FIX = Path(__file__).resolve().parent / "fixtures"


def fixture(name):
    return (FIX / name).read_text(encoding="utf-8")


class Guard(unittest.TestCase):
    def setUp(self):
        self.tmp = TempDir()
        self.rt = self.tmp.__enter__()

    def tearDown(self):
        self.tmp.__exit__()

    def test_lock_live_and_stale(self):
        G.acquire_lock(self.rt)
        self.assertIsNotNone(G.lock_info(self.rt))
        G.acquire_lock(self.rt)                                                    # same process may re-acquire
        G.release_lock(self.rt)
        self.assertIsNone(G.lock_info(self.rt))
        write_json(self.rt / "run.lock", {"pid": 999999, "started": "x"})          # a process that is gone: stale
        self.assertIsNone(G.lock_info(self.rt))
        G.acquire_lock(self.rt)
        self.assertEqual(G.lock_info(self.rt)["pid"], os.getpid())

    def test_lock_held_by_other_live_process(self):
        write_json(self.rt / "run.lock", {"pid": os.getppid(), "started": "09:00"})
        if G.lock_info(self.rt):
            with self.assertRaises(Refused):
                G.acquire_lock(self.rt)

    def test_cooldown(self):
        G.write_cooldown(self.rt, "captcha frame", hours=48)
        with self.assertRaises(Refused):
            G.check_cooldown(self.rt)
        with self.assertRaises(Refused):
            G.clear_cooldown(self.rt)                                              # before the end time
        later = dt.datetime.now() + dt.timedelta(hours=49)
        self.assertIn("cleared", G.clear_cooldown(self.rt, now=later))
        G.check_cooldown(self.rt)

    def test_block_texts_from_fixtures(self):
        self.assertTrue(G.classify_text("Just a moment...", "").startswith("hard:"))                       # title of an interstitial
        self.assertTrue(G.classify_text("", fixture("challenge_200.html")).startswith("hard:"))            # its short body
        self.assertEqual(G.classify_text("Application", fixture("dice_limit.html")), "limit")
        job_text = "Senior IAM Engineer. Monitor unusual activity in SIEM; triage access denied errors."
        self.assertEqual(G.classify_text("Senior IAM Engineer | Dice.com", job_text), "")                # never a block in a description
        self.assertTrue(G.classify_text("Access Denied", "").startswith("hard:"))                          # but it is as a page title

    def test_limit_text_in_a_long_job_description_is_not_a_limit(self):
        desc = "Senior API Engineer. " + "Build rate limiting. " * 200 + "Implement a per-user daily application limit reached alert."
        self.assertEqual(G.classify_text("Senior API Engineer | Dice.com", desc), "")
        self.assertEqual(G.classify_text("", "You have reached the maximum number of applications for today."), "limit")
        self.assertEqual(G.classify_text("Daily application limit", "x" * 5000), "limit")
        self.assertEqual(G.classify_frames(["https://www.google.com/recaptcha/api2/anchor?k=1"])[:5], "hard:")
        self.assertEqual(G.classify_frames(["https://www.dice.com/embedded"]), "")

    def test_hidden_recaptcha_helper_is_not_a_block(self):
        # seen live 2026-11-01 in the Dice wizard: Google's invisible reCAPTCHA helper, 0x0 and display:none
        self.assertEqual(G.classify_frames([("https://www.google.com/recaptcha/api2/aframe", 0, 0)]), "")
        self.assertEqual(G.classify_frames([("https://www.google.com/recaptcha/api2/anchor?k=1", 0, 0)]), "")   # hidden anchor
        self.assertEqual(G.classify_frames([("https://www.google.com/recaptcha/api2/aframe", 300, 300)]), "")   # helper even if sized
        self.assertTrue(G.classify_frames([("https://www.google.com/recaptcha/api2/bframe?k=1", 400, 580)]).startswith("hard:"))
        self.assertTrue(G.classify_frames([("https://www.google.com/recaptcha/api2/anchor?k=1", 304, 78)]).startswith("hard:"))

    def test_responses(self):
        self.assertEqual(G.classify_response("https://www.dice.com/job-detail/x", 403, "document"), "hard")
        self.assertEqual(G.classify_response("https://www.dice.com/job-detail/x", 429, "document"), "hard")
        self.assertEqual(G.classify_response("https://job-applications.dice.com/api/apply", 429, "fetch"), "hard")
        self.assertEqual(G.classify_response("https://job-applications.dice.com/api/apply", 403, "xhr"), "soft")
        self.assertEqual(G.classify_response("https://www.dice.com/x", 503, "document"), "soft")
        self.assertEqual(G.classify_response("https://www.dice.com/logo.png", 403, "image"), "")
        self.assertEqual(G.classify_response("https://cdn.example.com/x", 403, "document"), "")
        self.assertEqual(G.classify_response("https://notdice.com.evil.example/x", 429, "document"), "")

    def test_tracking_beacon_is_not_a_block(self):
        # seen live 2026-11-01 on a normal job page: a Google Ads beacon routed through dice.com answered 403
        url = "https://www.dice.com/vc7eq8/a?id=AW-1657927990&v=3&t=t&pid=737624451&gtm=45892e6as1h1za20kzb79074302zd79074302xf4&cv=1&r"
        for rt in ("document", "fetch", "xhr", "ping", "image", "other"):
            self.assertEqual(G.classify_response(url, 403, rt), "", rt)

    def test_soft_signals(self):
        s = G.SoftSignals()
        self.assertEqual(s.factor(1000), 1.0)
        self.assertFalse(s.add(1000))
        self.assertEqual(s.factor(1100), 2.0)
        self.assertFalse(s.add(1200))
        self.assertTrue(s.add(1300))
        s2 = G.SoftSignals()
        s2.add(0), s2.add(10)
        self.assertFalse(s2.add(4000))                                             # the first two are over an hour old

    def test_window(self):
        cfg = {"active_days": ["Mon", "Tue", "Wed", "Thu", "Fri"], "active_hours": ["08:00", "16:00"]}
        self.assertTrue(G.window_state(cfg, dt.datetime(2026, 11, 2, 9, 0))[0])    # Monday
        self.assertFalse(G.window_state(cfg, dt.datetime(2026, 11, 2, 16, 0))[0])
        self.assertFalse(G.window_state(cfg, dt.datetime(2026, 11, 1, 9, 0))[0])   # Sunday

    def test_daily_rebuild_and_limits(self):
        G.ledger_add(self.rt, "JOB-1", "g1", "applied")
        G.ledger_add(self.rt, "JOB-2", "g2", "unconfirmed")
        d = G.Daily(self.rt)
        self.assertEqual((d.get("submits"), d.get("unconfirmed")), (2, 1))
        d.add("wizard_opens", 3)
        self.assertEqual(json.loads((self.rt / "daily.json").read_text())["wizard_opens"], 3)
        self.assertEqual(G.limits({"max_submits_per_day": 40}), (40, 60))
        self.assertEqual(G.ledger_guids(self.rt), {"g1", "g2"})

    def test_stop_request(self):
        self.assertFalse(G.stop_requested(self.rt))
        G.request_stop(self.rt)
        self.assertTrue(G.stop_requested(self.rt))
        G.clear_stop(self.rt)
        self.assertFalse(G.stop_requested(self.rt))


if __name__ == "__main__":
    unittest.main()
