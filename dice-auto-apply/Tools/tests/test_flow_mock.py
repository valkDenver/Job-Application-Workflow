"""The Playwright page flow against hand-written mock Dice pages (design section 8). Headless Chrome; every dice.com request is answered
by page routing from tests/fixtures, so nothing reaches the real site. Skipped when Playwright or Google Chrome is not installed."""
import re
import unittest
from pathlib import Path

from helpers import TempDir, guid

try:
    from playwright.sync_api import sync_playwright  # noqa: F401
    import dice_apply_flow as F
    HAVE_PW = True
except Exception:  # Playwright not installed in this Python
    HAVE_PW = False
import dice_apply_platform as P

FIX = Path(__file__).resolve().parent / "fixtures"
APPLY_BTN = '<button onclick="location.href=\'/job-applications/{g}/wizard\'">Apply</button>'
EXTERNAL_BTN = '<button onclick="window.open(\'https://careers.example.com/job/1\')">Apply Now</button>'
CFG = {"stuck_timeout_sec": 6, "job_timeout_sec": 120, "job_page_dwell_sec": [0, 0], "slow_load_sec": 30,
       "delay_between_actions_ms": {"median": 50, "min": 10, "max": 80}, "typing_ms": [1, 2]}


class HeadlessSession(F.Session if HAVE_PW else object):
    def open(self):
        self.pw = F.sync_playwright().start()
        self.browser = self.pw.chromium.launch(channel="chrome", headless=True)
        self.ctx = self.browser.new_context()
        self.page = self.ctx.new_page()
        self.page.set_default_timeout(self.stuck_ms)
        self.ctx.on("response", self._on_response)
        return self

    def close(self):
        self.browser.close()
        self.pw.stop()


@unittest.skipUnless(HAVE_PW and P.chrome_installed(), "Playwright or Google Chrome is not installed")
class FlowMock(unittest.TestCase):
    def setUp(self):
        self.tmp = TempDir()
        d = self.tmp.__enter__()
        self.resume = d / "Pat_Workday HCM.docx"
        self.resume.write_bytes(b"resume")
        self.g = guid(7)
        self.job = {"Job ID": "JOB-7", "Job URL": f"https://www.dice.com/job-detail/{self.g}", "Job Title": "Workday HCM", "Company": "X"}
        self.mode = {"job": "apply", "question": False, "status": 200, "initial": "Old_Profile_Resume.pdf", "upload": True}
        self.s = HeadlessSession(CFG, d / "rt", d).open()
        self.s.ctx.route(re.compile(r"https://([a-z0-9-]+\.)*dice\.com/.*"), self.serve)
        self.s.ctx.route(re.compile(r"https://careers\.example\.com/.*"), lambda r: r.fulfill(body="<h1>Careers</h1>", content_type="text/html"))

    def tearDown(self):
        self.s.close()
        self.tmp.__exit__()

    def serve(self, route):
        url, m = route.request.url, self.mode
        if "/wizard/success" in url:
            body = "<h1>Your application is on its way!</h1>"
        elif "/wizard" in url:
            body = ((FIX / "mock_wizard.html").read_text(encoding="utf-8").replace("{{QUESTION}}", "true" if m["question"] else "false")
                    .replace("{{INITIAL}}", '"%s"' % m["initial"]).replace("{{UPLOAD}}", "true" if m["upload"] else "false"))
        elif "/job-detail/" in url:
            if m["job"] == "challenge":
                body = (FIX / "challenge_200.html").read_text(encoding="utf-8")
            elif m["job"] == "gone":
                body = "<h1>Workday HCM</h1><p>This job is no longer available.</p>"
            else:
                btn = {"apply": APPLY_BTN.format(g=self.g), "external": EXTERNAL_BTN, "applied": "<span>Applied</span>",
                       "continue": f'<a href="/job-applications/{self.g}/wizard">Continue Application</a>',
                       "login": '<a href="/login">Log in</a>' + APPLY_BTN.format(g=self.g)}[m["job"]]
                body = (FIX / "mock_job.html").read_text(encoding="utf-8").replace("{{APPLY}}", btn)
        else:
            body = "<html><body>home</body></html>"
        route.fulfill(status=m["status"], body=body, content_type="text/html")

    def apply(self, mode="dry-run", answers=None, gate=lambda: ""):
        return self.s.apply_job(self.job, self.resume, mode, answers or {}, gate)

    def test_dry_run_reaches_review_with_the_right_resume(self):
        o = self.apply()
        self.assertEqual(o.code, "DRY_RUN_OK", o.detail)
        self.assertTrue(o.wizard_opened)
        self.assertIn(self.resume.name, self.s.page.locator("body").inner_text())

    def test_submit_applies(self):
        o = self.apply("submit")
        self.assertEqual((o.code, o.submitted), ("APPLIED", True), o.detail)

    def test_gate_stops_before_submit(self):
        self.assertEqual(self.apply("submit", gate=lambda: "WINDOW_CLOSED").code, "WINDOW_CLOSED")

    def test_question_without_answer_needs_input(self):
        self.mode["question"] = True
        o = self.apply()
        self.assertEqual(o.code, "NEEDS_INPUT", o.detail)
        self.assertIn("Expected hourly rate?", o.detail)

    def test_question_answered_from_config(self):
        self.mode["question"] = True
        o = self.apply(answers={"expected hourly rate": "75", "willing to relocate": "No"})
        self.assertEqual(o.code, "DRY_RUN_OK", o.detail)

    def test_external_apply(self):
        self.mode["job"] = "external"
        o = self.apply()
        self.assertEqual(o.code, "EXTERNAL_APPLY", o.detail)
        self.assertIn("careers.example.com", o.detail)

    def test_failed_upload_with_the_same_name_is_caught(self):
        # another job (or a draft) already shows this exact file name and the upload silently fails: never submit the old file
        self.mode.update(initial=self.resume.name, upload=False)
        o = self.apply("submit")
        self.assertEqual((o.code, o.submitted), ("RESUME_NOT_ATTACHED", False), o.detail)

    def test_same_name_with_new_file_marker_is_fine(self):
        self.mode.update(initial=self.resume.name, upload=True)
        self.assertEqual(self.apply().code, "DRY_RUN_OK")

    def test_continue_application_draft(self):
        self.mode["job"] = "continue"  # seen live 2026-11-01 after an unfinished wizard
        self.assertEqual(self.apply().code, "DRY_RUN_OK")

    def test_already_applied(self):
        self.mode["job"] = "applied"
        self.assertEqual(self.apply().code, "ALREADY_APPLIED_ON_DICE")

    def test_job_gone(self):
        self.mode["job"] = "gone"
        self.assertEqual(self.apply().code, "JOB_GONE")

    def test_challenge_page_blocks(self):
        self.mode["job"] = "challenge"
        with self.assertRaises(F.Stop) as cm:
            self.apply()
        self.assertEqual(cm.exception.kind, "BLOCKED")

    def test_http_429_blocks(self):
        self.mode["status"] = 429
        with self.assertRaises(F.Stop) as cm:
            self.apply()
        self.assertEqual(cm.exception.kind, "BLOCKED")

    def test_logged_out(self):
        self.mode["job"] = "login"
        with self.assertRaises(F.Stop) as cm:
            self.apply()
        self.assertEqual(cm.exception.kind, "LOGIN_REQUIRED")

    def test_security_job_text_is_not_a_block(self):
        self.assertEqual(self.apply().code, "DRY_RUN_OK")  # the mock description says "unusual activity" on a long page


if __name__ == "__main__":
    unittest.main()
