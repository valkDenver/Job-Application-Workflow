"""Batched tracker writes (design section 11): one load/save per batch, user status kept, restart on a concurrent save, retries."""
import os
import time
import unittest
from unittest import mock

import openpyxl

from helpers import TempDir, job, make_tracker
from dice_apply_common import Refused
import tracker_io as T


def row(path, jid):
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        it = wb["Master Tracker"].iter_rows(values_only=True)
        head = [str(h) for h in next(it)]
        for r in it:
            if r[0] == jid:
                return {h: (r[i] if i < len(r) else None) for i, h in enumerate(head)}
    finally:
        wb.close()


class TrackerIO(unittest.TestCase):
    def setUp(self):
        self.tmp = TempDir()
        d = self.tmp.__enter__()
        self.rt = d / "rt"
        self.rt.mkdir()
        self.path = make_tracker(d / "Master Job Tracker - Test.xlsx",
                                 [job(n, **({"Resume Path": "C:/keep/me.docx"} if n == 3 else {})) for n in range(1, 7)])

    def tearDown(self):
        self.tmp.__exit__()

    def test_read_jobs(self):
        jobs = T.read_jobs(self.path)
        self.assertEqual(len(jobs), 6)
        self.assertEqual(jobs[0]["Job ID"], "JOB-1")
        self.assertEqual(jobs[0]["Source Portal"], "Dice")

    def test_batch_of_five_one_save(self):
        for n in range(1, 6):
            T.queue_add(self.rt, f"JOB-{n}", "applied", f"Applied on Dice by dice-auto-apply with r{n}.docx.",
                        resume_path=f"C:/r/r{n}.docx", resume_filename=f"r{n}.docx", applied_date="2026-11-02")
        with mock.patch.object(T, "save_atomic", wraps=T.save_atomic) as saver:
            n, rep = T.flush(self.path, self.rt)
        self.assertEqual((n, saver.call_count), (5, 1))
        r1, r3 = row(self.path, "JOB-1"), row(self.path, "JOB-3")
        self.assertEqual((r1["Pipeline Status"], r1["Applied Date"], r1["Resume Filename"]), ("APPLIED", "2026-11-02", "r1.docx"))
        self.assertIn("by dice-auto-apply", r1["Notes"])
        self.assertNotIn("by me", r1["Notes"])
        self.assertEqual(r3["Resume Path"], "C:/keep/me.docx")                    # never overwritten
        self.assertEqual(T.queue_items(self.rt), [])
        before = os.path.getmtime(self.path)
        self.assertEqual(T.flush(self.path, self.rt), (0, []))                    # nothing waiting: no save
        self.assertEqual(os.path.getmtime(self.path), before)

    def test_user_status_kept(self):
        wb = openpyxl.load_workbook(self.path)
        ws = wb["Master Tracker"]
        ws.cell(3, 17).value = "SKIPPED"                                           # JOB-2 is row 3; Pipeline Status is column 17
        wb.save(self.path)
        T.queue_add(self.rt, "JOB-2", "applied", "Applied on Dice by dice-auto-apply.", applied_date="2026-11-02")
        n, rep = T.flush(self.path, self.rt)
        r = row(self.path, "JOB-2")
        self.assertEqual(r["Pipeline Status"], "SKIPPED")
        self.assertEqual(r["Applied Date"], "2026-11-02")
        self.assertTrue(any("TRACKER_CHANGED" in x for x in rep))

    def test_note_only_and_no_duplicate_lines(self):
        T.queue_add(self.rt, "JOB-4", "note", "dice-auto-apply: EXTERNAL_APPLY url=https://x.example")
        T.flush(self.path, self.rt)
        T.queue_add(self.rt, "JOB-4", "note", "dice-auto-apply: EXTERNAL_APPLY url=https://x.example")
        T.flush(self.path, self.rt)
        r = row(self.path, "JOB-4")
        self.assertEqual(r["Pipeline Status"], "REVIEW")
        self.assertEqual(r["Notes"].count("EXTERNAL_APPLY"), 1)

    def test_concurrent_save_restarts(self):
        T.queue_add(self.rt, "JOB-5", "applied", "Applied.", applied_date="2026-11-02")
        real_apply, calls = T._apply, []

        def touching_apply(*a, **k):
            calls.append(1)
            if len(calls) == 1:                                                    # someone saves while the first load is open
                t = time.time() + 5
                os.utime(self.path, (t, t))
            return real_apply(*a, **k)

        with mock.patch.object(T, "_apply", side_effect=touching_apply), mock.patch.object(T.time, "sleep"):
            n, _ = T.flush(self.path, self.rt)
        self.assertEqual((n, len(calls)), (1, 2))
        self.assertEqual(row(self.path, "JOB-5")["Notes"].count("Applied."), 1)

    def test_excel_lock_keeps_queue(self):
        T.queue_add(self.rt, "JOB-6", "applied", "Applied.")
        lock = self.path.parent / ("~$" + self.path.name)
        lock.write_text("x")
        with mock.patch.object(T, "IS_WIN", False):                               # the macOS rule: a ~$ file means Excel has it open
            with self.assertRaises(Refused):
                T.flush(self.path, self.rt)
            self.assertIn("waiting", T.try_flush(self.path, self.rt))
        self.assertEqual(len(T.queue_items(self.rt)), 1)
        lock.unlink()
        self.assertEqual(T.flush(self.path, self.rt)[0], 1)

    def test_applied_date_is_the_submit_day_not_the_write_day(self):
        item = T.queue_add(self.rt, "JOB-2", "applied", "Applied.")
        rows = T.queue_items(self.rt)
        rows[0]["queued"] = "2026-10-31"                                           # submitted yesterday; Excel was open overnight
        from dice_apply_common import write_jsonl
        write_jsonl(T.queue_path(self.rt), rows)
        T.flush(self.path, self.rt)
        self.assertEqual(row(self.path, "JOB-2")["Applied Date"], "2026-10-31")

    def test_job_removed_from_tracker_is_dropped(self):
        T.queue_add(self.rt, "JOB-999", "note", "x")
        n, rep = T.flush(self.path, self.rt)
        self.assertTrue(any("no longer in the tracker" in x for x in rep))
        self.assertEqual(T.queue_items(self.rt), [])

    def test_backup_once_a_day(self):
        T.queue_add(self.rt, "JOB-1", "note", "a")
        T.flush(self.path, self.rt)
        T.queue_add(self.rt, "JOB-1", "note", "b")
        T.flush(self.path, self.rt)
        self.assertEqual(len(list((self.path.parent / "Backups").glob("*.xlsx"))), 1)


if __name__ == "__main__":
    unittest.main()
