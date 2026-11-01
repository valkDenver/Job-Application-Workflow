"""Resume resolution (design section 7): Resume Path first, Resume Filename search as the fallback, never guessing."""
import unittest
from pathlib import Path

from helpers import TempDir, make_resume
import resume_find as RF


class ParseCells(unittest.TestCase):
    def test_path_cell_kinds(self):
        self.assertEqual(RF.parse_path_cell("")[0], "empty")
        self.assertEqual(RF.parse_path_cell("https://onedrive.live.com/x")[0], "link")
        self.assertEqual(RF.parse_path_cell("unknown (saved as Pat_Workday.docx)"), ("name", "Pat_Workday.docx"))
        self.assertEqual(RF.parse_path_cell('"C:\\Users\\pat\\r.docx"')[0], "path")
        self.assertEqual(RF.parse_path_cell("/Users/pat/r.docx")[0], "path")
        self.assertEqual(RF.parse_path_cell("Pat_Workday.docx"), ("name", "Pat_Workday.docx"))

    def test_filename_cell_kinds(self):
        self.assertEqual(RF.parse_filename_cell("  'Pat_Workday.docx' "), ("name", "Pat_Workday.docx"))
        self.assertEqual(RF.parse_filename_cell("/Users/pat/Pat.docx")[0], "path")

    def test_same_name(self):
        self.assertTrue(RF._same_name("Pat_Workday.docx", "pat_workday.DOCX"))
        self.assertTrue(RF._same_name("Pat_Workday.docx", "Pat_Workday"))
        self.assertFalse(RF._same_name("Pat_Workday.docx", "Pat_Java.docx"))
        self.assertFalse(RF._same_name("Pat_Workday.docx", "Pat_Workday.pdf"))


class Resolve(unittest.TestCase):
    def setUp(self):
        self.tmp = TempDir()
        self.root = self.tmp.__enter__() / "Resumes"
        self.wd = make_resume(self.root, "Workday", "HCM", "JOB-1", "Pat_Workday HCM.docx")
        self.other = make_resume(self.root, "Workday", "HCM", "JOB-1", "Pat_Old.docx")
        self.java = make_resume(self.root, "Java", "", "JOB-2", "Pat_Java.docx")
        self.idx = RF.ResumeIndex(self.root)

    def tearDown(self):
        self.tmp.__exit__()

    def r(self, job, path="", name=""):
        return RF.resolve(job, path, name, self.idx)

    def test_valid_path_wins(self):
        res = self.r("JOB-1", str(self.wd), "Pat_Workday HCM.docx")
        self.assertTrue(res.ok)
        self.assertEqual((res.file, res.via), (self.wd, "path"))

    def test_missing_path_falls_back_to_name(self):
        res = self.r("JOB-1", str(self.root / "gone" / "Pat_Workday HCM.docx"), "")
        self.assertEqual((res.file, res.via), (self.wd, "name search"))

    def test_link_falls_back(self):
        res = self.r("JOB-1", "https://drive.google.com/abc", "Pat_Workday HCM.docx")
        self.assertEqual(res.via, "name search")

    def test_unknown_saved_as(self):
        res = self.r("JOB-1", "unknown (saved as Pat_Workday HCM.docx)", "")
        self.assertEqual(res.file, self.wd)

    def test_filename_without_extension_and_quotes(self):
        self.assertEqual(self.r("JOB-1", "", ' "Pat_Workday HCM" ').file, self.wd)

    def test_full_path_in_filename_column(self):
        self.assertEqual(self.r("JOB-2", "", str(self.java)).file, self.java)

    def test_mismatch(self):
        res = self.r("JOB-1", str(self.wd), "Pat_Old.docx")
        self.assertEqual(res.code, "RESUME_MISMATCH")

    def test_no_guessing_from_single_file(self):
        res = self.r("JOB-2", "", "Pat_Something_Else.docx")
        self.assertEqual(res.code, "RESUME_NOT_FOUND")
        self.assertIn("name not found", res.reason)

    def test_both_empty(self):
        self.assertEqual(self.r("JOB-1").code, "NO_RESUME")

    def test_folder_path(self):
        self.assertEqual(self.r("JOB-1", str(self.wd.parent), "Pat_Workday HCM.docx").file, self.wd)
        self.assertEqual(self.r("JOB-1", str(self.wd.parent), "").code, "RESUME_AMBIGUOUS")

    def test_ambiguous_name_search(self):
        make_resume(self.root, "Workday", "Finance", "JOB-1", "Pat_Workday HCM.docx")
        res = RF.resolve("JOB-1", "", "Pat_Workday HCM.docx", RF.ResumeIndex(self.root))
        self.assertEqual(res.code, "RESUME_AMBIGUOUS")

    def test_too_large_and_lock_file(self):
        big = make_resume(self.root, "Java", "", "JOB-3", "Big.docx", size=3 * 1024 * 1024)
        self.assertEqual(RF.resolve("JOB-3", str(big), "", RF.ResumeIndex(self.root)).code, "RESUME_TOO_LARGE")
        lock = make_resume(self.root, "Java", "", "JOB-4", "~$Pat.docx")
        self.assertEqual(RF.resolve("JOB-4", str(lock), "", self.idx).code, "RESUME_NOT_FOUND")

    def test_tech_of(self):
        self.assertEqual(RF.tech_of(self.wd, self.root), ("Workday", "HCM"))
        self.assertEqual(RF.tech_of(self.java, self.root), ("Java", ""))
        self.assertEqual(RF.tech_of(Path("C:/elsewhere/r.docx"), self.root), (None, None))


if __name__ == "__main__":
    unittest.main()
