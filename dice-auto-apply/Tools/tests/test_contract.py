"""The data contract with application-workflow (design section 5). Reads that folder's template and tracker_tool.py as text;
never imports its code. Skipped when application-workflow is not next to this folder."""
import ast
import re
import unittest
from pathlib import Path

import openpyxl

import helpers  # noqa: F401  (puts Tools/ on the import path)
import tracker_io as T

AW = Path(__file__).resolve().parents[3] / "application-workflow"
TEMPLATE = AW / "Master Job Tracker - TEMPLATE.xlsx"
TOOL = AW / "Tools" / "tracker_tool.py"
RESUME_TOOL = AW / "Tools" / "resume_path.py"


@unittest.skipUnless(TEMPLATE.exists() and TOOL.exists(), "application-workflow is not next to dice-auto-apply")
class Contract(unittest.TestCase):
    def test_columns_exist_in_template(self):
        wb = openpyxl.load_workbook(TEMPLATE, read_only=True)
        try:
            head = {str(h).strip().lower() for h in next(wb[T.MASTER].iter_rows(max_row=1, values_only=True)) if h}
        finally:
            wb.close()
        optional = {"email drafted date"}  # tracker_tool.py adds it when missing
        missing = [c for c in T.COLUMNS if c.lower() not in head and c.lower() not in optional]
        self.assertEqual(missing, [], "columns this flow reads are missing from the template")

    def test_status_rank_matches(self):
        src = TOOL.read_text(encoding="utf-8")
        m = re.search(r"^RANK = (\{.*?\})", src, re.S | re.M)
        self.assertIsNotNone(m)
        self.assertEqual(ast.literal_eval(re.sub(r"#.*", "", m.group(1))), T.RANK)

    def test_resume_layout_matches(self):
        src = RESUME_TOOL.read_text(encoding="utf-8")
        self.assertIn('ROOT_NAME = "Job Application Workflow"', src)
        self.assertIn('return os.path.join(documents_dir(), ROOT_NAME, "Resumes")', src)
        self.assertRegex(src, r'base == jid or base.startswith\(jid \+ " "\)')


if __name__ == "__main__":
    unittest.main()
