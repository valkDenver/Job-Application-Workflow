"""Tests for the Dashboard update (T12): the template's Applied and Skipped rows, extend-ranges bringing an older Dashboard up to
date, and setup_check noticing an older one. Offline, temporary folders. (The formulas' results were checked once in Excel:
Applied counts dated, FOLLOW-UP and APPLIED jobs and keeps CLOSED ones with an Applied Date.)

  python Tools/tests/test_dashboard.py      prints "all tests passed"
"""
import contextlib, io, os, shutil, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
TOOLS = os.path.dirname(HERE)
sys.path.insert(0, TOOLS)
import openpyxl
import tracker_tool as T
import setup_check as S

FAIL = []
TEMPLATE = os.path.join(T.APP_FOLDER, T.TEMPLATE_NAME)
OLD_APPLIED = "=COUNTIFS('Master Tracker'!$A$2:$A$1000,\"JOB-*\",'Master Tracker'!$Q$2:$Q$1000,\"APPLIED\")"
OLD_EXCLUDED = "=COUNTIF('Excluded Jobs'!$A$2:$A$1000,\"20*\")"


def check(name, ok, detail=""):
    if not ok:
        FAIL.append(name + (": " + str(detail)[:500] if detail else ""))


def dashboard(path):
    ws = openpyxl.load_workbook(path)["Dashboard"]
    return {ws.cell(r, 1).value: ws.cell(r, 2).value for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value}


def old_tracker(d):
    """A tracker made from an older template: the old Applied formula and the old Excluded Jobs row, ranges ending at row 1000."""
    rc = None
    saved = T.TRACKER_TXT
    T.TRACKER_TXT = os.path.join(d, "tracker_path.txt")
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            rc = T.main(["init", "--name", "Old One", "--folder", d])
    finally:
        T.TRACKER_TXT = saved
    assert rc == 0
    p = os.path.join(d, "Master Job Tracker - Old_One.xlsx")
    wb = openpyxl.load_workbook(p)
    ws = wb["Dashboard"]
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "Applied":
            ws.cell(r, 2).value = OLD_APPLIED
        if ws.cell(r, 1).value == "Skipped":
            ws.cell(r, 1).value, ws.cell(r, 2).value = T.OLD_EXCLUDED_LABEL, OLD_EXCLUDED
    wb.save(p)
    return p


def test_template_is_up_to_date():
    dash = dashboard(TEMPLATE)
    check("template Applied uses the Applied Date", "'Master Tracker'!$U$2:$U$100000,\"<>\"" in dash.get("Applied", ""), dash.get("Applied"))
    check("template has Skipped", "\"SKIPPED\"" in (dash.get("Skipped") or ""), dash.get("Skipped"))
    check("template has no old rows", T.old_dashboard_rows(openpyxl.load_workbook(TEMPLATE)) == [])


def test_extend_ranges_updates_an_old_dashboard_once():
    with tempfile.TemporaryDirectory() as d:
        p = old_tracker(d)
        check("setup_check notices it", "Dashboard is out of date" in S.check_tracker(p), S.check_tracker(p))
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            rc = T.main(["extend-ranges", p])
        dash = dashboard(p)
        check("Applied replaced", rc == 0 and "$U$2:$U$100000" in dash.get("Applied", ""), (buf.getvalue(), dash.get("Applied")))
        check("Excluded row became Skipped", T.OLD_EXCLUDED_LABEL not in dash and "SKIPPED" in (dash.get("Skipped") or ""), dash)
        check("other rows untouched", dash.get("Tier 1", "").startswith("=COUNTIFS") and "Follow-Up Required" in dash)
        check("setup_check is happy afterwards", S.check_tracker(p) == "", S.check_tracker(p))
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            T.main(["extend-ranges", p])
        check("second run changes nothing", "Nothing to change" in buf.getvalue(), buf.getvalue())


def test_columns_found_by_name():
    with tempfile.TemporaryDirectory() as d:
        p = old_tracker(d)
        wb = openpyxl.load_workbook(p)
        ws = wb[T.MASTER]
        head = [c.value for c in ws[1]]
        a, b = head.index("Applied Date") + 1, head.index("Follow-Up Date") + 1  # an older tracker with these two swapped
        ws.cell(1, a).value, ws.cell(1, b).value = "Follow-Up Date", "Applied Date"
        wb.save(p)
        with contextlib.redirect_stdout(io.StringIO()):
            T.main(["extend-ranges", p])
        from openpyxl.utils import get_column_letter as L
        check("Applied uses this tracker's Applied Date column", f"${L(b)}$2:${L(b)}$100000,\"<>\"" in dashboard(p)["Applied"], dashboard(p)["Applied"])


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
    if FAIL:
        print("FAILED:\n  " + "\n  ".join(FAIL)); sys.exit(1)
    print("all tests passed")
