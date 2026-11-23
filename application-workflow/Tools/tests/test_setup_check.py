"""Tests for Tools/setup_check.py. Offline, temporary folders; never installs anything and never touches the real tracker_path.txt.

  python Tools/tests/test_setup_check.py      prints "all tests passed"
"""
import contextlib, io, os, shutil, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
TOOLS = os.path.dirname(HERE)
FOLDER = os.path.dirname(TOOLS)
sys.path.insert(0, TOOLS)
import openpyxl
import setup_check as S

FAIL = []
TEMPLATE = os.path.join(FOLDER, "Master Job Tracker - TEMPLATE.xlsx")


def check(name, ok, detail=""):
    if not ok:
        FAIL.append(name + (": " + str(detail)[:400] if detail else ""))


def clean_copy(src, dst):
    """The template with its EXAMPLE rows deleted, as a candidate's own tracker."""
    wb = openpyxl.load_workbook(src)
    for s in ("Master Tracker", "JD Store", "Excluded Jobs"):
        ws = wb[s]
        for row in range(ws.max_row, 1, -1):
            if str(ws.cell(row, 1).value or "").strip().upper() == "EXAMPLE":
                ws.delete_rows(row)
    wb.save(dst)


def run(d, *args):
    """Runs main() with the folder constants pointed at <d>; returns (exit code, output)."""
    saved = S.FOLDER, S.TRACKER_TXT
    S.FOLDER, S.TRACKER_TXT = d, os.path.join(d, "tracker_path.txt")
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            rc = S.main(list(args))
    finally:
        S.FOLDER, S.TRACKER_TXT = saved
    return rc, buf.getvalue()


def test_tracker_checks():
    with tempfile.TemporaryDirectory() as d:
        renamed = os.path.join(d, "Master Job Tracker - Me.xlsx")
        shutil.copy(TEMPLATE, renamed)
        check("EXAMPLE rows are reported", "EXAMPLE rows left on Master Tracker, JD Store, Excluded Jobs" in S.check_tracker(renamed), S.check_tracker(renamed))
        check("the template itself is refused", "TEMPLATE" in S.check_tracker(TEMPLATE))
        for name in ("Master Job Tracker - SHARED.xlsx", "Master Job Tracker - LIVE.xlsx", "Job Feed - pm-ba-h1b.xlsx"):
            p = os.path.join(d, name); shutil.copy(TEMPLATE, p)
            check("feed or publisher file refused: " + name, "not your own" in S.check_tracker(p), S.check_tracker(p))
        mine = os.path.join(d, "Master Job Tracker - Mine.xlsx")
        clean_copy(TEMPLATE, mine)
        check("a clean tracker passes", S.check_tracker(mine) == "", S.check_tracker(mine))
        check("a missing file is reported", S.check_tracker(os.path.join(d, "nope.xlsx")) == "file not found")


def test_fix_saves_tracker_path_and_profile_checks():
    with tempfile.TemporaryDirectory() as d:
        mine = os.path.join(d, "Master Job Tracker - Mine.xlsx")
        clean_copy(TEMPLATE, mine)
        rc, out = run(d, "--tracker", mine)
        check("without --fix nothing is saved", not os.path.exists(os.path.join(d, "tracker_path.txt")) and "add --fix" in out, out)
        rc, out = run(d, "--tracker", mine, "--fix")
        check("--fix saves the tracker", open(os.path.join(d, "tracker_path.txt"), encoding="utf-8").read().strip() == mine, out)
        rc, out = run(d)
        check("saved tracker is used", "OK   tracker: " + mine in out, out)
        check("no profile is reported", "FIX  Candidate Profile  ->  none yet" in out, out)
        shutil.copy(os.path.join(FOLDER, "Candidate Profile - TEMPLATE.txt"), os.path.join(d, "Candidate Profile - TEMPLATE.txt"))
        prof = os.path.join(d, "Candidate Profile - Me.txt")
        shutil.copy(os.path.join(FOLDER, "Candidate Profile - TEMPLATE.txt"), prof)
        rc, out = run(d)
        check("blank Name is reported (TEMPLATE ignored)", "FIX  Candidate Profile: Candidate Profile - Me.txt" in out, out)
        text = open(prof, encoding="utf-8").read().replace("Name: [Full name]", "Name: Test Person")
        open(prof, "w", encoding="utf-8").write(text)
        rc, out = run(d)
        check("Name alone is not enough", "FIX  Candidate Profile: Candidate Profile - Me.txt" in out and "profile_check.py" in out, out)
        sys.path.insert(0, HERE)
        import test_profile_check as TP
        open(prof, "w", encoding="utf-8").write(TP.clean_unused(TP.fill(TP.TEMPLATE, TP.three_employers())))
        rc, out = run(d)
        check("a profile ready for resumes passes", "OK   Candidate Profile: Candidate Profile - Me.txt" in out, out)
        shutil.copy(prof, os.path.join(d, "Candidate Profile - Other.txt"))
        rc, out = run(d)
        check("two profiles need --profile", "more than one" in out, out)
        rc, out = run(d, "--profile", prof)
        check("--profile picks one", "OK   Candidate Profile: Candidate Profile - Me.txt" in out, out)


def test_tracker_must_be_in_google_drive():
    check("Windows Drive path", S.in_google_drive(r"G:\My Drive\Job Application Workflow\Master Job Tracker - A.xlsx"))
    check("Mac Drive path", S.in_google_drive("/Users/a/Library/CloudStorage/GoogleDrive-a@gmail.com/My Drive/Job Application Workflow/t.xlsx"))
    check("local folder is not Drive", not S.in_google_drive(r"C:\Users\a\Documents\Job-Application-Workflow\application-workflow\t.xlsx"))
    with tempfile.TemporaryDirectory() as d:
        local = os.path.join(d, "Master Job Tracker - Local.xlsx")
        clean_copy(TEMPLATE, local)
        rc, out = run(d, "--tracker", local)
        check("local tracker gets a FIX", "FIX  tracker is in Google Drive" in out and S.SHARE_WITH in out, out)
        drive = os.path.join(d, "My Drive", "Job Application Workflow")
        os.makedirs(drive)
        mine = os.path.join(drive, "Master Job Tracker - Drive.xlsx")
        clean_copy(TEMPLATE, mine)
        rc, out = run(d, "--tracker", mine)
        check("Drive tracker passes", "OK   tracker is in Google Drive" in out, out)


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
    if FAIL:
        print("FAILED:\n  " + "\n  ".join(FAIL)); sys.exit(1)
    print("all tests passed")
