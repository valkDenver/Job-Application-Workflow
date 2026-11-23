"""Tests for Submitted Date (T15): set --submitted keeps the first date, related shows submissions (same client through any vendor,
and same recruiter even without a resume), report counts them, edit (the viewer) can fill it. Offline, temporary folders.

  python Tools/tests/test_submitted.py      prints "all tests passed"
"""
import contextlib, io, json, os, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
TOOLS = os.path.dirname(HERE)
sys.path.insert(0, TOOLS)
import openpyxl
import tracker_tool as T

FAIL = []


def check(name, ok, detail=""):
    if not ok:
        FAIL.append(name + (": " + str(detail)[:500] if detail else ""))


def run(d, *args):
    saved = T.TRACKER_TXT
    T.TRACKER_TXT = os.path.join(d, "tracker_path.txt")
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            rc = T.main(list(args))
    finally:
        T.TRACKER_TXT = saved
    return rc, buf.getvalue()


ROWS = [
    {"Job ID": "JOB-1001", "Job Title": "Okta Engineer", "Company": "Acme Health", "Staffing Company": "Vendor One",
     "Recruiter Email": "a@vendorone.com", "Pipeline Status": "APPLIED"},
    {"Job ID": "JOB-1002", "Job Title": "Okta Engineer", "Company": "Acme Health, Inc.", "Staffing Company": "Vendor Two",
     "Recruiter Email": "b@vendortwo.com", "Pipeline Status": "REVIEW"},
    {"Job ID": "JOB-1003", "Job Title": "SailPoint Engineer", "Company": "Other Bank", "Staffing Company": "Vendor One",
     "Recruiter Email": "a@vendorone.com", "Pipeline Status": "REVIEW"},
]


def make(d):
    run(d, "init", "--name", "Sub Test", "--folder", d)
    p = os.path.join(d, "Master Job Tracker - Sub_Test.xlsx")
    wb = openpyxl.load_workbook(p)
    ws = wb[T.MASTER]
    head = [c.value for c in ws[1]]
    for i, row in enumerate(ROWS, 2):
        for k, v in row.items():
            ws.cell(i, head.index(k) + 1).value = v
    wb.save(p)
    return p


def test_submitted_date():
    with tempfile.TemporaryDirectory() as d:
        p = make(d)
        rc, o = run(d, "related", p, "JOB-1002")
        check("nothing submitted yet", "SUBMITTED TO THE SAME CLIENT" not in o, o)
        rc, o = run(d, "set", p, "JOB-1001", "--submitted", "2026-11-20")
        check("submission recorded", rc == 0 and f"{T.SUBMIT_COL} = 2026-11-20" in o, o)
        rc, o = run(d, "set", p, "JOB-1001", "--submitted", "2026-11-21")
        check("first date kept", "kept at 2026-11-20" in o, o)
        rc, o = run(d, "related", p, "JOB-1002")
        check("same client through another vendor is flagged", "SUBMITTED TO THE SAME CLIENT" in o and "JOB-1001" in o and "Vendor One" in o, o)
        rc, o = run(d, "related", p, "JOB-1003")
        check("same recruiter's submitted job listed even without a resume", "SAME RECRUITER\tJOB-1001" in o and "2026-11-20" in o, o)
        check("different client not flagged as same client", "SUBMITTED TO THE SAME CLIENT" not in o, o)
        rc, o = run(d, "report", p, "--since", "2026-11-15", "--until", "2026-11-21", "--json")
        check("report counts it", json.loads(o)["submitted"] == 1, o)
        rc, o = run(d, "edit", p, "JOB-1003", "--set", f"{T.SUBMIT_COL}=2026-11-21")
        check("the viewer's edit can fill it", rc == 0, o)
        rc, o = run(d, "set", p, "JOB-1003", "--submitted", "soon")
        check("bad date refused", rc == 2 and "--submitted needs a date" in o, o)


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
    if FAIL:
        print("FAILED:\n  " + "\n  ".join(FAIL)); sys.exit(1)
    print("all tests passed")
