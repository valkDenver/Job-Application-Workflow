"""Tests for Email Sent Date (T14): set --email-sent keeps the first send, edit (the viewer) can fill it, unsent lists the right jobs
per job (not per recruiter), report counts sends, and add-columns adds the column to an older tracker. Offline, temporary folders.

  python Tools/tests/test_email_sent.py      prints "all tests passed"
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


def tracker(d, rows, drop_sent_col=False):
    run(d, "init", "--name", "Sent Test", "--folder", d)
    p = os.path.join(d, "Master Job Tracker - Sent_Test.xlsx")
    wb = openpyxl.load_workbook(p)
    ws = wb[T.MASTER]
    head = [c.value for c in ws[1]]
    if drop_sent_col:  # an older tracker without the column
        ws.cell(1, head.index(T.SENT_COL) + 1).value = None
        head = [c.value for c in ws[1]]
    for i, row in enumerate(rows, 2):
        for k, v in row.items():
            ws.cell(i, head.index(k) + 1).value = v
    wb.save(p)
    return p


def value(p, job, col):
    ws = openpyxl.load_workbook(p)[T.MASTER]
    head = [c.value for c in ws[1]]
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value == job:
            return ws.cell(r, head.index(col) + 1).value if col in head else None


ROWS = [  # one recruiter, two jobs: only one of them is sent
    {"Job ID": "JOB-1001", "Job Title": "Okta Engineer", "Pipeline Status": "EMAIL DRAFTED", "Recruiter Email": "r@agency.com",
     "Email Drafted Date": "2026-11-18 (no resume)", "Notes": "2026-11-18, Situation 3, Version A, Gmail thread 18c0aa11bb"},
    {"Job ID": "JOB-1002", "Job Title": "SailPoint Engineer", "Pipeline Status": "EMAIL DRAFTED", "Recruiter Email": "r@agency.com",
     "Email Drafted Date": "2026-11-19 (no resume)", "Notes": "2026-11-19, Situation 3, Version B"},
    {"Job ID": "JOB-1003", "Job Title": "Skipped one", "Pipeline Status": "SKIPPED", "Recruiter Email": "x@y.com", "Email Drafted Date": "2026-11-18"},
    {"Job ID": "JOB-1004", "Job Title": "Not drafted", "Pipeline Status": "REVIEW", "Recruiter Email": "z@y.com"},
]


def test_set_keeps_the_first_send_and_unsent_is_per_job():
    with tempfile.TemporaryDirectory() as d:
        p = tracker(d, ROWS)
        rc, o = run(d, "unsent", p)
        check("unsent lists both drafted jobs of the recruiter, not skipped or undrafted ones",
              "2 job(s)" in o and "JOB-1001" in o and "JOB-1002" in o and "JOB-1003" not in o and "JOB-1004" not in o, o)
        check("thread shown when recorded", "18c0aa11bb" in o, o)
        rc, o = run(d, "set", p, "JOB-1001", "--email-sent", "2026-11-19")
        check("send recorded", rc == 0 and value(p, "JOB-1001", T.SENT_COL) == "2026-11-19", o)
        rc, o = run(d, "set", p, "JOB-1001", "--email-sent", "2026-11-21")
        check("first send kept", value(p, "JOB-1001", T.SENT_COL) == "2026-11-19" and "kept at 2026-11-19" in o, o)
        rc, o = run(d, "unsent", p)
        check("the other job of the same recruiter is still unsent", "1 job(s)" in o and "JOB-1002" in o and "JOB-1001" not in o, o)
        rc, o = run(d, "set", p, "JOB-1002", "--email-sent", "19/11/2026")
        check("bad date refused", rc == 2 and "--email-sent needs a date" in o, o)
        rc, o = run(d, "report", p, "--since", "2026-11-15", "--until", "2026-11-21", "--json")
        check("report counts sends", json.loads(o)["emails_sent"] == 1, o)


def test_older_tracker_and_viewer_edit():
    with tempfile.TemporaryDirectory() as d:
        p = tracker(d, ROWS, drop_sent_col=True)
        rc, o = run(d, "set", p, "JOB-1002", "--email-sent", "2026-11-20")
        check("column created on first write", rc == 0 and value(p, "JOB-1002", T.SENT_COL) == "2026-11-20", o)
        other = os.path.join(d, "other")
        os.makedirs(other)
        p2 = tracker(other, ROWS, drop_sent_col=True)
        rc, o = run(d, "add-columns", p2)
        check("add-columns adds it", T.SENT_COL in [c.value for c in openpyxl.load_workbook(p2)[T.MASTER][1]], o)
        rc, o = run(d, "edit", p2, "JOB-1001", "--set", f"{T.SENT_COL}=2026-11-18")
        check("the viewer's edit can fill it", rc == 0 and value(p2, "JOB-1001", T.SENT_COL) == "2026-11-18", o)
        rc, o = run(d, "edit", p2, "JOB-1001", "--set", f"{T.SENT_COL}=tomorrow")
        check("edit checks the date", rc == 2, o)


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
    if FAIL:
        print("FAILED:\n  " + "\n  ".join(FAIL)); sys.exit(1)
    print("all tests passed")
