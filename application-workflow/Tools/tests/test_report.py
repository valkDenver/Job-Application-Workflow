"""Tests for tracker_tool.py report and dated set --note lines (T13). Offline, temporary folders.

  python Tools/tests/test_report.py      prints "all tests passed"
"""
import contextlib, datetime as dt, io, json, os, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
TOOLS = os.path.dirname(HERE)
sys.path.insert(0, TOOLS)
import openpyxl
import tracker_tool as T

FAIL = []
TODAY = dt.date.today().isoformat()


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


def tracker(d, rows):
    """rows: [{column: value}] added to a fresh tracker made by init."""
    run(d, "init", "--name", "Report Test", "--folder", d)
    p = os.path.join(d, "Master Job Tracker - Report_Test.xlsx")
    wb = openpyxl.load_workbook(p)
    ws = wb[T.MASTER]
    head = [c.value for c in ws[1]]
    for i, row in enumerate(rows, 2):
        for k, v in row.items():
            ws.cell(i, head.index(k) + 1).value = v
    wb.save(p)
    return p


ROWS = [
    {"Job ID": "JOB-1001", "Pipeline Status": "FOLLOW-UP", "Date Found": "2026-11-16", "Vendor Call Date": "2026-11-17",
     "Email Drafted Date": "2026-11-17 (resume attached)", "Applied Date": "2026-11-18",
     "Notes": "2026-11-17: Resume generated\n\n2026-11-19: Follow-up #1 (no reply): Version B, rate ask, Gmail thread x"},
    {"Job ID": "JOB-1002", "Pipeline Status": "EMAIL DRAFTED", "Date Found": "2026-11-18", "Email Drafted Date": "2026-11-18 (no resume)",
     "Notes": "Resume generated\n\n2026-11-18, Situation 3, Version A"},  # an old undated line is not counted
    {"Job ID": "JOB-1003", "Pipeline Status": "SKIPPED", "Date Found": "2026-11-19",
     "Notes": "2026-11-19: Skipped by me (was REVIEW): rate too low"},
    {"Job ID": "JOB-1004", "Pipeline Status": "RESUME CREATED", "Date Found": "2026-11-10",
     "Notes": "2026-11-12: Resume reused from the library\n2026-11-20: Resume generated"},
    {"Job ID": "JOB-1005", "Pipeline Status": "VENDOR CALLED", "Date Found": dt.datetime(2026, 11, 20), "Vendor Call Date": dt.datetime(2026, 11, 22)},
    {"Job ID": "JOB-1006", "Pipeline Status": "REVIEW", "Date Found": "2026-11-21"},
]


def test_report_counts_a_period():
    with tempfile.TemporaryDirectory() as d:
        p = tracker(d, ROWS)
        rc, o = run(d, "report", p, "--since", "2026-11-16", "--until", "2026-11-21", "--json")
        n = json.loads(o)
        want = {"jobs_received": 5, "vendor_calls": 1, "resumes_new": 2, "resumes_reused": 0, "emails_with_resume": 1,
                "emails_without_resume": 1, "followups": 1, "applied": 1, "skipped": 1, "first_contacts": 2}
        check("counts", rc == 0 and all(n[k] == v for k, v in want.items()), {k: (n.get(k), v) for k, v in want.items() if n.get(k) != v})
        check("days to first contact", n["days_to_first_contact"] == {"same day": 1, "1 day": 1, "2 or more days": 0}, n["days_to_first_contact"])
        rc, o = run(d, "report", p, "--since", "2026-11-22", "--until", "2026-11-22")
        check("Excel dates count too", "Vendor calls:      1" in o and "2 or more days 1" in o, o)
        rc, o = run(d, "report", p, "--since", "2026-11-20", "--until", "2026-11-19")
        check("bad period refused", rc == 2 and "after" in o, o)


def test_resume_notes_count_on_their_date_wherever_it_sits():
    with tempfile.TemporaryDirectory() as d:
        p = tracker(d, [
            {"Job ID": "JOB-1001", "Pipeline Status": "RESUME CREATED", "Notes": "Dice employment type: Contract | Resume generated 2026-11-17"},
            {"Job ID": "JOB-1002", "Pipeline Status": "RESUME CREATED", "Notes": "Resume reused 2026-11-18 (Module 3 library check): copied from JOB-1001"},
            {"Job ID": "JOB-1003", "Pipeline Status": "RESUME CREATED", "Notes": "2026-11-18 Resume reused from JOB-1001 (ATS 89%)"},
            {"Job ID": "JOB-1004", "Pipeline Status": "RESUME CREATED", "Notes": "Reused JOB-1001 payroll resume (ATS ~98%)"},  # no date: not in a period
            {"Job ID": "JOB-1005", "Pipeline Status": "CLOSED", "Notes": "Module 3 (2026-11-18): core requirement not on resume"},  # not a resume
        ])
        rc, o = run(d, "report", p, "--since", "2026-11-16", "--until", "2026-11-21", "--json")
        n = json.loads(o)
        check("dated resume notes in any position", (n["resumes_new"], n["resumes_reused"]) == (1, 2), (n["resumes_new"], n["resumes_reused"]))
        check("resume_note", T.resume_note("Reused JOB-1001 payroll resume") == ("reused", None) and T.resume_note("not on resume") is None)


def test_reopened_skip_not_counted():
    with tempfile.TemporaryDirectory() as d:
        p = tracker(d, [{"Job ID": "JOB-1001", "Pipeline Status": "REVIEW",
                         "Notes": "2026-11-19: Skipped by me (was REVIEW): rate\n\n2026-11-19: Reopened by me (was skipped)"},
                        {"Job ID": "JOB-1002", "Pipeline Status": "SKIPPED", "Notes": "2026-11-19: Skipped by me (was REVIEW): visa"}])
        rc, o = run(d, "report", p, "--since", "2026-11-19", "--until", "2026-11-19", "--json")
        check("only jobs still skipped are counted", json.loads(o)["skipped"] == 1, o)


def test_set_note_is_dated():
    with tempfile.TemporaryDirectory() as d:
        p = tracker(d, [{"Job ID": "JOB-1001", "Pipeline Status": "REVIEW"}])
        run(d, "set", p, "JOB-1001", "--note", "Resume generated")
        run(d, "set", p, "JOB-1001", "--note", "2026-11-01: Follow-up #1 (no reply): Version A")
        notes = openpyxl.load_workbook(p)[T.MASTER].cell(2, 23).value
        check("undated note gets today's date", f"{TODAY}: Resume generated" in notes, notes)
        check("a dated note is kept as written", "2026-11-01: Follow-up #1" in notes and "Follow-up" not in notes.split(TODAY)[0][-5:], notes)
        rc, o = run(d, "report", p, "--days", "1", "--json")
        check("today's resume is counted", json.loads(o)["resumes_new"] == 1, o)


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
    if FAIL:
        print("FAILED:\n  " + "\n  ".join(FAIL)); sys.exit(1)
    print("all tests passed")
