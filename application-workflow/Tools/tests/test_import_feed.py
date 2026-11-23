"""Tests for tracker_tool.py import-feed. Offline: synthetic feed and tracker files in a temporary folder; never touches the
real tracker_path.txt, feed_path.txt or any real feed.

  python Tools/tests/test_import_feed.py      prints "all tests passed"
"""
import contextlib, io, os, sys, tempfile

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
    saved = T.TRACKER_TXT, T.FEED_TXT
    T.TRACKER_TXT, T.FEED_TXT = os.path.join(d, "tracker_path.txt"), os.path.join(d, "feed_path.txt")
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            rc = T.main(list(args))
    finally:
        T.TRACKER_TXT, T.FEED_TXT = saved
    return rc, buf.getvalue()


def make_tracker(d):
    rc, out = run(d, "init", "--name", "Test Person", "--folder", d)
    assert rc == 0, out
    return os.path.join(d, "Master Job Tracker - Test_Person.xlsx")


def make_feed(d, numbers, extra_col=False, name="Master Job Tracker - SHARED.xlsx"):
    """A feed built like feed_builder.py: the template's sheets, one Master Tracker and one JD Store row per job."""
    p = os.path.join(d, name)
    if os.path.exists(p):
        wb = openpyxl.load_workbook(p)
    else:
        wb = openpyxl.load_workbook(os.path.join(T.APP_FOLDER, T.TEMPLATE_NAME))
        for s in T.EXAMPLE_SHEETS:
            for r in reversed(T.example_rows(wb[s])):
                wb[s].delete_rows(r)
    m, j = wb[T.MASTER], wb[T.JDS]
    head = [c.value for c in m[1]]
    if extra_col and "Feed Extra" not in head:
        m.cell(1, len([h for h in head if h]) + 1).value = "Feed Extra"
        head = [c.value for c in m[1]]
    for n in numbers:
        vals = {"Job ID": f"JOB-{n:04d}", "Priority": "Tier 1", "Job Title": f"Engineer {n}", "Job URL": f"https://example.com/jobs/{n}",
                "Pipeline Status": "REVIEW", "Recruiter Email": f"r{n}@example.com", "Feed Extra": f"x{n}"}
        m.append([vals.get(h) for h in head])
        j.append([f"JOB-{n:04d}", f"Engineer {n}", f"Description of job {n}"])
    wb.save(p)
    return p


def master(path):
    wb = openpyxl.load_workbook(path)
    ws = wb[T.MASTER]
    head = [c.value for c in ws[1]]
    rows = [{h: ws.cell(r, i + 1).value for i, h in enumerate(head) if h} for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value]
    links = {ws.cell(r, head.index("Job URL") + 1).hyperlink.target if ws.cell(r, head.index("Job URL") + 1).hyperlink else None
             for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value}
    jd = {ws2[0].value: ws2[2].value for ws2 in wb[T.JDS].iter_rows(min_row=2) if ws2[0].value}
    return rows, links, jd


def test_first_run_needs_a_start_then_imports_only_new_jobs():
    with tempfile.TemporaryDirectory() as d:
        trk = make_tracker(d)
        feed = make_feed(d, [1001, 1005, 1010, 1020])
        rc, out = run(d, "import-feed", trk, "--feed", feed)
        check("first run without a start is refused", rc == 2 and "--after" in out, out)
        rc, out = run(d, "import-feed", trk, "--feed", feed, "--after", "JOB-1005", "--dry-run")
        check("dry run reports", rc == 0 and "Would add 2 job(s), JOB-1010 to JOB-1020" in out, out)
        check("dry run changes nothing", master(trk)[0] == [] and not os.path.exists(os.path.join(d, "feed_path.txt")))
        rc, out = run(d, "import-feed", trk, "--feed", feed, "--after", "JOB-1005")
        rows, links, jd = master(trk)
        check("adds the jobs above the start", rc == 0 and [r["Job ID"] for r in rows] == ["JOB-1010", "JOB-1020"], out)
        check("descriptions added", jd == {"JOB-1010": "Description of job 1010", "JOB-1020": "Description of job 1020"}, jd)
        check("links kept", links == {"https://example.com/jobs/1010", "https://example.com/jobs/1020"}, links)
        check("feed remembered", open(os.path.join(d, "feed_path.txt"), encoding="utf-8").read().strip() == feed)
        rc, out = run(d, "import-feed", trk)
        check("second run: nothing new, no --feed needed", rc == 0 and "No new jobs" in out, out)
        make_feed(d, [1031, 1040])
        rc, out = run(d, "import-feed", trk)
        check("later runs add only the newer jobs", [r["Job ID"] for r in master(trk)[0]] == ["JOB-1010", "JOB-1020", "JOB-1031", "JOB-1040"], out)
        rc, out = run(d, "import-feed", trk, "--from", "JOB-1001")
        rows = master(trk)[0]
        check("--from brings back older jobs once, never a duplicate", sorted(r["Job ID"] for r in rows) ==
              ["JOB-1001", "JOB-1005", "JOB-1010", "JOB-1020", "JOB-1031", "JOB-1040"] and "4 already there" in out, out)
        check("feed file unchanged by import", openpyxl.load_workbook(feed)[T.MASTER].max_row == 7)


def test_columns_matched_by_name_and_progress_kept():
    with tempfile.TemporaryDirectory() as d:
        trk = make_tracker(d)
        wb = openpyxl.load_workbook(trk)  # an older tracker: Job Title and Priority swapped, and one job of my own
        ws = wb[T.MASTER]
        head = [c.value for c in ws[1]]
        a, b = head.index("Priority") + 1, head.index("Job Title") + 1
        ws.cell(1, a).value, ws.cell(1, b).value = "Job Title", "Priority"
        ws.append(["JOB-0007"] + [None] * (len(head) - 1))
        wb.save(trk)
        feed = make_feed(d, [1001, 1002], extra_col=True)
        rc, out = run(d, "import-feed", trk, "--feed", feed, "--from", "JOB-1001")
        rows = {r["Job ID"]: r for r in master(trk)[0]}
        check("values land by column name", rows["JOB-1002"]["Job Title"] == "Engineer 1002" and rows["JOB-1002"]["Priority"] == "Tier 1", rows.get("JOB-1002"))
        check("my own job kept", "JOB-0007" in rows)
        check("missing feed column added at the end", rows["JOB-1001"].get("Feed Extra") == "x1001" and "Feed Extra" in out, out)


def test_refusals():
    with tempfile.TemporaryDirectory() as d:
        trk = make_tracker(d)
        rc, out = run(d, "import-feed", trk)
        check("no feed known", rc == 2 and "--feed" in out, out)
        rc, out = run(d, "import-feed", trk, "--feed", os.path.join(d, "nope.xlsx"), "--after", "JOB-1001")
        check("missing feed file", rc == 2 and "not found" in out, out)
        rc, out = run(d, "import-feed", trk, "--feed", trk, "--after", "JOB-1001")
        check("tracker given as feed", rc == 2 and "your own tracker" in out, out)
        feed = make_feed(d, [1001])
        rc, out = run(d, "import-feed", trk, "--feed", feed, "--from", "JOB-1001", "--after", "JOB-1001")
        check("both starts", rc == 2 and "not both" in out, out)


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
    if FAIL:
        print("FAILED:\n  " + "\n  ".join(FAIL)); sys.exit(1)
    print("all tests passed")
