"""Builds a realistic test tracker from the user's TestSample file and checks every resume and recruiter case against it (plan, Phase 0).

Usage:  python make_test_tracker.py "<path to Master Job Tracker - TestSample.xlsx>" [--check] [--today YYYY-MM-DD]

Copies TestSample to Tools/testdata/tracker_copy.xlsx (git-ignored; same format and volume as the real tracker), then fills 22 Dice rows
with made-up resume cases and recruiters, and creates their resume files under Tools/testdata/Resumes/ (a path with spaces).
Writes Tools/testdata/expected.json. With --check it runs the selection rules and prints one line per case: PASS or FAIL.
Re-running rebuilds everything from scratch. Never touches the TestSample file itself.
"""
import argparse
import datetime as dt
import json
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
import openpyxl  # noqa: E402

import dice_apply_select as SEL  # noqa: E402
import resume_find as RF  # noqa: E402
import tracker_io as T  # noqa: E402

DATA = HERE.parent / "testdata"
ROOT = DATA / "Resumes"
COPY = DATA / "tracker_copy.xlsx"


def mk(tech, spec, jid, name, size=4000):
    folder = ROOT / tech / spec / f"{jid} - Test Client (via Test Staffing)" if spec else ROOT / tech / f"{jid} - Test Client (via Test Staffing)"
    folder.mkdir(parents=True, exist_ok=True)
    p = folder / name
    p.write_bytes(b"test resume " * (size // 12))
    return p


def build(sample, today):
    if DATA.exists():
        shutil.rmtree(DATA)
    DATA.mkdir(parents=True)
    shutil.copyfile(sample, COPY)
    wb = openpyxl.load_workbook(COPY)
    ws = wb[T.MASTER]
    hm = T.header_map([c.value for c in ws[1]])
    col = lambda n: T.need(hm, n) + 1
    rows = [r for r in range(2, ws.max_row + 1)
            if str(ws.cell(r, col("Source Portal")).value or "").lower() == "dice"
            and str(ws.cell(r, col("Pipeline Status")).value or "").upper() == "REVIEW"
            and not ws.cell(r, col("Email Drafted Date")).value][:22]
    jid = [str(ws.cell(r, col("Job ID")).value).upper() for r in rows]
    exp = {}

    def setc(i, **kw):
        for k, v in kw.items():
            ws.cell(rows[i], col(k.replace("_", " "))).value = v

    for i, r in enumerate(rows):  # made-up, unique recruiters on every case row
        setc(i, Recruiter_Name=f"Case Recruiter{i}", Recruiter_Email=f"case{i}@staffing{i}.example", Staffing_Company=f"Case Staffing {i}")

    p0 = mk("Workday", "HCM", jid[0], "Case_Workday HCM.docx"); setc(0, Resume_Path=str(p0)); exp[jid[0]] = ["OK", "path"]
    mk("Java", "Backend", jid[1], "Case_Java.docx")
    setc(1, Resume_Path=f"/Users/someone/Documents/Job Application Workflow/Resumes/Java/Backend/{jid[1]} - X/Case_Java.docx")
    exp[jid[1]] = ["OK", "name search"] if not Path("/Users/someone").exists() else ["OK", "path"]
    p2 = mk("DevOps", "", jid[2], "Case_DevOps.docx"); setc(2, Resume_Path=str(p2.parent)); exp[jid[2]] = ["OK", "path folder"]
    mk("AWS", "", jid[3], "Case_AWS.docx"); setc(3, Resume_Path="https://onedrive.live.com/case3", Resume_Filename="Case_AWS.docx")
    exp[jid[3]] = ["OK", "name search"]
    mk("Python", "", jid[4], "Case_Python.docx"); setc(4, Resume_Path="unknown (saved as Case_Python.docx)"); exp[jid[4]] = ["OK", "name search"]
    mk("Azure", "", jid[5], "Case_Azure.docx"); setc(5, Resume_Filename="Case_Azure.docx"); exp[jid[5]] = ["OK", "name search"]
    mk("GCP", "", jid[6], "Case_GCP.docx"); setc(6, Resume_Filename="Case_GCP"); exp[jid[6]] = ["OK", "name search"]
    mk("React", "", jid[7], "Case React.docx"); setc(7, Resume_Filename=' "Case React.docx" '); exp[jid[7]] = ["OK", "name search"]
    p8 = mk("Node", "", jid[8], "Case_Node.docx"); setc(8, Resume_Filename=str(p8)); exp[jid[8]] = ["OK", "path (in Resume Filename)"]
    p9 = mk("Salesforce", "", jid[9], "Case_SF.docx"); setc(9, Resume_Path=str(p9), Resume_Filename="Case_SF.docx"); exp[jid[9]] = ["OK", "path"]
    p10 = mk("ServiceNow", "", jid[10], "Case_SN.docx"); setc(10, Resume_Path=str(p10), Resume_Filename="Case_Other.docx")
    exp[jid[10]] = ["RESUME_MISMATCH"]
    setc(11, Resume_Path=str(ROOT / "Gone" / "Case_Missing.docx")); exp[jid[11]] = ["RESUME_NOT_FOUND"]
    mk("Workday", "HCM", jid[12], "Case_Twice.docx"); mk("Workday", "Finance", jid[12], "Case_Twice.docx")
    setc(12, Resume_Filename="Case_Twice.docx"); exp[jid[12]] = ["RESUME_AMBIGUOUS"]
    same = dict(Recruiter_Name="Sam Caserecruiter", Recruiter_Email="sam.case@samestaffing.example", Staffing_Company="Same Staffing")
    mk("Workday", "Integration", jid[15], "Case_Earlier.docx")
    setc(15, Resume_Filename="Case_Earlier.docx", Pipeline_Status="APPLIED", Applied_Date="2026-10-20", **same); exp[jid[15]] = ["STATUS_APPLIED"]
    mk("Workday", "Integration", jid[13], "Case_Similar.docx"); setc(13, Resume_Filename="Case_Similar.docx", **same); exp[jid[13]] = ["OK", "name search"]
    mk("Java", "", jid[14], "Case_Different.docx"); setc(14, Resume_Filename="Case_Different.docx", **same)
    exp[jid[14]] = ["SAME_RECRUITER " + jid[15]]
    mk("IAM", "", jid[16], "Case_Mailed.docx")
    setc(16, Resume_Filename="Case_Mailed.docx", Email_Drafted_Date="2026-10-29 (resume attached)", Pipeline_Status="EMAIL DRAFTED")
    exp[jid[16]] = ["EMAILED_RESUME"]
    mk("IAM", "", jid[17], "Case_NoResumeMail.docx")
    setc(17, Resume_Filename="Case_NoResumeMail.docx", Email_Drafted_Date="2026-10-29 (no resume)", Pipeline_Status="EMAIL DRAFTED")
    exp[jid[17]] = ["OK", "name search"]
    for i, st in ((18, "CLOSED"), (19, "SKIPPED"), (20, "FOLLOW-UP")):
        mk("API", "", jid[i], f"Case_{i}.docx"); setc(i, Resume_Filename=f"Case_{i}.docx", Pipeline_Status=st)
        exp[jid[i]] = ["STATUS_" + st]
    mk("Scala", "", jid[21], "Case_Today.docx")
    setc(21, Resume_Filename="Case_Today.docx", Run_Timestamp=f"{today} 07:00 MST"); exp[jid[21]] = ["OK", "name search"]
    wb.save(COPY)
    (DATA / "expected.json").write_text(json.dumps({"today": today, "run_date_today": [jid[21]], "cases": exp}, indent=2), encoding="utf-8")
    print(f"Built {COPY} with {len(exp)} cases; resumes under {ROOT}")


def check(today):
    exp = json.loads((DATA / "expected.json").read_text(encoding="utf-8"))
    jobs = T.read_jobs(COPY)
    cfg = {"paces": {}, "max_resume_mb": 2}
    day = dt.date.fromisoformat(exp["today"])
    p = SEL.build_plan(jobs, {}, cfg, set(), set(), RF.ResumeIndex(ROOT), today=day)
    got = {j["Job ID"]: ["OK", r.via] for j, r, _ in p.picked}
    got.update({j["Job ID"]: [c] for j, c, _ in p.skipped})
    fails = 0
    for jid, want in exp["cases"].items():
        ok = got.get(jid) == want
        fails += not ok
        print(f"{'PASS' if ok else 'FAIL'}  {jid:<9} want {want}  got {got.get(jid)}")
    p2 = SEL.build_plan(jobs, {"run_date": "today"}, cfg, set(), set(), RF.ResumeIndex(ROOT), today=day)
    ids = [j["Job ID"] for j, _, _ in p2.picked]
    ok = ids == exp["run_date_today"]
    fails += not ok
    print(f"{'PASS' if ok else 'FAIL'}  --run-date today picks {ids}")
    print(f"{len(exp['cases']) + 1 - fails} passed, {fails} failed; {len(p.picked)} picked of {p.counts.get('dice_rows')} Dice rows")
    return 1 if fails else 0


def main(argv):
    a = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    a.add_argument("sample")
    a.add_argument("--check", action="store_true")
    a.add_argument("--today", default=dt.date.today().isoformat())
    a = a.parse_args(argv)
    build(Path(a.sample), a.today)
    return check(a.today) if a.check else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
