"""Builds small synthetic trackers and resume folders for the unit tests. Made-up names only (design section 14)."""
import os
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))

HEADERS = ["Job ID", "Priority", "Source Portal", "Run Timestamp", "Job Title", "Company", "Staffing Company",
           "Location / Work Arrangement", "Hourly Rate", "Rate Verified?", "C2C Status", "Posted / Posting Age", "Job URL",
           "Recruiter Name", "Recruiter Email", "Red Flags", "Pipeline Status", "Vendor Call Date", "Resume Filename", "Resume Path",
           "Applied Date", "Follow-Up Date", "Notes", "Applicants", "Recruiter Phone", "Date Found", "Seniority", "Verification Notes",
           "Why Niche / Potentially Niche", "Skip Reason", "Tier Reason", "Email Drafted Date"]


def guid(n):
    return f"{n:08x}-0000-4000-8000-{n:012x}"


def job(n, **kw):
    d = {"Job ID": f"JOB-{n}", "Priority": "Tier 2", "Source Portal": "Dice", "Run Timestamp": "2026-11-02 07:00 MST",
         "Job Title": f"Role {n}", "Company": f"Client {n}", "Staffing Company": f"Agency {n}",
         "Job URL": f"https://www.dice.com/job-detail/{guid(n)}", "Recruiter Name": f"Pat Example{n}",
         "Recruiter Email": f"pat{n}@agency{n}.example", "Pipeline Status": "REVIEW", "Date Found": f"2026-10-{10 + n % 15:02d}"}
    d.update(kw)
    return d


def make_tracker(path, rows):
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active.title = "Read Me"
    ws = wb.create_sheet("Master Tracker")
    ws.append(HEADERS)
    for r in rows:
        ws.append([r.get(h, "") for h in HEADERS])
    wb.create_sheet("JD Store").append(["Job ID", "Job Description"])
    wb.save(path)
    return Path(path)


def make_resume(root, tech, spec, job_id, name, size=2000):
    folder = Path(root) / tech / spec / f"{job_id} - Client (via Agency)" if spec else Path(root) / tech / f"{job_id} - Client (via Agency)"
    folder.mkdir(parents=True, exist_ok=True)
    p = folder / name
    p.write_bytes(b"x" * size)
    return p


class TempDir:
    def __enter__(self):
        self.td = tempfile.TemporaryDirectory(prefix="dice apply test ")  # a space in the path on purpose
        return Path(self.td.name)

    def __exit__(self, *exc):
        try:
            self.td.cleanup()
        except OSError:
            pass


def cfg(**kw):
    import dice_apply_config as C
    base = C.load()
    base.update(kw)
    return base


def env_runtime(path):
    os.environ["DICE_APPLY_RUNTIME"] = str(path)
