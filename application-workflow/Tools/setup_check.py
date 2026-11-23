"""Checks that this computer is ready for the workflow (Modules 02-04), one OK or FIX line per item, like Dice's doctor.

Usage (from the application-workflow folder; python3 on a Mac):
  python Tools/setup_check.py [--tracker <file.xlsx>] [--profile <file.txt>] [--fix]

Checks:
  Python 3.9 or later
  the packages in requirements.txt (openpyxl, python-docx)
  the page-fit engine for three-page resumes: Microsoft Word on Windows, LibreOffice on a Mac or Linux
  the tracker (--tracker, else the one saved in tracker_path.txt): opens, has the Master Tracker and JD Store sheets,
    is not the TEMPLATE or TestSample file nor a job feed or the publisher's tracker, and has no EXAMPLE rows left; and it is in
    Google Drive (required: it is shared with valktech.den@gmail.com as Viewer; the sharing itself cannot be checked from here)
  the Candidate Profile (--profile, else the one "Candidate Profile - <name>.txt" in this folder): ready for the first resume or
    email (profile_check.py's resume level: personal details, employers and their dates, no template text left)
  the Resumes folder (Documents/Job Application Workflow/Resumes) can be written to

Without --fix it changes nothing. --fix only:
  installs the missing packages from requirements.txt with this Python (pip; --user on a Mac or Linux outside a virtual environment)
  saves --tracker in tracker_path.txt when that tracker passes its check (the tracker viewer and Dice read it from there)
Exit code 0 = all OK, 1 = something to fix.
"""
import argparse, glob, importlib, importlib.util, os, re, shutil, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
FOLDER = os.path.dirname(HERE)
REQUIREMENTS = os.path.join(FOLDER, "requirements.txt")
TRACKER_TXT = os.path.join(FOLDER, "tracker_path.txt")
IMPORT_NAME = {"python-docx": "docx"}  # pip name -> import name, where they differ
NOT_MINE = ("TEMPLATE", "TestSample")
PY = "python3" if sys.platform == "darwin" else "python"


class Report:
    def __init__(self):
        self.problems = 0

    def line(self, ok, what, fix=""):
        if not ok:
            self.problems += 1
        print(("OK   " if ok else "FIX  ") + what + (("  ->  " + fix) if fix and not ok else ""))


def requirements():
    """[(pip name, pinned version or None)] from requirements.txt."""
    out = []
    try:
        for ln in open(REQUIREMENTS, encoding="utf-8"):
            ln = ln.split("#")[0].strip()
            if ln:
                name, _, ver = ln.partition("==")
                out.append((name.strip(), ver.strip() or None))
    except OSError:
        pass
    return out


def missing_packages():
    importlib.invalidate_caches()
    return [n for n, _ in requirements() if importlib.util.find_spec(IMPORT_NAME.get(n, n)) is None]


def install_requirements():
    cmd = [sys.executable, "-m", "pip", "install", "-r", REQUIREMENTS]
    in_venv = sys.prefix != getattr(sys, "base_prefix", sys.prefix)
    if not sys.platform.startswith("win") and not in_venv:
        cmd.append("--user")
    print("Installing: " + " ".join(cmd[2:]))
    return subprocess.run(cmd).returncode == 0


def layout_engine():
    """(found, what) for the engine resume_render.py uses to fit resumes to three pages."""
    if sys.platform.startswith("win"):  # resume_render.py uses Word on Windows, never LibreOffice
        try:
            import winreg
            winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"))
            return True, "Microsoft Word"
        except OSError:
            return False, "Microsoft Word"
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe and os.path.exists("/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        exe = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    return bool(exe), "LibreOffice"


def saved_tracker():
    try:
        return open(TRACKER_TXT, encoding="utf-8-sig").read().strip().strip('"') or None
    except OSError:
        return None


def check_tracker(path):
    """'' when the tracker is usable, else what is wrong and how to fix it."""
    if not os.path.isfile(path):
        return "file not found"
    base = os.path.basename(path)
    if any(w.lower() in base.lower() for w in NOT_MINE):
        return "this is the TEMPLATE or TestSample file; copy the template and rename it Master Job Tracker - <your name>.xlsx"
    if re.search(r"\b(SHARED|LIVE)\b|^Job Feed - ", base):
        return "this is a job feed or the publisher's tracker, not your own; add its jobs to your own tracker with tracker_tool.py import-feed"
    try:
        import openpyxl
    except ImportError:
        return "openpyxl is not installed (see above)"
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
    except Exception as e:  # noqa: BLE001
        return "cannot open it (%s); close it in Excel and try again" % type(e).__name__
    try:
        for s in ("Master Tracker", "JD Store"):
            if s not in wb.sheetnames:
                return 'no "%s" sheet; start from Master Job Tracker - TEMPLATE.xlsx' % s
        left = [ws.title for ws in wb.worksheets if ws.title in ("Master Tracker", "JD Store", "Excluded Jobs")
                and any(r and str(r[0] or "").strip().upper() == "EXAMPLE" for r in ws.iter_rows(min_row=2, max_row=50, values_only=True))]
        if left:
            return "EXAMPLE rows left on %s; delete them (a missed one shows up as a fake job in the viewer), or make a fresh tracker with tracker_tool.py init" % ", ".join(left)
        sys.path.insert(0, HERE)
        import tracker_tool
        if tracker_tool.old_dashboard_rows(wb):
            return ("its Dashboard is out of date (Applied shrinks when a job moves to FOLLOW-UP); close it in Excel and run: "
                    "%s Tools/tracker_tool.py extend-ranges \"%s\"" % (PY, path))
        return ""
    finally:
        wb.close()


SHARE_WITH = "valktech.den@gmail.com"  # the candidate's tracker is shared with this address as Viewer (onboarding requirement)


def in_google_drive(path):
    """True when the file sits in a Google Drive for desktop folder: a "My Drive" or "Shared drives" folder in its path (G:\\My Drive\\...,
    ~/Library/CloudStorage/GoogleDrive-<account>/My Drive/...), or a GoogleDrive-<account> / "Google Drive" folder."""
    parts = [p.lower() for p in re.split(r"[\\/]+", os.path.abspath(path)) if p]
    return any(p in ("my drive", "shared drives", "google drive") or p.startswith("googledrive") for p in parts)


def find_profile(given):
    """(path or None, problem)."""
    if given:
        return (given, "") if os.path.isfile(given) else (None, "file not found: %s" % given)
    mine = [p for p in glob.glob(os.path.join(FOLDER, "Candidate Profile - *.txt"))
            if not any(w.lower() in os.path.basename(p).lower() for w in NOT_MINE)]
    if not mine:
        return None, "none yet; copy Candidate Profile - TEMPLATE.txt to Candidate Profile - <your name>.txt and fill it in"
    if len(mine) > 1:
        return None, "more than one (%s); give the one to use with --profile" % ", ".join(os.path.basename(p) for p in mine)
    return mine[0], ""


def resumes_folder():
    """(ok, folder, note)."""
    sys.path.insert(0, HERE)
    try:
        import resume_path
        root = resume_path.resumes_root()
    except Exception:  # noqa: BLE001
        root = os.path.join(os.path.expanduser("~"), "Documents", "Job Application Workflow", "Resumes")
    if os.path.isdir(root):
        return os.access(root, os.W_OK), root, ""
    p = root
    while p and not os.path.isdir(p):
        parent = os.path.dirname(p)
        if parent == p:
            break
        p = parent
    return os.access(p, os.W_OK), root, " (created with your first resume)"


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--tracker")
    ap.add_argument("--profile")
    ap.add_argument("--fix", action="store_true")
    a = ap.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass
    r = Report()
    print("Checking with %s (Python %s)" % (sys.executable, sys.version.split()[0]))

    r.line(sys.version_info[:2] >= (3, 9), "Python %s" % sys.version.split()[0], "install Python 3.9 or later from python.org (Windows: tick \"Add to PATH\")")

    if not requirements():
        r.line(False, "requirements.txt", "it is missing from the application-workflow folder; get the folder again")
    else:
        miss = missing_packages()
        if miss and a.fix and install_requirements():
            miss = missing_packages()
        names = ", ".join(n for n, _ in requirements())
        r.line(not miss, "packages: " + names if not miss else "packages missing: " + ", ".join(miss),
               "%s Tools/setup_check.py --fix  (or: %s -m pip install -r requirements.txt)" % (PY, PY))

    found, what = layout_engine()
    r.line(found, "%s, to fit resumes onto three pages" % what,
           ("install Microsoft Word" if what == "Microsoft Word" else "install LibreOffice (free) from libreoffice.org")
           + "; without it resumes are still saved, but not fitted to three pages")

    tracker = a.tracker or saved_tracker()
    if not tracker:
        r.line(False, "tracker", "create it with: %s Tools/tracker_tool.py init --name \"<your name>\"  (or, for a tracker you already have: "
                                 "%s Tools/setup_check.py --tracker \"<that file>\" --fix)" % (PY, PY))
    else:
        tracker = os.path.abspath(tracker)
        problem = check_tracker(tracker)
        r.line(not problem, "tracker: %s" % tracker, problem)
        if not problem and a.tracker and a.fix and os.path.normcase(saved_tracker() or "") != os.path.normcase(tracker):
            with open(TRACKER_TXT, "w", encoding="utf-8") as f:
                f.write(tracker + "\n")
            print("     saved in tracker_path.txt")
        elif not problem and a.tracker and os.path.normcase(saved_tracker() or "") != os.path.normcase(tracker):
            print("     not saved in tracker_path.txt yet (add --fix to save it)")
        if os.path.isfile(tracker):
            r.line(in_google_drive(tracker), "tracker is in Google Drive (shared with %s as Viewer, so your progress is followed)" % SHARE_WITH,
                   "required: move it into your Google Drive folder (Google Drive for desktop), for example \"<My Drive>/Job Application "
                   "Workflow/\", then run %s Tools/setup_check.py --tracker \"<new location>\" --fix and share it with %s as Viewer"
                   % (PY, SHARE_WITH))

    prof, problem = find_profile(a.profile)
    if not prof:
        r.line(False, "Candidate Profile", problem)
    else:
        sys.path.insert(0, HERE)
        import profile_check
        probs = profile_check.check_resume(profile_check.read(prof))
        r.line(not probs, "Candidate Profile: %s" % os.path.basename(prof),
               "%d thing(s) to fill or fix (first: %s); see them all with %s Tools/profile_check.py, or say \"Update my profile\""
               % (len(probs), probs[0] if probs else "", PY))

    ok, root, note = resumes_folder()
    r.line(ok, "Resumes folder: %s%s" % (root, note), "this folder cannot be written to; check its permissions or OneDrive")

    print("All good." if not r.problems else "%d thing(s) to fix." % r.problems)
    return 0 if not r.problems else 1


if __name__ == "__main__":
    sys.exit(main())
