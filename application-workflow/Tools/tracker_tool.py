# NOTE: application-workflow/Tools/tracker_tool.py and publisher/Tools/tracker_tool.py are separate copies so the two flows stay independent. If you change the cleaning rules or shared helpers in one, make the same change in the other.
"""Small helper for the Master Job Tracker so a module reads or changes one job without loading the whole workbook
into the conversation. Reads print only what is asked for; writes touch only the cells named.

Usage:  python tracker_tool.py <command> <tracker.xlsx> [arguments]

Read commands (never change the file):
  get <JOB-####>                 every non-empty field of that job's Master Tracker row, then its full JD from the JD Store sheet
  related <JOB-####> [--limit 8] earlier resumes or submissions (a Resume Filename, Resume Path or Submitted Date is filled) of the same
                                 recruiter (same email, or same name at the same staffing company) and of the same staffing company with
                                 a different recruiter; first, a "SUBMITTED TO THE SAME CLIENT" line for jobs at the same end client
                                 (Company) that have a Submitted Date, through any vendor. Module 03 uses it for its soft warning before
                                 it creates a new resume, Module 02 for "already submitted to this client?"
  followups [--jobs JOB-####,...] [--on YYYY-MM-DD] [--due] [--limit 10]
                                 jobs whose email may need a follow-up (Module 04 follow-up mode), grouped by recruiter email.
                                 --jobs: exactly those jobs; --on: jobs whose email was drafted that day; --due (default): jobs whose
                                 Follow-Up Date has arrived, or, with no Follow-Up Date, whose email is at least 1 business day old.
                                 A job counts when its status is EMAIL DRAFTED, it has an Email Drafted Date, or Notes holds a
                                 Situation tag. SKIPPED and CLOSED jobs are left out. Prints at most --limit jobs (0 = all)
  next-id                        the next available Job ID (highest JOB number found + 1)
  drafted-notes [--exclude JOB-####]
                                 for jobs that have an email drafted (Pipeline Status EMAIL DRAFTED, an Email Drafted Date, or a Situation tag in Notes):
                                 only the style-tag lines of their Notes
                                 (Module 04 uses this to avoid repeating an earlier email style)
  last-run                       newest Run Timestamp (Master Tracker) and newest Date and Time of Run (Excluded Jobs)
  summary                        total jobs, first and last Job ID, count by Pipeline Status, and how many are skipped
  unsent [--limit 10]            jobs with an Email Drafted Date but no Email Sent Date (not SKIPPED or CLOSED), grouped by recruiter
                                 (--limit recruiters; 0 = all), with the Gmail thread of the first email when Notes records it.
                                 Module 04's "Check my sent emails" uses it
  report [--since YYYY-MM-DD] [--until YYYY-MM-DD] [--days 7] [--json]
                                 what was done in a period (default: the last 7 days up to today): jobs received (Date Found),
                                 vendor calls (Vendor Call Date), resumes new and reused and follow-ups drafted (dated Notes lines),
                                 emails sent (Email Sent Date), submitted (Submitted Date),
                                 emails drafted with and without a resume (Email Drafted Date; drafts, not sends), applied (Applied
                                 Date), skipped by me (dated Notes lines), and days from Date Found to first contact for the jobs first
                                 contacted in the period. --json prints the same numbers as one JSON object (for a report across trackers)
  skipped [--all] [--limit 50]   how many jobs are skipped (by me / automatic), then the ones skipped by me with their reasons
  check-contacts [--since JOB-####] [--limit 60]
                                 finds recruiter cells that hold description text, glued-on words (name@co.comPhone), notes inside the email cell,
                                 a phone that is not a number, or a name that does not match the address. --since <Job ID> checks only that job and the
                                 ones after it (for example the rows you just copied in). Prints the cell, the problem and the suggested value; changes nothing
  list [--status "RESUME NEEDED"] [--jobs JOB-####,...] [--limit 0]
                                 one line per job: Job ID, Job Title, Company, Pipeline Status, Resume Filename and the start of
                                 Why Niche (never the JD). --jobs lists exactly those jobs (and names any not found); without it,
                                 every job with one of the --status values. The batch resume run uses this to pick and group jobs
  recheck-list [--days 14] [--status REVIEW,SHORTLIST,SELECTED,FOUND] [--limit 200]
                                 first prints how many jobs match (per portal), then up to --limit rows (Job ID, portal, URL,
                                 status, Date Found); lists tracked jobs worth re-checking; --limit 0 lists all

Write commands (refuse to run while the workbook is open in Excel):
  set <JOB-####> [--set "Field=Value" ...] [--note "text"] [--confirmed-closed] [--email-drafted attached|none] [--email-sent YYYY-MM-DD] [--submitted YYYY-MM-DD]
                                 only these fields can be set: Pipeline Status, Vendor Call Date, Resume Filename, Resume Path,
                                 Follow-Up Date. Resume Path holds the resume's folder, not the file name: give the folder (it must hold
                                 a file named like the Resume Filename), or the full file path (checked, then its folder is stored), or a
                                 link. When only Resume Filename is given, the file is looked up in that job's folder under
                                 Documents/Job Application Workflow/Resumes and its folder is stored. A replaced path is noted.
                                 --email-sent records the day the first email for the job was actually sent, as Module 04 found
                                 it in Gmail sent mail (Email Sent Date, created if missing); an Email Sent Date already filled is kept.
                                 --submitted records the day the recruiter confirmed submitting me to the client (Submitted Date,
                                 created if missing); a Submitted Date already filled is kept.
                                 --note appends a paragraph to Notes (existing Notes are never replaced), starting with today's
                                 date unless it already starts with one (YYYY-MM-DD), so the report can count it.
                                 --email-drafted records the draft: Email Drafted Date = today plus "(resume attached)" or "(no resume)".
                                 Module 04 passes it with Pipeline Status=EMAIL DRAFTED; the date is written even when the status
                                 itself cannot move (already further along). The column is created at the end if it is missing.
                                 Pipeline Status only moves forward: a backward move is skipped and reported, the other fields
                                 are still written. A job whose status is SKIPPED or CLOSED is not changed unless --confirmed-closed.
  prune-excluded [--days 14]     delete Excluded Jobs rows older than N days, always keeping every row of the most recent run
  edit <JOB-####> [--set "Field=Value" ...] [--note "text"] [--reason "text"] [--if-unchanged-since EPOCH]
                                 what the user does by hand (used by the tracker viewer's edit panel). Fields: Pipeline Status (any
                                 value; SKIPPED needs --reason), Vendor Call Date, Resume Filename, Resume Path, Applied Date,
                                 Follow-Up Date, Email Drafted Date, Email Sent Date, Submitted Date, Recruiter Name, Recruiter Email, Recruiter Phone. An empty value clears a field.
                                 Setting VENDOR CALLED or EMAIL DRAFTED fills Vendor Call Date or Email Drafted Date with today
                                 when it is empty (never overwrites). Clearing a call date, email date, resume filename or path
                                 writes a dated Notes line. Editing the Email Drafted Date keeps its "(resume attached)" tag.
                                 Writes the dated Notes lines the modules rely on (status change, skipped/reopened, "recruiter email
                                 supplied by me"). Setting APPLIED fills Applied Date with today when it is empty.
                                 --if-unchanged-since refuses when the file was saved by someone else after that time.
  skip <JOB-####> --reason "text" [--force]
                                 the user decided not to pursue this job: Pipeline Status = SKIPPED, Skip Reason = "Manual: <text>",
                                 and a dated line in Notes (with the status it had). Adds the Skip Reason column if missing.
                                 Jobs already APPLIED or in FOLLOW-UP need --force. The row stays in the tracker, so the job is never
                                 re-added or re-researched.
  reopen <JOB-####> [--status REVIEW] [--note "text"]
                                 undo a skip: back to the status it had before (or --status), Skip Reason cleared, dated Notes line
  add-columns                    add the Skip Reason, Tier Reason, Email Drafted Date, Email Sent Date and Submitted Date columns to the Master Tracker if they are missing (safe to repeat)
  extend-ranges [--rows 100000]  raise the end row of the filters, dropdown lists and Dashboard formulas (older templates stop at
                                 row 1000) so they cover new jobs, and bring an older Dashboard up to date (Applied counts every job
                                 with an Applied Date, so it no longer shrinks when a job moves to FOLLOW-UP; the always-empty
                                 "Excluded Jobs (last 14 days)" row becomes Skipped). Safe to repeat; changes nothing if already up to date.

Setup command (no tracker argument; candidate-only, not in the publisher's copy of this file):
  init --name "<your name>" [--folder <folder>]
                                 creates your own tracker from "Master Job Tracker - TEMPLATE.xlsx" (left unchanged) as
                                 "Master Job Tracker - <Your_Name>.xlsx" in the application-workflow folder (or --folder): the name is
                                 written with underscores instead of spaces and plain letters (José Rao -> Jose_Rao). The EXAMPLE rows
                                 on Master Tracker, JD Store and Excluded Jobs are removed; every sheet, format, dropdown and Dashboard
                                 formula is kept. Saves the new tracker's location in tracker_path.txt (the tracker viewer and Dice
                                 read it). Refuses if that file already exists.

Job feed command (candidate-only):
  import-feed <tracker.xlsx> [--feed "<feed file>"] [--from JOB-#### | --after JOB-####] [--dry-run]
                                 adds the job feed's new jobs to your tracker: the feed's Master Tracker rows and their JD Store rows,
                                 for every feed Job ID above the highest feed Job ID (JOB-1001 and up) already in your tracker. The
                                 first time, give where to start: --after the Job ID you were given ("copy the jobs above JOB-####"), or
                                 --from the first Job ID of a day (that job included). Columns are matched by name (a feed column your
                                 tracker lacks is added at the end), Job URL links are kept, and a Job ID already in your tracker is never
                                 added twice, so it is safe to run again. The feed file is only read. --feed is remembered in
                                 feed_path.txt, so later runs need only the tracker. --dry-run prints what it would add.

Exit code 0 = done, 2 = refused or not found (the message says why).
"""
import argparse, datetime as dt, os, re, sys, tempfile
import openpyxl
from openpyxl.worksheet.cell_range import MultiCellRange

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contact_clean as cc

MASTER, JDS, EXCL = "Master Tracker", "JD Store", "Excluded Jobs"
RANK = {"FOUND": 0, "REVIEW": 0, "SELECTED": 1, "SHORTLIST": 1, "CALL READY": 2, "VENDOR CALLED": 3, "RESUME NEEDED": 4,
        "RESUME CREATED": 5, "EMAIL DRAFTED": 6, "APPLIED": 7, "FOLLOW-UP": 8, "CLOSED": 9}  # SKIPPED is handled separately
SETTABLE = ["Pipeline Status", "Vendor Call Date", "Resume Filename", "Resume Path", "Follow-Up Date"]
EMAIL_COL = "Email Drafted Date"
SENT_COL = "Email Sent Date"  # the day my first email for the job was actually sent (Module 04 checks Gmail; the draft date is EMAIL_COL)
SUBMIT_COL = "Submitted Date"  # the day a recruiter confirmed submitting me to the client for this job
MILESTONE_FIELDS = ("Vendor Call Date", "Email Drafted Date", "Resume Filename", "Resume Path")


class Refused(Exception):
    pass


_OPEN = []  # read-only workbooks keep the file open until closed; close them all when a command ends (a long-running helper would otherwise block later saves)


def track(wb):
    _OPEN.append(wb)
    return wb


def close_all():
    while _OPEN:
        try:
            _OPEN.pop().close()
        except Exception:
            pass


def out(s=""):
    sys.stdout.write(str(s) + "\n")


def norm(h):
    return re.sub(r"\s+", " ", str(h or "")).strip().lower()


def text(v):
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def header_map(row):
    return {norm(h): i for i, h in enumerate(row) if norm(h)}


def need(hm, name, sheet):
    i = hm.get(norm(name))
    if i is None:
        raise Refused(f'The "{sheet}" sheet has no "{name}" column.')
    return i


def open_read(path):
    if not os.path.exists(path):
        raise Refused(f"File not found: {path}")
    return track(openpyxl.load_workbook(path, read_only=True, data_only=True))


def sheet(wb, name):
    if name not in wb.sheetnames:
        raise Refused(f'The workbook has no "{name}" sheet.')
    return wb[name]


def rows_of(ws):
    it = ws.iter_rows(values_only=True)
    try:
        head = next(it)
    except StopIteration:
        raise Refused(f'The "{ws.title}" sheet is empty.')
    return head, it


DATE_START = re.compile(r"^\d{4}-\d{2}-\d{2}")


def dated(note):
    """A Notes line that starts with a date: today's is added unless the note already starts with one (YYYY-MM-DD), so every
    activity the modules record (resume generated, follow-up drafted...) keeps its day for the report."""
    note = note.strip()
    return note if DATE_START.match(note) else f"{dt.date.today().isoformat()}: {note}"


def job_no(v):
    m = re.match(r"^JOB-(\d+)$", text(v), re.I)
    return int(m.group(1)) if m else None


# ---------- read commands ----------
def cmd_get(a):
    wb = open_read(a.tracker)
    head, it = rows_of(sheet(wb, MASTER))
    hm = header_map(head)
    idc = need(hm, "Job ID", MASTER)
    want = a.job.upper()
    row = next((r for r in it if text(r[idc]).upper() == want), None)
    if row is None:
        raise Refused(f"{want} is not in the Master Tracker.")
    out(f"== {want} (Master Tracker row) ==")
    for i, h in enumerate(head):
        if h and i < len(row) and text(row[i]):
            out(f"{h}: {text(row[i])}")
    out(f"== {want} (full JD from the JD Store sheet) ==")
    jd = None
    if JDS in wb.sheetnames:
        jh, jit = rows_of(wb[JDS])
        jm = header_map(jh)
        ji, jt = need(jm, "Job ID", JDS), need(jm, "Full Job Description", JDS)
        for r in jit:
            if text(r[ji]).upper() == want:
                jd = text(r[jt])
                break
    out(jd if jd else "(no JD stored for this Job ID)")


# tracker wording for a missing value, e.g. "Not named in posting", "Not in posting text", "Not publicly verified",
# "Not displayed (client not named)"
_NOT_GIVEN = re.compile(r"^not (shown|displayed|named|disclosed|specified|listed|stated|provided|given|available|"
                        r"in posting|publicly)\b")


def _clean(v):
    t = text(v)
    return "" if norm(t) in ("", "-", "n/a", "na", "none", "unknown", "direct client", "direct") or _NOT_GIVEN.match(norm(t)) else t


def _agency_key(v):
    t = re.sub(r"\([^)]*\)", " ", _clean(v).lower())
    words = [w for w in re.split(r"[^a-z0-9]+", t) if w and w not in ("inc", "llc", "ltd", "corp", "corporation", "co", "company", "lp", "llp", "plc", "the")]
    return " ".join(words)


def _name_key(v):
    return " ".join(w for w in re.split(r"[^a-z0-9]+", re.sub(r"\([^)]*\)", " ", _clean(v).lower())) if w)


GENERIC_INBOX = {"info", "hr", "careers", "career", "jobs", "job", "recruiting", "recruitment", "talent", "us", "contact", "accommodations",
                 "admin", "support", "resumes", "resume", "apply", "hiring", "staffing", "team", "sales", "hello", "office", "mail", "noreply", "no-reply"}


def _email_key(v):
    e = _clean(v).lower()
    if "@" not in e:
        return ""
    local = e.split("@", 1)[0]
    return "" if local in GENERIC_INBOX else e


def cmd_related(a):
    wb = open_read(a.tracker)
    head, it = rows_of(sheet(wb, MASTER))
    hm = header_map(head)
    idc = need(hm, "Job ID", MASTER)
    col = lambda n: hm.get(norm(n))
    cell = lambda r, n: text(r[col(n)]) if col(n) is not None and col(n) < len(r) else ""
    want = a.job.upper()
    rows = [r for r in it if job_no(r[idc]) is not None]
    me = next((r for r in rows if text(r[idc]).upper() == want), None)
    if me is None:
        raise Refused(f"{want} is not in the Master Tracker.")
    my_email, my_name, my_agency = _email_key(cell(me, "Recruiter Email")), _name_key(cell(me, "Recruiter Name")), _agency_key(cell(me, "Staffing Company"))
    one = lambda t, n=140: re.sub(r"\s+", " ", t).strip()[:n]
    out(f"{want}\t{one(cell(me, 'Job Title'))}\t{cell(me, 'Priority')}\tStaffing: {cell(me, 'Staffing Company') or '-'}\tRecruiter: {cell(me, 'Recruiter Name') or '-'} {cell(me, 'Recruiter Email')}".rstrip())
    out(f"Why niche: {one(cell(me, 'Why Niche / Potentially Niche')) or '-'}")
    my_client = _agency_key(cell(me, "Company"))
    same_client = [r for r in rows if text(r[idc]).upper() != want and _day(cell(r, SUBMIT_COL)) and my_client
                   and _agency_key(cell(r, "Company")) == my_client]
    if same_client:
        out(f"SUBMITTED TO THE SAME CLIENT ({cell(me, 'Company')}), any vendor: " + "; ".join(
            f"{text(r[idc])} {one(cell(r, 'Job Title'), 60)} on {cell(r, SUBMIT_COL)} via {cell(r, 'Staffing Company') or '-'}" for r in same_client))
    found = []
    for r in rows:
        jid = text(r[idc]).upper()
        if jid == want:
            continue
        if not (_clean(cell(r, "Resume Filename")) or _clean(cell(r, "Resume Path")) or _day(cell(r, SUBMIT_COL))):
            continue
        e, n, g = _email_key(cell(r, "Recruiter Email")), _name_key(cell(r, "Recruiter Name")), _agency_key(cell(r, "Staffing Company"))
        if (my_email and e == my_email) or (my_name and my_agency and n == my_name and g == my_agency):
            kind = "SAME RECRUITER"
        elif my_agency and g == my_agency:
            kind = "SAME AGENCY"
        else:
            continue
        found.append((0 if kind == "SAME RECRUITER" else 1, -job_no(r[idc]), kind, r))
    found.sort(key=lambda x: (x[0], x[1]))
    if not found:
        out("No earlier resumes or submissions for this recruiter or staffing company.")
        return
    strong = sum(1 for f in found if f[2] == "SAME RECRUITER")
    out(f"{len(found)} earlier resume(s) or submission(s): {strong} same recruiter, {len(found) - strong} same agency (different recruiter)")
    out("Match\tJob ID\tJob Title\tPriority\tStatus\tSubmitted\tRecruiter\tResume Filename\tResume Path\tWhy niche")
    shown = found if not a.limit else found[:a.limit]
    for _, _, kind, r in shown:
        g = lambda n: one(cell(r, n)).replace(chr(9), " ")
        out(chr(9).join([kind, text(r[idc]), g("Job Title"), g("Priority"), g("Pipeline Status"), g(SUBMIT_COL) or "-", g("Recruiter Name") or "-", g("Resume Filename") or "-", g("Resume Path") or "-", one(cell(r, "Why Niche / Potentially Niche"), 120) or "-"]))
    if len(shown) < len(found):
        out(f"(listed the first {len(shown)} of {len(found)}; use --limit 0 for all)")


def _first_date(v):
    m = re.search(r"\d{4}-\d{2}-\d{2}", text(v))
    return m.group(0) if m else ""


def _next_business_day(d):
    d += dt.timedelta(days=1)
    while d.weekday() >= 5:
        d += dt.timedelta(days=1)
    return d


def cmd_followups(a):
    wb = open_read(a.tracker)
    head, it = rows_of(sheet(wb, MASTER))
    hm = header_map(head)
    idc = need(hm, "Job ID", MASTER)

    def cell(r, n):
        i = hm.get(norm(n))
        return text(r[i]) if i is not None and i < len(r) else ""

    rows = [r for r in it if job_no(r[idc]) is not None]
    today = dt.date.today()
    picked, notes_out = [], []
    wanted = [x.strip().upper() for x in (a.jobs or "").split(",") if x.strip()]
    by_id = {}
    for r in rows:
        by_id.setdefault(text(r[idc]).upper(), r)

    def evidence(r):
        notes = cell(r, "Notes")
        return (cell(r, "Pipeline Status").upper() == "EMAIL DRAFTED" or bool(_first_date(cell(r, EMAIL_COL)))
                or bool(re.search(r"\bSituation [1-4]\b", notes)))

    def last_email_date(r):
        d = _first_date(cell(r, EMAIL_COL))
        if d:
            return d
        best = ""
        for p in re.split(r"\n\s*\n", cell(r, "Notes")):
            m = re.match(r"\s*(\d{4}-\d{2}-\d{2}):", p)
            if m and re.search(r"\bSituation [1-4]\b", p) and m.group(1) > best:
                best = m.group(1)
        return best

    def due_date(r):
        fu = _first_date(cell(r, "Follow-Up Date"))
        if fu:
            return fu
        le = last_email_date(r)
        if not le:
            return ""
        return _next_business_day(dt.date.fromisoformat(le)).isoformat()

    if wanted:
        for w in wanted:
            r = by_id.get(w)
            if r is None:
                notes_out.append(f"{w}: not in the Master Tracker")
                continue
            st = cell(r, "Pipeline Status").upper()
            if st in ("SKIPPED", "CLOSED"):
                notes_out.append(f"{w}: left out because it is {st}")
                continue
            picked.append(r)
    elif a.on:
        try:
            dt.date.fromisoformat(a.on)
        except ValueError:
            raise Refused(f"--on must be a date like 2026-10-22 (got {a.on}).")
        for r in rows:
            if cell(r, "Pipeline Status").upper() in ("SKIPPED", "CLOSED") or not evidence(r):
                continue
            if last_email_date(r) == a.on:
                picked.append(r)
    else:
        for r in rows:
            if cell(r, "Pipeline Status").upper() in ("SKIPPED", "CLOSED") or not evidence(r):
                continue
            dd = due_date(r)
            if dd and dd <= today.isoformat():
                picked.append(r)

    def sortkey(r):
        return (cell(r, "Recruiter Email").lower() or "~", job_no(r[idc]))

    picked.sort(key=sortkey)
    total = len(picked)
    shown = picked if not a.limit else picked[:a.limit]
    how = f"--jobs {','.join(wanted)}" if wanted else (f"--on {a.on}" if a.on else "due today or earlier")
    out(f"{total} job(s) for follow-up ({how}; today is {today.isoformat()}, {today.strftime('%A')})")
    for n in notes_out:
        out("NOTE: " + n)
    if shown:
        out("\t".join(["Recruiter Email", "Job ID", "Job Title", "Status", "Recruiter Name", "Last email date", "Email tag", "Follow-Up Date", "Due",
                       "Resume file", "Hourly Rate", "Rate Verified", "C2C Status", "Follow-ups so far", "Gmail thread"]))
    for r in shown:
        notes = cell(r, "Notes")
        nums = [int(x) for x in re.findall(r"Follow-up #(\d+)", notes)]
        tids = re.findall(r"Gmail thread ([A-Za-z0-9]+)", notes)
        ev = cell(r, EMAIL_COL)
        tag = re.search(r"\(([^)]*)\)", ev)
        g = lambda n: re.sub(r"\s+", " ", cell(r, n)).replace("\t", " ")
        out("\t".join([g("Recruiter Email") or "(none)", text(r[idc]), g("Job Title"), g("Pipeline Status"), g("Recruiter Name") or "-",
                       last_email_date(r) or "-", tag.group(1) if tag else "-", _first_date(cell(r, "Follow-Up Date")) or "-", due_date(r) or "-",
                       "yes" if (cell(r, "Resume Filename") or cell(r, "Resume Path")) else "no", g("Hourly Rate") or "-", g("Rate Verified?") or "-",
                       g("C2C Status")[:60] or "-", str(max(nums) if nums else 0), tids[-1] if tids else "-"]))
    if len(shown) < total:
        out(f"(listed the first {len(shown)} of {total}; tell the user and ask which to do, or use --limit 0 for all)")
    if not total and not notes_out:
        out("Nothing is due. Name a job (--jobs) or a day (--on) to follow up anyway.")


def cmd_next_id(a):
    wb = open_read(a.tracker)
    head, it = rows_of(sheet(wb, MASTER))
    idc = need(header_map(head), "Job ID", MASTER)
    top = max([n for n in (job_no(r[idc]) for r in it) if n is not None] or [0])
    out(f"JOB-{top + 1:04d}")


def cmd_drafted_notes(a):
    wb = open_read(a.tracker)
    head, it = rows_of(sheet(wb, MASTER))
    hm = header_map(head)
    idc, stc, ntc = need(hm, "Job ID", MASTER), need(hm, "Pipeline Status", MASTER), need(hm, "Notes", MASTER)
    ec = hm.get(norm(EMAIL_COL))
    skip = (a.exclude or "").upper()
    found = 0
    for r in it:
        jid = text(r[idc])
        if not jid or jid.upper() == skip:
            continue
        paras = [p.strip() for p in re.split(r"\n\s*\n", text(r[ntc]))]
        tags = [p for p in paras if re.search(r"\bVersion [ABC]\b|\bSituation [1-4]\b", p)]
        has_date = ec is not None and ec < len(r) and bool(text(r[ec]))
        has_sit = any(re.search(r"\bSituation [1-4]\b", p) for p in paras)
        if text(r[stc]).upper() != "EMAIL DRAFTED" and not has_date and not has_sit:
            continue
        found += 1
        out(f"{jid}: " + (" | ".join(re.sub(r"\s+", " ", t) for t in tags) if tags else "(no style tag recorded)"))
    if not found:
        out("(no jobs with Pipeline Status EMAIL DRAFTED)")


def stamp(v):
    m = re.match(r"(\d{4}-\d{2}-\d{2})(?:[ T](\d{2}:\d{2}))?", text(v))
    return (m.group(1), m.group(2) or "00:00") if m else None


def cmd_last_run(a):
    wb = open_read(a.tracker)
    best = {}
    for name, col in ((MASTER, "Run Timestamp"), (EXCL, "Date and Time of Run")):
        if name not in wb.sheetnames:
            best[name] = None
            continue
        head, it = rows_of(wb[name])
        c = need(header_map(head), col, name)
        vals = [(stamp(r[c]), text(r[c])) for r in it if stamp(r[c])]
        best[name] = max(vals) if vals else None
    out(f"Master Tracker newest Run Timestamp: {best[MASTER][1] if best[MASTER] else '(none)'}")
    out(f"Excluded Jobs newest Date and Time of Run: {best[EXCL][1] if best[EXCL] else '(none)'}")
    later = max([b for b in best.values() if b], default=None)
    out(f"Later of the two: {later[1] if later else '(none, this looks like a brand-new tracker)'}")


def cmd_summary(a):
    wb = open_read(a.tracker)
    head, it = rows_of(sheet(wb, MASTER))
    hm = header_map(head)
    idc, stc = need(hm, "Job ID", MASTER), need(hm, "Pipeline Status", MASTER)
    nums, by = [], {}
    for r in it:
        n = job_no(r[idc])
        if n is None:
            continue
        nums.append(n)
        s = text(r[stc]) or "(blank)"
        by[s] = by.get(s, 0) + 1
    out(f"Total jobs: {len(nums)}")
    if by.get("SKIPPED"):
        out("Skipped: " + describe_skips(open_read(a.tracker)))
    if nums:
        out(f"First Job ID: JOB-{min(nums):04d}   Last Job ID: JOB-{max(nums):04d}")
        out("By Pipeline Status: " + ", ".join(f"{k} {v}" for k, v in sorted(by.items(), key=lambda kv: -kv[1])))


FIRST_DATE = re.compile(r"(\d{4}-\d{2}-\d{2})")
NOTE_LINE = re.compile(r"^(\d{4}-\d{2}-\d{2})[:,]?\s*(.*)$")


def _day(v):
    """The first YYYY-MM-DD in a cell (a date, '2026-11-20 (resume attached)', '2026-11-20 09:00 MST'), or None."""
    m = FIRST_DATE.search(text(v))
    if not m:
        return None
    try:
        return dt.date.fromisoformat(m.group(1))
    except ValueError:
        return None


def activity(tracker, since, until):
    """The report's numbers for since..until (both included), as a dict."""
    wb = open_read(tracker)
    head, it = rows_of(sheet(wb, MASTER))
    hm = header_map(head)
    idc = need(hm, "Job ID", MASTER)

    def cell(r, name):
        i = hm.get(norm(name))
        return r[i] if i is not None and i < len(r) else None

    inside = lambda d: d is not None and since <= d <= until
    n = {k: 0 for k in ("jobs_received", "vendor_calls", "resumes_new", "resumes_reused", "emails_with_resume", "emails_without_resume",
                        "emails_sent", "followups", "applied", "submitted", "skipped")}
    gaps = []
    for r in it:
        if job_no(r[idc] if idc < len(r) else None) is None:
            continue
        found = _day(cell(r, "Date Found"))
        n["jobs_received"] += inside(found)
        call, applied = _day(cell(r, "Vendor Call Date")), _day(cell(r, "Applied Date"))
        n["vendor_calls"] += inside(call)
        n["applied"] += inside(applied)
        em_raw = text(cell(r, EMAIL_COL))
        em = _day(em_raw)
        if inside(em):
            n["emails_with_resume" if "resume attached" in em_raw.lower() else "emails_without_resume"] += 1
        n["emails_sent"] += inside(_day(cell(r, SENT_COL)))
        n["submitted"] += inside(_day(cell(r, SUBMIT_COL)))
        for line in text(cell(r, "Notes")).splitlines():
            m = NOTE_LINE.match(line.strip())
            if not m:
                continue
            try:
                d = dt.date.fromisoformat(m.group(1))
            except ValueError:
                continue
            if not inside(d):
                continue
            rest = m.group(2).lower()
            if rest.startswith("resume generated"):
                n["resumes_new"] += 1
            elif rest.startswith("resume reused"):
                n["resumes_reused"] += 1
            elif rest.startswith("follow-up #"):
                n["followups"] += 1
            elif rest.startswith("skipped by me") and text(cell(r, "Pipeline Status")).upper() == "SKIPPED":  # not if reopened since
                n["skipped"] += 1
        first = min([d for d in (call, em, applied) if d], default=None)
        if inside(first) and found and first >= found:
            gaps.append((first - found).days)
    gaps.sort()
    n["first_contacts"] = len(gaps)
    n["days_to_first_contact"] = {"same day": sum(g == 0 for g in gaps), "1 day": sum(g == 1 for g in gaps),
                                  "2 or more days": sum(g >= 2 for g in gaps)}
    n["median_days_to_first_contact"] = gaps[len(gaps) // 2] if gaps else None
    return n


MY_NOTE = re.compile(r"^(resume generated|resume reused|follow-up #|situation|no email drafted|.*\bby me\b)", re.I)  # notes the candidate's
# own work writes (the feed's imported Notes lines never count as activity)
ACTIVITY_COLS = ("Vendor Call Date", EMAIL_COL, SENT_COL, SUBMIT_COL, "Applied Date")


def attention_facts(tracker, day=None):
    """For a daily status across candidates (publisher side, read only): {'last_active': date or None, 'drafts_not_sent': n (drafted
    before `day`, no Email Sent Date), 'followups_due': n (Follow-Up Date on or before `day`, job not SKIPPED or CLOSED)}."""
    day = day or dt.date.today()
    wb = open_read(tracker)
    try:
        head, it = rows_of(sheet(wb, MASTER))
        hm = header_map(head)
        idc = need(hm, "Job ID", MASTER)

        def cell(r, name):
            i = hm.get(norm(name))
            return r[i] if i is not None and i < len(r) else None
        last, unsent, due = None, 0, 0
        for r in it:
            if job_no(r[idc] if idc < len(r) else None) is None:
                continue
            dates = [d for d in (_day(cell(r, c)) for c in ACTIVITY_COLS) if d and d <= day]
            for line in text(cell(r, "Notes")).splitlines():
                m = NOTE_LINE.match(line.strip())
                if m and MY_NOTE.match(m.group(2)):
                    try:
                        d = dt.date.fromisoformat(m.group(1))
                    except ValueError:
                        continue
                    if d <= day:
                        dates.append(d)
            if dates:
                last = max([d for d in (last,) if d] + dates)
            status = text(cell(r, "Pipeline Status")).upper()
            if status in ("SKIPPED", "CLOSED"):
                continue
            drafted = _day(cell(r, EMAIL_COL))
            if drafted and drafted < day and not text(cell(r, SENT_COL)):
                unsent += 1
            fu = _day(cell(r, "Follow-Up Date"))
            if fu and fu <= day:
                due += 1
        return {"last_active": last, "drafts_not_sent": unsent, "followups_due": due}
    finally:
        close_all()


THREAD = re.compile(r"Gmail thread\s+([A-Za-z0-9_-]+)")


def cmd_unsent(a):
    """Jobs whose email was drafted but whose send is not recorded yet, grouped by recruiter, for Module 04's send check."""
    wb = open_read(a.tracker)
    head, it = rows_of(sheet(wb, MASTER))
    hm = header_map(head)
    idc = need(hm, "Job ID", MASTER)

    def cell(r, name):
        i = hm.get(norm(name))
        return text(r[i]) if i is not None and i < len(r) else ""

    jobs = []
    for r in it:
        if job_no(r[idc] if idc < len(r) else None) is None:
            continue
        drafted = cell(r, EMAIL_COL)
        if not _day(drafted) or cell(r, SENT_COL) or cell(r, "Pipeline Status").upper() in ("SKIPPED", "CLOSED"):
            continue
        threads = THREAD.findall(cell(r, "Notes"))
        jobs.append((cell(r, "Recruiter Email").lower(), cell(r, "Job ID"), cell(r, "Job Title"), drafted, threads[0] if threads else ""))
    jobs.sort(key=lambda x: (x[0], x[3]))
    recruiters = list(dict.fromkeys(j[0] for j in jobs))
    shown = recruiters if not a.limit else recruiters[:a.limit]
    out(f"{len(jobs)} job(s) with an email drafted and no Email Sent Date, {len(recruiters)} recruiter(s)"
        + (f"; listing the first {len(shown)} recruiters" if len(shown) < len(recruiters) else ""))
    out("Recruiter Email\tJob ID\tJob Title\tEmail Drafted Date\tGmail thread of the first email (blank: match by subject and date)")
    for j in jobs:
        if j[0] in shown:
            out("\t".join(j))


def cmd_report(a):
    today = dt.date.today()
    try:
        until = dt.date.fromisoformat(a.until) if a.until else today
        since = dt.date.fromisoformat(a.since) if a.since else until - dt.timedelta(days=max(1, a.days) - 1)
    except ValueError:
        raise Refused("Dates must look like 2026-11-20.")
    if since > until:
        raise Refused("--since is after --until.")
    n = activity(a.tracker, since, until)
    if a.json:
        import json
        out(json.dumps(dict(n, tracker=os.path.basename(a.tracker), since=since.isoformat(), until=until.isoformat())))
        return
    out(f"Activity {since.isoformat()} to {until.isoformat()} ({os.path.basename(a.tracker)})")
    out(f"Jobs received:     {n['jobs_received']}")
    out(f"Vendor calls:      {n['vendor_calls']}")
    out(f"Resumes:           {n['resumes_new']} new, {n['resumes_reused']} reused")
    out(f"Emails drafted:    {n['emails_with_resume'] + n['emails_without_resume']} ({n['emails_with_resume']} with resume, "
        f"{n['emails_without_resume']} without)")
    out(f"Emails sent:       {n['emails_sent']} (as recorded by Module 04's sent-mail check)")
    out(f"Follow-ups:        {n['followups']}")
    out(f"Applied:           {n['applied']}")
    out(f"Submitted:         {n['submitted']} (a recruiter confirmed submitting you to the client)")
    out(f"Skipped by me:     {n['skipped']}")
    d = n["days_to_first_contact"]
    if n["first_contacts"]:
        out(f"First contact:     {n['first_contacts']} job(s); after Date Found: same day {d['same day']}, 1 day {d['1 day']}, "
            f"2 or more days {d['2 or more days']} (median {n['median_days_to_first_contact']} day(s))")
    else:
        out("First contact:     none in this period")
    out("Emails drafted are drafts; emails sent are the ones Module 04 found in Gmail sent mail. Resumes and follow-ups count dated Notes lines only.")


def cmd_list(a):
    wb = open_read(a.tracker)
    head, it = rows_of(sheet(wb, MASTER))
    hm = header_map(head)
    idc = need(hm, "Job ID", MASTER)
    need(hm, "Pipeline Status", MASTER)

    def cell(r, n):
        i = hm.get(norm(n))
        return text(r[i]).replace("\t", " ").replace("\n", " ") if i is not None and i < len(r) else ""

    wanted = []
    for x in (a.jobs or "").split(","):  # matched by number, as get does (JOB-01047 = JOB-1047); a repeated ID is listed once
        if x.strip() and x.strip().upper() not in wanted:
            wanted.append(x.strip().upper())
    want_no = {job_no(x): x for x in wanted if job_no(x) is not None}
    want_status = {re.sub(r"\s+", " ", s).strip().upper() for s in a.status.split(",") if s.strip()}
    found = {}
    for r in it:
        n = job_no(r[idc])
        if n is None or n in found:
            continue
        if wanted and n not in want_no:
            continue
        if not wanted and re.sub(r"\s+", " ", cell(r, "Pipeline Status")).upper() not in want_status:
            continue
        found[n] = [text(r[idc]), cell(r, "Job Title"), cell(r, "Company"), cell(r, "Pipeline Status"),
                      cell(r, "Resume Filename"), cell(r, "Why Niche / Potentially Niche")[:120]]
    order = [n for n in dict.fromkeys(job_no(x) for x in wanted) if n in found]
    rows = [found[n] for n in order] if wanted else list(found.values())
    out(f"{len(rows)} jobs" + (f" (status {', '.join(sorted(want_status))})" if not wanted else ""))
    missing = [x for x in wanted if job_no(x) not in found]
    if missing:
        out("Not found in the tracker: " + ", ".join(missing))
    shown = rows if not a.limit else rows[:a.limit]
    out("Job ID\tJob Title\tCompany\tPipeline Status\tResume Filename\tWhy Niche")
    for m in shown:
        out("\t".join(m))
    if len(shown) < len(rows):
        out(f"(listed the first {len(shown)} only; {len(rows) - len(shown)} more match; --limit 0 lists all)")


def cmd_check_contacts(a):
    wb = open_read(a.tracker)
    head, it = rows_of(sheet(wb, MASTER))
    hm = header_map(head)
    idc = need(hm, "Job ID", MASTER)
    nc, ec, pc = (need(hm, n, MASTER) for n in ("Recruiter Name", "Recruiter Email", "Recruiter Phone"))
    first = job_no(a.since) if a.since else None
    found, checked = [], 0
    for r in it:
        n = job_no(r[idc])
        if n is None or (first is not None and n < first):
            continue
        checked += 1
        get = lambda i: r[i] if i < len(r) else None
        for col, why, fix in cc.check_row(get(nc), get(ec), get(pc)):
            found.append((text(r[idc]), col, why, text(get({"Recruiter Name": nc, "Recruiter Email": ec, "Recruiter Phone": pc}.get(col, nc)))[:70], fix))
    bad_rows = len({f[0] for f in found})
    out(f"Checked {checked} job(s); {bad_rows} with a problem in Recruiter Name / Email / Phone." + ("" if found else " All clean."))
    if found:
        out("Job ID\tCell\tProblem\tCurrent value\tSuggested value (empty = look by hand)")
        for f in found[: a.limit if a.limit else None]:
            out("\t".join(f))
        out("Fix a cell with: python tracker_tool.py edit <tracker> JOB-#### --set \"Recruiter Email=...\" (the tool refuses text that is not a plain address/name/phone).")


def cmd_recheck_list(a):
    wb = open_read(a.tracker)
    head, it = rows_of(sheet(wb, MASTER))
    hm = header_map(head)
    cols = [need(hm, n, MASTER) for n in ("Job ID", "Source Portal", "Job URL", "Pipeline Status", "Date Found")]
    want = {s.strip().upper() for s in a.status.split(",") if s.strip()}
    cutoff = (dt.date.today() - dt.timedelta(days=a.days)).isoformat()
    matches = []
    for r in it:
        if job_no(r[cols[0]]) is None or text(r[cols[3]]).upper() not in want:
            continue
        d = text(r[cols[4]])[:10]
        if not d or d < cutoff:
            continue
        matches.append([text(r[c]).replace("\t", " ").replace("\n", " ") for c in cols])
    per = {}
    for m in matches:
        k = m[1] or "(no portal)"
        per[k] = per.get(k, 0) + 1
    out(f"{len(matches)} jobs match (" + ", ".join(f"{k} {v}" for k, v in sorted(per.items())) + ")")
    shown = matches if not a.limit else matches[:a.limit]
    out("Job ID\tSource Portal\tJob URL\tPipeline Status\tDate Found")
    for m in shown:
        out("\t".join(m))
    if len(shown) < len(matches):
        out(f"(listed the first {len(shown)} only; {len(matches) - len(shown)} more match. Tell the user the total and the estimated cost, and ask whether to check all of them (--limit 0) or only a number of the newest.)")


# ---------- write commands ----------
def guard_writable(path):
    if not os.path.exists(path):
        raise Refused(f"File not found: {path}")
    lock = os.path.join(os.path.dirname(os.path.abspath(path)), "~$" + os.path.basename(path))
    if not os.path.exists(lock):
        return
    if os.name == "nt":
        # Excel keeps the workbook itself locked while it is open; a lock file left behind by a crash does not.
        try:
            with open(path, "r+b"):
                pass
        except OSError:
            raise Refused("The workbook is open in Excel (" + os.path.basename(lock) + " exists and the file is locked). Save and close it, then run this again. Nothing was changed.")
        return  # stale lock file: the workbook is not open anywhere
    raise Refused("The workbook looks open in Excel (found " + os.path.basename(lock) + "). Save and close it, then run this again. If Excel is not open, delete that ~$ file. Nothing was changed.")


def save_atomic(wb, path):
    d = os.path.dirname(os.path.abspath(path))
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=d, prefix=".trk_")
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    except PermissionError:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise Refused("Could not write the workbook (is it open in Excel?). Nothing was changed.")
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def cmd_set(a):
    guard_writable(a.tracker)
    changes = {}
    for item in a.set or []:
        if "=" not in item:
            raise Refused(f'Use "Field=Value", got: {item}')
        k, v = item.split("=", 1)
        match = next((s for s in SETTABLE if norm(s) == norm(k)), None)
        if not match:
            raise Refused(f'"{k.strip()}" cannot be set with this tool. Allowed: {", ".join(SETTABLE)} (and --note for Notes).')
        changes[match] = v.strip()
    email_flag = getattr(a, "email_drafted", None)
    want_email = email_flag is not None or changes.get("Pipeline Status", "").upper() == "EMAIL DRAFTED"
    firsts = {}  # date columns recorded once: the first send, the first submission
    for col, val, flag in ((SENT_COL, getattr(a, "email_sent", None), "--email-sent"), (SUBMIT_COL, getattr(a, "submitted", None), "--submitted")):
        if val:
            try:
                firsts[col] = dt.date.fromisoformat(val.strip()).isoformat()
            except ValueError:
                raise Refused(f'{flag} needs a date like 2026-11-20, got "{val}".')
    if not changes and not a.note and not want_email and not firsts:
        raise Refused("Nothing to do: give --set, --note, --email-sent or --submitted.")
    wb = openpyxl.load_workbook(a.tracker)  # formula mode on purpose: keeps every formula
    ws = sheet(wb, MASTER)
    if want_email:
        ensure_column(ws, EMAIL_COL, 28)
    for col in firsts:
        ensure_column(ws, col, 18)
    head = [c.value for c in ws[1]]
    hm = header_map(head)
    idc = need(hm, "Job ID", MASTER) + 1
    want = a.job.upper()
    r = next((x for x in range(2, ws.max_row + 1) if text(ws.cell(x, idc).value).upper() == want), None)
    if r is None:
        raise Refused(f"{want} is not in the Master Tracker.")
    report = []
    st_col = need(hm, "Pipeline Status", MASTER) + 1
    cur = text(ws.cell(r, st_col).value).upper()
    if cur in ("SKIPPED", "CLOSED") and not a.confirmed_closed:
        raise Refused(f"{want} is {cur}. Ask the user before working on it, then repeat with --confirmed-closed. Nothing was changed.")
    for f, v in list(changes.items()):
        if f == "Pipeline Status":
            new = v.upper()
            if new not in RANK and new != "SKIPPED":
                raise Refused(f'"{v}" is not a Pipeline Status value.')
            if new == cur:
                del changes[f]
                report.append(f"Pipeline Status already {cur}")
                continue
            if cur == "SKIPPED" and new != "SKIPPED":
                del changes[f]
                report.append(f"Pipeline Status kept at SKIPPED (setting {new} would reopen a dropped job; only the user may do that)")
                continue
            if new != "SKIPPED" and cur in RANK and RANK[new] < RANK[cur]:
                del changes[f]
                report.append(f"Pipeline Status kept at {cur} (setting {new} would move it backward)")
                continue
            changes[f] = new
    notes = []
    if "Resume Path" in changes or changes.get("Resume Filename"):
        notes += resume_path_change(ws, hm, r, want, changes, report)
    if a.note:
        notes.append(dated(a.note))
    first_vals = {}
    for col, val in firsts.items():
        have = text(ws.cell(r, need(hm, col, MASTER) + 1).value)
        if have:
            report.append(f"{col} kept at {have} (the first date is recorded once)")
        else:
            first_vals[col] = val
    if not changes and not notes and not want_email and not first_vals:
        out(f"{want}: no change needed (" + "; ".join(report) + "). File not saved.")
        return
    for f, v in changes.items():
        c = need(hm, f, MASTER) + 1
        ws.cell(r, c).value = v
        if f == "Resume Path":
            ws.cell(r, c).hyperlink = v if re.match(r"^https?://", v, re.I) else None
        report.append(f"{f} = {v}")
    email_val = None
    if want_email:
        email_val = dt.date.today().isoformat() + {"attached": " (resume attached)", "none": " (no resume)"}.get(email_flag, "")
        ws.cell(r, need(hm, EMAIL_COL, MASTER) + 1).value = email_val
        report.append(f"{EMAIL_COL} = {email_val}")
    for col, val in first_vals.items():
        ws.cell(r, need(hm, col, MASTER) + 1).value = val
        report.append(f"{col} = {val}")
    note_before, note_text = None, "\n\n".join(notes)
    if note_text:
        nc = need(hm, "Notes", MASTER) + 1
        old = text(ws.cell(r, nc).value)
        note_before = old
        ws.cell(r, nc).value = (old + "\n\n" if old else "") + note_text
        report.append("Notes: appended " + str(len(note_text)) + " characters")
    save_atomic(wb, a.tracker)
    # read back to confirm
    chk = track(openpyxl.load_workbook(a.tracker, read_only=True, data_only=False))
    cws = chk[MASTER]
    row = next(cws.iter_rows(min_row=r, max_row=r, values_only=True))
    for f, v in changes.items():
        if text(row[need(hm, f, MASTER)]) != v:
            raise Refused(f"Read-back check failed for {f}. Check the file before continuing.")
    if email_val is not None and text(row[need(hm, EMAIL_COL, MASTER)]) != email_val:
        raise Refused(f"Read-back check failed for {EMAIL_COL}. Check the file before continuing.")
    for col, val in first_vals.items():
        if text(row[need(hm, col, MASTER)]) != val:
            raise Refused(f"Read-back check failed for {col}. Check the file before continuing.")
    if note_text:
        got = text(row[need(hm, "Notes", MASTER)])
        if not got.startswith(note_before) or not got.endswith(note_text):
            raise Refused("Read-back check failed for Notes. Check the file before continuing.")
    out(f"{want} updated: " + "; ".join(report))


def clean_resume_path(v, filename):
    """The value Module 03 gives for Resume Path -> the folder to store (Resume Path holds the job's resume folder; the file is
    <Resume Path>/<Resume Filename>). Accepts the whole 'RESUME PATH: ...' line and quotes around it. A local path must be absolute: a
    folder must hold a file named like the Resume Filename, and a full file path must have that name (its folder is stored). A link is
    kept as it is."""
    v = re.sub(r"^\s*RESUME PATH:\s*", "", v, flags=re.I).strip().strip('"').strip("'").strip()
    if re.match(r"^https?://", v, re.I):
        return v
    if not os.path.isabs(v):
        raise Refused(f"Resume Path must be the full path of the resume's folder (got: {v}). Nothing was changed.")
    if os.path.isfile(v):
        if filename and os.path.basename(v).lower() != filename.strip().lower():
            raise Refused(f"Resume Path points to {os.path.basename(v)}, but the Resume Filename is {filename}. Nothing was changed.")
        return os.path.dirname(os.path.normpath(v))
    if not os.path.isdir(v):
        raise Refused(f"No folder at the Resume Path: {v}. Save the resume first, then repeat. Nothing was changed.")
    if filename and not os.path.isfile(os.path.join(v, filename.strip())):
        raise Refused(f"The folder {v} holds no file named {filename}. Nothing was changed.")
    return os.path.normpath(v)


def resume_path_change(ws, hm, r, job, changes, report):
    """Fill Resume Path (the resume's folder) whenever a resume is recorded. A given path is checked; with only a Resume Filename, the
    file is looked up in the job's folder under Documents/Job Application Workflow/Resumes (resume_path.py find). Returns the Notes
    lines to add."""
    filename = changes.get("Resume Filename") or text(ws.cell(r, need(hm, "Resume Filename", MASTER) + 1).value)
    col = hm.get(norm("Resume Path"))
    old = text(ws.cell(r, col + 1).value) if col is not None else ""
    if "Resume Path" in changes:
        changes["Resume Path"] = clean_resume_path(changes["Resume Path"], filename)
    elif col is None:
        report.append("Resume Path not filled: the tracker has no Resume Path column")
        return []
    else:
        import resume_path
        hits = resume_path.find_resumes(job, filename)
        if len(hits) != 1:
            report.append(f"Resume Path not filled: {len(hits) or 'no'} file(s) named {filename} in the {job} folders under "
                          f"{resume_path.resumes_root()}; give it with --set \"Resume Path=<full path>\"")
            return []
        changes["Resume Path"] = os.path.dirname(hits[0])
    if changes["Resume Path"] == old:
        del changes["Resume Path"]
        report.append("Resume Path already set")
        return []
    return [f"{dt.date.today().isoformat()}: Resume Path replaced (was {old})"] if old else []


def cmd_prune_excluded(a):
    guard_writable(a.tracker)
    wb = openpyxl.load_workbook(a.tracker)
    ws = sheet(wb, EXCL)
    head = [c.value for c in ws[1]]
    c = need(header_map(head), "Date and Time of Run", EXCL) + 1
    stamps = {x: stamp(ws.cell(x, c).value) for x in range(2, ws.max_row + 1)}
    known = [s for s in stamps.values() if s]
    if not known:
        out("Excluded Jobs has no dated rows; nothing to prune.")
        return
    newest = max(known)
    cutoff = (dt.date.today() - dt.timedelta(days=a.days)).isoformat()
    drop = [x for x, s in stamps.items() if s and s[0] < cutoff and s != newest]
    if not drop:
        out(f"Pruned 0 Excluded Jobs rows older than {a.days} days.")
        return
    blocks, start, prev = [], None, None
    for x in sorted(drop):
        if start is None:
            start = prev = x
        elif x == prev + 1:
            prev = x
        else:
            blocks.append((start, prev - start + 1))
            start = prev = x
    blocks.append((start, prev - start + 1))
    for first, count in reversed(blocks):  # bottom-up so earlier row numbers stay valid
        ws.delete_rows(first, count)
    save_atomic(wb, a.tracker)
    out(f"Pruned {len(drop)} Excluded Jobs rows older than {a.days} days (kept every row of the latest run, {newest[0]} {newest[1]}).")


SKIP_COL = "Skip Reason"
MANUAL_PREFIX = re.compile(r"^manual:\s*", re.I)


def strip_manual(reason):
    return MANUAL_PREFIX.sub("", reason or "")


def skip_counts(wb):
    """(skipped by me, skipped automatically, [rows of mine]) from the Master Tracker."""
    head, it = rows_of(sheet(wb, MASTER))
    hm = header_map(head)
    idc, stc = need(hm, "Job ID", MASTER), need(hm, "Pipeline Status", MASTER)
    rc = hm.get(norm(SKIP_COL))
    tc, cc, sc = hm.get(norm("Job Title")), hm.get(norm("Company")), hm.get(norm("Staffing Company"))
    mine, auto = [], 0
    for r in it:
        if job_no(r[idc]) is None or text(r[stc]).upper() != "SKIPPED":
            continue
        reason = text(r[rc]) if rc is not None and rc < len(r) else ""
        if reason.lower().startswith("manual"):
            g = lambda i: text(r[i]) if i is not None and i < len(r) else ""
            mine.append((text(r[idc]), g(tc), g(sc) or g(cc), re.sub(r"^manual:\s*", "", reason, flags=re.I)))
        else:
            auto += 1
    return len(mine), auto, mine


def describe_skips(wb):
    m, auto, _ = skip_counts(wb)
    return f"{m + auto} ({m} by me, {auto} automatic)"


def cmd_skipped(a):
    wb = open_read(a.tracker)
    m, auto, mine = skip_counts(wb)
    out(f"Skipped: {m + auto} ({m} by me, {auto} automatic: closed, W2-only and other system drops)")
    shown = mine if a.all or not a.limit else mine[:a.limit]
    for jid, title, comp, reason in shown:
        out(f"{jid}\t{title}\t{comp}\t{reason}")
    if len(shown) < len(mine):
        out(f"(listed {len(shown)} of {len(mine)} skipped by me; use --limit 0 for all)")


TIER_COL = "Tier Reason"


def ensure_skip_column(ws):
    return ensure_column(ws, SKIP_COL, 42)


def ensure_column(ws, name, width):
    """Column number (1-based) of a column; adds it in the first empty header cell if it is missing."""
    hm = header_map([c.value for c in ws[1]])
    if norm(name) in hm:
        return hm[norm(name)] + 1, False
    col = 1
    while ws.cell(1, col).value not in (None, ""):
        col += 1
    cell, left = ws.cell(1, col), ws.cell(1, col - 1)
    cell.value = name
    if not cell.has_style or cell.fill.fill_type is None:  # copy the header look from its left neighbour
        from copy import copy
        cell.font, cell.fill, cell.alignment, cell.border = copy(left.font), copy(left.fill), copy(left.alignment), copy(left.border)
    letter = cell.column_letter
    ws.column_dimensions[letter].width = width
    ref = ws.auto_filter.ref
    m = re.match(r"^([A-Z]{1,3}\d+):([A-Z]{1,3})(\d+)$", str(ref or ""))
    if m:
        from openpyxl.utils import column_index_from_string
        if column_index_from_string(m.group(2)) < col:
            ws.auto_filter.ref = f"{m.group(1)}:{letter}{m.group(3)}"
    return col, True


def find_job_row(ws, hm, want):
    idc = need(hm, "Job ID", MASTER) + 1
    row = next((x for x in range(2, ws.max_row + 1) if text(ws.cell(x, idc).value).upper() == want), None)
    if row is None:
        raise Refused(f"{want} is not in the Master Tracker.")
    return row


def append_note(ws, row, hm, line):
    nc = need(hm, "Notes", MASTER) + 1
    old = text(ws.cell(row, nc).value)
    ws.cell(row, nc).value = (old + "\n\n" if old else "") + line


def cmd_add_columns(a):
    guard_writable(a.tracker)
    wb = openpyxl.load_workbook(a.tracker)
    ws = sheet(wb, MASTER)
    done = []
    for name, width in ((SKIP_COL, 42), (TIER_COL, 60), (EMAIL_COL, 28), (SENT_COL, 18), (SUBMIT_COL, 18)):
        col, added = ensure_column(ws, name, width)
        if added:
            done.append(f'"{name}" (column {openpyxl.utils.get_column_letter(col)})')
    if not done:
        out(f'The "{SKIP_COL}", "{TIER_COL}", "{EMAIL_COL}", "{SENT_COL}" and "{SUBMIT_COL}" columns already exist. Nothing to change.')
        return
    save_atomic(wb, a.tracker)
    out("Added to the Master Tracker: " + ", ".join(done) + ".")


def cmd_skip(a):
    guard_writable(a.tracker)
    reason = (a.reason or "").strip()
    if not reason:
        raise Refused("Give a reason: --reason \"...\" (a few words, for example: rate too low).")
    wb = openpyxl.load_workbook(a.tracker)
    ws = sheet(wb, MASTER)
    col, added = ensure_skip_column(ws)
    hm = header_map([c.value for c in ws[1]])
    want = a.job.upper()
    row = find_job_row(ws, hm, want)
    stc = need(hm, "Pipeline Status", MASTER) + 1
    cur = text(ws.cell(row, stc).value).upper()
    if cur == "SKIPPED":
        old = text(ws.cell(row, col).value) or "(no reason recorded)"
        out(f"{want} is already SKIPPED ({old}). Nothing changed.")
        return
    if cur == "CLOSED":
        raise Refused(f"{want} is CLOSED already, so there is nothing to skip.")
    if cur in ("APPLIED", "FOLLOW-UP") and not a.force:
        raise Refused(f"{want} is {cur}. Skipping an application in progress needs the user's OK; repeat with --force if they confirm.")
    ws.cell(row, stc).value = "SKIPPED"
    ws.cell(row, col).value = "Manual: " + reason
    append_note(ws, row, hm, f"{dt.date.today().isoformat()}: Skipped by me (was {cur or 'blank'}): {reason}")
    save_atomic(wb, a.tracker)
    chk = track(openpyxl.load_workbook(a.tracker, read_only=True))
    got = next(chk[MASTER].iter_rows(min_row=row, max_row=row, values_only=True))
    if text(got[stc - 1]) != "SKIPPED" or text(got[col - 1]) != "Manual: " + reason:
        raise Refused("Read-back check failed. Check the file before continuing.")
    out(f"{want} skipped (was {cur or 'blank'}): {reason}." + (f' Added the "{SKIP_COL}" column.' if added else ""))


def cmd_reopen(a):
    guard_writable(a.tracker)
    wb = openpyxl.load_workbook(a.tracker)
    ws = sheet(wb, MASTER)
    hm = header_map([c.value for c in ws[1]])
    want = a.job.upper()
    row = find_job_row(ws, hm, want)
    stc = need(hm, "Pipeline Status", MASTER) + 1
    if text(ws.cell(row, stc).value).upper() != "SKIPPED":
        raise Refused(f"{want} is not SKIPPED, so there is nothing to reopen.")
    nc = need(hm, "Notes", MASTER) + 1
    was = re.findall(r"Skipped by me \(was ([A-Z][A-Z -]*?)\)", text(ws.cell(row, nc).value))
    new = (a.status or (was[-1] if was and was[-1] in RANK else "REVIEW")).upper()
    if new not in RANK:
        raise Refused(f'"{new}" is not a Pipeline Status value.')
    rc = hm.get(norm(SKIP_COL))
    old_reason = text(ws.cell(row, rc + 1).value) if rc is not None else ""
    ws.cell(row, stc).value = new
    if rc is not None:
        ws.cell(row, rc + 1).value = None
    line = f"{dt.date.today().isoformat()}: Reopened by me (was skipped" + (": " + strip_manual(old_reason) if old_reason else "") + f"); Pipeline Status set to {new}"
    append_note(ws, row, hm, line + (". " + a.note.strip() if a.note else ""))
    save_atomic(wb, a.tracker)
    out(f"{want} reopened: Pipeline Status {new}, Skip Reason cleared.")


USER_FIELDS = ["Pipeline Status", "Vendor Call Date", "Resume Filename", "Resume Path", "Applied Date", "Follow-Up Date",
               "Email Drafted Date", SENT_COL, SUBMIT_COL, "Recruiter Name", "Recruiter Email", "Recruiter Phone"]
DATE_FIELDS = {"Vendor Call Date", "Applied Date", "Follow-Up Date", SENT_COL, SUBMIT_COL}


def cmd_edit(a):
    guard_writable(a.tracker)
    if a.if_unchanged_since is not None and abs(os.path.getmtime(a.tracker) - a.if_unchanged_since) > 1.0:
        raise Refused("The workbook was saved by someone else (a module or Excel) after you loaded it. Reload and try again. Nothing was changed.")
    changes = {}
    for item in a.set or []:
        if "=" not in item:
            raise Refused(f'Use "Field=Value", got: {item}')
        k, v = item.split("=", 1)
        f = next((x for x in USER_FIELDS if norm(x) == norm(k)), None)
        if not f:
            raise Refused(f'"{k.strip()}" cannot be edited here. Allowed: {", ".join(USER_FIELDS)}.')
        v = v.strip()
        if f in DATE_FIELDS and v:
            try:
                dt.date.fromisoformat(v)
            except ValueError:
                raise Refused(f"{f} must be a date like 2026-10-22 (got {v}).")
        if f == EMAIL_COL and v:
            m = re.match(r"^(\d{4}-\d{2}-\d{2})(\s*\([^)]*\))?$", v)
            try:
                if not m:
                    raise ValueError
                dt.date.fromisoformat(m.group(1))
            except ValueError:
                raise Refused(f"{f} must be a date like 2026-10-22 (got {v}).")
        if f == "Recruiter Email" and v:
            good, why = cc.clean_email(v)
            if not good:
                raise Refused(f"Recruiter Email does not look like an email address: {v}")
            if why:  # extra or glued text: refuse instead of guessing, so the wrong text never reaches the tracker
                raise Refused(f"Recruiter Email must be the address alone ({why}). Use: {good}")
            v = good
        if f == "Recruiter Name" and v and not cc.PLACEHOLDER.match(v):
            good, _ = cc.clean_name(v)
            if not good:
                raise Refused(f"Recruiter Name must be a person's name, optionally followed by bracketed notes (got: {v}).")
            v = good
        if f == "Recruiter Phone" and v and not cc.PLACEHOLDER.match(v):
            good, _ = cc.clean_phone(v)
            if not good:
                raise Refused(f"Recruiter Phone must contain a phone number (got: {v}).")
        if f == "Pipeline Status":
            v = v.upper()
            if v not in RANK and v != "SKIPPED":
                raise Refused(f'"{v}" is not a Pipeline Status value.')
        if f == "Resume Path" and os.path.isfile(v):  # Resume Path holds the folder; a pasted file path keeps only its folder
            v = os.path.dirname(os.path.normpath(v))
        changes[f] = v
    if not changes and not (a.note or "").strip():
        raise Refused("Nothing to do: give --set and/or --note.")
    wb = openpyxl.load_workbook(a.tracker)
    ws = sheet(wb, MASTER)
    if EMAIL_COL in changes or changes.get("Pipeline Status") == "EMAIL DRAFTED":
        ensure_column(ws, EMAIL_COL, 28)
    for col in (SENT_COL, SUBMIT_COL):
        if col in changes:
            ensure_column(ws, col, 18)
    hm = header_map([c.value for c in ws[1]])
    want = a.job.upper()
    row = find_job_row(ws, hm, want)
    today = dt.date.today().isoformat()
    notes, report = [], []

    def cell(name):
        return ws.cell(row, need(hm, name, MASTER) + 1)

    cur = text(cell("Pipeline Status").value).upper()
    if "Pipeline Status" in changes:
        new = changes["Pipeline Status"]
        if new == cur:
            del changes["Pipeline Status"]
        elif new == "SKIPPED":
            reason = (a.reason or "").strip()
            if not reason:
                raise Refused("Skipping a job needs a reason (--reason).")
            col, _ = ensure_skip_column(ws)
            hm = header_map([c.value for c in ws[1]])
            ws.cell(row, col).value = "Manual: " + reason
            notes.append(f"{today}: Skipped by me (was {cur or 'blank'}): {reason}")
        elif cur == "SKIPPED":
            rc = hm.get(norm(SKIP_COL))
            old = text(ws.cell(row, rc + 1).value) if rc is not None else ""
            if rc is not None:
                ws.cell(row, rc + 1).value = None
            notes.append(f"{today}: Reopened by me (was skipped" + (": " + strip_manual(old) if old else "") + f"); Pipeline Status set to {new}")
        else:
            notes.append(f"{today}: Pipeline Status set to {new} by me (was {cur or 'blank'})")
        if new == "APPLIED" and "Applied Date" not in changes and not text(cell("Applied Date").value):
            changes["Applied Date"] = today
        if new != cur and new == "VENDOR CALLED" and "Vendor Call Date" not in changes and not text(cell("Vendor Call Date").value):
            changes["Vendor Call Date"] = today
        if new != cur and new == "EMAIL DRAFTED" and EMAIL_COL not in changes and not text(cell(EMAIL_COL).value):
            changes[EMAIL_COL] = today
    if changes.get("Recruiter Email"):
        if changes.get("Recruiter Name"):
            notes.append(f"{today}: recruiter name and email supplied by me ({changes['Recruiter Name']}, {changes['Recruiter Email']}); the name belongs to this email.")
        else:
            notes.append(f"{today}: recruiter email {changes['Recruiter Email']} supplied by me (not found in the posting; the name is not confirmed to belong to it).")
    if changes.get(EMAIL_COL):
        old_email = text(cell(EMAIL_COL).value)
        sfx = re.search(r"\s*\([^)]*\)\s*$", old_email)
        if sfx and "(" not in changes[EMAIL_COL]:
            changes[EMAIL_COL] = changes[EMAIL_COL] + " " + sfx.group(0).strip()
    for f in MILESTONE_FIELDS:
        if f in changes and not changes[f]:
            was = text(cell(f).value)
            if was:
                notes.append(f"{today}: {f} cleared by me (was {was})")
    for f, v in changes.items():
        c = cell(f)
        c.value = v if v else None
        if f == "Resume Path":
            c.hyperlink = v if re.match(r"^https?://", v, re.I) else None
        report.append(f"{f} = {v if v else '(cleared)'}")
    if (a.note or "").strip():
        notes.append(f"{today}: {a.note.strip()}")
    for line in notes:
        nc = cell("Notes")
        old = text(nc.value)
        nc.value = (old + "\n\n" if old else "") + line
    if notes:
        report.append(f"Notes: added {len(notes)} dated line(s)")
    if not report:
        out(f"{want}: no change needed. File not saved.")
        return
    save_atomic(wb, a.tracker)
    chk = track(openpyxl.load_workbook(a.tracker, read_only=True))
    got = next(chk[MASTER].iter_rows(min_row=row, max_row=row, values_only=True))
    for f, v in changes.items():
        if text(got[need(hm, f, MASTER)]) != v:
            raise Refused(f"Read-back check failed for {f}. Check the file before continuing.")
    out(f"{want} updated: " + "; ".join(report))


OLD_APPLIED = re.compile(r'^=COUNTIFS\(.+,"APPLIED"\)$')  # counted only the jobs whose status is still APPLIED
OLD_EXCLUDED_LABEL = "Excluded Jobs (last 14 days)"


def applied_formula(wb, end=100000):
    """Dashboard 'Applied': every job with an Applied Date, plus APPLIED or FOLLOW-UP jobs without one (the rule Dice uses), so it
    never shrinks when a job moves on. Columns are found by header name."""
    from openpyxl.utils import get_column_letter as col
    hm = header_map([c.value for c in wb[MASTER][1]]) if MASTER in wb.sheetnames else {}
    idc, st, ad = (col(hm.get(norm(h), d) + 1) for h, d in (("Job ID", 0), ("Pipeline Status", 16), ("Applied Date", 20)))
    rng = lambda c: f"'{MASTER}'!${c}$2:${c}${end}"
    base = f'{rng(idc)},"JOB-*"'
    return (f'=COUNTIFS({base},{rng(ad)},"<>")+COUNTIFS({base},{rng(ad)},"",{rng(st)},"APPLIED")'
            f'+COUNTIFS({base},{rng(ad)},"",{rng(st)},"FOLLOW-UP")')


def skipped_formula(wb, end=100000):
    from openpyxl.utils import get_column_letter as col
    hm = header_map([c.value for c in wb[MASTER][1]]) if MASTER in wb.sheetnames else {}
    idc, st = (col(hm.get(norm(h), d) + 1) for h, d in (("Job ID", 0), ("Pipeline Status", 16)))
    return f"=COUNTIFS('{MASTER}'!${idc}$2:${idc}${end},\"JOB-*\",'{MASTER}'!${st}$2:${st}${end},\"SKIPPED\")"


def old_dashboard_rows(wb):
    """[(row, what)] of Dashboard rows that still have the old Applied formula or the old Excluded Jobs row."""
    if "Dashboard" not in wb.sheetnames:
        return []
    found = []
    for r, vals in enumerate(wb["Dashboard"].iter_rows(max_col=2, values_only=True), 1):  # works in read-only mode too
        label, f = text(vals[0] if vals else None), (vals[1] if vals and len(vals) > 1 else None)
        if label == "Applied" and isinstance(f, str) and OLD_APPLIED.match(f.replace(" ", "")) and "Applied Date" not in f:
            found.append((r, "applied"))
        elif label == OLD_EXCLUDED_LABEL:
            found.append((r, "excluded"))
    return found


def update_dashboard(wb, end=100000):
    """Replaces the old Applied formula and the old Excluded Jobs row (always 0 in a candidate's tracker) with Skipped. Returns
    what changed."""
    done = []
    for r, what in old_dashboard_rows(wb):
        ws = wb["Dashboard"]
        if what == "applied":
            ws.cell(r, 2).value = applied_formula(wb, end)
            done.append("Dashboard: Applied now counts every job with an Applied Date (and APPLIED or FOLLOW-UP ones without one)")
        else:
            ws.cell(r, 1).value = "Skipped"
            ws.cell(r, 2).value = skipped_formula(wb, end)
            done.append(f'Dashboard: "{OLD_EXCLUDED_LABEL}" (always 0 in your tracker) replaced by Skipped')
    return done


def cmd_extend_ranges(a):
    guard_writable(a.tracker)
    n = a.rows
    wb = openpyxl.load_workbook(a.tracker)
    done = update_dashboard(wb, n)

    def grow(ref):
        # "A1:AC1000" -> "A1:AC100000" (only when the end row is below n)
        m = re.match(r"^([A-Z]{1,3}\d+):([A-Z]{1,3})(\d+)$", str(ref))
        if m and int(m.group(3)) < n:
            return f"{m.group(1)}:{m.group(2)}{n}"
        return None

    for ws in wb:
        if ws.auto_filter.ref:
            new = grow(ws.auto_filter.ref)
            if new:
                done.append(f"{ws.title}: filter {ws.auto_filter.ref} -> {new}")
                ws.auto_filter.ref = new
        if ws.data_validations:
            for dv in ws.data_validations.dataValidation:
                refs, changed = [], False
                for rng in str(dv.sqref).split():
                    new = grow(rng)
                    refs.append(new or rng)
                    changed = changed or bool(new)
                if changed:
                    done.append(f"{ws.title}: dropdown {dv.sqref} -> {' '.join(refs)}")
                    dv.sqref = MultiCellRange(" ".join(refs))
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    nv = re.sub(r"(:\$[A-Z]{1,3}\$)(\d+)", lambda m: m.group(1) + str(n) if 1000 <= int(m.group(2)) < n else m.group(0), v)
                    if nv != v:
                        c.value = nv
                        done.append(f"{ws.title}!{c.coordinate}: formula range extended")
    if not done:
        out(f"Nothing to change: filters, dropdowns and formulas already reach row {n} or beyond, and the Dashboard is up to date.")
        return
    save_atomic(wb, a.tracker)
    chk = openpyxl.load_workbook(a.tracker)
    bad = [ws.title for ws in chk if ws.auto_filter.ref and grow(ws.auto_filter.ref)]
    if bad:
        raise Refused("Read-back check failed for " + ", ".join(bad) + ". Check the file before continuing.")
    kinds = {}
    for d in done:
        if d.startswith("Dashboard:"):
            out(d + ".")
            continue
        k = "formulas" if "formula" in d else "dropdowns" if "dropdown" in d else "filters"
        kinds[k] = kinds.get(k, 0) + 1
    if kinds:
        out(f"Extended to row {n}: " + ", ".join(f"{v} {k}" for k, v in kinds.items()) + ".")


# ---------- setup ----------
APP_FOLDER = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_NAME = "Master Job Tracker - TEMPLATE.xlsx"
TRACKER_TXT = os.path.join(APP_FOLDER, "tracker_path.txt")
FEED_TXT = os.path.join(APP_FOLDER, "feed_path.txt")
FEED_FIRST_ID = 1001  # feed Job IDs start here; a candidate's own jobs are below it
EXAMPLE_SHEETS = (MASTER, JDS, EXCL)


def file_safe_name(name):
    """'José  Rao-Kumar' -> 'Jose_Rao_Kumar': plain letters and digits joined by underscores (no spaces or accents, so the
    path survives the viewer launchers and tracker_path.txt on Windows and Mac)."""
    import unicodedata
    plain = unicodedata.normalize("NFKD", str(name or "")).encode("ascii", "ignore").decode("ascii")
    return "_".join(re.findall(r"[A-Za-z0-9]+", plain))


def example_rows(ws):
    return [r for r in range(2, ws.max_row + 1) if text(ws.cell(r, 1).value).upper() == "EXAMPLE"]


def cmd_init(a):
    part = file_safe_name(a.name)
    if not part:
        raise Refused("Give your name with --name, for example --name \"Asha Rao\" (letters or digits).")
    if part.upper() in ("TEMPLATE", "TESTSAMPLE", "SHARED", "LIVE"):
        raise Refused(f'"{a.name}" is a reserved word in tracker names; use your own name.')
    folder = os.path.abspath(a.folder or APP_FOLDER)
    template = os.path.join(APP_FOLDER, TEMPLATE_NAME)
    if not os.path.isfile(template):
        raise Refused(f"The template is missing: {template}. Get the application-workflow folder again.")
    target = os.path.join(folder, f"Master Job Tracker - {part}.xlsx")
    if os.path.exists(target):
        raise Refused(f"{target} already exists; nothing was changed. To use it, run: python Tools/setup_check.py --tracker \"{target}\" --fix")
    if not os.path.isdir(folder):
        raise Refused(f"Folder not found: {folder}")
    wb = openpyxl.load_workbook(template)
    removed = []
    for name in EXAMPLE_SHEETS:
        if name in wb.sheetnames:
            rows = example_rows(wb[name])
            for r in reversed(rows):
                wb[name].delete_rows(r)
            if rows:
                removed.append(name)
    save_atomic(wb, target)
    chk = openpyxl.load_workbook(target)  # read-back check: same sheets, no EXAMPLE rows left
    left = [n for n in EXAMPLE_SHEETS if n in chk.sheetnames and example_rows(chk[n])]
    if chk.sheetnames != wb.sheetnames or left:
        raise Refused(f"Read-back check failed for {target} ({', '.join(left) or 'sheets differ'}). Delete that file and run init again.")
    txt = TRACKER_TXT
    try:
        before = open(txt, encoding="utf-8-sig").read().strip()
    except OSError:
        before = ""
    with open(txt, "w", encoding="utf-8", newline="\n") as f:
        f.write(target + "\n")
    out(f"Created {target}")
    out(f"EXAMPLE rows removed from: {', '.join(removed) or 'none (the template had none)'}; the template is unchanged.")
    out("Saved its location in tracker_path.txt" + (f" (it was: {before})" if before and before != target else "") + ".")
    out("Next: python Tools/setup_check.py")


def _read_txt(path):
    try:
        return open(path, encoding="utf-8-sig").read().strip().strip('"')
    except OSError:
        return ""


def _last_job_row(ws, idc):
    """Last row (1-based) whose Job ID cell is filled; 1 when there are no data rows (formatted empty rows do not count)."""
    last = 1
    for r in range(2, ws.max_row + 1):
        if text(ws.cell(r, idc).value):
            last = r
    return last


def cmd_import_feed(a):
    feed = a.feed or _read_txt(FEED_TXT)
    if not feed:
        raise Refused("Give your job feed file with --feed \"<feed file>\" (the Master Job Tracker - SHARED.xlsx or Job Feed - <name>.xlsx shared with you). "
                      "It is remembered after the first run.")
    feed = os.path.abspath(feed)
    if not os.path.isfile(feed):
        raise Refused(f"Feed file not found: {feed}. If it was downloaded again it may have a new name; give it with --feed.")
    if os.path.normcase(feed) == os.path.normcase(os.path.abspath(a.tracker)):
        raise Refused("--feed is your own tracker; give the job feed file shared with you.")
    if a.from_id and a.after_id:
        raise Refused("Give --from or --after, not both.")
    guard_writable(a.tracker)

    wb = openpyxl.load_workbook(a.tracker)
    tm, tj = sheet(wb, MASTER), sheet(wb, JDS)
    thm = header_map([c.value for c in tm[1]])
    t_id = need(thm, "Job ID", MASTER) + 1
    mine = {text(tm.cell(r, t_id).value).upper() for r in range(2, tm.max_row + 1)} - {""}
    mine_nos = {n for n in (job_no(x) for x in mine) if n is not None}  # matched by number, as get and list do (JOB-01047 = JOB-1047)
    have_feed = [n for n in mine_nos if n >= FEED_FIRST_ID]
    if a.from_id or a.after_id:
        n = job_no(a.from_id or a.after_id)
        if n is None:
            raise Refused(f"{a.from_id or a.after_id} is not a Job ID like JOB-1234.")
        floor = n if a.from_id else n + 1
    elif have_feed:
        floor = max(have_feed) + 1
    else:
        raise Refused("Your tracker has no feed jobs yet, so say where to start: --after JOB-#### (the starting Job ID you were given: "
                      "the jobs above it are added) or --from JOB-#### (that job and the ones after it).")

    fw = open_read(feed)
    head, it = rows_of(sheet(fw, MASTER))
    fhm = header_map(head)
    f_id = need(fhm, "Job ID", MASTER + " sheet of the feed")
    feed_cols = [h for h in head if text(h)]
    new, skipped, feed_max = [], 0, None
    for r in it:
        n = job_no(r[f_id] if f_id < len(r) else None)
        if n is None:
            continue
        feed_max = max(feed_max or n, n)
        if n < floor:
            continue
        if n in mine_nos:
            skipped += 1
            continue
        new.append((n, {h: (r[fhm[norm(h)]] if fhm[norm(h)] < len(r) else None) for h in feed_cols}))
    new.sort(key=lambda x: x[0])
    want = {f"JOB-{n:04d}" for n, _ in new} | {text(row["Job ID"]).upper() for _, row in new}
    jd = {}
    if JDS in fw.sheetnames:
        jhead, jit = rows_of(fw[JDS])
        jhm = header_map(jhead)
        j_id = jhm.get(norm("Job ID"))
        if j_id is not None:
            jcols = [h for h in jhead if text(h)]
            for r in jit:
                k = text(r[j_id] if j_id < len(r) else None).upper()
                if k in want and k not in jd:
                    jd[k] = {h: (r[jhm[norm(h)]] if jhm[norm(h)] < len(r) else None) for h in jcols}
    close_all()

    if not new:
        out(f"No new jobs: the feed's last job is {('JOB-%04d' % feed_max) if feed_max else 'none'}, and you already have every job "
            f"from JOB-{floor:04d} on." + (f" ({skipped} already in your tracker.)" if skipped else ""))
        _remember_feed(a, feed)
        return
    first, last = f"JOB-{new[0][0]:04d}", f"JOB-{new[-1][0]:04d}"
    missing_jd = [text(row["Job ID"]) for _, row in new if text(row["Job ID"]).upper() not in jd]
    if a.dry_run:
        out(f"Would add {len(new)} job(s), {first} to {last}, with {len(new) - len(missing_jd)} description(s)"
            + (f"; {skipped} already in your tracker" if skipped else "") + ". Nothing was changed.")
        return

    added_cols = []
    for h in feed_cols:
        if norm(h) not in thm:
            col, added = ensure_column(tm, h, 20)
            if added:
                added_cols.append(h)
    thm = header_map([c.value for c in tm[1]])
    from copy import copy
    row = _last_job_row(tm, t_id)
    style = {c: copy(tm.cell(row, c)._style) for c in range(1, tm.max_column + 1)} if row >= 2 else {}
    for _, vals in new:
        row += 1
        for h, v in vals.items():
            c = tm.cell(row, thm[norm(h)] + 1)
            if thm[norm(h)] + 1 in style:
                c._style = copy(style[thm[norm(h)] + 1])
            c.value = v if v not in ("",) else None
            if norm(h) == norm("Job URL") and isinstance(v, str) and re.match(r"^https?://", v.strip(), re.I):
                c.hyperlink = v.strip()
    jhm2 = header_map([c.value for c in tj[1]])
    j_idc = need(jhm2, "Job ID", JDS) + 1
    jrow = _last_job_row(tj, j_idc)
    jd_have = {job_no(tj.cell(x, j_idc).value) for x in range(2, tj.max_row + 1)} - {None}  # descriptions already there are kept
    jstyle = {c: copy(tj.cell(jrow, c)._style) for c in range(1, tj.max_column + 1)} if jrow >= 2 else {}
    jd_added = 0
    for _, vals in new:
        rec = jd.get(text(vals["Job ID"]).upper())
        if not rec or job_no(vals["Job ID"]) in jd_have:
            continue
        jrow += 1
        jd_added += 1
        for h, v in rec.items():
            if norm(h) in jhm2:
                c = tj.cell(jrow, jhm2[norm(h)] + 1)
                if jhm2[norm(h)] + 1 in jstyle:
                    c._style = copy(jstyle[jhm2[norm(h)] + 1])
                c.value = v
    save_atomic(wb, a.tracker)
    chk = openpyxl.load_workbook(a.tracker, read_only=True)  # read-back: every new Job ID is in the saved file
    try:
        ids = {text(r[t_id - 1]).upper() for r in chk[MASTER].iter_rows(min_row=2, values_only=True) if r and len(r) >= t_id}
    finally:
        chk.close()
    lost = [text(v["Job ID"]) for _, v in new if text(v["Job ID"]).upper() not in ids]
    if lost:
        raise Refused(f"Read-back check failed: {len(lost)} job(s) not found after saving ({lost[0]}...). Check the tracker before continuing.")
    out(f"Added {len(new)} job(s) to your tracker, {first} to {last}, with {jd_added} description(s) in the JD Store"
        + (f"; {skipped} already there were left as they were" if skipped else "") + ".")
    if missing_jd:
        out(f"No description in the feed for {len(missing_jd)} job(s): {', '.join(missing_jd[:5])}{' ...' if len(missing_jd) > 5 else ''}.")
    if added_cols:
        out("Columns added at the end of your Master Tracker: " + ", ".join(added_cols) + ".")
    _remember_feed(a, feed)


def _remember_feed(a, feed):
    if a.feed and not a.dry_run and os.path.normcase(_read_txt(FEED_TXT)) != os.path.normcase(feed):
        with open(FEED_TXT, "w", encoding="utf-8", newline="\n") as f:
            f.write(feed + "\n")
        out("Saved the feed's location in feed_path.txt; next time run it without --feed.")


def main(argv):
    p = argparse.ArgumentParser(prog="tracker_tool.py", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)

    def add(name, fn, job=False):
        s = sub.add_parser(name)
        s.add_argument("tracker")
        if job:
            s.add_argument("job")
        s.set_defaults(fn=fn)
        return s

    add("get", cmd_get, True)
    s = add("related", cmd_related, True)
    s.add_argument("--limit", type=int, default=8)
    s = add("followups", cmd_followups)
    s.add_argument("--jobs")
    s.add_argument("--on")
    s.add_argument("--due", action="store_true")
    s.add_argument("--limit", type=int, default=10)
    add("next-id", cmd_next_id)
    add("last-run", cmd_last_run)
    add("summary", cmd_summary)
    s = add("unsent", cmd_unsent)
    s.add_argument("--limit", type=int, default=10)
    s = add("report", cmd_report)
    s.add_argument("--since")
    s.add_argument("--until")
    s.add_argument("--days", type=int, default=7)
    s.add_argument("--json", action="store_true")
    s = add("drafted-notes", cmd_drafted_notes)
    s.add_argument("--exclude")
    s = add("list", cmd_list)
    s.add_argument("--status", default="RESUME NEEDED")
    s.add_argument("--jobs")
    s.add_argument("--limit", type=int, default=0)
    s = add("recheck-list", cmd_recheck_list)
    s.add_argument("--days", type=int, default=14)
    s.add_argument("--status", default="REVIEW,SHORTLIST,SELECTED,FOUND")
    s.add_argument("--limit", type=int, default=200)
    s = add("check-contacts", cmd_check_contacts)
    s.add_argument("--since")
    s.add_argument("--limit", type=int, default=60)
    s = add("set", cmd_set, True)
    s.add_argument("--email-drafted", choices=["attached", "none"], default=None)
    s.add_argument("--email-sent", dest="email_sent")
    s.add_argument("--submitted")
    s.add_argument("--set", action="append")
    s.add_argument("--note")
    s.add_argument("--confirmed-closed", action="store_true")
    s = add("prune-excluded", cmd_prune_excluded)
    s.add_argument("--days", type=int, default=14)
    s = add("edit", cmd_edit, True)
    s.add_argument("--set", action="append")
    s.add_argument("--note")
    s.add_argument("--reason")
    s.add_argument("--if-unchanged-since", type=float, dest="if_unchanged_since")
    s = add("skip", cmd_skip, True)
    s.add_argument("--reason")
    s.add_argument("--force", action="store_true")
    s = add("reopen", cmd_reopen, True)
    s.add_argument("--status")
    s.add_argument("--note")
    add("add-columns", cmd_add_columns)
    s = add("skipped", cmd_skipped)
    s.add_argument("--all", action="store_true")
    s.add_argument("--limit", type=int, default=50)
    s = add("extend-ranges", cmd_extend_ranges)
    s.add_argument("--rows", type=int, default=100000)
    s = add("import-feed", cmd_import_feed)
    s.add_argument("--feed")
    s.add_argument("--from", dest="from_id")
    s.add_argument("--after", dest="after_id")
    s.add_argument("--dry-run", action="store_true", dest="dry_run")
    s = sub.add_parser("init")
    s.add_argument("--name", required=True)
    s.add_argument("--folder")
    s.set_defaults(fn=cmd_init)
    a = p.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    try:
        a.fn(a)
    except Refused as e:
        out("REFUSED: " + str(e))
        return 2
    finally:
        close_all()
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
