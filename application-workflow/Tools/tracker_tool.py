# NOTE: application-workflow/Tools/tracker_tool.py and publisher/Tools/tracker_tool.py are separate copies so the two flows stay independent. If you change the cleaning rules or shared helpers in one, make the same change in the other.
"""Small helper for the Master Job Tracker so a module reads or changes one job without loading the whole workbook
into the conversation. Reads print only what is asked for; writes touch only the cells named.

Usage:  python tracker_tool.py <command> <tracker.xlsx> [arguments]

Read commands (never change the file):
  get <JOB-####>                 every non-empty field of that job's Master Tracker row, then its full JD from the JD Store sheet
  related <JOB-####> [--limit 8] earlier resumes (a Resume Filename or Resume Path is filled) of the same recruiter (same email, or same
                                 name at the same staffing company) and of the same staffing company with a different recruiter.
                                 Module 03 uses this for its soft warning before it creates a new resume
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
  skipped [--all] [--limit 50]   how many jobs are skipped (by me / automatic), then the ones skipped by me with their reasons
  check-contacts [--since JOB-####] [--limit 60]
                                 finds recruiter cells that hold description text, glued-on words (name@co.comPhone), notes inside the email cell,
                                 a phone that is not a number, or a name that does not match the address. --since <Job ID> checks only that job and the
                                 ones after it (for example the rows you just copied in). Prints the cell, the problem and the suggested value; changes nothing
  recheck-list [--days 14] [--status REVIEW,SHORTLIST,SELECTED,FOUND] [--limit 200]
                                 first prints how many jobs match (per portal), then up to --limit rows (Job ID, portal, URL,
                                 status, Date Found); lists tracked jobs worth re-checking; --limit 0 lists all

Write commands (refuse to run while the workbook is open in Excel):
  set <JOB-####> [--set "Field=Value" ...] [--note "text"] [--confirmed-closed] [--email-drafted attached|none]
                                 only these fields can be set: Pipeline Status, Vendor Call Date, Resume Filename, Resume Path,
                                 Follow-Up Date. Resume Path must be the full path of an existing file named like the Resume Filename
                                 (or a link); when only Resume Filename is given, the file is looked up in that job's folder under
                                 Documents/Job Application Workflow/Resumes and Resume Path is filled with it. A replaced path is noted.
                                 --note appends a paragraph to Notes (existing Notes are never replaced).
                                 --email-drafted records the draft: Email Drafted Date = today plus "(resume attached)" or "(no resume)".
                                 Module 04 passes it with Pipeline Status=EMAIL DRAFTED; the date is written even when the status
                                 itself cannot move (already further along). The column is created at the end if it is missing.
                                 Pipeline Status only moves forward: a backward move is skipped and reported, the other fields
                                 are still written. A job whose status is SKIPPED or CLOSED is not changed unless --confirmed-closed.
  prune-excluded [--days 14]     delete Excluded Jobs rows older than N days, always keeping every row of the most recent run
  edit <JOB-####> [--set "Field=Value" ...] [--note "text"] [--reason "text"] [--if-unchanged-since EPOCH]
                                 what the user does by hand (used by the tracker viewer's edit panel). Fields: Pipeline Status (any
                                 value; SKIPPED needs --reason), Vendor Call Date, Resume Filename, Resume Path, Applied Date,
                                 Follow-Up Date, Email Drafted Date, Recruiter Name, Recruiter Email, Recruiter Phone. An empty value clears a field.
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
  add-columns                    add the Skip Reason, Tier Reason and Email Drafted Date columns to the Master Tracker if they are missing (safe to repeat)
  extend-ranges [--rows 100000]  raise the end row of the filters, dropdown lists and Dashboard formulas (older templates stop at
                                 row 1000) so they cover new jobs. Safe to repeat; changes nothing if already large enough.

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
    found = []
    for r in rows:
        jid = text(r[idc]).upper()
        if jid == want:
            continue
        if not (_clean(cell(r, "Resume Filename")) or _clean(cell(r, "Resume Path"))):
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
        out("No earlier resumes for this recruiter or staffing company.")
        return
    strong = sum(1 for f in found if f[2] == "SAME RECRUITER")
    out(f"{len(found)} earlier resume(s): {strong} same recruiter, {len(found) - strong} same agency (different recruiter)")
    out("Match\tJob ID\tJob Title\tPriority\tStatus\tRecruiter\tResume Filename\tResume Path\tWhy niche")
    shown = found if not a.limit else found[:a.limit]
    for _, _, kind, r in shown:
        g = lambda n: one(cell(r, n)).replace(chr(9), " ")
        out(chr(9).join([kind, text(r[idc]), g("Job Title"), g("Priority"), g("Pipeline Status"), g("Recruiter Name") or "-", g("Resume Filename") or "-", g("Resume Path") or "-", one(cell(r, "Why Niche / Potentially Niche"), 120) or "-"]))
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
    if not changes and not a.note and not want_email:
        raise Refused("Nothing to do: give --set and/or --note.")
    wb = openpyxl.load_workbook(a.tracker)  # formula mode on purpose: keeps every formula
    ws = sheet(wb, MASTER)
    if want_email:
        ensure_column(ws, EMAIL_COL, 28)
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
        notes.append(a.note.strip())
    if not changes and not notes and not want_email:
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
    if note_text:
        got = text(row[need(hm, "Notes", MASTER)])
        if not got.startswith(note_before) or not got.endswith(note_text):
            raise Refused("Read-back check failed for Notes. Check the file before continuing.")
    out(f"{want} updated: " + "; ".join(report))


def clean_resume_path(v, filename):
    """The value Module 03 gives for Resume Path -> the path to store. Accepts the whole 'RESUME PATH: ...' line and quotes around it.
    A local path must be absolute and point to an existing file whose name is the job's Resume Filename; a link is kept as it is."""
    v = re.sub(r"^\s*RESUME PATH:\s*", "", v, flags=re.I).strip().strip('"').strip("'").strip()
    if re.match(r"^https?://", v, re.I):
        return v
    if not os.path.isabs(v):
        raise Refused(f"Resume Path must be the full path of the saved file (got: {v}). Nothing was changed.")
    if not os.path.isfile(v):
        raise Refused(f"No file at the Resume Path: {v}. Save the resume first, then repeat. Nothing was changed.")
    if filename and os.path.basename(v).lower() != filename.strip().lower():
        raise Refused(f"Resume Path points to {os.path.basename(v)}, but the Resume Filename is {filename}. Nothing was changed.")
    return os.path.normpath(v)


def resume_path_change(ws, hm, r, job, changes, report):
    """Fill Resume Path whenever a resume is recorded. A given path is checked; with only a Resume Filename, the file is looked up in the
    job's folder under Documents/Job Application Workflow/Resumes (resume_path.py find). Returns the Notes lines to add."""
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
        changes["Resume Path"] = hits[0]
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
    for name, width in ((SKIP_COL, 42), (TIER_COL, 60), (EMAIL_COL, 28)):
        col, added = ensure_column(ws, name, width)
        if added:
            done.append(f'"{name}" (column {openpyxl.utils.get_column_letter(col)})')
    if not done:
        out(f'The "{SKIP_COL}", "{TIER_COL}" and "{EMAIL_COL}" columns already exist. Nothing to change.')
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
               "Email Drafted Date", "Recruiter Name", "Recruiter Email", "Recruiter Phone"]
DATE_FIELDS = {"Vendor Call Date", "Applied Date", "Follow-Up Date"}


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
        changes[f] = v
    if not changes and not (a.note or "").strip():
        raise Refused("Nothing to do: give --set and/or --note.")
    wb = openpyxl.load_workbook(a.tracker)
    ws = sheet(wb, MASTER)
    if EMAIL_COL in changes or changes.get("Pipeline Status") == "EMAIL DRAFTED":
        ensure_column(ws, EMAIL_COL, 28)
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


def cmd_extend_ranges(a):
    guard_writable(a.tracker)
    n = a.rows
    wb = openpyxl.load_workbook(a.tracker)
    done = []

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
        out(f"Nothing to change: filters, dropdowns and formulas already reach row {n} or beyond.")
        return
    save_atomic(wb, a.tracker)
    chk = openpyxl.load_workbook(a.tracker)
    bad = [ws.title for ws in chk if ws.auto_filter.ref and grow(ws.auto_filter.ref)]
    if bad:
        raise Refused("Read-back check failed for " + ", ".join(bad) + ". Check the file before continuing.")
    kinds = {}
    for d in done:
        k = "formulas" if "formula" in d else "dropdowns" if "dropdown" in d else "filters"
        kinds[k] = kinds.get(k, 0) + 1
    out(f"Extended to row {n}: " + ", ".join(f"{v} {k}" for k, v in kinds.items()) + ".")


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
    s = add("drafted-notes", cmd_drafted_notes)
    s.add_argument("--exclude")
    s = add("recheck-list", cmd_recheck_list)
    s.add_argument("--days", type=int, default=14)
    s.add_argument("--status", default="REVIEW,SHORTLIST,SELECTED,FOUND")
    s.add_argument("--limit", type=int, default=200)
    s = add("check-contacts", cmd_check_contacts)
    s.add_argument("--since")
    s.add_argument("--limit", type=int, default=60)
    s = add("set", cmd_set, True)
    s.add_argument("--email-drafted", choices=["attached", "none"], default=None)
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
