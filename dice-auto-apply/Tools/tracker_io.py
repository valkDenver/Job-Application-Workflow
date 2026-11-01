# NOTE: guard_writable(), save_atomic(), header_map()/need(), the Notes append rule and the RANK list are copies from
# application-workflow/Tools/tracker_tool.py, so the two flows stay independent (design section 5). If the tracker's columns, status
# values, Excel-lock handling or Notes format change there, make the same change here (tests/test_contract.py checks the template).
"""This flow's only reader and writer of the Master Job Tracker (design section 11).

Reading: read_jobs() returns every Master Tracker row as a dict (read-only load, nothing printed).
Writing: outcomes are queued in pending_tracker.jsonl the moment they happen; flush() writes the whole queue in ONE load and save,
restarts if someone else saved the file in between, never moves Pipeline Status backwards, never overwrites a status the user set,
reads every written cell back, and removes from the queue only what was written.
"""
import datetime as dt
import os
import shutil
import tempfile
import time
import uuid
from pathlib import Path

from dice_apply_common import Refused, append_jsonl, norm, read_jsonl, text, today, write_jsonl

MASTER = "Master Tracker"
RANK = {"FOUND": 0, "REVIEW": 0, "SELECTED": 1, "SHORTLIST": 1, "CALL READY": 2, "VENDOR CALLED": 3, "RESUME NEEDED": 4,
        "RESUME CREATED": 5, "EMAIL DRAFTED": 6, "APPLIED": 7, "FOLLOW-UP": 8, "CLOSED": 9}  # SKIPPED is handled separately
USER_FINAL = {"APPLIED", "FOLLOW-UP", "CLOSED", "SKIPPED"}
COLUMNS = ["Job ID", "Priority", "Source Portal", "Run Timestamp", "Job Title", "Company", "Staffing Company", "Job URL",
           "Recruiter Name", "Recruiter Email", "Pipeline Status", "Resume Filename", "Resume Path", "Applied Date", "Notes",
           "Date Found", "Email Drafted Date"]
REQUIRED = ["Job ID", "Source Portal", "Job URL", "Pipeline Status", "Resume Filename", "Resume Path", "Applied Date", "Notes"]
QUEUE = "pending_tracker.jsonl"
IS_WIN = os.name == "nt"


def header_map(row):
    return {norm(h): i for i, h in enumerate(row) if norm(h)}


def need(hm, name):
    i = hm.get(norm(name))
    if i is None:
        raise Refused(f'The "{MASTER}" sheet has no "{name}" column.')
    return i


# ---------- reading ----------
def read_jobs(path):
    """All Master Tracker rows that have a Job ID, as dicts keyed by the COLUMNS names ('' when a column is missing)."""
    import openpyxl
    if not os.path.exists(path):
        raise Refused(f"Tracker not found: {path}")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if MASTER not in wb.sheetnames:
            raise Refused(f'The workbook has no "{MASTER}" sheet.')
        it = wb[MASTER].iter_rows(values_only=True)
        hm = header_map(next(it, ()))
        for c in REQUIRED:
            need(hm, c)
        idx = {c: hm.get(norm(c)) for c in COLUMNS}
        jobs = []
        for n, r in enumerate(it, start=2):
            g = lambda c: text(r[idx[c]]) if idx[c] is not None and idx[c] < len(r) else ""
            if not g("Job ID").upper().startswith("JOB-"):
                continue
            j = {c: g(c) for c in COLUMNS}
            j["Job ID"] = j["Job ID"].upper()
            j["_row"] = n
            jobs.append(j)
        return jobs
    finally:
        wb.close()


# ---------- the queue ----------
def queue_path(rt):
    return Path(rt) / QUEUE


def queue_add(rt, job_id, kind, note, applied_date=None, resume_path=None, resume_filename=None, guid=None):
    """kind: 'applied' (status -> APPLIED), or 'note' (a dated Notes line only)."""
    item = {"id": uuid.uuid4().hex, "job": job_id, "kind": kind, "note": note, "applied_date": applied_date,
            "resume_path": resume_path, "resume_filename": resume_filename, "guid": guid, "queued": today()}
    append_jsonl(queue_path(rt), item)
    return item


def queue_items(rt):
    return read_jsonl(queue_path(rt))


def queued_applied_jobs(rt):
    return {i["job"] for i in queue_items(rt) if i.get("kind") == "applied"}


def _queue_remove(rt, ids):
    keep = [i for i in read_jsonl(queue_path(rt)) if i.get("id") not in ids]  # re-read: a round may have appended meanwhile
    write_jsonl(queue_path(rt), keep)


# ---------- writing ----------
def guard_writable(path):
    if not os.path.exists(path):
        raise Refused(f"Tracker not found: {path}")
    lock = os.path.join(os.path.dirname(os.path.abspath(path)), "~$" + os.path.basename(path))
    if not os.path.exists(lock):
        return
    if IS_WIN:
        # Excel keeps the workbook itself locked while it is open; a lock file left behind by a crash does not.
        try:
            with open(path, "r+b"):
                pass
        except OSError:
            raise Refused("The tracker is open in Excel. Close it and the updates will be written.")
        return
    raise Refused("The tracker looks open in Excel (found " + os.path.basename(lock) + "). Close it and the updates will be written.")


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
        raise Refused("Could not write the tracker (is it open in Excel?).")
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def backup_once_a_day(path, rt):
    mark = Path(rt) / f".backup_{today()}"
    if mark.exists():
        return None
    src = Path(path)
    dest_dir = src.parent / "Backups"
    dest_dir.mkdir(exist_ok=True)
    dest = dest_dir / f"{src.stem} - {dt.datetime.now().strftime('%Y-%m-%d_%H%M%S')}.xlsx"
    shutil.copy2(src, dest)
    mark.parent.mkdir(parents=True, exist_ok=True)
    mark.write_text(str(dest), encoding="utf-8")
    return dest


def _apply(ws, hm, row, item, report):
    """Apply one queued item to one row. Returns {column: expected value} for the read-back check."""
    def cell(name):
        return ws.cell(row, need(hm, name) + 1)

    def set_if_empty(name, v):
        if v and norm(name) in hm and not text(cell(name).value):
            cell(name).value = v
            written[name] = v

    written, notes, jid = {}, [], item["job"]
    cur = text(cell("Pipeline Status").value).upper()
    if item["kind"] == "applied":
        if cur in USER_FINAL:
            if cur != "APPLIED":
                report.append(f"{jid} TRACKER_CHANGED: status is {cur}, kept")
        else:
            cell("Pipeline Status").value = "APPLIED"
            written["Pipeline Status"] = "APPLIED"
        set_if_empty("Applied Date", item.get("applied_date") or item.get("queued") or today())  # the submit day, not the write day
        set_if_empty("Resume Path", item.get("resume_path"))
        set_if_empty("Resume Filename", item.get("resume_filename"))
    if item.get("note"):
        line = f"{item.get('queued') or today()}: {item['note']}"
        if item["kind"] == "applied" and "Pipeline Status" in written:
            line += f" Pipeline Status set to APPLIED by dice-auto-apply (was {cur or 'blank'})."
        notes.append(line)
    for line in notes:
        nc = cell("Notes")
        old = text(nc.value)
        if line in old:
            continue  # the same line is never written twice (a flush retried after a crash)
        nc.value = (old + "\n\n" if old else "") + line
    return written


def flush(path, rt, tries=3):
    """Write every queued item in one load and save. Returns (written_count, report_lines). Raises Refused when it cannot write."""
    import openpyxl
    items = queue_items(rt)
    if not items:
        return 0, []
    last_err = None
    for _ in range(tries):
        guard_writable(path)
        before = os.path.getmtime(path)
        wb = openpyxl.load_workbook(path)  # formula mode on purpose: keeps every formula
        if MASTER not in wb.sheetnames:
            raise Refused(f'The workbook has no "{MASTER}" sheet.')
        ws = wb[MASTER]
        hm = header_map([c.value for c in ws[1]])
        idc = need(hm, "Job ID") + 1
        rows = {text(ws.cell(r, idc).value).upper(): r for r in range(2, ws.max_row + 1)}
        report, checks, done, dropped = [], [], [], []
        for it in items:
            row = rows.get(it["job"].upper())
            if row is None:
                report.append(f"{it['job']} is no longer in the tracker; its update was dropped")
                dropped.append(it["id"])
                continue
            checks.append((row, _apply(ws, hm, row, it, report)))
            done.append(it["id"])
        if abs(os.path.getmtime(path) - before) > 0.001:
            last_err = "the tracker was saved by someone else while updating; trying again"
            time.sleep(1)
            continue  # someone saved in between: throw this load away and start again
        backup_once_a_day(path, rt)
        save_atomic(wb, path)
        chk = openpyxl.load_workbook(path, read_only=True, data_only=False)
        try:
            cache = {}
            for r in chk[MASTER].iter_rows(min_row=1, values_only=True):
                cache[len(cache) + 1] = r
            for row, written in checks:
                got = cache.get(row, ())
                for name, v in written.items():
                    i = need(hm, name)
                    if text(got[i] if i < len(got) else None) != text(v):
                        raise Refused(f"Read-back check failed for row {row} ({name}). The queue was kept; check the file.")
        finally:
            chk.close()
        _queue_remove(rt, set(done) | set(dropped))
        return len(done), report
    raise Refused(last_err or "Could not write the tracker.")


def try_flush(path, rt):
    """Quiet flush used before plan/status/report: returns a short message, never raises."""
    if not queue_items(rt):
        return ""
    try:
        n, _ = flush(path, rt)
        return f"Wrote {n} waiting tracker update(s)."
    except Refused as e:
        return f"{len(queue_items(rt))} tracker update(s) waiting: {e}"
    except Exception as e:  # a damaged or unreadable workbook must not break status/report
        return f"{len(queue_items(rt))} tracker update(s) waiting: {type(e).__name__}"
