"""Which jobs a round takes, in which order, and why the others are skipped (design section 6).

Selection is driven by the resume: a job is eligible when Resume Path or Resume Filename is filled and resolves to a file. Pipeline
Status is used only to leave jobs out. The same-recruiter rule (8) is simulated in queue order, so `plan` shows what the round will do.
"""
import datetime as dt
import re

from dice_apply_common import norm, text
import dice_apply_selectors as S
import resume_find as RF
from tracker_io import USER_FINAL

# ---------- recruiter keys (copies of tracker_tool.py related, see the NOTE in tracker_io.py) ----------
_NOT_GIVEN = re.compile(r"^not (shown|displayed|named|disclosed|specified|listed|stated|provided|given|available|in posting|publicly)\b")
GENERIC_INBOX = {"info", "hr", "careers", "career", "jobs", "job", "recruiting", "recruitment", "talent", "us", "contact", "accommodations",
                 "admin", "support", "resumes", "resume", "apply", "hiring", "staffing", "team", "sales", "hello", "office", "mail", "noreply", "no-reply"}


def _clean(v):
    t = text(v)
    return "" if norm(t) in ("", "-", "n/a", "na", "none", "unknown", "direct client", "direct") or _NOT_GIVEN.match(norm(t)) else t


def agency_key(v):
    t = re.sub(r"\([^)]*\)", " ", _clean(v).lower())
    words = [w for w in re.split(r"[^a-z0-9]+", t) if w and w not in ("inc", "llc", "ltd", "corp", "corporation", "co", "company", "lp", "llp", "plc", "the")]
    return " ".join(words)


def name_key(v):
    return " ".join(w for w in re.split(r"[^a-z0-9]+", re.sub(r"\([^)]*\)", " ", _clean(v).lower())) if w)


def email_key(v):
    e = _clean(v).lower()
    if "@" not in e:
        return ""
    return "" if e.split("@", 1)[0] in GENERIC_INBOX else e


def same_recruiter(a, b):
    ea, eb = email_key(a["Recruiter Email"]), email_key(b["Recruiter Email"])
    if ea and ea == eb:
        return True
    na, nb = name_key(a["Recruiter Name"]), name_key(b["Recruiter Name"])
    ga, gb = agency_key(a["Staffing Company"]), agency_key(b["Staffing Company"])
    return bool(na and ga and na == nb and ga == gb)


def has_recruiter(j):
    return bool(email_key(j["Recruiter Email"]) or (name_key(j["Recruiter Name"]) and agency_key(j["Staffing Company"])))


# ---------- small parsers ----------
def guid_of(url):
    m = S.JOB_URL.match(text(url))
    return m.group(1).lower() if m else ""


def date_part(v):
    m = re.match(r"^(\d{4}-\d{2}-\d{2})", text(v))
    return m.group(1) if m else ""


def title_key(j):
    return re.sub(r"[^a-z0-9]+", " ", (j["Job Title"] + " | " + j["Company"]).lower()).strip()


def emailed_resume(j):
    return "resume attached" in j["Email Drafted Date"].lower()


def applied_evidence(j, ledger_guids, queued):
    """Rules 7/8: has this job been applied to by any route?"""
    return (j["Pipeline Status"].upper() in ("APPLIED", "FOLLOW-UP") or bool(j["Applied Date"]) or emailed_resume(j)
            or guid_of(j["Job URL"]) in ledger_guids or j["Job ID"] in queued)


# ---------- filters ----------
def _day(word, today):
    w = word.strip().lower()
    if w == "today":
        return today
    if w == "yesterday":
        return today - dt.timedelta(days=1)
    return dt.date.fromisoformat(w)


def apply_filters(jobs, f, today):
    """f: dict of the filter flags (design section 6). Returns (kept jobs, notes)."""
    notes = []
    kept = jobs
    if f.get("jobs"):
        want = [x.strip().upper() for x in f["jobs"].split(",") if x.strip()]
        kept = [j for j in kept if j["Job ID"] in want]
        missing = [w for w in want if w not in {j["Job ID"] for j in kept}]
        if missing:
            notes.append("not in the tracker: " + ", ".join(missing))
    if f.get("status"):
        st = {s.strip().upper() for s in f["status"].split(",") if s.strip()}
        kept = [j for j in kept if j["Pipeline Status"].upper() in st]
    if f.get("run_date"):
        d = _day(f["run_date"], today).isoformat()
        no_ts = sum(1 for j in kept if not date_part(j["Run Timestamp"]))
        kept = [j for j in kept if date_part(j["Run Timestamp"]) == d]
        if no_ts:
            notes.append(f"{no_ts} job(s) have no Run Timestamp and cannot match a run date")
    if f.get("run_latest"):
        stamps = [j["Run Timestamp"] for j in jobs if date_part(j["Run Timestamp"])]
        latest = max(stamps, key=lambda s: s[:16]) if stamps else ""
        kept = [j for j in kept if j["Run Timestamp"] == latest]
        notes.append(f"latest run: {latest or 'none found'}")
    if f.get("found"):
        w = f["found"].strip().lower()
        if w.startswith("last:"):
            n = int(re.sub(r"\D", "", w) or "1")
            since = (today - dt.timedelta(days=n - 1)).isoformat()
            kept = [j for j in kept if date_part(j["Date Found"]) >= since]
        else:
            d = _day(w, today).isoformat()
            kept = [j for j in kept if date_part(j["Date Found"]) == d]
    if f.get("priority"):
        tiers = {re.sub(r"\D", "", p) for p in f["priority"].split(",") if re.sub(r"\D", "", p)}
        kept = [j for j in kept if re.sub(r"\D", "", j["Priority"]) in tiers]
    if f.get("title_contains"):
        words = [w.lower() for w in f["title_contains"]]
        kept = [j for j in kept if any(w in j["Job Title"].lower() for w in words)]
    return kept, notes


# ---------- order ----------
def order(jobs):
    """Oldest Date Found first, then spread so the same company is not back to back where that can be avoided."""
    base = sorted(jobs, key=lambda j: (date_part(j["Date Found"]) or "9999", int(re.sub(r"\D", "", j["Job ID"]) or 0)))
    out, rest = [], list(base)
    while rest:
        last = agency_key(out[-1]["Company"]) if out else None
        pick = next((i for i, j in enumerate(rest[:8]) if agency_key(j["Company"]) != last), 0)
        out.append(rest.pop(pick))
    return out


# ---------- the plan ----------
class Plan:
    def __init__(self):
        self.picked = []        # (job, Resolution, note)
        self.skipped = []       # (job, code, detail)
        self.left = []          # eligible but over the limit
        self.notes = []
        self.limit_note = ""
        self.counts = {}

    def as_dict(self):
        def j(job):
            return {"job": job["Job ID"], "title": job["Job Title"], "company": job["Company"]}
        return {
            "picked": [dict(j(x), resume=str(r.file), via=r.via, note=n) for x, r, n in self.picked],
            "skipped": [dict(j(x), code=c, detail=d) for x, c, d in self.skipped],
            "left_for_next_round": len(self.left), "notes": self.notes, "limit_note": self.limit_note, "counts": self.counts,
        }


def build_plan(all_jobs, filters, cfg, ledger_guids, queued, index, today=None, limit=None, allow=(), allow_duplicate=False,
               remaining_today=None, override_resume=None):
    """Design section 6 rules 1-8 on the jobs that pass the filters; the same-recruiter rule runs in queue order."""
    today = today or dt.date.today()
    p = Plan()
    dice = [j for j in all_jobs if text(j["Source Portal"]).lower() == "dice"]
    p.counts["dice_rows"] = len(dice)
    cands, notes = apply_filters(dice, filters, today)
    p.notes += notes
    if filters.get("jobs"):
        non_dice = {x.strip().upper() for x in filters["jobs"].split(",")} - {j["Job ID"] for j in dice}
        for jid in sorted(non_dice & {j["Job ID"] for j in all_jobs}):
            p.skipped.append(({"Job ID": jid, "Job Title": "", "Company": ""}, "NOT_DICE", "Source Portal is not Dice"))
    max_bytes = int(float(cfg.get("max_resume_mb") or 2) * 1048576)
    applied = [j for j in all_jobs if applied_evidence(j, ledger_guids, queued)]
    applied_titles = {title_key(j): j["Job ID"] for j in applied}
    allow = {a.upper() for a in allow}

    eligible = []
    for j in cands:
        st = j["Pipeline Status"].upper()
        if not guid_of(j["Job URL"]):
            p.skipped.append((j, "NOT_DICE", "Job URL is not a dice.com/job-detail link"))
        elif not (j["Resume Path"] or j["Resume Filename"]) and not override_resume:
            p.skipped.append((j, "NO_RESUME", "Resume Path and Resume Filename are both empty"))
        elif st in USER_FINAL:
            p.skipped.append((j, "STATUS_" + st.replace(" ", "_"), f"Pipeline Status is {st}"))
        elif j["Applied Date"]:
            p.skipped.append((j, "ALREADY_APPLIED", f"Applied Date {j['Applied Date']}"))
        elif emailed_resume(j):
            p.skipped.append((j, "EMAILED_RESUME", "the resume was emailed to the recruiter (Email Drafted Date: resume attached)"))
        elif guid_of(j["Job URL"]) in ledger_guids or j["Job ID"] in queued:
            p.skipped.append((j, "IN_LEDGER", "already submitted by dice-auto-apply"))
        elif not allow_duplicate and title_key(j) in applied_titles and applied_titles[title_key(j)] != j["Job ID"]:
            p.skipped.append((j, "DUPLICATE_OF " + applied_titles[title_key(j)], "same title and company already applied"))
        else:
            r = RF.resolve(j["Job ID"], j["Resume Path"], j["Resume Filename"], index, override_resume, max_bytes)
            if not r.ok:
                p.skipped.append((j, r.code, r.reason))
            else:
                eligible.append((j, r))

    # rule 8 in queue order; jobs picked earlier in this round count as applied
    cap = [n for n in (limit, cfg.get("max_per_round"), remaining_today) if n is not None]
    cap_n = min(cap) if cap else None
    if cap_n is not None and limit is not None and cap_n < limit:
        why = "max_per_round" if cfg.get("max_per_round") == cap_n else "what is left of today's limit"
        p.limit_note = f"asked for {limit}, trimmed to {cap_n} ({why})"
    by_job = {j["Job ID"]: j for j in all_jobs}
    tech_cache = {}

    def tech(job, res=None):
        if job["Job ID"] not in tech_cache:
            r = res or RF.resolve(job["Job ID"], job["Resume Path"], job["Resume Filename"], index, None, 10 ** 12)
            tech_cache[job["Job ID"]] = RF.tech_of(r.file, index.root) if r.file else (None, None)
        return tech_cache[job["Job ID"]]

    earlier = [j for j in applied if has_recruiter(j)]
    no_recruiter = 0
    round_titles = {}
    ordered = order([e[0] for e in eligible])
    res_of = {j["Job ID"]: r for j, r in eligible}
    for j in ordered:
        r = res_of[j["Job ID"]]
        if cap_n is not None and len(p.picked) >= cap_n:
            p.left.append(j)
            continue
        note = ""
        if not allow_duplicate and title_key(j) in round_titles:  # rule 7 inside the round: the same posting listed twice
            p.skipped.append((j, "DUPLICATE_OF " + round_titles[title_key(j)], "same title and company earlier in this round"))
            continue
        if not has_recruiter(j):
            no_recruiter += 1
        elif j["Job ID"] not in allow:
            same = [e for e in earlier if e["Job ID"] != j["Job ID"] and same_recruiter(e, j)]
            verdict = None
            for e in same:
                t1, t2 = tech(j, r), tech(e)
                if t1[0] is None or t2[0] is None:
                    verdict = ("SAME_RECRUITER " + e["Job ID"], f"role unclear: {j['Job Title']} vs {e['Job Title']}")
                    break
                similar = norm(t1[0]) == norm(t2[0]) and (not t1[1] or not t2[1] or norm(t1[1]) == norm(t2[1]))
                if not similar:
                    verdict = ("SAME_RECRUITER " + e["Job ID"], f"different role: {'/'.join(x for x in t1 if x)} vs {'/'.join(x for x in t2 if x)}")
                    break
                note = f"same recruiter as {e['Job ID']}, similar role ({t1[0]})"
            if verdict:
                p.skipped.append((j, verdict[0], verdict[1]))
                continue
        p.picked.append((j, r, note))
        earlier.append(j)
        round_titles.setdefault(title_key(j), j["Job ID"])
    if no_recruiter:
        p.notes.append(f"{no_recruiter} job(s) have no recruiter name or email, so the same-recruiter check could not run for them")
    p.counts.update(picked=len(p.picked), skipped=len(p.skipped), left=len(p.left), matched=len(cands))
    return p


def estimate_minutes(n, cfg, pace):
    pc = (cfg.get("paces") or {}).get(pace) or {}
    per_job = (cfg.get("work_sec_per_job") or 45) + (sum((cfg.get("job_page_dwell_sec") or [6, 20])) / 2)
    gap = pc.get("gap_median", 90) * 1.15  # log-normal mean is above the median
    every = sum(pc.get("break_every", [20, 35])) / 2
    brk = sum(pc.get("break_min", [5, 15])) / 2 * 60
    total = n * per_job + max(0, n - 1) * gap + int(max(0, n - 1) / every) * brk
    return int(round(total / 60))
