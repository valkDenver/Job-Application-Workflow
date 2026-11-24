"""Checks a Candidate Profile before it is used, and is the one place the tools read the profile's fields from.

Usage (from the application-workflow folder; python3 on a Mac):
  python Tools/profile_check.py ["<Candidate Profile.txt>"] [--level resume|call]

  --level resume (default): ready for the first resume or email (Modules 03 and 04):
      Name, Location, Phone, Email and Total experience are filled in; no "[...]" template text is left;
      every employer block in use has its name, job title, Dates ("Mon YYYY - Mon YYYY" or "Mon YYYY - Present",
      start not after end, nothing in the future, most recent employer first), Domain, Domain guidance and Seniority notes;
      unused employer blocks are left completely empty (they are ignored); Resume template is blank or a known template.
  --level call: everything above, plus the Vendor Call Facts Module 02 needs before a vendor call (the same rules as its
      FACT CHECK): each blank fact (N/A counts as answered), "Facts last confirmed on" missing or older than 14 days, and a
      date in the facts that has passed or is within 14 days.

Without a file it uses the one "Candidate Profile - <name>.txt" in the application-workflow folder.
Prints one FIX line per problem (call level also BLANK and STALE lines), then a summary. Never changes the file.
Exit code 0 = ready, 1 = something to fix or fill, 2 = no profile found.
"""
import argparse, datetime as dt, json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
EMPLOYER_FIELDS = ("Employer / client name", "Job title", "Dates", "Domain", "Domain guidance", "Seniority notes")
PERSONAL_REQUIRED = ("Name", "Location", "Phone", "Email", "Total experience")
CALL_OPTIONAL = {"Total US work experience", "Current compensation", "Anything else"}  # blank is a valid answer for these
STALE_DAYS = 14
FULL_MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"]
LABEL = re.compile(r"^([A-Za-z][^:\[\]]{0,80}?):[ \t]*(.*?)[ \t]*$")  # "Label: value" on ONE line; a blank value stays blank


def is_placeholder(v):
    return v.startswith("[")


def parse(txt):
    """{'fields': {label: value} for the lines outside employer blocks (first one wins), 'employers': [{label: value}, ...] for
    the employer blocks in use (all-empty blocks are dropped), 'call': [(section, label, value)] for the Vendor Call Facts part,
    'placeholders': [label] left in the resume part}. A blank field stays blank (it never takes the next field's line), but a value
    may continue on the lines right below its label (the template invites "one per line" for Domain guidance and Where each job was
    worked): those lines, up to a blank line, a heading or the next label, are joined to it with "; "."""
    fields, employers, call, placeholders = {}, [], [], []
    block, in_call, section = None, False, ""
    target = None  # where the last label's value went, so the lines below it can continue it

    def add_more(more):
        kind, where, key = target
        old = where[key][2] if kind == "call" else where[key]
        new = more if not old or is_placeholder(old) else old + "; " + more
        if kind == "call":
            where[key] = (where[key][0], where[key][1], new)
        else:
            where[key] = new
    for raw in txt.splitlines():
        line = raw.strip()
        if not line:
            target = None
            continue
        if line.startswith("VENDOR CALL FACTS"):
            in_call, block, section, target = True, None, "VENDOR CALL FACTS", None
            continue
        if re.match(r"^EMPLOYER \d+\b", line) and not in_call:
            block, target = {}, None
            employers.append(block)
            continue
        if in_call and (re.match(r"^\d+\.\s+[A-Z]", line) or line.startswith("REFERENCES FOR")):
            section, target = line.split("(")[0].strip(), None
            continue
        if line.isupper() or line.startswith("CANDIDATE PROFILE"):  # PERSONAL DETAILS, RESUME LAYOUT, OPTIONAL...
            block, target = None, None
            continue
        m = LABEL.match(line) if not line.startswith("-") else None
        if not m:
            if target is not None:  # a continuation line ("Healthcare", "- Sony: US remote, 2025 to now")
                add_more(line.lstrip("-* ").strip())
            continue
        label, value = m.group(1).strip(), m.group(2).strip()
        if in_call:
            call.append((section, label, value))
            target = ("call", call, len(call) - 1)
        elif block is not None and label in EMPLOYER_FIELDS:
            block[label] = value
            target = ("dict", block, label)
        else:
            if label in fields:  # only the first occurrence counts; do not continue a later one
                target = None
                continue
            fields[label] = value
            target = ("dict", fields, label)
    placeholders = [k for k, v in fields.items() if v and is_placeholder(v)]  # after continuations, which can replace them
    used = []
    for b in employers:
        vals = [b.get(f, "") for f in EMPLOYER_FIELDS]
        if any(v and not is_placeholder(v) for v in vals[:3]):  # a block counts as used once its name, title or dates are filled
            used.append(b)
    return {"fields": fields, "employers": used, "call": call, "placeholders": placeholders}


def read(path):
    with open(path, encoding="utf-8-sig", errors="ignore") as f:
        return parse(f.read())


def month(s):
    """'Jan 2025' or 'January 2025' -> (2025, 1); None when not in that form."""
    m = re.match(r"^([A-Za-z]{3,9})\.?\s+(\d{4})$", s.strip())
    if not m:
        return None
    word = m.group(1).lower()
    full = [n for n in FULL_MONTHS if n.startswith(word)]  # Jan, Janu..., January, Sept
    if not full:
        return None
    return int(m.group(2)), FULL_MONTHS.index(full[0]) + 1


def date_range(s):
    """(start, end or None for Present) or None."""
    parts = re.split(r"\s*[-–—]\s*|\s+to\s+", s.strip(), flags=re.I)  # "Jan 2020 - Mar 2023", "Jan 2020-Mar 2023", "Jan 2020 to Present"
    if len(parts) != 2:
        return None
    start = month(parts[0])
    end = None if parts[1].strip().lower() == "present" else month(parts[1])
    if start is None or (end is None and parts[1].strip().lower() != "present"):
        return None
    return start, end


def known_templates():
    try:
        with open(os.path.join(HERE, "resume_templates.json"), encoding="utf-8") as f:
            return [k for k in json.load(f) if not k.startswith("_")]
    except (OSError, ValueError):
        return []


def check_resume(p, today=None):
    """[problem] for the resume level."""
    today = today or dt.date.today()
    now = (today.year, today.month)
    out = []
    f = p["fields"]
    for label in PERSONAL_REQUIRED:
        v = f.get(label, "")
        if not v or is_placeholder(v):
            out.append(f"{label}: fill it in (PERSONAL DETAILS)")
    for label in p["placeholders"]:
        if label not in PERSONAL_REQUIRED:
            out.append(f"{label}: still has the template's [...] text; fill it in or leave it empty")
    if not p["employers"]:
        out.append("no employer filled in; add at least your most recent employer (most recent first)")
    starts = []
    for i, e in enumerate(p["employers"], 1):
        who = e.get("Employer / client name") or f"employer {i}"
        for label in EMPLOYER_FIELDS:
            v = e.get(label, "")
            if not v or is_placeholder(v):
                out.append(f"{who}: {label} is missing")
        d = e.get("Dates", "")
        if d and not is_placeholder(d):
            r = date_range(d)
            if not r:
                out.append(f'{who}: Dates "{d}" must look like "Jan 2020 - Mar 2023" or "Jan 2025 - Present" (month and year at both ends)')
            else:
                start, end = r
                if start > now or (end and end > now):
                    out.append(f'{who}: Dates "{d}" are in the future')
                elif end and start > end:
                    out.append(f'{who}: Dates "{d}" start after they end')
                starts.append((who, start))
    for (a, sa), (b, sb) in zip(starts, starts[1:]):
        if sb > sa:
            out.append(f"employers are not most recent first: {b} started after {a}; put the most recent employer first")
    t = f.get("Resume template", "")
    names = known_templates()
    if t and not is_placeholder(t) and names and t.strip().lower() not in names:
        out.append(f'Resume template "{t}" is not one of: {", ".join(names)} (or leave it blank for classic)')
    return out


def check_call(p, today=None):
    """([blank fact], [stale fact]) for the call level, the same rules as Module 02's FACT CHECK."""
    today = today or dt.date.today()
    blank, stale = [], []
    confirmed = None
    for section, label, value in p["call"]:
        if label == "Facts last confirmed on":
            m = re.search(r"\d{4}-\d{2}-\d{2}", value)
            confirmed = dt.date.fromisoformat(m.group(0)) if m else None
            continue
        if label in CALL_OPTIONAL:
            continue
        if not value or is_placeholder(value):
            blank.append(f"{label} ({section})" if section else label)
            continue
        for m in re.finditer(r"\d{4}-\d{2}-\d{2}", value):
            try:
                d = dt.date.fromisoformat(m.group(0))
            except ValueError:
                continue
            if label == "First US entry date":
                continue
            if d < today:
                stale.append(f"{label}: {m.group(0)} has passed; update it")
            elif (d - today).days <= STALE_DAYS:
                stale.append(f"{label}: {m.group(0)} is {(d - today).days} day(s) away")
    if confirmed is None:
        stale.insert(0, "Facts last confirmed on: missing; review the Vendor Call Facts and set today's date")
    elif (today - confirmed).days > STALE_DAYS:
        stale.insert(0, f"Facts last confirmed on: {confirmed.isoformat()} is {(today - confirmed).days} days old; review the facts and set today's date")
    return blank, stale


def find_profile():
    sys.path.insert(0, HERE)
    import setup_check
    return setup_check.find_profile(None)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("profile", nargs="?")
    ap.add_argument("--level", choices=("resume", "call"), default="resume")
    a = ap.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass
    path = a.profile
    if not path:
        path, problem = find_profile()
        if not path:
            print("No profile: " + problem)
            return 2
    if not os.path.isfile(path):
        print("No profile: file not found: " + path)
        return 2
    p = read(path)
    print("Profile: " + os.path.basename(path))
    fixes = check_resume(p)
    for x in fixes:
        print("FIX    " + x)
    blank = stale = []
    if a.level == "call":
        blank, stale = check_call(p)
        for x in stale:
            print("STALE  " + x)
        for x in blank:
            print("BLANK  " + x + "  (answered with \"let me confirm\" on a call)")
    if not fixes:
        print("Ready for resumes and emails." if a.level == "resume" else
              ("Ready for vendor calls." if not blank and not stale else
               "Ready for resumes and emails; vendor calls: %d blank, %d stale." % (len(blank), len(stale))))
    else:
        print("%d thing(s) to fix before your first resume or email." % len(fixes))
    return 0 if not fixes and not blank and not stale else 1


if __name__ == "__main__":
    sys.exit(main())
