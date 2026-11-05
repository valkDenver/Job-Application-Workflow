"""Builds the save location for a generated resume, the same way on Windows and Mac, so Module 03 never has to
guess folder names. It never reads or prints anything except what is asked for.

Usage:  python resume_path.py <command> [arguments]        (on a Mac: python3 resume_path.py ...)

Layout:  Documents/Job Application Workflow/Resumes/<Technology>/<Specialty, optional>/<Job folder>/<Candidate>_<Title>.docx

Commands:
  categories                     the technology list (with keyword hints and specialties) the model must choose from
  sync [--dry-run]               adds to the technology list every folder already in Resumes/ that is missing from it: each top folder
                                 as a technology, each folder inside it that is not a job folder (JOB-... or NO-ID_...) as its specialty.
                                 Only adds; never removes, renames or moves anything. A folder missing from the list is never searched
                                 by the library check, so run this once after folders were made by hand or the list was reset.
                                 --dry-run prints what would be added without saving
  path --tech T [--specialty S] --name "Full Name" --title "Position Title" [--job JOB-#### --tracker FILE] [--new] [--dry-run]
                                 creates the folders and prints the folder, file name and full path.
                                 --tech must be an existing technology (close spellings are matched, for example "Dev Ops" -> DevOps);
                                 an unknown one is refused unless --new is given, which creates it and adds it to the list.
                                 --specialty is added to the technology's list the same way (it is created when new).
                                 --job with --tracker names the job folder from the tracker's Company and Staffing Company columns;
                                 without --job the folder is NO-ID_<date>.
                                 If that exact file already exists in the job folder, nothing is overwritten: the command lists the
                                 files already there and exits with code 3 so you can choose a different title/name (for example one that
                                 reflects the new tool set) and run it again.
  list --tech T [--specialty S] [--new] [--limit 10]
                                 the resumes already saved in that technology folder (with --specialty: only that specialty's folder;
                                 without it: the whole technology folder, specialties included), newest first, at most --limit of them
                                 (the total is always shown). Creates the folder when missing. Prints path, size and date for each .docx.
                                 An unknown technology is refused (with the closest names) unless --new is given, as for path.
                                 Module 03 uses it for the library check.
  reuse --source FILE --tech T [--specialty S] [--job JOB-#### --tracker FILE] [--new]
                                 copies an existing resume (the library check found it good enough) into this job's folder, built the
                                 same way as `path`, keeping the same file name. The original is never moved or changed, and nothing is
                                 overwritten (exit code 3 if that file name is already in the job folder).

  find --job JOB-#### [--filename F]
                                 locates the saved resume of that job: looks through Resumes/<Technology>/<Specialty>/ for job folders that start with the
                                 Job ID and lists the .docx files in them (only the one named F when --filename is given). Module 04 uses it to find the file
                                 to attach. Exit code 2 when nothing is found.
  evidence get --employer E --tool T
  evidence set --employer E --tool T --outcome used|competitor|inconclusive [--note "text"]
  evidence list                  the saved results of Module 03's employer tool evidence check (did this employer really use this enterprise
                                 platform?), kept in employer_tools_cache.json so the same web search is not repeated for later jobs.
                                 get prints the saved outcome or "none"; delete the file to clear everything.

The technology list is kept in resume_categories.json next to this script and grows as new technologies are added.
Exit code 0 = done, 2 = refused or not found, 3 = file already exists.
"""
import argparse, datetime as dt, difflib, json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(HERE, "resume_categories.json")
EVIDENCE = os.path.join(HERE, "employer_tools_cache.json")
ROOT_NAME = "Job Application Workflow"
ILLEGAL = r'\/:*?"<>|'
SUFFIXES = {"inc", "llc", "corp", "corporation", "ltd", "limited", "llp", "co", "company", "lp", "plc"}

SEED = {
    "technologies": {
        "DevOps": {"hints": ["Ansible", "Terraform", "Jenkins", "Kubernetes", "Docker", "CI/CD", "Puppet", "Chef"], "specialties": ["Automation", "CI-CD", "Cloud-Infra"]},
        "AWS": {"hints": ["AWS", "Lambda", "S3", "EC2", "DynamoDB"], "specialties": []},
        "Azure": {"hints": ["Azure", "Azure DevOps", "Functions"], "specialties": []},
        "GCP": {"hints": ["GCP", "BigQuery", "Cloud Run", "Pub/Sub"], "specialties": []},
        "Java": {"hints": ["Java", "Spring Boot", "Hibernate", "Microservices"], "specialties": ["Backend", "Frontend", "Full Stack"]},
        ".NET": {"hints": [".NET", "C#", "ASP.NET"], "specialties": []},
        "Python": {"hints": ["Python", "Django", "Flask", "FastAPI"], "specialties": []},
        "Scala": {"hints": ["Scala", "Spark", "Akka"], "specialties": []},
        "Angular": {"hints": ["Angular", "RxJS", "TypeScript"], "specialties": []},
        "React": {"hints": ["React", "Redux", "Next.js"], "specialties": []},
        "Node": {"hints": ["Node.js", "Express", "NestJS"], "specialties": []},
        "Android": {"hints": ["Android", "Kotlin"], "specialties": []},
        "React Native": {"hints": ["React Native", "mobile"], "specialties": []},
        "API": {"hints": ["REST API", "GraphQL", "API Gateway", "Integration"], "specialties": []},
        "Data Engineering": {"hints": ["ETL", "Airflow", "Snowflake", "Databricks", "Kafka"], "specialties": []},
        "AI-ML": {"hints": ["Machine Learning", "LLM", "Bedrock", "MLOps"], "specialties": []},
        "IAM": {"hints": ["IAM", "Okta", "SailPoint", "Active Directory", "SSO"], "specialties": []},
        "IDR": {"hints": ["Incident Detection", "Security Monitoring", "SIEM"], "specialties": []},
        "ServiceNow": {"hints": ["ServiceNow", "ITSM"], "specialties": []},
        "Workday": {"hints": ["Workday"], "specialties": ["Finance", "HCM", "Integration", "Functional"]},
        "Salesforce": {"hints": ["Salesforce", "Apex", "Lightning"], "specialties": []},
        "QA-Test Automation": {"hints": ["Selenium", "Cypress", "Test Automation", "QA"], "specialties": []},
        "Architect": {"hints": ["Solution Architect", "Enterprise Architect", "Technical Architect"], "specialties": []},
        "Unqork": {"hints": ["Unqork", "low-code"], "specialties": []},
        "Alfresco": {"hints": ["Alfresco", "ECM"], "specialties": []},
    }
}


class Refused(Exception):
    pass


def out(s=""):
    sys.stdout.write(str(s) + "\n")


def key(s):
    return re.sub(r"[^a-z0-9#+.]", "", str(s or "").lower())


# ---------- locations ----------
def documents_dir():
    """The real Documents folder (Windows may redirect it to OneDrive; Mac is ~/Documents)."""
    if sys.platform.startswith("win"):
        try:
            import ctypes
            from ctypes import wintypes
            guid = ctypes.c_byte * 16
            # FOLDERID_Documents {FDD39AD0-238F-46AF-ADB4-6C85480369C7}
            fid = (ctypes.c_ubyte * 16)(0xD0, 0x9A, 0xD3, 0xFD, 0x8F, 0x23, 0xAF, 0x46, 0xAD, 0xB4, 0x6C, 0x85, 0x48, 0x03, 0x69, 0xC7)
            buf = ctypes.c_wchar_p()
            if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(fid), 0, None, ctypes.byref(buf)) == 0 and buf.value:
                p = buf.value
                ctypes.windll.ole32.CoTaskMemFree(buf)
                return p
        except Exception:
            pass
    return os.path.join(os.path.expanduser("~"), "Documents")


def resumes_root():
    return os.path.join(documents_dir(), ROOT_NAME, "Resumes")


# ---------- category list ----------
def load_config():
    if not os.path.exists(CONFIG):
        save_config(SEED)
        return json.loads(json.dumps(SEED))
    with open(CONFIG, encoding="utf-8") as f:
        data = json.load(f)
    data.setdefault("technologies", {})
    return data


def save_config(data):
    tmp = CONFIG + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(tmp, CONFIG)


def match_name(name, options):
    """Exact (ignoring case/spaces/punctuation) or a very close spelling; None when nothing fits."""
    k = key(name)
    by_key = {key(o): o for o in options}
    if k in by_key:
        return by_key[k]
    close = difflib.get_close_matches(k, list(by_key), n=1, cutoff=0.86)
    return by_key[close[0]] if close else None


def clean_part(s, limit=None, spaces=True):
    s = "".join(" " if c in ILLEGAL else c for c in str(s or "") if c.isprintable())
    s = re.sub(r"\s+", " ", s).strip(" .")
    if limit and len(s) > limit:
        cut = s[:limit].rsplit(" ", 1)[0]
        s = (cut if len(cut) >= limit // 2 else s[:limit]).strip()
    return s if spaces else s.replace(" ", "_")


PLACEHOLDERS = {"", "-", "n/a", "na", "none", "unknown", "not displayed", "not disclosed", "confidential", "not specified", "tbd"}
# tracker wording such as "Not named in posting (clinical R&D IT client)" or "Not displayed (client not named)"
PLACEHOLDER_RE = re.compile(r"^not (displayed|named|disclosed|specified|listed|stated|provided|given|shown|available|in posting)\b", re.I)


def clean_company(s):
    if str(s or "").strip().lower() in PLACEHOLDERS or PLACEHOLDER_RE.match(str(s or "").strip()):
        return ""
    words = [w for w in re.sub(r"[,]", " ", str(s or "")).split()]
    while words and re.sub(r"[.]", "", words[-1]).lower() in SUFFIXES:
        words.pop()
    return clean_part(" ".join(words), 25)


def file_name(name, title):
    n = re.sub(r"[^\w+#.-]+", "_", clean_part(name), flags=re.U).strip("_")
    t = re.sub(r"[^\w+#.-]+", "_", clean_part(title, 50), flags=re.U).strip("_")
    return f"{n}_{t}.docx"


# ---------- tracker ----------
def job_companies(tracker, job):
    sys.path.insert(0, HERE)
    import tracker_tool as tt
    wb = tt.open_read(tracker)
    try:
        ws = tt.sheet(wb, tt.MASTER)
        head, rows = tt.rows_of(ws)
        hm = tt.header_map(head)
        idc = tt.need(hm, "Job ID", tt.MASTER)
        want = tt.job_no(job)
        if want is None:
            raise Refused(f"{job} is not a Job ID like JOB-0123.")
        for r in rows:
            if tt.job_no(r[idc]) == want:
                def g(n):
                    i = hm.get(tt.norm(n))
                    return tt.text(r[i]) if i is not None and i < len(r) else ""
                return g("Company"), g("Staffing Company")
        raise Refused(f"{job} was not found in the tracker.")
    finally:
        tt.close_all()


def job_folder(job, client, staffing):
    jid = job.upper()
    c, s = clean_company(client), clean_company(staffing)
    if c and s and key(c) != key(s):
        return f"{jid} - {c} (via {s})"
    if c or s:
        return f"{jid} - {c or s}"
    return jid


# ---------- commands ----------
def cmd_categories(a):
    data = load_config()
    out("Choose one technology from this list. A new technology is created only with --new. Hints help match the job; specialties are optional.")
    for t, v in sorted(data["technologies"].items(), key=lambda kv: key(kv[0])):
        sp = ", ".join(v.get("specialties") or []) or "-"
        out(f"{t}\thints: {', '.join(v.get('hints') or []) or '-'}\tspecialties: {sp}")


JOB_DIR = re.compile(r"^(JOB-\d+|NO-ID_)", re.I)


def cmd_sync(a):
    data = load_config()
    techs = data["technologies"]
    root = resumes_root()
    if not os.path.isdir(root):
        out(f"No Resumes folder yet, nothing to add: {root}")
        return 0
    subdirs = lambda p: sorted(n for n in os.listdir(p) if os.path.isdir(os.path.join(p, n)) and not n.startswith((".", "~")))
    added_t, added_s, notes = [], [], []
    for name in subdirs(root):
        if JOB_DIR.match(name):
            notes.append(f"job folder directly under Resumes, not a technology (left as it is): {name}")
            continue
        by_key = {key(t): t for t in techs}
        tech = by_key.get(key(name))
        if tech is None:
            tech = name  # the folder's own name, so new resumes and the library check use this same folder
            techs[tech] = {"hints": [], "specialties": []}
            added_t.append(tech)
        elif tech.lower() != name.lower():
            notes.append(f'folder "{name}" is listed as "{tech}": new resumes go to "{tech}", and the library check does not search "{name}"')
        specs = techs[tech].setdefault("specialties", [])
        for sub in subdirs(os.path.join(root, name)):
            if JOB_DIR.match(sub) or key(sub) in {key(s) for s in specs}:
                continue
            specs.append(sub)
            added_s.append(f"{tech} / {sub}")
    if not added_t and not added_s:
        out(f"Nothing to add: every folder in {root} is already in the list ({len(techs)} technologies).")
    else:
        out(("Would add" if a.dry_run else "Added") + f" {len(added_t)} technolog{'y' if len(added_t) == 1 else 'ies'} and {len(added_s)} specialt{'y' if len(added_s) == 1 else 'ies'}:")
        for t in added_t:
            out(f"  technology: {t}")
        for s in added_s:
            out(f"  specialty:  {s}")
        if not a.dry_run:
            save_config(data)
    for n in notes:
        out("Note: " + n)
    return 0


def resolve(a, data):
    techs = data["technologies"]
    tech = match_name(a.tech, techs)
    created = False
    if tech is None:
        if not a.new:
            near = difflib.get_close_matches(key(a.tech), [key(t) for t in techs], n=3, cutoff=0.5)
            ex = ", ".join(t for t in techs if key(t) in near) or "none close"
            raise Refused(f'"{a.tech}" is not in the technology list (closest: {ex}). Pick an existing one, or repeat with --new to add it.')
        tech = clean_part(a.tech)
        techs[tech] = {"hints": [], "specialties": []}
        created = True
    spec = None
    if getattr(a, "specialty", None):
        specs = techs[tech].setdefault("specialties", [])
        spec = match_name(a.specialty, specs)
        if spec is None:
            spec = clean_part(a.specialty)
            specs.append(spec)
            created = True
    return tech, spec, created


def target_folder(a, data):
    tech, spec, created = resolve(a, data)
    if a.job:
        if not a.tracker:
            raise Refused("--job needs --tracker (the tracker file) to read the company names.")
        client, staffing = job_companies(a.tracker, a.job)
        jf = job_folder(a.job, client, staffing)
    else:
        jf = "NO-ID_" + dt.date.today().strftime("%Y-%m-%d")
    parts = [resumes_root(), tech] + ([spec] if spec else []) + [jf]
    return tech, spec, created, os.path.join(*parts)


def cmd_path(a):
    data = load_config()
    tech, spec, created, folder = target_folder(a, data)
    fname = file_name(a.name, a.title)
    full = os.path.join(folder, fname)
    if os.path.exists(full):
        out(f"File already exists, nothing overwritten: {full}")
        out("Files already in this job folder:")
        for f in sorted(os.listdir(folder)):
            out("  " + f)
        out("Choose a different title or name that reflects what is different about this resume (for example its main tool set) and run again.")
        return 3
    if not a.dry_run:
        os.makedirs(folder, exist_ok=True)
        if created:
            save_config(data)
    out(f"TECHNOLOGY: {tech}" + (f" / {spec}" if spec else "") + ("  (new, added to the list)" if created and not a.dry_run else ""))
    out(f"FOLDER: {folder}")
    out(f"FILE NAME: {fname}")
    out(f"FULL PATH: {full}")
    if a.dry_run:
        out("(dry run: nothing created)")
    elif os.path.isdir(folder):
        others = sorted(f for f in os.listdir(folder) if f.lower().endswith(".docx"))
        if others:
            out("Other resumes already in this job folder: " + ", ".join(others))
    return 0


def cmd_list(a):
    data = load_config()
    tech, spec, created = resolve(a, data)
    if created:
        save_config(data)
    folder = os.path.join(resumes_root(), tech, *([spec] if spec else []))
    if not os.path.isdir(folder):
        os.makedirs(folder, exist_ok=True)
        out(f"No resumes yet: folder created for future use: {folder}")
        return 0
    found = []
    for dp, _, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(".docx") and not f.startswith("~$"):
                p = os.path.join(dp, f)
                st = os.stat(p)
                found.append((p, st.st_size, st.st_mtime))
    if not found:
        out(f"No resumes yet in: {folder}")
        return 0
    found.sort(key=lambda x: x[2], reverse=True)
    import hashlib
    seen, unique = set(), []
    for item in found:  # a resume copied into another job's folder (reuse) is the same resume: show it once, newest copy first
        with open(item[0], "rb") as fh:
            h = hashlib.md5(fh.read()).hexdigest()
        if h not in seen:
            seen.add(h)
            unique.append(item)
    found = unique
    shown = found[:a.limit] if a.limit > 0 else found
    out(f"{len(found)} resume(s) in {folder}" + (f" (showing the {len(shown)} newest)" if len(shown) < len(found) else ""))
    out("Path	Size KB	Modified")
    for p, sz, m in shown:
        out(f"{p}	{sz // 1024}	{dt.datetime.fromtimestamp(m).strftime('%Y-%m-%d')}")
    return 0


def cmd_reuse(a):
    import shutil
    src = os.path.abspath(a.source)
    if not os.path.isfile(src) or not src.lower().endswith(".docx"):
        raise Refused(f"Source resume not found (a .docx file is needed): {a.source}")
    data = load_config()
    tech, spec, created, folder = target_folder(a, data)
    dest = os.path.join(folder, os.path.basename(src))
    if os.path.abspath(dest) == src:
        out(f"Already in this job folder, nothing copied: {dest}")
        out(f"FULL PATH: {dest}")
        return 0
    if os.path.exists(dest):
        out(f"File already exists in the job folder, nothing overwritten: {dest}")
        return 3
    os.makedirs(folder, exist_ok=True)
    shutil.copy2(src, dest)
    if created:
        save_config(data)
    out(f"COPIED: {src}")
    out(f"FOLDER: {folder}")
    out(f"FILE NAME: {os.path.basename(src)}")
    out(f"FULL PATH: {dest}")
    return 0


def find_resumes(job, filename=None):
    """Full paths of the .docx files in the job folders of that Job ID (only the one named filename when given). tracker_tool uses it
    to fill Resume Path when Module 03 records a Resume Filename."""
    jid = job.strip().upper()
    root = resumes_root()
    hits = []
    for dp, dirs, files in os.walk(root):
        base = os.path.basename(dp).upper()
        if base == jid or base.startswith(jid + " "):
            for f in files:
                if f.lower().endswith(".docx") and not f.startswith("~$"):
                    if not filename or key(f) == key(filename):
                        hits.append(os.path.join(dp, f))
    return sorted(hits)


def cmd_find(a):
    jid = a.job.strip().upper()
    root = resumes_root()
    if not os.path.isdir(root):
        raise Refused(f"No Resumes folder yet: {root}")
    hits = find_resumes(jid, a.filename)
    if not hits:
        raise Refused(f"No resume found for {jid}" + (f" named {a.filename}" if a.filename else "") + f" under {root}")
    for h in sorted(hits):
        out(h)
    return 0


def load_evidence():
    if not os.path.exists(EVIDENCE):
        return {}
    try:
        with open(EVIDENCE, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def cmd_evidence(a):
    data = load_evidence()
    if a.action == "list":
        if not data:
            out("No saved results.")
        for k, v in sorted(data.items()):
            out(f"{v['employer']}	{v['tool']}	{v['outcome']}	{v.get('date', '')}	{v.get('note', '')}")
        return 0
    if not a.employer or not a.tool:
        raise Refused("--employer and --tool are required.")
    k = key(a.employer) + "||" + key(a.tool)
    if a.action == "get":
        v = data.get(k)
        out(f"{v['outcome']}	{v.get('date', '')}	{v.get('note', '')}" if v else "none")
        return 0
    if a.outcome not in ("used", "competitor", "inconclusive"):
        raise Refused("--outcome must be used, competitor or inconclusive.")
    data[k] = {"employer": a.employer, "tool": a.tool, "outcome": a.outcome, "date": dt.date.today().strftime("%Y-%m-%d"), "note": a.note or ""}
    tmp = EVIDENCE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(tmp, EVIDENCE)
    out(f"Saved: {a.employer} / {a.tool} = {a.outcome}")
    return 0


def main(argv):
    p = argparse.ArgumentParser(prog="resume_path.py", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("categories")
    s.set_defaults(fn=cmd_categories)
    s = sub.add_parser("sync")
    s.add_argument("--dry-run", action="store_true")
    s.set_defaults(fn=cmd_sync)
    s = sub.add_parser("path")
    s.add_argument("--tech", required=True)
    s.add_argument("--specialty")
    s.add_argument("--name", required=True)
    s.add_argument("--title", required=True)
    s.add_argument("--job")
    s.add_argument("--tracker")
    s.add_argument("--new", action="store_true")
    s.add_argument("--dry-run", action="store_true")
    s.set_defaults(fn=cmd_path)
    s = sub.add_parser("list")
    s.add_argument("--tech", required=True)
    s.add_argument("--specialty")
    s.add_argument("--new", action="store_true")
    s.add_argument("--limit", type=int, default=10)
    s.set_defaults(fn=cmd_list)
    s = sub.add_parser("reuse")
    s.add_argument("--source", required=True)
    s.add_argument("--tech", required=True)
    s.add_argument("--specialty")
    s.add_argument("--job")
    s.add_argument("--tracker")
    s.add_argument("--new", action="store_true")
    s.set_defaults(fn=cmd_reuse)
    s = sub.add_parser("find")
    s.add_argument("--job", required=True)
    s.add_argument("--filename")
    s.set_defaults(fn=cmd_find)
    s = sub.add_parser("evidence")
    s.add_argument("action", choices=["get", "set", "list"])
    s.add_argument("--employer")
    s.add_argument("--tool")
    s.add_argument("--outcome")
    s.add_argument("--note")
    s.set_defaults(fn=cmd_evidence)
    a = p.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    try:
        return a.fn(a) or 0
    except Refused as e:
        sys.stderr.write(str(e) + "\n")
        return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
