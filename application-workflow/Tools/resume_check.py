"""Checks a finished resume DOCX against the mechanical rules of Module 03, so the model fixes what this reports instead of
counting by eye. It only reads the file; it never changes anything. It needs nothing but Python (no extra packages).

Usage:  python resume_check.py <resume.docx> [--profile "<Candidate Profile.txt>"] [--tools "Ansible,Splunk,RHEL"]
                                [--headline "Ansible"] [--exempt "Java"] [--required "Ansible,Cassandra|MongoDB"] [--preferred "Jenkins,Linux"]
                                (on a Mac: python3 resume_check.py ...)

The layout is checked against the template the resume was rendered with (Tools/resume_render.py records it in the file;
--template overrides it; a file without one is checked as classic). Templates are in Tools/resume_templates.json.

What it checks (each line is PASS, WARN or FAIL):
  layout     the template's font, name and contact alignment, heading style, margins and body size within the template's
             auto-fit limits, no tables (except a table-based template), no text boxes/pictures/columns/page breaks, hidden
             word-break characters, markdown asterisks, bullet alignment, headline length
  bullets    count per employer, word bands (short 10-15, medium 16-25, long 26-35), the 25/50/25 mix, at least one short and one long,
             no more than 3 in a row, first 3-4 bullets of employers 1 and 2 not short, no bold, opening verbs, banned phrases,
             stiff phrasing, first-person pronouns, repeated connecting phrases, employer name inside bullets
  use cases  two sentences, word range per employer, at most 2 tools (taken from the Tech Stack), each tool also named in a bullet
  stacks     every Tech Stack item and every Technical Skills item appears in the bullets; Tech Stack label bold and line italic;
             skills labels bold
  summary    one paragraph, technologies bold at most once and named in the bullets
  length     total words against 1,450-1,650 (three pages)
  --profile  name, contact line, employer names, titles and dates match the Candidate Profile exactly
  --tools    each tool named in at most 2 employers (3 for --headline; --exempt tools are skipped, and are not counted as tools in Project Use Cases)
  --required/--preferred  the ATS match score with Module 03's credits (use A|B for either-or keywords)

Exit code 0 = no FAIL, 1 = at least one FAIL, 2 = the file could not be read.
"""
import argparse, json, os, re, sys, zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TEMPLATES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "resume_templates.json")
STAMP = "resume-template:"  # resume_render.py writes "resume-template:<key>" into the file's Category property
WORDS_MIN, WORDS_MAX = 1450, 1650
RANGES = [(9, 12), (8, 11), (6, 9), (6, 9), (5, 7)]
USE_CASE_WORDS = [(30, 45), (30, 45), (25, 40), (25, 40), (20, 35)]
BANNED = ["responsible for", "worked on", "involved in", "utilized", "leveraged", "played a key role", "successfully", "various",
          "robust solutions", "cutting-edge solutions", "seamless solutions"]
STIFF = [r"\blived in\b", r"\bsat in\b", r"\bis used to\b", r"\bare used to\b"]
CONNECTORS = [r"\bso that\b", r"\bwhich helped\b", r"\bensuring\b", r"\bso\b"]
PRONOUNS = re.compile(r"(?<![\w/])(I|Me|My|We|Our|Us|me|my|we|our|us)\b(?!/)")  # not I/O or PL/I

results = []


def add(level, area, msg):
    results.append((level, area, msg))


def on(el, tag):
    return el.find(W + tag) is not None


def flag(rpr, tag):
    if rpr is None:
        return False
    e = rpr.find(W + tag)
    if e is None:
        return False
    return e.get(W + "val", "true") not in ("0", "false", "none")


class Run:
    def __init__(self, r):
        rpr = r.find(W + "rPr")
        parts = []
        for ch in r:
            if ch.tag == W + "t":
                parts.append(ch.text or "")
            elif ch.tag == W + "tab":
                parts.append("\t")
            elif ch.tag in (W + "br", W + "cr") and ch.get(W + "type") != "page":
                parts.append(" ")
        self.text = "".join(parts)
        self.bold = flag(rpr, "b")
        self.italic = flag(rpr, "i")
        u = rpr.find(W + "u") if rpr is not None else None
        self.underline = u is not None and u.get(W + "val", "single") != "none"
        sz = rpr.find(W + "sz") if rpr is not None else None
        self.size = int(sz.get(W + "val")) / 2 if sz is not None else None
        rf = rpr.find(W + "rFonts") if rpr is not None else None
        self.font = rf.get(W + "ascii") if rf is not None else None


class Para:
    def __init__(self, p):
        self.runs = [Run(r) for r in p.iter(W + "r")]
        self.text = "".join(r.text for r in self.runs)
        ppr = p.find(W + "pPr")
        self.style = ""
        self.bullet = False
        self.jc = ""
        if ppr is not None:
            ps = ppr.find(W + "pStyle")
            self.style = ps.get(W + "val", "") if ps is not None else ""
            self.bullet = ppr.find(W + "numPr") is not None or self.style in ("ListBullet", "ListParagraph")
            jc = ppr.find(W + "jc")
            self.jc = jc.get(W + "val", "") if jc is not None else ""
        self.page_break = any(b.get(W + "type") == "page" for b in p.iter(W + "br"))

    def all_bold(self):
        rs = [r for r in self.runs if r.text.strip()]
        return bool(rs) and all(r.bold for r in rs)

    def any_bold(self):
        return any(r.bold and r.text.strip() for r in self.runs)

    def all_italic(self):
        rs = [r for r in self.runs if r.text.strip()]
        return bool(rs) and all(r.italic for r in rs)


def load(path):
    try:
        z = zipfile.ZipFile(path)
        root = ET.fromstring(z.read("word/document.xml"))
    except Exception as e:
        sys.stderr.write(f"Cannot read {path}: {e}\n")
        sys.exit(2)
    body = root.find(W + "body")
    paras = [Para(p) for p in body.iter(W + "p")]  # document order, including text inside table cells
    xml = z.read("word/document.xml").decode("utf-8", "ignore")
    sect = body.find(W + "sectPr")
    stamp = None
    try:
        core = z.read("docProps/core.xml").decode("utf-8", "ignore")
        m = re.search(re.escape(STAMP) + r"([A-Za-z0-9_]+)", core)
        stamp = m.group(1) if m else None
    except KeyError:
        pass
    return paras, xml, sect, stamp


def load_template(key):
    """The template preset for key, or classic when the key or the templates file is missing."""
    try:
        with open(TEMPLATES_FILE, encoding="utf-8") as f:
            data = {k: v for k, v in json.load(f).items() if not k.startswith("_")}
    except (OSError, ValueError):
        add("WARN", "layout", f"cannot read {TEMPLATES_FILE}; layout checked against classic defaults")
        data = {}
    if key and key in data:
        return key, data[key]
    if key:
        add("WARN", "layout", f"unknown template '{key}'; layout checked as classic")
    else:
        add("WARN", "layout", "the file names no template (not rendered with resume_render.py); layout checked as classic")
    return "classic", data.get("classic", {"font": "Arial", "body_pt": 11, "margins_in": 0.65,
                                            "name": {"pt": 20, "align": "center", "bold": True},
                                            "contact": {"align": "center"},
                                            "heading": {"underline": True}})


def words(s):
    return len(s.split())


def band(n):
    return "S" if 10 <= n <= 15 else "M" if 16 <= n <= 25 else "L" if 26 <= n <= 35 else "X"


def norm(s):
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def mentions(text, tok):
    pat = r"(?<![A-Za-z0-9])" + re.escape(tok.strip()) + r"(?![A-Za-z0-9])"
    return len(re.findall(pat, text, re.I))


def split_sections(paras):
    idx = {}
    for i, p in enumerate(paras):
        t = p.text.strip()
        if t in ("PROFESSIONAL SUMMARY", "TECHNICAL SKILLS", "PROFESSIONAL EXPERIENCE") and t not in idx:
            idx[t] = i
    return idx


LABELS = ("Domain:", "Tech Stack:", "Project Use Case:")


def parse_employers(paras, start):
    """Employer blocks. A header is either 'Employer | Title | Dates' on one line, or 'Employer<tab>Dates' with the
    job title on the next line (split style); both are stored as hdr_text 'Employer | Title | Dates'."""
    emps, cur, i = [], None, start
    n = len(paras)
    while i < n:
        p, t = paras[i], paras[i].text.strip()
        if not t:
            i += 1
            continue
        plain = not p.bullet and not t.startswith(LABELS)
        if plain and re.match(r"^.+ \| .+ \| .+$", t):
            cur = {"hdr": p, "hdr_text": t, "domain": None, "use": None, "bullets": [], "stack": None, "start": i}
            emps.append(cur)
        elif plain and "\t" in t:
            company, _, dates = t.partition("\t")
            j = i + 1
            while j < n and not paras[j].text.strip():
                j += 1
            title = paras[j].text.strip() if j < n and not paras[j].text.strip().startswith(LABELS) else ""
            cur = {"hdr": p, "hdr_text": f"{company.strip()} | {title} | {dates.strip()}", "domain": None, "use": None,
                   "bullets": [], "stack": None, "start": i}
            emps.append(cur)
            if title:
                i = j
        elif cur is not None:
            if t.startswith("Domain:"):
                cur["domain"] = p
            elif t.startswith("Project Use Case:"):
                cur["use"] = p
            elif t.startswith("Tech Stack:"):
                cur["stack"] = p
            elif p.bullet:
                cur["bullets"].append(p)
        i += 1
    return emps


def find_contact(paras):
    """Index of the contact line: the first of lines 2-4 that holds an email address (a headline may come before it)."""
    for i in range(1, min(4, len(paras))):
        if "@" in paras[i].text:
            return i
    return 1


def body_size(emps, t):
    sizes = [r.size for e in emps for b in e["bullets"] for r in b.runs if r.text.strip() and r.size]
    return max(set(sizes), key=sizes.count) if sizes else t.get("body_pt", 11)


def align_ok(jc, want):
    want = want or "left"
    return {"center": jc == "center", "right": jc in ("right", "end"),
            "left": jc in ("", "left", "start"), "justify": jc == "both"}[want]


def check_layout(paras, xml, sect, idx, key, t, emps):
    a = "layout"
    if not paras:
        add("FAIL", a, "document is empty")
        return
    add("INFO", a, f"template: {key}")
    tn, tc, th = t.get("name", {}), t.get("contact", {}), t.get("heading", {})
    ci = find_contact(paras)
    name, contact = paras[0], paras[ci] if len(paras) > ci else None
    ok = align_ok(name.jc, tn.get("align", "center")) and (name.all_bold() or not tn.get("bold", True)) \
        and (not tn.get("pt") or any(r.size == tn["pt"] for r in name.runs))
    add("PASS" if ok else "FAIL", a, f"candidate name: {tn.get('align', 'center')}, {tn.get('pt')} pt"
        + (", bold" if tn.get("bold", True) else "") + ("" if ok else " is required by the template"))
    if contact is not None and align_ok(contact.jc, tc.get("align", "center")):
        add("PASS", a, f"contact line {tc.get('align', 'center')}-aligned")
    else:
        add("FAIL", a, f"contact line must be {tc.get('align', 'center')}-aligned")
    for p in paras[1:ci]:
        n = words(p.text)
        if n > 14:
            add("WARN", a, f"headline is {n} words; keep it to one short line (about 8-14 words)")
    order = ["PROFESSIONAL SUMMARY", "TECHNICAL SKILLS", "PROFESSIONAL EXPERIENCE"]
    if [idx.get(k, -1) for k in order] == sorted(idx.get(k, -1) for k in order) and all(k in idx for k in order):
        add("PASS", a, "the three section titles are present, in order")
    else:
        add("FAIL", a, "section titles PROFESSIONAL SUMMARY / TECHNICAL SKILLS / PROFESSIONAL EXPERIENCE missing or out of order")
    for k in order:
        if k in idx:
            p = paras[idx[k]]
            ok = p.all_bold() and (all(r.underline for r in p.runs if r.text.strip()) or not th.get("underline"))
            if not ok:
                add("FAIL", a, f"{k} must be bold" + (" and underlined" if th.get("underline") else ""))
    font = t.get("font", "Arial")
    fonts = {r.font for p in paras for r in p.runs if r.text.strip() and r.font}
    if fonts and fonts != {font}:
        add("FAIL", a, f"the {key} template uses {font} only; found {sorted(fonts)}")
    else:
        add("PASS", a, f"{font} throughout")
    if sect is not None:
        mar = sect.find(W + "pgMar")
        if mar is not None:
            vals = [int(mar.get(W + k, 0)) / 1440 for k in ("top", "bottom", "left", "right")]
            hi = t.get("margins_in", 0.65) + 0.1
            if all(0.5 - 0.005 <= v <= hi + 0.005 for v in vals):
                add("PASS", a, f"margins {vals[0]:.2f} inch (template range 0.50-{hi:.2f})")
            else:
                add("FAIL", a, f"margins must be 0.50-{hi:.2f} inch for the {key} template; found {[round(v, 2) for v in vals]}")
    bs = body_size(emps, t)
    lo = t.get("min_body_pt", 10)
    if bs < lo:
        add("FAIL", a, f"bullet text is {bs} pt; the {key} template's minimum is {lo} pt")
    else:
        add("PASS", a, f"bullet text {bs} pt (minimum {lo})")
    table_ok = t.get("layout") == "hanging"
    for tag, label in (("<w:tbl>", "tables"), ("<w:drawing", "pictures or shapes"), ("<w:pict", "text boxes or shapes")):
        if tag in xml:
            if tag == "<w:tbl>" and table_ok:
                add("INFO", a, f"the {key} template lays the page out in a table (ATS caution)")
            else:
                add("FAIL", a, f"{label} found; the {key} template is a single-column text document")
    if sect is not None:
        cols = sect.find(W + "cols")
        if cols is not None and int(cols.get(W + "num", "1")) > 1:
            add("FAIL", a, "multiple columns found; the resume must be a single-column text document")
    if any(p.page_break for p in paras):
        add("FAIL", a, "manual page break found; keep continuous flow")
    full = "\n".join(p.text for p in paras)
    bad = [c for c in ("­", "​", "‌", "‍", "⁠") if c in full]
    if bad:
        add("FAIL", a, "soft hyphen / zero-width characters found")
    if "**" in full:
        add("FAIL", a, "visible markdown asterisks found")


def check_summary_skills(paras, idx, emps, all_bullet_text):
    s_i, k_i, e_i = idx.get("PROFESSIONAL SUMMARY"), idx.get("TECHNICAL SKILLS"), idx.get("PROFESSIONAL EXPERIENCE")
    if None in (s_i, k_i, e_i):
        return
    summ = [p for p in paras[s_i + 1:k_i] if p.text.strip()]
    if len(summ) == 1:
        add("PASS", "summary", f"one paragraph, {words(summ[0].text)} words")
        bolds = [r.text.strip() for r in summ[0].runs if r.bold and r.text.strip()]
        dup = sorted({b for b in bolds if bolds.count(b) > 1})
        if dup:
            add("FAIL", "summary", f"technologies bold more than once: {dup}")
        missing = [b for b in bolds if not mentions(all_bullet_text, b)]
        if missing:
            add("FAIL", "summary", f"bold technologies not named in any bullet: {missing}")
        else:
            add("PASS", "summary", "every bold technology is named in the bullets")
    else:
        add("FAIL", "summary", f"summary must be one paragraph; found {len(summ)}")
    skills = [p for p in paras[k_i + 1:e_i] if p.text.strip()]
    if not skills:
        add("FAIL", "skills", "Technical Skills is empty")
        return
    miss, unbold = [], []
    for p in skills:
        if ":" not in p.text:
            add("WARN", "skills", f"line without a category label: {p.text[:40]}")
            continue
        label, items = p.text.split(":", 1)
        lab_runs = [r for r in p.runs if r.text.strip() and r.text.strip().rstrip(":") and r.bold]
        if not lab_runs:
            unbold.append(label)
        for r in p.runs:
            if r.bold and r.text.strip() and r.text.strip() not in (label + ":", label, ":"):
                add("FAIL", "skills", f"items must not be bold: {r.text.strip()[:30]}")
                break
        for it in re.split(r",\s*|\s+[·•]\s+", items):
            it = it.strip()
            if not it:
                continue
            base = re.sub(r"\s*\(.*?\)", "", it).strip()
            inner = re.findall(r"\((.*?)\)", it)
            cands = [base] + inner
            if not any(mentions(all_bullet_text, c) or norm(c) in norm(all_bullet_text) for c in cands):
                miss.append(it)
    if unbold:
        add("FAIL", "skills", f"category labels must be bold: {unbold}")
    if miss:
        add("WARN", "skills", f"Technical Skills items not found in any bullet (check each is a real practice or tool): {miss}")
    else:
        add("PASS", "skills", "every Technical Skills item appears in the bullets; labels bold")


def check_employers(emps, exempt, t):
    if not emps:
        add("FAIL", "experience", "no employer blocks found")
        return
    bs = body_size(emps, t)
    justify = t.get("justify", True)
    for k, e in enumerate(emps):
        rng = RANGES[min(k, 4)]
        uc = USE_CASE_WORDS[min(k, 4)]
        nm = e["hdr_text"].split("|")[0].strip()
        tag = f"E{k + 1} {nm}"
        first = next((r for r in e["hdr"].runs if r.text.strip()), None)  # the employer name
        if first is None or not first.bold or not first.size or first.size < bs + 0.5:
            add("FAIL", tag, f"employer name must be bold and at least 0.5 pt larger than the bullet text ({bs} pt)")
        for key, label in (("domain", "Domain:"), ("use", "Project Use Case:")):
            p = e[key]
            if p is None:
                add("FAIL", tag, f"missing '{label}' line")
                continue
            lab = [r for r in p.runs if r.text.strip()]
            if not (lab and lab[0].bold and not any(r.bold for r in lab[1:])):
                add("FAIL", tag, f"'{label}' label must be bold and its text plain")
        if e["use"] is not None:
            txt = e["use"].text.split(":", 1)[1].strip()
            n = words(txt)
            sent = len([s for s in re.split(r"(?<=[.!?])\s+", txt) if s.strip()])
            if not (uc[0] <= n <= uc[1]):
                add("FAIL", tag, f"Project Use Case is {n} words; range is {uc[0]}-{uc[1]}")
            if sent != 2:
                add("FAIL", tag, f"Project Use Case has {sent} sentences; it needs 2")
            e["use_text"] = txt
        bl = e["bullets"]
        nb = len(bl)
        if not (rng[0] <= nb <= rng[1]):
            add("FAIL", tag, f"{nb} bullets; range is {rng[0]}-{rng[1]}")
        else:
            add("PASS", tag, f"{nb} bullets (range {rng[0]}-{rng[1]})")
        texts = [b.text.strip() for b in bl]
        ws = [words(t) for t in texts]
        bd = [band(n) for n in ws]
        if "X" in bd:
            add("FAIL", tag, "bullets outside 10-35 words: " + ", ".join(f"#{i + 1}={ws[i]}" for i, x in enumerate(bd) if x == "X"))
        cnt = {b: bd.count(b) for b in "SML"}
        if nb:
            if cnt["S"] < 1 or cnt["L"] < 1:
                add("FAIL", tag, f"needs at least one short and one long bullet (S={cnt['S']} M={cnt['M']} L={cnt['L']})")
            sh, lg = cnt["S"] / nb, cnt["L"] / nb
            if not (0.15 <= sh <= 0.35 and 0.15 <= lg <= 0.35):
                add("WARN", tag, f"mix is S={cnt['S']} M={cnt['M']} L={cnt['L']}; target is about 25/50/25")
            else:
                add("PASS", tag, f"mix S={cnt['S']} M={cnt['M']} L={cnt['L']} ({''.join(bd)})")
        run = 1
        for i in range(1, len(bd)):
            run = run + 1 if bd[i] == bd[i - 1] else 1
            if run > 3:
                add("FAIL", tag, f"more than 3 bullets in a row in band {bd[i]} (ending at #{i + 1})")
                break
        if k < 2 and "S" in bd[:4]:
            add("FAIL", tag, "the first 3-4 bullets must be medium or long (a short bullet is in the top 4)")
        if any(b.any_bold() for b in bl):
            add("FAIL", tag, "bullets must contain no bold")
        if justify and any(b.jc != "both" for b in bl):
            add("FAIL", tag, "every bullet must be justified (the template justifies bullets)")
        if not justify and any(b.jc == "both" for b in bl):
            add("FAIL", tag, "bullets must be left-aligned in this template")
        verbs = [t.split()[0].lower().strip(",.;:") for t in texts if t]
        rep = sorted({v for v in verbs if verbs.count(v) > 2})
        adj = sorted({verbs[i] for i in range(1, len(verbs)) if verbs[i] == verbs[i - 1]})
        if rep:
            add("FAIL", tag, f"opening verb used more than twice: {rep}")
        if adj:
            add("FAIL", tag, f"same opening verb in neighboring bullets: {adj}")
        nmain = nm.lower()
        inb = sum(1 for t in texts if nmain and nmain in t.lower())
        if inb > 1:
            add("WARN", tag, f"employer name appears in {inb} bullets (at most one)")
        # Tech Stack
        st = e["stack"]
        if st is None:
            add("FAIL", tag, "missing Tech Stack line")
        else:
            lab = next((r for r in st.runs if r.text.strip()), None)
            if not (lab is not None and lab.bold and st.all_italic()):
                add("FAIL", tag, "Tech Stack line must be italic with a bold 'Tech Stack:' label")
            btxt = " ".join(texts)
            items = [x.strip() for x in st.text.split(":", 1)[1].split(",") if x.strip()]
            miss = [x for x in items if not (mentions(btxt, x) or norm(x) in norm(btxt))]
            if miss:
                add("WARN", tag, f"Tech Stack items not found in this employer's bullets: {miss}")
            else:
                add("PASS", tag, "every Tech Stack item appears in this employer's bullets")
            if e.get("use_text"):
                named = [x for x in items if mentions(e["use_text"], x) and x.lower() not in exempt]
                if len(named) > 2:
                    add("FAIL", tag, f"Project Use Case names more than 2 tools (databases, cloud platforms, frameworks and servers count; languages listed in --exempt do not): {named}")
                unbacked = [x for x in named if not mentions(btxt, x)]
                if unbacked:
                    add("FAIL", tag, f"Project Use Case names tools no bullet names: {unbacked}")


def check_text_rules(emps):
    allb = [b.text.strip() for e in emps for b in e["bullets"]]
    joined = "\n".join(allb)
    low = joined.lower()
    hits = [b for b in BANNED if re.search(r"\b" + re.escape(b) + r"\b", low)]
    if hits:
        add("FAIL", "style", f"banned phrases found: {hits}")
    else:
        add("PASS", "style", "no banned phrases")
    stiff = [s for s in STIFF if re.search(s, low)]
    if stiff:
        add("FAIL", "style", f"stiff or passive phrasing found: {stiff}")
    fp = sorted(set(PRONOUNS.findall(joined)))
    if fp:
        add("FAIL", "style", f"first-person pronouns found: {fp}")
    else:
        add("PASS", "style", "no first-person pronouns")
    for c in CONNECTORS:
        n = sum(1 for t in allb if re.search(c, t.lower()))
        if n > 2:
            add("FAIL", "style", f"connecting phrase '{c.strip(chr(92) + 'b')}' used in {n} bullets (limit 2)")


def check_profile(path, paras, emps):
    txt = open(path, encoding="utf-8", errors="ignore").read()

    def field(label):
        m = re.search(r"^\s*" + re.escape(label) + r":\s*(.+)$", txt, re.M)
        return m.group(1).strip() if m else ""
    contact = " | ".join(x for x in (field("Location"), field("Phone"), field("Email"), field("LinkedIn")) if x)
    if paras and paras[0].text.strip() != field("Name"):
        add("FAIL", "profile", f"name differs from the profile ({field('Name')})")
    else:
        add("PASS", "profile", "name matches")
    parts = lambda s: [x for x in re.split(r"\s*[|•·]\s*", s.strip()) if x]  # templates may use • or · between items
    ci = find_contact(paras)
    if len(paras) > ci and parts(paras[ci].text) != parts(contact):
        add("FAIL", "profile", f"contact line differs; expected: {contact}")
    else:
        add("PASS", "profile", "contact line matches")
    names = re.findall(r"^\s*Employer / client name:\s*(.+)$", txt, re.M)
    titles = re.findall(r"^\s*Job title:\s*(.+)$", txt, re.M)
    dates = re.findall(r"^\s*Dates:\s*(.+)$", txt, re.M)
    want = [f"{a.strip()} | {b.strip()} | {c.strip()}" for a, b, c in zip(names, titles, dates)]
    got = [e["hdr_text"] for e in emps]
    if got != want:
        add("FAIL", "profile", f"employer headers differ from the profile.\n      expected: {want}\n      found:    {got}")
    else:
        add("PASS", "profile", "employer names, titles and dates match, in order")


def check_tools(emps, tools, headline, exempt):
    stacks = []
    for e in emps:
        t = " ".join(b.text for b in e["bullets"]) + " " + (e["stack"].text if e["stack"] else "")
        stacks.append(t)
    for t in tools:
        if t.lower() in exempt:
            continue
        where = [k + 1 for k, s in enumerate(stacks) if mentions(s, t)]
        limit = 3 if t.lower() in headline else 2
        if len(where) > limit:
            add("FAIL", "tools", f"{t} is named in {len(where)} employers {where}; limit is {limit} (add --exempt {t} if the JD's years requirement justifies it)")
        else:
            add("PASS", "tools", f"{t}: employers {where or 'none'} (limit {limit})")


def ats_score(paras, idx, emps, required, preferred):
    s_i, k_i, e_i = idx.get("PROFESSIONAL SUMMARY"), idx.get("TECHNICAL SKILLS"), idx.get("PROFESSIONAL EXPERIENCE")
    summ_skills = " ".join(p.text for p in paras[s_i + 1:e_i]) if None not in (s_i, e_i) else ""
    bullets = [b.text for e in emps for b in e["bullets"]]
    stacks = " ".join(e["stack"].text for e in emps if e["stack"])
    total = got = 0.0
    rows = []
    for words_, w in ((required, 2), (preferred, 1)):
        for kw in words_:
            alts = [a.strip() for a in kw.split("|") if a.strip()]
            n = sum(1 for b in bullets if any(mentions(b, a) for a in alts))
            if n >= 2:
                c = 1.0
            elif n == 1 or any(mentions(summ_skills, a) for a in alts):
                c = 0.5
            elif any(mentions(stacks, a) for a in alts):
                c = 0.25
            else:
                c = 0.0
            total += w
            got += w * c
            rows.append((kw, "required" if w == 2 else "preferred", n, c))
    if total:
        add("INFO", "ats", f"ATS match score {round(got / total * 100)}% (synonym credit of 0.25 is not computed here)")
        for kw, kind, n, c in rows:
            if c < 1:
                add("INFO", "ats", f"  {kw} ({kind}): {n} bullet mention(s), credit {c}")


def main(argv):
    ap = argparse.ArgumentParser(prog="resume_check.py", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("resume")
    ap.add_argument("--profile")
    ap.add_argument("--tools", default="")
    ap.add_argument("--headline", default="")
    ap.add_argument("--exempt", default="")
    ap.add_argument("--required", default="")
    ap.add_argument("--preferred", default="")
    ap.add_argument("--template", default="", help="check the layout against this template instead of the one the file names")
    a = ap.parse_args(argv)
    csv = lambda s: [x.strip() for x in s.split(",") if x.strip()]
    paras, xml, sect, stamp = load(a.resume)
    key, tpl = load_template(a.template.strip().lower() or stamp)
    idx = split_sections(paras)
    emps = parse_employers(paras, idx.get("PROFESSIONAL EXPERIENCE", len(paras)) + 1)
    check_layout(paras, xml, sect, idx, key, tpl, emps)
    allb = " ".join(b.text for e in emps for b in e["bullets"])
    check_summary_skills(paras, idx, emps, allb)
    check_employers(emps, {x.lower() for x in csv(a.exempt)}, tpl)
    check_text_rules(emps)
    total = sum(words(p.text) for p in paras)
    if WORDS_MIN <= total <= WORDS_MAX:
        add("PASS", "length", f"{total} words (target {WORDS_MIN:,}-{WORDS_MAX:,})")
    else:
        add("FAIL", "length", f"{total} words; target is {WORDS_MIN:,}-{WORDS_MAX:,} (three pages)")
    if a.profile:
        check_profile(a.profile, paras, emps)
    if a.tools:
        check_tools(emps, csv(a.tools), {x.lower() for x in csv(a.headline)}, {x.lower() for x in csv(a.exempt)})
    if a.required or a.preferred:
        ats_score(paras, idx, emps, csv(a.required), csv(a.preferred))
    order = {"FAIL": 0, "WARN": 1, "INFO": 2, "PASS": 3}
    fails = sum(1 for r in results if r[0] == "FAIL")
    warns = sum(1 for r in results if r[0] == "WARN")
    for lvl, area, msg in sorted(results, key=lambda r: order[r[0]]):
        sys.stdout.write(f"{lvl:5} [{area}] {msg}\n")
    sys.stdout.write(f"\n{fails} FAIL, {warns} WARN, {sum(1 for r in results if r[0] == 'PASS')} PASS\n")
    return 1 if fails else 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    sys.exit(main(sys.argv[1:]))
