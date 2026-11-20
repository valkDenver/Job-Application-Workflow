"""Render a Module 03 resume to DOCX in a chosen template.

The AI writes the resume content once as JSON; this script lays it out. The same
content can be rendered in any template in Tools/resume_templates.json. A template
picks a page layout and a style for each part (header, headings, employers, Tech
Stack), plus fonts, sizes, colours and spacing. The words never change.

Usage (python3 on a Mac):
  python Tools/resume_render.py templates
  python Tools/resume_render.py render <content.json> <out.docx> [--template classic]

Content JSON:
  {
    "name": "Jane Doe",
    "headline": "Optional positioning line under the name, built from the Target Role",
    "contact": "City, ST | 555-0100 | jane@example.com | linkedin.com/in/jane",
    "summary": "One paragraph. **Technology** in double asterisks is bolded.",
    "skills": [["Languages", "Python, SQL"], ["Cloud", "AWS, Azure"]],
    "employers": [
      {"header": "Employer | Job title | Jan 2024 - Present",
       "domain": "Domain text",
       "usecase": "Project Use Case text.",
       "bullets": ["Bullet one.", "Bullet two."],
       "stack": "Python, AWS, CI/CD"}
    ]
  }

Template keys (see resume_templates.json for examples):
  layout          single | hanging (section titles in a narrow left column)
  header.style    simple | band (name block on a shaded full-width band)
  heading.style   text (underline and/or rule) | bar (shaded bar with a left accent)
  employer.style  inline (Company | Title | Dates on one line)
                  | split (company left, dates right, job title on the line below)
  justify         false left-aligns body text (used in narrow columns)
  stack           how the Tech Stack line looks: bold, italic, color
  gap_stack_pt, gap_employer_pt   space before each Tech Stack / between employers
  footer          name and "Page X of Y" from page 2 on
"""
import argparse
import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

TEMPLATES_FILE = Path(__file__).with_name("resume_templates.json")
HEADINGS = ("PROFESSIONAL SUMMARY", "TECHNICAL SKILLS", "PROFESSIONAL EXPERIENCE")
ALIGN = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
         "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}
PAGE_WIDTH_IN = 8.5

# Word requires child elements in schema order. These are the full orders for a paragraph's
# pPr and a run's rPr, used to insert borders, shading and character spacing in the right place.
PPR_ORDER = ("pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr "
             "suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap "
             "overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid "
             "spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection "
             "textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange").split()
RPR_ORDER = ("rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss "
             "imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs "
             "highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout "
             "specVanish oMath").split()


def insert_ordered(parent, el, tag, order):
    after = tuple("w:" + t for t in order[order.index(tag) + 1:])
    parent.insert_element_before(el, *after)


def load_templates():
    data = json.loads(TEMPLATES_FILE.read_text(encoding="utf-8"))
    return {k: v for k, v in data.items() if not k.startswith("_")}


def split_header(h):
    parts = [x.strip() for x in h.split(" | ")]
    while len(parts) < 3:
        parts.append("")
    return parts[0], parts[1], " | ".join(parts[2:])


def keep_next(p, together=False):
    """Keep this paragraph on the same page as the next one (no stranded headers)."""
    p.paragraph_format.keep_with_next = True
    if together:
        p.paragraph_format.keep_together = True
    return p


class Renderer:
    def __init__(self, t):
        self.t = t
        self.font = t["font"]
        self.body = t["body_pt"]
        self.margin = t["margins_in"]
        self.body_align = "justify" if t.get("justify", True) else "left"
        self.doc = Document()
        for s in self.doc.sections:
            m = Inches(self.margin)
            s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = m
            s.footer_distance = Inches(max(0.25, self.margin - 0.3))
        for name in ("Normal", "List Bullet"):
            st = self.doc.styles[name]
            st.font.name = self.font
            st.font.size = Pt(self.body)
            st.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), self.font)
            pf = st.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = t.get("line_spacing", 1.0)
            pf.widow_control = True
        self.target = self.doc
        self.width = PAGE_WIDTH_IN - 2 * self.margin  # text width of the current target
        self._fresh = None  # the empty first paragraph of a new table cell, reused once

    # ---------- low-level helpers ----------
    def color(self, key):
        """A colour value is a hex string or the name of a preset colour ('accent', 'muted')."""
        if not key:
            return None
        return self.t.get(key, key)

    def into(self, cell, width_in):
        prev = (self.target, self.width, self._fresh)
        self.target, self.width, self._fresh = cell, width_in, cell.paragraphs[0]
        return prev

    def back(self, prev):
        self.target, self.width, self._fresh = prev

    def para(self, align=None, before=0, after=0, style=None):
        if self._fresh is not None:
            p, self._fresh = self._fresh, None
            if style:
                p.style = self.doc.styles[style]
        else:
            p = self.target.add_paragraph(style=style) if style else self.target.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        pf.line_spacing = self.t.get("line_spacing", 1.0)
        if align is not None:
            p.alignment = ALIGN[align] if isinstance(align, str) else align
        return p

    def run(self, p, text, bold=False, italic=False, underline=False, pt=None,
            color=None, spacing=None):
        r = p.add_run(text)
        r.bold, r.italic, r.underline = bold, italic, underline
        r.font.size = Pt(pt or self.body)
        r.font.name = self.font
        rpr = r._element.get_or_add_rPr()
        rpr.get_or_add_rFonts().set(qn("w:eastAsia"), self.font)
        c = self.color(color)
        if c:
            r.font.color.rgb = RGBColor.from_string(c)
        if spacing:  # character spacing in points (Word stores twentieths of a point)
            sp = OxmlElement("w:spacing")
            sp.set(qn("w:val"), str(int(spacing * 20)))
            insert_ordered(rpr, sp, "spacing", RPR_ORDER)
        return r

    def field(self, p, code, **kw):
        """A Word field such as PAGE or NUMPAGES, updated by Word when the file opens or prints."""
        for kind, text in (("begin", None), ("instr", f" {code} "), ("separate", None),
                           ("text", "1"), ("end", None)):
            r = self.run(p, "" if kind != "text" else text, **kw)
            if kind == "instr":
                r.text = ""
                it = OxmlElement("w:instrText")
                it.set(qn("xml:space"), "preserve")
                it.text = text
                r._element.append(it)
            elif kind != "text":
                fc = OxmlElement("w:fldChar")
                fc.set(qn("w:fldCharType"), kind)
                r._element.append(fc)

    def rule(self, p, color, sides=("bottom",), size=6, space=2):
        """Thin paragraph border: a horizontal line that ATS parsers ignore."""
        bdr = OxmlElement("w:pBdr")
        for side in ("top", "left", "bottom", "right"):
            if side in sides:
                e = OxmlElement(f"w:{side}")
                e.set(qn("w:val"), "single")
                e.set(qn("w:sz"), str(size))
                e.set(qn("w:space"), str(space))
                e.set(qn("w:color"), self.color(color))
                bdr.append(e)
        insert_ordered(p._p.get_or_add_pPr(), bdr, "pBdr", PPR_ORDER)

    def shade(self, p, fill):
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), self.color(fill))
        insert_ordered(p._p.get_or_add_pPr(), shd, "shd", PPR_ORDER)

    def right_tab(self, p):
        p.paragraph_format.tab_stops.add_tab_stop(Inches(self.width - 0.02), WD_TAB_ALIGNMENT.RIGHT)

    def cell_border(self, cell, side, color, size=6):
        tcPr = cell._tc.get_or_add_tcPr()
        b = OxmlElement("w:tcBorders")
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(size))
        e.set(qn("w:color"), self.color(color))
        b.append(e)
        tcPr.insert(1 if tcPr.find(qn("w:tcW")) is not None else 0, b)

    def table(self, widths, rows=1):
        """A borderless fixed-width table in the document body."""
        tbl = self.doc.add_table(rows=rows, cols=len(widths))
        tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        tbl.autofit = False
        layout = OxmlElement("w:tblLayout")
        layout.set(qn("w:type"), "fixed")
        tbl._tbl.tblPr.append(layout)
        for row in tbl.rows:
            for cell, w in zip(row.cells, widths):
                cell.width = Inches(w)
        return tbl

    def labelled(self, label, text, align=None, before=0, after=0, label_color=None):
        p = self.para(align, before, after)
        self.run(p, label, bold=True, color=label_color)
        self.run(p, text)
        return p

    # ---------- parts ----------
    def footer(self, c):
        """Name and 'Page X of Y' on pages 2 onward; page 1 has no footer."""
        f = self.t.get("footer")
        if not f:
            return
        sec = self.doc.sections[0]
        sec.different_first_page_header_footer = True
        p = sec.footer.paragraphs[0]
        p.alignment = ALIGN[f.get("align", "right")]
        kw = dict(pt=f.get("pt", 8.5), color=f.get("color", "808080"))
        self.run(p, c["name"] + "   ·   Page ", **kw)
        self.field(p, "PAGE", **kw)
        self.run(p, " of ", **kw)
        self.field(p, "NUMPAGES", **kw)

    def header(self, c):
        t = self.t
        n, ct = t["name"], t["contact"]
        hs = t.get("header", {})
        band = hs.get("style") == "band"
        contact = c["contact"]
        if t.get("contact_sep"):
            contact = contact.replace(" | ", t["contact_sep"])

        p = self.para(n["align"], 0, n.get("after", 0))
        self.run(p, c["name"], bold=n.get("bold", True), pt=n["pt"],
                 color="FFFFFF" if band else n.get("color"), spacing=n.get("spacing"))
        if band:
            self.shade(p, hs["fill"])

        hl = t.get("headline")
        if hl and c.get("headline"):
            p = self.para(n["align"], hl.get("before", 0), hl.get("after", 1))
            self.run(p, c["headline"], bold=hl.get("bold", False), pt=hl["pt"],
                     color=hl.get("color"), spacing=hl.get("spacing"))
            if band:
                self.shade(p, hs["fill"])

        p = self.para(ct["align"], 0, ct.get("after", 0))
        self.run(p, contact, pt=ct["pt"], color=ct.get("color"))
        if band:
            self.shade(p, hs["fill"])
            p.paragraph_format.line_spacing = 1.2
        if hs.get("rule"):
            self.rule(p, hs["rule"], size=12, space=4)

    def heading(self, text, first=False):
        h = self.t["heading"]
        before = h.get("first_before", h.get("before", 8)) if first else h.get("before", 8)
        p = keep_next(self.para(h.get("align"), before, h.get("after", 3)))
        if h.get("style") == "bar":
            p.paragraph_format.left_indent = Inches(0.05)
        self.run(p, text, bold=h.get("bold", True), underline=h.get("underline", False),
                 pt=h["pt"], color=h.get("color"), spacing=h.get("spacing"))
        if h.get("style") == "bar":
            self.shade(p, h["fill"])
            if h.get("rule"):
                self.rule(p, h["rule"], sides=("left",), size=24)
        elif h.get("rule"):
            self.rule(p, h["rule"], sides=("top", "bottom") if h.get("rule_above") else ("bottom",))
        return p

    def summary(self, c):
        p = self.para(self.body_align)
        for i, part in enumerate(re.split(r"\*\*", c["summary"])):
            if part:
                self.run(p, part, bold=(i % 2 == 1))

    def skills(self, c):
        t = self.t
        sep = t.get("skill_sep")
        for label, items in c["skills"]:
            self.labelled(label + ": ", items.replace(", ", sep) if sep else items,
                          after=t.get("skill_after", 0), label_color=t.get("skill_label_color"))

    def employer_head(self, e, before):
        et = self.t["employer"]
        company, title, dates = split_header(e["header"])
        if et.get("style") == "split":
            p = keep_next(self.para(before=before))
            self.right_tab(p)
            self.run(p, company, bold=True, pt=et["pt"], color=et.get("color"))
            self.run(p, "\t" + dates, bold=et.get("dates_bold", True),
                     pt=et.get("dates_pt", self.body), color=et.get("dates_color"))
            p = keep_next(self.para(after=et.get("after", 1)))
            self.run(p, title, bold=et.get("title_bold", False), italic=et.get("title_italic", True),
                     pt=et.get("title_pt", self.body), color=et.get("title_color"))
        else:
            p = keep_next(self.para(before=before, after=et.get("after", 0)))
            self.run(p, e["header"], bold=True, pt=et["pt"], color=et.get("color"))

    def employer(self, e, first):
        t = self.t
        self.employer_head(e, 0 if first else t.get("gap_employer_pt", 10))
        la, lc = t.get("label_after", 0), t.get("label_color")
        keep_next(self.labelled("Domain: ", e["domain"], after=la, label_color=lc))
        keep_next(self.labelled("Project Use Case: ", e["usecase"], align=self.body_align,
                                after=t.get("usecase_after", 2), label_color=lc), together=True)
        last = len(e["bullets"]) - 1
        for i, b in enumerate(e["bullets"]):
            p = self.para(self.body_align, 0, t["bullet"]["after"], style="List Bullet")
            self.run(p, b)
            if i == 0 or i == last:  # first bullet stays with the use case; last with the Tech Stack
                keep_next(p, together=True)
        s = t.get("stack", {})
        p = self.para(before=t.get("gap_stack_pt", 6))
        if s.get("label_bold") is not None:
            self.run(p, "Tech Stack: ", bold=s["label_bold"], italic=s.get("italic", True),
                     color=s.get("label_color", s.get("color")))
            self.run(p, e["stack"], bold=s.get("bold", True), italic=s.get("italic", True),
                     color=s.get("color"))
        else:
            self.run(p, "Tech Stack: " + e["stack"], bold=s.get("bold", True),
                     italic=s.get("italic", True), color=s.get("color"))
        if s.get("fill"):
            self.shade(p, s["fill"])

    def experience(self, c):
        for i, e in enumerate(c["employers"]):
            self.employer(e, i == 0)

    # ---------- page layouts ----------
    def render(self, c):
        self.footer(c)
        getattr(self, "layout_" + self.t.get("layout", "single"))(c)
        return self.doc

    def layout_single(self, c):
        self.header(c)
        self.heading(HEADINGS[0], first=True)
        self.summary(c)
        self.heading(HEADINGS[1])
        self.skills(c)
        self.heading(HEADINGS[2])
        self.experience(c)

    def layout_hanging(self, c):
        """Section titles sit in a narrow left column beside their content."""
        t = self.t
        self.header(c)
        hw = t.get("hang_col_in", 1.3)
        full = self.width
        tbl = self.table([hw, full - hw], rows=3)
        parts = ((HEADINGS[0], self.summary), (HEADINGS[1], self.skills), (HEADINGS[2], self.experience))
        h = t["heading"]
        for i, (row, (title, fill)) in enumerate(zip(tbl.rows, parts)):
            top = h.get("first_before", 4) if i == 0 else h.get("before", 8)
            prev = self.into(row.cells[0], hw)
            for j, w in enumerate(title.split(" ", 1)):
                p = self.para(before=top if j == 0 else 0)
                self.run(p, w, bold=True, pt=h["pt"], color=h.get("color"), spacing=h.get("spacing"))
            self.back(prev)
            if i and t.get("hang_rule"):
                for cell in row.cells:
                    self.cell_border(cell, "top", t["hang_rule"])
            prev = self.into(row.cells[1], full - hw - 0.15)
            self.para(before=top).paragraph_format.line_spacing = Pt(1)
            fill(c)
            self.para(after=2).paragraph_format.line_spacing = Pt(2)
            self.back(prev)


# ---------- auto-fit to a page count ----------
# Scale levels, loosest first. Each level scales the body and heading sizes, the spacing and
# the margins of a template by this factor (the name keeps its size). Fit picks the loosest
# level that ends on the target page count, which also leaves the last page as full as it can be.
FIT_LEVELS = (1.10, 1.07, 1.04, 1.0, 0.97, 0.94, 0.91, 0.88, 0.85)
FONT_KEYS = ("pt", "dates_pt", "title_pt")
GAP_KEYS = ("before", "after", "first_before", "gap_stack_pt", "gap_employer_pt", "skill_after",
            "label_after", "usecase_after")


def scaled(t, s):
    """A copy of template t with sizes, gaps and margins scaled by s, within readable limits."""
    t = json.loads(json.dumps(t))
    half = lambda v: round(v * 2) / 2  # Word sizes go in half points
    lo = t.get("min_body_pt", 10)
    hi = t["body_pt"] + 1
    t["body_pt"] = min(hi, max(lo, half(t["body_pt"] * s)))
    t["margins_in"] = round(min(t["margins_in"] + 0.1, max(0.5, t["margins_in"] * s)), 2)
    for k in GAP_KEYS:
        if k in t:
            t[k] = half(t[k] * s)
    for part in ("headline", "contact", "heading", "employer", "bullet"):
        d = t.get(part, {})
        for k in FONT_KEYS:
            if k in d:
                d[k] = max(lo, half(d[k] * s))
        for k in GAP_KEYS:
            if k in d:
                d[k] = half(d[k] * s)
    return t


def count_pages(paths):
    """[(pages, last_page_fill 0-1 or None)] for each DOCX, laid out by Word or LibreOffice.

    Uses Word through pywin32 or PowerShell on Windows, or LibreOffice (soffice) anywhere.
    Returns None when no layout engine is available (fit is then skipped)."""
    paths = [str(Path(p).resolve()) for p in paths]
    if sys.platform == "win32":
        try:
            import win32com.client  # noqa: F401
            return _count_word_pywin32(paths)
        except ImportError:
            return _count_word_powershell(paths)
    return _count_soffice(paths)


def _count_word_pywin32(paths):
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    out = []
    try:
        for p in paths:
            doc = word.Documents.Open(p, False, True, False)
            pages = doc.ComputeStatistics(2)
            end = doc.Content
            end.Collapse(0)
            sec = doc.Sections(1).PageSetup
            usable = sec.PageHeight - sec.TopMargin - sec.BottomMargin
            fill = (end.Information(6) - sec.TopMargin) / usable
            out.append((pages, max(0.0, min(1.0, fill))))
            doc.Close(False)
    finally:
        word.Quit()
    return out


def _count_word_powershell(paths):
    import subprocess
    import tempfile
    ps = r"""
$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
try { foreach ($p in $args) {
  $d = $w.Documents.Open($p, $false, $true, $false); $n = $d.ComputeStatistics(2)
  $r = $d.Content; $r.Collapse(0); $s = $d.Sections.Item(1).PageSetup
  $f = ($r.Information(6) - $s.TopMargin) / ($s.PageHeight - $s.TopMargin - $s.BottomMargin)
  "$n $f"; $d.Close($false) } } finally { $w.Quit() }
"""
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8") as f:
        f.write(ps)
    try:
        res = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", f.name,
                              *paths], capture_output=True, text=True, timeout=600)
    except (OSError, subprocess.TimeoutExpired):
        return None
    finally:
        Path(f.name).unlink(missing_ok=True)
    rows = [ln.split() for ln in res.stdout.splitlines() if ln.strip()]
    if len(rows) != len(paths):
        return None
    return [(int(n), max(0.0, min(1.0, float(fl.replace(",", "."))))) for n, fl in rows]


def _count_soffice(paths):
    import shutil
    import subprocess
    import tempfile
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    mac = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if not exe and Path(mac).exists():
        exe = mac
    if not exe:
        return None
    out = []
    with tempfile.TemporaryDirectory() as tmp:
        try:
            subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", tmp, *paths],
                           capture_output=True, timeout=600)
        except (OSError, subprocess.TimeoutExpired):
            return None
        for p in paths:
            pdf = Path(tmp) / (Path(p).stem + ".pdf")
            if not pdf.exists():
                return None
            data = pdf.read_bytes()
            out.append((len(re.findall(rb"/Type\s*/Page[^s]", data)), None))
    return out


def render_fit(t, content, out, pages):
    """Render at every scale level, keep the loosest one that ends on `pages` pages."""
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        files, seen = [], set()
        for s in FIT_LEVELS:
            ts = scaled(t, s)
            sig = json.dumps(ts, sort_keys=True)
            if sig in seen:
                continue
            seen.add(sig)
            f = Path(tmp) / f"fit_{s:.2f}.docx"
            Renderer(ts).render(content).save(f)
            files.append((s, ts, f))
        counts = count_pages([f for _, _, f in files])
        if counts is None:
            Renderer(t).render(content).save(out)
            return "FIT skipped: no layout engine found (needs Microsoft Word on Windows, or LibreOffice); rendered at the template's normal size"
        fitting = [(s, ts, f, c) for (s, ts, f), c in zip(files, counts) if c[0] <= pages]
        if fitting:
            s, ts, f, (n, fill) = fitting[0]  # loosest level that fits
            note = ""
        else:
            s, ts, f, (n, fill) = files[-1] + (counts[-1],)
            note = f"; WARNING the content needs more than {pages} pages even at the tightest level"
        out.write_bytes(f.read_bytes())
    full = f", last page {round(fill * 100)}% full" if fill is not None else ""
    return (f"fit: {n} page(s){full}; scale {s:.2f}, body {ts['body_pt']} pt, "
            f"margins {ts['margins_in']} in{note}")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("templates", help="list the available templates")
    r = sub.add_parser("render", help="render content JSON to DOCX")
    r.add_argument("content")
    r.add_argument("out")
    r.add_argument("--template", default="classic")
    r.add_argument("--pages", type=int, default=3,
                   help="fit the resume onto this many pages (default 3; 0 = no fitting)")
    a = ap.parse_args(argv)

    templates = load_templates()
    if a.cmd == "templates":
        for key, t in templates.items():
            ats = "ATS-safe" if t.get("ats_safe", True) else "ATS caution"
            print(f"{key:<10} {t['label']} [{ats}]: {t['description']}")
        return 0

    key = a.template.strip().lower()
    if key not in templates:
        print(f"ERROR unknown template '{a.template}'. Choose one of: {', '.join(templates)}")
        return 2
    content = json.loads(Path(a.content).read_text(encoding="utf-8"))
    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    if a.pages > 0:
        note = render_fit(templates[key], content, out, a.pages)
    else:
        Renderer(templates[key]).render(content).save(out)
        note = "no fitting"
    print(f"saved {out} (template: {key}); {note}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
