"""Rebuild the sample resumes in the "Resume Templates" folder: one PDF per template, made from the made-up
candidate in "Resume Templates/sample_content.json", plus an overview of every template's first page.

Run it after adding or changing a template in Tools/resume_templates.json (python3 on a Mac):
  python Tools/resume_samples.py

It needs Microsoft Word (Windows) or LibreOffice to make PDFs. The overview also needs PyMuPDF
(pip install pymupdf); without it the overview is skipped. Old sample PDFs are replaced.
"""
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import resume_render as rr  # noqa: E402

FOLDER = Path(__file__).resolve().parent.parent / "Resume Templates"
CONTENT = FOLDER / "sample_content.json"
OVERVIEW = "00 - All templates at a glance.pdf"


def overview(items, out):
    """A letter-size sheet with the first page of each sample, four to a page, labelled."""
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf
        except ImportError:
            return False
    sheet = pymupdf.open()
    W, H = 612, 792
    margin, gap, label_h = 30, 18, 34
    cw = (W - 2 * margin - gap) / 2
    ch = cw * 11 / 8.5
    for i, (label, note, pdf) in enumerate(items):
        if i % 4 == 0:
            page = sheet.new_page(width=W, height=H)
        col, row = i % 2, (i % 4) // 2
        x = margin + col * (cw + gap)
        y = margin + row * (ch + label_h + gap)
        page.insert_text((x, y + 12), label, fontsize=12, fontname="hebo")
        page.insert_text((x, y + 26), note, fontsize=8, fontname="helv", color=(0.35, 0.35, 0.35))
        rect = pymupdf.Rect(x, y + label_h, x + cw, y + label_h + ch)
        with pymupdf.open(pdf) as src:
            page.show_pdf_page(rect, src, 0)
        page.draw_rect(rect, color=(0.75, 0.75, 0.75), width=0.6)
    sheet.save(out)
    return True


def main():
    templates = rr.load_templates()
    content = rr.json.loads(CONTENT.read_text(encoding="utf-8"))
    with tempfile.TemporaryDirectory() as tmp:
        docs, made = [], []
        for n, (key, t) in enumerate(templates.items(), 1):
            name = f"{n:02d} - {t['label']} ({key})"
            d = Path(tmp) / f"{name}.docx"
            note = rr.render_fit(t, content, d, 3, key)
            print(f"{name}: {note}")
            docs.append(d)
            made.append((name, key, t))
        pdfs = rr.to_pdf(docs, tmp)
        if pdfs is None:
            print("ERROR could not make PDFs: Microsoft Word (Windows) or LibreOffice is needed")
            return 1
        for old in FOLDER.glob("*.pdf"):
            old.unlink()
        items = []
        for (name, key, t), pdf in zip(made, pdfs):
            dest = FOLDER / f"{name}.pdf"
            shutil.copy(pdf, dest)
            ats = "ATS-safe" if t.get("ats_safe", True) else "ATS caution: uses a table"
            items.append((f"{t['label']}  ({key})", ats, str(dest)))
        if overview(items, FOLDER / OVERVIEW):
            print(f"saved {OVERVIEW}")
        else:
            print("overview skipped: PyMuPDF is not installed (pip install pymupdf)")
    print(f"samples saved in {FOLDER}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
