# NOTE: documents_dir(), resumes_root(), key() and the job-folder search below are copies from application-workflow/Tools/resume_path.py,
# so the two flows stay independent (design section 5). If the Resumes folder layout or those helpers change there, make the same change here.
"""Finds the resume file to attach for a job (design section 7).

Resume Path is the primary source; the Resume Filename search is the fallback. Never guesses: a job folder that holds one resume whose
name does not match is not used.

  python resume_find.py <JOB-####> [--path "<Resume Path cell>"] [--filename "<Resume Filename cell>"] [--root <Resumes folder>]
"""
import argparse
import os
import re
import sys
from pathlib import Path

from dice_apply_common import out

ROOT_NAME = "Job Application Workflow"
RESUME_EXT = (".docx", ".pdf", ".doc", ".rtf", ".txt")
DEFAULT_MAX_BYTES = 2 * 1024 * 1024  # Dice shows 2MB for the cover letter; assumed for the resume until verified on the page


def key(s):
    return re.sub(r"[^a-z0-9#+.]", "", str(s or "").lower())


def documents_dir():
    """The real Documents folder (Windows may redirect it to OneDrive; Mac is ~/Documents)."""
    if sys.platform.startswith("win"):
        try:
            import ctypes
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


# ---------- reading the two cells ----------
_QUOTES = "\"'“”‘’ \t"
_UNKNOWN = re.compile(r"^unknown\s*\(\s*saved\s+as\s+(.+?)\s*\)\s*$", re.I)


def _strip(v):
    return str(v or "").strip().strip(_QUOTES).strip()


def _looks_like_path(s):
    return bool(re.match(r"^([A-Za-z]:[\\/]|\\\\|/|~[\\/])", s))


def parse_path_cell(v):
    """-> (kind, value): ('empty', ''), ('link', url), ('name', file name), ('path', Path)."""
    s = _strip(v)
    if not s:
        return "empty", ""
    if re.match(r"^https?://", s, re.I):
        return "link", s
    m = _UNKNOWN.match(s)
    if m:
        return "name", _strip(m.group(1))
    if _looks_like_path(s):
        return "path", Path(os.path.expanduser(s.replace("\\", "/") if not sys.platform.startswith("win") else s))
    return "name", s  # a bare file name typed into the path column


def parse_filename_cell(v):
    """-> (kind, value): ('empty', ''), ('name', file name), ('path', Path) when a full path was pasted into this column."""
    s = _strip(v)
    if not s:
        return "empty", ""
    if _looks_like_path(s):
        return "path", Path(os.path.expanduser(s))
    return "name", s


def _same_name(a, b):
    """File names match when equal ignoring case/punctuation, or when one has no extension and the stems match."""
    a, b = os.path.basename(a), os.path.basename(b)
    if key(a) == key(b):
        return True
    sa, ea = os.path.splitext(a)
    sb, eb = os.path.splitext(b)
    ea = ea.lower() if ea.lower() in RESUME_EXT else ""
    eb = eb.lower() if eb.lower() in RESUME_EXT else ""
    if not ea or not eb:
        return key(sa if ea else a) == key(sb if eb else b)
    return False


# ---------- the Resumes folder index ----------
class ResumeIndex:
    """Walks the Resumes folder once and maps each job folder (JOB-#### ...) to its resume files."""

    def __init__(self, root=None):
        self.root = Path(root or resumes_root())
        self.by_job = {}
        if not self.root.is_dir():
            return
        for dp, _dirs, files in os.walk(self.root):
            base = os.path.basename(dp).upper()
            m = re.match(r"^(JOB-\d+)(\s|$)", base)
            if not m:
                continue
            for f in files:
                if f.lower().endswith(RESUME_EXT) and not f.startswith("~$"):
                    self.by_job.setdefault(m.group(1), []).append(Path(dp) / f)

    def find(self, job_id, filename):
        """Files in the job's folders whose name matches filename (the fallback search)."""
        hits = [p for p in self.by_job.get(job_id.upper(), []) if _same_name(p.name, filename)]
        if not hits and "." not in os.path.basename(filename):
            hits = [p for p in self.by_job.get(job_id.upper(), []) if key(p.stem) == key(filename)]
        return sorted(set(hits))


def tech_of(path, root=None):
    """(Technology, Specialty or '') from Resumes/<Technology>/<Specialty>/<JOB-#### ...>/<file>, or (None, None) when unknown."""
    if not path:
        return None, None
    parts = list(Path(path).parts)
    rootp = Path(root or resumes_root())
    try:
        rel = list(Path(path).resolve().relative_to(rootp.resolve()).parts)
    except (ValueError, OSError):
        idx = max((i for i, p in enumerate(parts) if p.lower() == "resumes"), default=None)
        if idx is None:
            return None, None
        rel = parts[idx + 1:]
    if len(rel) < 3:
        return None, None
    folders = rel[:-1]
    job_i = next((i for i, p in enumerate(folders) if re.match(r"^JOB-\d+", p, re.I)), None)
    if job_i is None or job_i == 0:
        return None, None
    tech = folders[0]
    spec = folders[1] if job_i >= 2 else ""
    return tech, spec


# ---------- resolution ----------
class Resolution:
    def __init__(self, file=None, via="", code="", reason=""):
        self.file, self.via, self.code, self.reason = file, via, code, reason

    @property
    def ok(self):
        return self.file is not None and not self.code

    def __repr__(self):
        return f"Resolution(file={self.file}, via={self.via!r}, code={self.code!r}, reason={self.reason!r})"


def _check_file(p, max_bytes):
    if p.name.startswith("~$"):
        return "RESUME_NOT_FOUND", "the path is a Word lock file (~$)"
    if p.suffix.lower() not in RESUME_EXT:
        return "RESUME_NOT_FOUND", f"{p.suffix or 'no extension'} is not a resume type Dice accepts (.pdf .doc .docx .txt .rtf)"
    try:
        size = p.stat().st_size
    except OSError:
        return "RESUME_NOT_FOUND", "file not readable"
    if size > max_bytes:
        return "RESUME_TOO_LARGE", f"{size / 1048576:.1f} MB is over the {max_bytes / 1048576:.0f} MB limit"
    return "", ""


def resolve(job_id, path_cell, filename_cell, index, override=None, max_bytes=DEFAULT_MAX_BYTES):
    """Design section 7: Resume Path first, then the Resume Filename search. Returns a Resolution."""
    if override:
        p = Path(override).expanduser()
        if not p.is_file():
            return Resolution(code="RESUME_NOT_FOUND", reason=f"--resume file not found: {p}")
        code, why = _check_file(p, max_bytes)
        return Resolution(p, "--resume", code, why)

    pk, pv = parse_path_cell(path_cell)
    fk, fv = parse_filename_cell(filename_cell)
    if pk == "empty" and fk == "empty":
        return Resolution(code="NO_RESUME", reason="Resume Path and Resume Filename are both empty")

    # the name the Filename column asks for (a full path pasted there counts as a path to try, plus its name)
    fname = fv if fk == "name" else (fv.name if fk == "path" else "")
    path_name = pv.name if pk == "path" and pv.suffix else (pv if pk == "name" else "")
    if fname and path_name and not _same_name(path_name, fname):
        return Resolution(code="RESUME_MISMATCH", reason=f"Resume Path names {path_name} but Resume Filename is {fname}")

    reasons = []
    # 1. Resume Path (primary)
    if pk == "path":
        if pv.is_file():
            code, why = _check_file(pv, max_bytes)
            return Resolution(pv, "path", code, why)
        if pv.is_dir():
            files = [f for f in pv.iterdir() if f.is_file() and f.suffix.lower() in RESUME_EXT and not f.name.startswith("~$")]
            if fname:
                files = [f for f in files if _same_name(f.name, fname)]
            if len(files) == 1:
                code, why = _check_file(files[0], max_bytes)
                return Resolution(files[0], "path folder", code, why)
            if len(files) > 1:
                return Resolution(code="RESUME_AMBIGUOUS", reason="the Resume Path folder holds " + ", ".join(f.name for f in sorted(files)))
            reasons.append("no matching resume in the Resume Path folder")
        else:
            reasons.append("file not at path")
    elif pk == "link":
        reasons.append("path is a link")
    if fk == "path":
        if fv.is_file():
            code, why = _check_file(fv, max_bytes)
            return Resolution(fv, "path (in Resume Filename)", code, why)
        reasons.append("file not at the path in Resume Filename")

    # 2. Resume Filename search (fallback)
    name = fname or (path_name if isinstance(path_name, str) else "") or (pv.name if pk == "path" and pv.suffix else "")
    if not name:
        return Resolution(code="RESUME_NOT_FOUND", reason="; ".join(reasons) or "no file name to search for")
    hits = index.find(job_id, name)
    if len(hits) == 1:
        code, why = _check_file(hits[0], max_bytes)
        return Resolution(hits[0], "name search", code, why)
    if len(hits) > 1:
        return Resolution(code="RESUME_AMBIGUOUS", reason="name search found " + ", ".join(str(h) for h in hits))
    reasons.append("name not found in Resumes")
    return Resolution(code="RESUME_NOT_FOUND", reason="; ".join(reasons))


def main(argv):
    p = argparse.ArgumentParser(prog="resume_find.py", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("job")
    p.add_argument("--path", default="")
    p.add_argument("--filename", default="")
    p.add_argument("--root")
    a = p.parse_args(argv)
    r = resolve(a.job.upper(), a.path, a.filename, ResumeIndex(a.root))
    out(f"{r.code or 'OK'}\t{r.via or '-'}\t{r.file or '-'}\t{r.reason}")
    return 0 if r.ok else 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
