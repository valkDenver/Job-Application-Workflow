"""Small helpers shared by the dice-auto-apply tools: the Refused error, plain output, text normalising, JSON-lines files.

No third-party imports, so every other module can use it, including under a Python without Playwright.
"""
import datetime as dt
import json
import os
import re
import sys
import tempfile


class Refused(Exception):
    """A plain-English reason to stop. The CLI prints it as 'REFUSED: <text>' and exits with code 2."""


class Usage(Refused):
    """Bad arguments or config. The CLI exits with code 3."""


def out(s=""):
    sys.stdout.write(str(s) + "\n")
    sys.stdout.flush()


def norm(h):
    return re.sub(r"\s+", " ", str(h or "")).strip().lower()


def text(v):
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def label_key(s):
    """Normalised screening-question label: lower case, no punctuation, single spaces."""
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]", " ", str(s or "").lower())).strip()


def today():
    return dt.date.today().isoformat()


def now_iso():
    return dt.datetime.now().replace(microsecond=0).isoformat(sep=" ")


def read_json(path, default=None):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return default
    except (OSError, ValueError):
        return default


def write_json(path, data):
    d = os.path.dirname(os.path.abspath(path))
    os.makedirs(d, exist_ok=True)
    fd, tmp = tempfile.mkstemp(suffix=".json", dir=d, prefix=".tmp_")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(tmp, path)


def read_jsonl(path):
    rows = []
    try:
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        rows.append(json.loads(line))
                    except ValueError:
                        pass  # a half-written last line after a crash is skipped, never fatal
    except FileNotFoundError:
        pass
    return rows


def append_jsonl(path, obj):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(obj, ensure_ascii=False) + "\n")
        f.flush()
        os.fsync(f.fileno())


def write_jsonl(path, rows):
    d = os.path.dirname(os.path.abspath(path))
    os.makedirs(d, exist_ok=True)
    fd, tmp = tempfile.mkstemp(suffix=".jsonl", dir=d, prefix=".tmp_")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    os.replace(tmp, path)
