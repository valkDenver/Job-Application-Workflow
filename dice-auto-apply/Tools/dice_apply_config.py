"""Settings: one shared file, Tools/dice_apply_config.json (committed; the same defaults for every candidate).

A candidate's own changes never touch that file (no git conflicts, no personal paths committed): `config set` saves them in
<runtime folder>/settings.json, and screening answers in <runtime folder>/screening_answers.json. The runtime folder is next to the
Dice login, outside the repo (~/.job-workflow/dice-apply/ on a Mac, %USERPROFILE%\\.job-workflow\\dice-apply\\ on Windows).
load() = the shared file with the personal settings on top. Every value is validated by `config set` / `config answer`.
"""
import os
import re
from pathlib import Path

from dice_apply_common import Usage, label_key, read_json, write_json
from dice_apply_platform import HERE, runtime_dir

CONFIG = Path(os.environ.get("DICE_APPLY_CONFIG") or HERE / "dice_apply_config.json")  # the env var is for tests
DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def shared():
    cfg = read_json(CONFIG)
    if not isinstance(cfg, dict) or not cfg.get("paces"):
        raise Usage(f"The settings file is missing or damaged: {CONFIG}. Restore it from git (git checkout -- Tools/dice_apply_config.json).")
    return cfg


def personal_path():
    return runtime_dir(shared()) / "settings.json"  # runtime_dir comes from the shared file only, never from a personal setting


def personal():
    return read_json(personal_path(), {}) or {}


def load():
    cfg = shared()
    cfg.update({k: v for k, v in personal().items() if k in SETTABLE})
    return cfg


def _int(lo, hi, allow_none=False):
    def f(v):
        if allow_none and v.strip().lower() in ("", "none", "null", "off"):
            return None
        try:
            n = int(v)
        except ValueError:
            raise Usage(f"expected a whole number, got {v!r}")
        if not lo <= n <= hi:
            raise Usage(f"must be between {lo} and {hi}")
        return n
    return f


def _days(v):
    parts = [p.strip()[:3].title() for p in re.split(r"[,\s]+", v) if p.strip()]
    if v.strip().lower() in ("weekdays", "mon-fri"):
        parts = DAYS[:5]
    if v.strip().lower() in ("all", "every day", "everyday"):
        parts = list(DAYS)
    bad = [p for p in parts if p not in DAYS]
    if not parts or bad:
        raise Usage("use day names like Mon,Tue,Wed,Thu,Fri (or 'weekdays' / 'all')")
    return parts


def _hours(v):
    m = re.match(r"^\s*(\d{1,2}):?(\d{2})?\s*-\s*(\d{1,2}):?(\d{2})?\s*$", v)
    if not m:
        raise Usage("use a range like 08:00-16:00")
    a = f"{int(m.group(1)):02d}:{m.group(2) or '00'}"
    b = f"{int(m.group(3)):02d}:{m.group(4) or '00'}"
    if not ("00:00" <= a < b <= "23:59"):
        raise Usage("the start must be before the end, within one day")
    return [a, b]


def _pace(v):
    v = v.strip().lower()
    if v not in (load().get("paces") or {}):
        raise Usage("pace must be brisk, normal or relaxed")
    return v


def _path(v):
    v = v.strip().strip('"')
    return str(Path(v).expanduser()) if v.lower() not in ("", "none", "null") else None


SETTABLE = {
    "max_submits_per_day": _int(1, 1000, allow_none=True),
    "max_per_round": _int(1, 1000, allow_none=True),
    "max_wizard_opens_per_day": _int(1, 2000, allow_none=True),
    "active_days": _days,
    "active_hours": _hours,
    "default_pace": _pace,
    "cooldown_hours": _int(1, 24 * 14),
    "stuck_timeout_sec": _int(5, 120),
    "job_timeout_sec": _int(60, 900),
    "tracker_batch_size": _int(1, 50),
    "resumes_root": _path,
}  # runtime_dir is deliberately not settable: moving it would leave the cooldown, ledger and run lock behind


def set_value(key, value):
    k = key.strip()
    if k not in SETTABLE:
        raise Usage(f'"{k}" cannot be set. Settable: {", ".join(SETTABLE)}')
    try:
        v = SETTABLE[k](value)
    except Usage as e:
        raise Usage(f"{k}: {e}")
    mine = personal()
    mine[k] = v
    write_json(personal_path(), mine)
    return v


# ---------- screening answers (personal: outside the repo) ----------
def answers_path():
    return runtime_dir(shared()) / "screening_answers.json"


def set_answer(question, answer):
    q = label_key(question)
    if not q:
        raise Usage("the question text is empty")
    ans = read_json(answers_path(), {}) or {}
    if answer is None or str(answer).strip() == "":
        ans.pop(q, None)
    else:
        ans[q] = str(answer).strip()
    write_json(answers_path(), ans)
    return q


def answers(cfg=None):
    return {label_key(k): v for k, v in (read_json(answers_path(), {}) or {}).items()}
