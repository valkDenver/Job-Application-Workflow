"""Fails when a file with personal data is about to be committed from dice-auto-apply/ (design section 14). Run before every commit.

Usage:  python Tools/tests/check_no_personal_data.py
Looks at every file git would add or has changed in dice-auto-apply/ (tracked, staged or untracked, ignored files excluded) and refuses
tracker paths, test data, spreadsheets, resumes, run output and HTML outside Tools/tests/fixtures/. The shared settings file
Tools/dice_apply_config.json is committed on purpose, but must never hold screening answers (they live outside the repo).
"""
import re
import subprocess
import sys
from pathlib import Path

FOLDER = Path(__file__).resolve().parents[2]
BAD = [
    (r"(^|/)tracker_path\.txt$", "your tracker location"),
    (r"(^|/)screening_answers\.json$", "your screening answers"),
    (r"(^|/)Tools/testdata/", "test copies of the tracker and test resumes"),
    (r"\.(xlsx|xlsm|xls|csv|docx|doc|pdf|rtf)$", "a spreadsheet or resume"),
    (r"(^|/)runs/|\.png$|\.jsonl$", "round output (screenshots, results)"),
    (r"(^|/)browser-profile/", "the Dice login session"),
]
HTML_OK = re.compile(r"(^|/)Tools/tests/fixtures/[^/]+\.html$")


def main():
    r = subprocess.run(["git", "status", "--porcelain", "--untracked-files=all", "--", str(FOLDER)], capture_output=True, text=True,
                       cwd=FOLDER)
    if r.returncode != 0:
        print("Could not run git status: " + r.stderr.strip())
        return 2
    problems = []
    for line in r.stdout.splitlines():
        path = line[3:].strip().strip('"').replace("\\", "/")
        if " -> " in path:
            path = path.split(" -> ", 1)[1]
        if line[:2].strip() == "D":
            continue
        for pat, why in BAD:
            if re.search(pat, path, re.I):
                problems.append(f"{path}: {why}")
                break
        else:
            if path.lower().endswith((".html", ".htm")) and not HTML_OK.search(path):
                problems.append(f"{path}: HTML outside Tools/tests/fixtures (could be a saved Dice page)")
    cfg = FOLDER / "Tools" / "dice_apply_config.json"
    text = cfg.read_text(encoding="utf-8") if cfg.exists() else ""
    if re.search(r'"(screening_answers|answers)"\s*:\s*\{\s*"', text):
        problems.append("Tools/dice_apply_config.json: holds screening answers; they belong in screening_answers.json outside the repo")
    for key in ("resumes_root", "runtime_dir"):
        if re.search(r'"%s"\s*:\s*"' % key, text):
            problems.append(f"Tools/dice_apply_config.json: {key} holds a personal path; keep it null (personal values go in settings.json)")
    if problems:
        print("Personal data could be committed. Remove these from the commit (or add them to .gitignore):")
        for p in problems:
            print("  " + p)
        return 1
    print("OK: nothing personal is about to be committed from dice-auto-apply/.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
