"""Tests for the commands (Claude Code skills) in application-workflow/.claude/skills/<name>/SKILL.md (T19): every command exists, has a description, and points
only to prompts, tools, tracker_tool subcommands, Module 04 sections and other commands that really exist; no name clashes with a
Claude Code built-in command. Offline; reads files only.

  python Tools/tests/test_commands.py      prints "all tests passed"
"""
import os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
TOOLS = os.path.dirname(HERE)
FOLDER = os.path.dirname(TOOLS)
SKILLS = os.path.join(FOLDER, ".claude", "skills")
sys.path.insert(0, TOOLS)
import tracker_tool as T

FAIL = []
EXPECTED = ["setup", "profile", "get-jobs", "check-setup", "call-prep", "build-resume", "email-draft", "email-follow-up", "email-sent", "my-progress"]
BUILT_IN = {"add-dir", "agents", "bug", "clear", "compact", "config", "context", "cost", "doctor", "exit", "export", "help", "hooks", "ide",
            "init", "install-github-app", "login", "logout", "mcp", "memory", "model", "output-style", "permissions", "plugin", "pr-comments",
            "release-notes", "resume", "review", "rewind", "security-review", "status", "statusline", "terminal-setup", "upgrade", "usage", "vim"}


def check(name, ok, detail=""):
    if not ok:
        FAIL.append(name + (": " + str(detail)[:400] if detail else ""))


def read(name):
    return open(os.path.join(SKILLS, name, "SKILL.md"), encoding="utf-8").read()


def subcommands():
    """tracker_tool.py's subcommand names, from its argument parser."""
    import argparse
    names = set()
    orig = argparse.ArgumentParser.parse_args
    try:
        def grab(self, args=None, namespace=None):
            for act in self._actions:
                if isinstance(act, argparse._SubParsersAction):
                    names.update(act.choices)
            raise SystemExit(0)
        argparse.ArgumentParser.parse_args = grab
        try:
            T.main(["summary", "x"])
        except SystemExit:
            pass
    finally:
        argparse.ArgumentParser.parse_args = orig
    return names


def test_commands():
    have = sorted(d for d in os.listdir(SKILLS) if os.path.isfile(os.path.join(SKILLS, d, "SKILL.md")))
    check("no old .claude/commands folder (a skill and a command of the same name clash)",
          not os.path.exists(os.path.join(FOLDER, ".claude", "commands")))
    check("exactly the expected commands", have == sorted(EXPECTED), have)
    subs = subcommands()
    check("tracker_tool subcommands found", {"import-feed", "summary", "report", "unsent"} <= subs, sorted(subs))
    module04 = open(os.path.join(FOLDER, "4 - Recruiter Follow-up Email.txt"), encoding="utf-8").read()
    for name in have:
        text = read(name)
        check(f"/{name} is not a built-in command", name not in BUILT_IN)
        check(f"/{name} name is lowercase letters, digits and dashes", re.fullmatch(r"[a-z0-9]+(-[a-z0-9]+)*", name) is not None)
        fm = re.match(r"^---\n(.*?)\n---\n", text, re.S)
        check(f"/{name} has front matter", fm is not None)
        if fm:
            check(f"/{name} has a description", re.search(r"^description: \S.{10,}", fm.group(1), re.M) is not None, fm.group(1))
        for ref in re.findall(r"`([^`]+\.(?:txt|md|py))`", text):  # files named in backticks
            if ref in ("tracker_path.txt", "feed_path.txt"):  # each candidate's own files, made by setup and get-jobs
                continue
            check(f"/{name}: {ref} exists", os.path.exists(os.path.join(FOLDER, ref)), ref)
        for tool in re.findall(r"Tools/([a-z_]+\.py)", text):
            check(f"/{name}: Tools/{tool} exists", os.path.isfile(os.path.join(TOOLS, tool)), tool)
        for sub in re.findall(r"tracker_tool\.py ([a-z][a-z-]+)", text):
            check(f"/{name}: tracker_tool.py {sub} exists", sub in subs, sub)
        for sec in re.findall(r"Section (10[A-C])", text):
            check(f"/{name}: Module 04 Section {sec} exists", re.search(r"^" + sec + r"\. ", module04, re.M) is not None, sec)
        for other in re.findall(r"(?<![\w/])/([a-z][a-z0-9-]+)\b", text):
            if other in ("hr",):  # "$/hr" style text is not a command
                continue
            check(f"/{name} mentions /{other}, which exists", other in EXPECTED, other)


# Writes Claude may run without asking (approved 2026-11-25): each tool protects the tracker or file itself
# (status only forward, first dates kept, no Job ID added twice, no file overwritten except the resume being rendered).
SAFE_WRITES = {"tracker_tool.py": {"set", "import-feed"}, "resume_render.py": {"render"},
               "resume_path.py": {"path", "reuse", "list", "evidence set", "sync"}}
# Always ask, even in auto mode: the candidate's decisions and one-time changes.
MUST_ASK = ["tracker_tool.py " + s + " *" for s in ("edit", "skip", "reopen", "prune-excluded", "init", "add-columns", "extend-ranges")] \
    + ["tracker_tool.py set *--confirmed-closed*", "setup_check.py *--fix*"]
READS = {"tracker_tool.py": {"get", "list", "related", "followups", "drafted-notes", "summary", "report", "unsent", "skipped",
                             "check-contacts", "next-id", "last-run", "recheck-list"},
         "setup_check.py": {""}, "profile_check.py": {""}, "resume_check.py": {""}, "email_check.py": {"brief", "check", "batch"},
         "resume_render.py": {"templates"}, "resume_path.py": {"categories", "find", "evidence get", "evidence list"}}


def test_settings():
    """.claude/settings.json: Claude runs the read-only scripts, the approved safe writes and the skills without asking; every
    other write is not allowed, and the candidate's decisions (edit, skip, reopen...) always ask."""
    import json
    p = json.load(open(os.path.join(FOLDER, ".claude", "settings.json"), encoding="utf-8"))["permissions"]
    subs = subcommands()
    for rule in p["allow"]:
        m = re.fullmatch(r"Bash\(python3? Tools/([a-z_]+\.py)(?: ([a-z-]+(?: [a-z]+)?))? \*\)", rule)
        s = re.fullmatch(r"Skill\(([a-z-]+)(?: \*)?\)", rule)
        check(f"allow rule has a known form: {rule}", bool(m or s), rule)
        if m:
            script, sub = m.group(1), m.group(2) or ""
            check(f"{rule}: Tools/{script} exists", os.path.isfile(os.path.join(TOOLS, script)), script)
            check(f"{rule}: a read or an approved safe write", sub in READS.get(script, set()) | SAFE_WRITES.get(script, set()), sub)
            if script == "tracker_tool.py":
                check(f"{rule}: tracker_tool.py {sub} exists", sub in subs, sub)
        if s:
            check(f"{rule}: skill exists", s.group(1) in EXPECTED, s.group(1))
    for py in ("python", "python3"):
        for script, names in list(READS.items()) + list(SAFE_WRITES.items()):
            for sub in names:
                rule = f"Bash({py} Tools/{script}{' ' + sub if sub else ''} *)"
                check(f"{rule} is allowed", rule in p["allow"])
        for r in MUST_ASK:
            check(f"{py} Tools/{r} asks", f"Bash({py} Tools/{r})" in p.get("ask", []))
    check("every skill is allowed", all(f"Skill({n})" in p["allow"] for n in EXPECTED))
    check("nothing denied by accident", not p.get("deny"), p.get("deny"))


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
    if FAIL:
        print("FAILED:\n  " + "\n  ".join(FAIL)); sys.exit(1)
    print("all tests passed")
