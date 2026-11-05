"""Totals the tokens of a Claude Code session and of each helper (subagent) it launched, from the session logs, so a batch
resume run can be compared with an earlier one. It only reads the logs; it never changes anything.

Usage:  python token_usage.py <session.jsonl>         one session (the main conversation's log file)
        python token_usage.py <project folder>        the newest session in that folder that launched helpers
        python token_usage.py --recent [N]            the N newest sessions (default 10) that launched helpers, any project
        options: --resumes N   resumes made in the session, when the helpers' replies do not carry BATCH_RUN.md result lines
                 --calls       also count the helpers' tool calls by tool, and Bash calls by the script they run
        (on a Mac: python3 token_usage.py ...)

The logs are in ~/.claude/projects/<project>/<session id>.jsonl, and each helper's log in <session id>/subagents/.
Per helper it prints: calls, uncached input, cache write, cache read, output, total, the largest context of any one call
(peak), the context of its first call (first: the base every call carries, before the helper has read
anything: system prompt, tool definitions and the launch text), whether the helper's context was summarized (compacted: its earlier text, Module 03 included, was replaced by a
summary), and the outcomes counted from its BATCH_RUN.md result lines (JOB-#### | BUILT | ...).
Total = all four token kinds added together, the same measure as the token brief. Cache reads cost about a tenth of
normal input, so the breakdown matters when comparing cost rather than volume.
"""
import argparse, glob, json, os, re, sys

ROOT = os.path.join(os.environ.get("CLAUDE_CONFIG_DIR") or os.path.join(os.path.expanduser("~"), ".claude"), "projects")
SCRIPT = re.compile(r"([\w.-]+\.(?:py|sh))\b")
LINE = re.compile(r"^\s*(JOB-\d+)\s*\|\s*(BUILT|REUSED|GATE-STOP|ERROR)\b", re.I | re.M)


def read_log(path):
    """Usage per API call (deduplicated by message id), peak context, compaction, and the last assistant text."""
    calls, compacted, last_text, tools = {}, False, "", {}
    with open(path, encoding="utf-8") as f:
        for line in f:
            try:
                d = json.loads(line)
            except ValueError:
                continue
            if not isinstance(d, dict):
                continue
            if d.get("isCompactSummary") or d.get("subtype") == "compact_boundary":
                compacted = True
            if d.get("type") != "assistant":
                continue
            m = d.get("message") or {}
            u = m.get("usage")
            if u:
                calls[m.get("id") or d.get("requestId") or len(calls)] = u
            for c in m.get("content") or []:
                if isinstance(c, dict) and c.get("type") == "tool_use":  # a streamed call is logged more than once; the last copy has the full input
                    key = c.get("name", "?")
                    if key == "Bash":
                        sc = SCRIPT.search(str((c.get("input") or {}).get("command", "")))
                        key = "Bash " + (sc.group(1) if sc else "(other)")
                    tools[c.get("id")] = key
                    if key == "SubagentHandback":  # a helper may hand its final reply back through this tool instead of as text
                        last_text = str((c.get("input") or {}).get("message", "")) or last_text
            texts = [c.get("text", "") for c in m.get("content") or [] if isinstance(c, dict) and c.get("type") == "text"]
            if any(t.strip() for t in texts):
                last_text = "\n".join(texts)
    tot = {"calls": len(calls), "input": 0, "cache_write": 0, "cache_read": 0, "output": 0, "peak": 0, "first": 0}
    for u in calls.values():
        i, cw, cr = u.get("input_tokens", 0), u.get("cache_creation_input_tokens", 0), u.get("cache_read_input_tokens", 0)
        tot["input"] += i
        tot["cache_write"] += cw
        tot["cache_read"] += cr
        tot["output"] += u.get("output_tokens", 0)
        tot["peak"] = max(tot["peak"], i + cw + cr)
        if not tot["first"]:
            tot["first"] = i + cw + cr
    tot["total"] = tot["input"] + tot["cache_write"] + tot["cache_read"] + tot["output"]
    tot["compacted"] = compacted
    tot["tools"] = list(tools.values())
    tot["outcomes"] = {}
    for _, kind in LINE.findall(last_text):
        tot["outcomes"][kind.upper()] = tot["outcomes"].get(kind.upper(), 0) + 1
    return tot


def helpers_of(session):
    folder = os.path.join(os.path.splitext(session)[0], "subagents")
    return sorted(glob.glob(os.path.join(folder, "*.jsonl")), key=os.path.getmtime)


def description(agent_log):
    try:
        with open(os.path.splitext(agent_log)[0] + ".meta.json", encoding="utf-8") as f:
            return json.load(f).get("description", "")
    except (OSError, ValueError):
        return ""


def k(n):
    return f"{n / 1000:,.0f}K" if n >= 1000 else str(n)


def report(session, resumes=None, calls=False):
    print(f"Session: {session}")
    rows, made = [], {}
    for h in helpers_of(session):
        t = read_log(h)
        rows.append((os.path.basename(h)[:-6], description(h), t))
        for kind, n in t["outcomes"].items():
            made[kind] = made.get(kind, 0) + n
    main = read_log(session)
    print(f"{'helper':<24} {'calls':>5} {'input':>7} {'c.write':>8} {'c.read':>8} {'output':>7} {'total':>8} {'first':>6} {'peak':>7}  compacted  outcomes / description")
    for name, desc, t in rows:
        oc = ", ".join(f"{n} {kind}" for kind, n in sorted(t["outcomes"].items())) or desc[:40]
        print(f"{name[:24]:<24} {t['calls']:>5} {k(t['input']):>7} {k(t['cache_write']):>8} {k(t['cache_read']):>8} "
              f"{k(t['output']):>7} {k(t['total']):>8} {k(t['first']):>6} {k(t['peak']):>7}  {'YES' if t['compacted'] else 'no':<9}  {oc}")
    htot = sum(t["total"] for _, _, t in rows)
    print(f"\nHelpers: {len(rows)}, {k(htot)} tokens in total"
          + (f"; {sum(t['calls'] for _, _, t in rows)} calls; first-call context {k(min(t['first'] for _, _, t in rows))}-{k(max(t['first'] for _, _, t in rows))}"
             f"; largest peak context {k(max(t['peak'] for _, _, t in rows))}" if rows else "")
          + (f"; {sum(1 for _, _, t in rows if t['compacted'])} compacted" if any(t["compacted"] for _, _, t in rows) else ""))
    print(f"Main session: {k(main['total'])} tokens ({main['calls']} calls, peak context {k(main['peak'])})")
    if calls and rows:
        count = {}
        for _, _, t in rows:
            for key in t["tools"]:
                count[key] = count.get(key, 0) + 1
        print("Helper tool calls by kind: " + ", ".join(f"{key} {n}" for key, n in sorted(count.items(), key=lambda kv: -kv[1])))
    if made:
        print("Outcomes from result lines: " + ", ".join(f"{n} {kind}" for kind, n in sorted(made.items())))
    count = resumes if resumes is not None else made.get("BUILT", 0) + made.get("REUSED", 0)
    if count:
        print(f"Per resume ({count} built or reused): helpers {k(htot / count)}, helpers + main session {k((htot + main['total']) / count)}")
    else:
        print("Per resume: unknown (no BATCH_RUN.md result lines found; pass --resumes N)")


def sessions_with_helpers(folder):
    found = [s for s in glob.glob(os.path.join(folder, "*.jsonl")) if helpers_of(s)]
    return sorted(found, key=os.path.getmtime, reverse=True)


def main(argv):
    ap = argparse.ArgumentParser(prog="token_usage.py", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("target", nargs="?")
    ap.add_argument("--recent", nargs="?", const=10, type=int)
    ap.add_argument("--resumes", type=int)
    ap.add_argument("--calls", action="store_true")
    a = ap.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        pass
    if a.recent is not None:
        if not os.path.isdir(ROOT):
            print(f"No Claude Code logs found in {ROOT}.")
            return 1
        found = []
        for p in glob.glob(os.path.join(ROOT, "*")):
            found += sessions_with_helpers(p)
        if not found:
            print(f"No session in {ROOT} launched helpers.")
        for s in sorted(found, key=os.path.getmtime, reverse=True)[:a.recent or None]:
            print(f"{len(helpers_of(s)):>3} helpers  {s}")
        return 0
    if not a.target:
        ap.error("give a session .jsonl file, a project folder, or --recent")
    target = a.target
    if os.path.isdir(target):
        found = sessions_with_helpers(target)
        if not found:
            print(f"No session in {target} launched helpers.")
            return 1
        target = found[0]
    if not os.path.isfile(target):
        print(f"File not found: {target}")
        return 1
    report(target, a.resumes, a.calls)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
