"""Helper for Module 04 (Recruiter Email). Standard library only; works on any machine with Python 3.

Usage:
  python email_check.py brief --count N [--date YYYY-MM-DD] [--seed TEXT]
      Hands out a different intro line, sign-off, ask style, email shape, years wording and duty pick
      to each of N emails, so a batch does not read alike. The model writes the actual wording.

  python email_check.py check EMAIL.txt --years "8+ years" --tools "PeopleCode,Application Engine,Integration Broker"
        [--min 90] [--max 120] [--no-resume]
      Checks one email (plain text: greeting, body, sign-off word, first name, email address).
      Fails when the years wording or the tools are missing, the word count is out of range,
      a banned phrase or character is used, or the shape is wrong.
      --years none     skips the years check (the JD gave no years and no total is used).

  python email_check.py batch EMAIL1.txt EMAIL2.txt ...
      Checks that the emails of one run do not share an intro, sign-off, opening words,
      ask wording or a stock reaction phrase.

Exit code 0 = all good, 1 = something to fix (the lines starting with FAIL say what).
"""
import argparse
import datetime
import random
import re
import sys

INTRO_PLAIN = [
    "Hope your week is going well.", "Hope you're having a good week.", "Hope your day is going well.",
    "Hope you're doing well.", "Hope you're having a smooth week.", "Hope all is well on your end.",
    "Hope the week's treating you well.", "Hope you're having a productive week.", "Hope things are going well.",
    "Hope your week's been good so far.", "Hope your day's been a good one.", "Hope the week's been kind so far.",
    "Hope you're having a steady week.", "Hope your week is off to a good start.",
    "Hope things are moving along well on your side.", "Hope the day is treating you well.",
    "Hope you've had a good week so far.", "Hope work is going well.", "Hope you're settling into the week well.",
]
INTRO_MONDAY = ["Happy Monday.", "Happy Monday, hope the week is off to a good start."]
INTRO_FRIDAY = ["Happy Friday.", "Hope your week's been good."]
SIGNOFFS = [
    "Thanks", "Best", "Talk soon", "Appreciate it", "Cheers", "Take care", "Thanks a lot", "All the best",
    "Speak soon", "Much appreciated", "Chat soon", "Have a good one", "Thanks for your time", "Be well",
    "Catch you soon",
]
ASKS = [
    "a plain question: is there a good time this week for a quick call",
    "ask for the best way to connect",
    "ask for 15 minutes this week and say you will make it fit",
    "say you are free for a call whenever suits them",
    "ask whether the role is still open, then offer a call",
    "ask what the next step is on their side",
    "offer to start by email and ask if that is easier",
    "ask for the best time to call",
    "ask them to send a time and say you will confirm right away",
    "ask whether a short conversation fits their schedule this week",
    "ask whether they prefer a call or email to start",
    "invite them to say how they like to start",
    "ask for a short intro call to cover the details",
    "ask whether now is a good moment to talk this week",
]
SHAPES = [
    "intro, ask first, then fit with years, then interest sentence",
    "intro, role mention, fit with years and tools, interest sentence, ask",
    "intro, question opener about the role, fit, ask",
    "intro, why you are writing, fit, interest sentence, ask as a statement",
    "years folded into the opening sentence, tools, interest sentence, ask",
    "intro, interest sentence first, then fit with years, then ask",
    "intro, short role mention, fit paragraph, one-line ask",
]
YEARS_STYLES = [
    "N+ years on X", "been working in X for N+ years", "N+ years of X behind me", "X has been my focus for N+ years",
    "after N years in X", "fold the years into the opening sentence", "N+ years under my belt in X",
    "I have spent N+ years in X",
]

BANNED = [
    "hope this finds you well", "i am writing to", "i wanted to reach out", "i am excited", "i'm excited",
    "great fit", "leverage", "utilize", "proven track record", "dynamic", "passionate", "esteemed", "synergy",
    "do not hesitate", "please find attached", "at your earliest convenience", "look forward to the opportunity",
    "just following up", "bumping this", "circling back", "per my last email", "touching base", "as per",
    "trust everything", "thank you for reading", "thanks for the introduction", "learn more",
]
STOCK_REACTION = ["caught my eye", "stood out", "caught my attention", "jumped out", "stuck with me", "resonated",
                  "stands out to me", "got my attention", "feels familiar", "familiar territory", "right up my alley"]


def words(text):
    return re.findall(r"[A-Za-z0-9'+/#.\-]+", text)


def parse_email(text):
    lines = [l.rstrip() for l in text.strip().splitlines()]
    nonempty = [i for i, l in enumerate(lines) if l.strip()]
    if len(nonempty) < 5:
        return None
    greeting = lines[nonempty[0]].strip()
    email_line = lines[nonempty[-1]].strip()
    name_line = lines[nonempty[-2]].strip()
    signoff_line = lines[nonempty[-3]].strip()
    body_lines = lines[nonempty[0] + 1: nonempty[-3]]
    body = "\n".join(body_lines).strip()
    return greeting, body, signoff_line, name_line, email_line


def sentences(body):
    flat = re.sub(r"\s+", " ", body)
    return [s.strip() for s in re.split(r"(?<=[.?!])\s+", flat) if s.strip()]


def cmd_brief(a):
    d = datetime.date.fromisoformat(a.date) if a.date else datetime.date.today()
    rng = random.Random(f"{d.isoformat()}|{a.seed or ''}")
    wd = d.weekday()
    special = INTRO_MONDAY if wd == 0 else INTRO_FRIDAY if wd == 4 else []
    plain = INTRO_PLAIN[:]
    rng.shuffle(plain)
    intros = (special[:2] if special else []) + plain
    signs, asks, shapes, ystyles = SIGNOFFS[:], ASKS[:], SHAPES[:], YEARS_STYLES[:]
    for lst in (signs, asks, shapes, ystyles):
        rng.shuffle(lst)

    def pick(lst, i):
        return lst[i % len(lst)]

    print(f"Date {d.isoformat()} ({d.strftime('%A')}). One line per email. The model writes the wording.")
    for i in range(a.count):
        seed = rng.randint(1, 99)
        print(f"#{i+1}: intro=\"{pick(intros, i)}\" | sign-off=\"{pick(signs, i)}\" | ask={pick(asks, i)} | "
              f"shape={pick(shapes, i)} | years-wording={pick(ystyles, i)} | duty-seed={seed} "
              f"(JD duty number = (seed mod number of duties) + 1)")
    if a.count > len(SIGNOFFS):
        print(f"NOTE: {a.count} emails but only {len(SIGNOFFS)} sign-offs; repeats appear, so keep intros and asks different.")


def cmd_check(a):
    text = open(a.file, encoding="utf-8").read()
    p = parse_email(text)
    fails = []
    if not p:
        print("FAIL: could not read the email (need greeting, body, sign-off word, first name, email address).")
        return 1
    greeting, body, signoff, name, email = p
    low = text.lower()
    if not re.match(r"^Hi( [A-Z][\w'-]+)?,$", greeting):
        fails.append(f"greeting should be 'Hi Name,' or 'Hi,' (found '{greeting}')")
    if "@" not in email or len(name.split()) != 1 or not signoff.endswith(","):
        fails.append("closing block should be: sign-off word with comma, first name alone, then email address")
    n = len(words(body))
    if n < a.min or n > a.max:
        fails.append(f"body is {n} words; needs {a.min} to {a.max}")
    paras = [b for b in re.split(r"\n\s*\n", body) if b.strip()]
    if not 2 <= len(paras) <= 4:
        fails.append(f"body has {len(paras)} paragraphs; needs 2 to 4")
    if a.years and a.years.lower() != "none":
        m = re.search(r"(\d+)", a.years)
        num = m.group(1) if m else a.years
        if not re.search(rf"\b{re.escape(num)}\s*\+?\s*(?:plus\s+)?years?\b", body, re.I):
            fails.append(f"years claim '{a.years}' is missing from the body")
    elif not a.years:
        fails.append("pass --years (use --years none if the JD gives no years and no total is stated)")
    if a.tools:
        tools = [t.strip() for t in a.tools.split(",") if t.strip()]
        found = [t for t in tools if t.lower() in body.lower()]
        need = min(2, len(tools))
        if len(found) < need:
            fails.append(f"only {len(found)} of the JD tools found in the body ({', '.join(found) or 'none'}); needs at least {need}")
        if len(tools) > 3:
            fails.append("pass at most 3 tools; the email names at most 3")
    else:
        fails.append("pass --tools with the 2 or 3 JD tools the email must name")
    for ph in BANNED:
        if re.search(r"" + re.escape(ph) + r"", low):
            fails.append(f"banned phrase: '{ph}'")
    if "—" in text or ";" in text:
        fails.append("em dash or semicolon found")
    if re.search(r"^\s*([-*•]|\d+\.)\s", body, re.M) or "**" in text:
        fails.append("bullets or bold found")
    if text.count("!") > 1:
        fails.append("more than one exclamation mark")
    if re.search(r"\b\d{3}[-. ]\d{3}[-. ]\d{4}\b", text) or "linkedin.com" in low:
        fails.append("phone number or LinkedIn link found")
    if re.search(r"\b(senior )?software engineer\b|\blead engineer\b|\bprincipal\b", low):
        fails.append("own job title or seniority mentioned")
    if a.no_resume and re.search(r"\b(resume|résumé|cv|attached|attachment)\b", low):
        fails.append("resume or attachment wording in a no-resume email")
    hits = [s for s in STOCK_REACTION if s in low]
    if len(hits) > 1:
        fails.append(f"more than one stock reaction phrase: {', '.join(hits)}")
    for f in fails:
        print("FAIL:", f)
    if not fails:
        print(f"OK: {n} words, {len(paras)} paragraphs, years and tools present.")
    return 1 if fails else 0


def cmd_batch(a):
    items = []
    for f in a.files:
        p = parse_email(open(f, encoding="utf-8").read())
        if not p:
            print(f"FAIL: could not read {f}")
            return 1
        greeting, body, signoff, name, email = p
        s = sentences(body)
        items.append(dict(f=f, intro=s[0].lower() if s else "", second=" ".join(s[1].lower().split()[:2]) if len(s) > 1 else "",
                          signoff=signoff.lower(), last=s[-1].lower() if s else "", low=body.lower(),
                          years=_years_snippet(body)))
    fails = 0

    def jac(x, y):
        a_, b_ = set(re.findall(r"\w+", x)), set(re.findall(r"\w+", y))
        return len(a_ & b_) / max(1, len(a_ | b_))

    for i in range(len(items)):
        for j in range(i + 1, len(items)):
            A, B = items[i], items[j]
            pair = f"{A['f']} vs {B['f']}"
            for key, label in (("intro", "same intro line"), ("signoff", "same sign-off"),
                               ("second", "same first two words in the opening line")):
                if A[key] and A[key] == B[key]:
                    print(f"FAIL: {pair}: {label}")
                    fails += 1
            if jac(A["last"], B["last"]) >= 0.6:
                print(f"FAIL: {pair}: ask wording too similar")
                fails += 1
            if A["years"] and A["years"] == B["years"]:
                print(f"FAIL: {pair}: same years wording ('{A['years']}')")
                fails += 1
    for s in STOCK_REACTION:
        users = [it["f"] for it in items if s in it["low"]]
        if len(users) > 1:
            print(f"FAIL: stock phrase '{s}' used in more than one email: {', '.join(users)}")
            fails += 1
    if not fails:
        print(f"OK: {len(items)} emails are different enough from each other.")
    return 1 if fails else 0


def _years_snippet(body):
    m = re.search(r"((?:\w+\s+){0,3})\d+\s*\+?\s*years?", body, re.I)
    return re.sub(r"\d+", "N", m.group(0).lower()) if m else ""


def main():
    ap = argparse.ArgumentParser(description="Module 04 helper")
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("brief")
    b.add_argument("--count", type=int, required=True)
    b.add_argument("--date")
    b.add_argument("--seed")
    c = sub.add_parser("check")
    c.add_argument("file")
    c.add_argument("--years")
    c.add_argument("--tools")
    c.add_argument("--min", type=int, default=90)
    c.add_argument("--max", type=int, default=120)
    c.add_argument("--no-resume", action="store_true")
    t = sub.add_parser("batch")
    t.add_argument("files", nargs="+")
    a = ap.parse_args()
    sys.exit({"brief": lambda x: cmd_brief(x) or 0, "check": cmd_check, "batch": cmd_batch}[a.cmd](a))


if __name__ == "__main__":
    main()
