"""Tests for Tools/profile_check.py and resume_check.py's profile reading. Offline; profiles are built from the template in a
temporary folder.

  python Tools/tests/test_profile_check.py      prints "all tests passed"
"""
import datetime as dt, os, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
TOOLS = os.path.dirname(HERE)
sys.path.insert(0, TOOLS)
import profile_check as P
import resume_check as R

FAIL = []
TEMPLATE = open(os.path.join(os.path.dirname(TOOLS), "Candidate Profile - TEMPLATE.txt"), encoding="utf-8").read()
TODAY = dt.date(2026, 11, 21)


def check(name, ok, detail=""):
    if not ok:
        FAIL.append(name + (": " + str(detail)[:500] if detail else ""))


def fill(txt, pairs):
    """Replaces each 'Label: <anything>' line, in order, with the given value; pairs = [(label, value), ...] applied to the next
    line with that label that has not been replaced yet."""
    lines = txt.splitlines()
    done = set()
    for label, value in pairs:
        for i, ln in enumerate(lines):
            if i not in done and ln.strip().startswith(label + ":"):
                lines[i] = f"{label}: {value}"
                done.add(i)
                break
    return "\n".join(lines) + "\n"


def employer(name, title, dates, domain="Telecom / Data"):
    return [("Employer / client name", name), ("Job title", title), ("Dates", dates), ("Domain", domain),
            ("Domain guidance", "Telecom; data platforms"), ("Seniority notes", "Senior ownership")]


def personal(linkedin=""):
    return [("Name", "Test Person"), ("Location", "Denver, CO"), ("Phone", "555-0100"), ("Email", "test@example.com"),
            ("LinkedIn", linkedin), ("Total experience", "Around 8 years in data")]


def three_employers():
    return (personal() + employer("Acme", "Senior Engineer", "Jan 2024 - Present") + employer("Beta", "Engineer", "Mar 2021 - Dec 2023")
            + employer("Gamma", "Analyst", "June 2018 - Feb 2021") + [("Seniority notes", ""), ("Resume template", ""),
                                                                    ("Education", ""), ("Certifications", "")])


def clean_unused(txt):
    """Blanks every remaining employer field the way a candidate leaves unused blocks: label with nothing after it."""
    out = []
    for ln in txt.splitlines():
        s = ln.strip()
        if any(s.startswith(f + ":") for f in P.EMPLOYER_FIELDS) and "[" in s:
            ln = s.split(":")[0] + ":"
        out.append(ln)
    return "\n".join(out) + "\n"


def write(d, txt):
    p = os.path.join(d, "Candidate Profile - Test.txt")
    open(p, "w", encoding="utf-8").write(txt)
    return p


def test_blank_fields_never_take_the_next_line():
    with tempfile.TemporaryDirectory() as d:
        path = write(d, clean_unused(fill(TEMPLATE, three_employers())))
        name, contact, want = R.profile_expectations(path)
        check("name", name == "Test Person", name)
        check("blank LinkedIn stays out of the contact line", contact == "Denver, CO | 555-0100 | test@example.com", contact)
        check("only the 3 employers in use", want == ["Acme | Senior Engineer | Jan 2024 - Present", "Beta | Engineer | Mar 2021 - Dec 2023",
                                                      "Gamma | Analyst | June 2018 - Feb 2021"], want)
        check("resume level passes", P.check_resume(P.read(path), TODAY) == [], P.check_resume(P.read(path), TODAY))


def test_template_text_and_missing_fields():
    p = P.parse(TEMPLATE)
    probs = P.check_resume(p, TODAY)
    for label in P.PERSONAL_REQUIRED:
        check("required reported: " + label, any(x.startswith(label + ":") for x in probs), probs)
    check("LinkedIn template text reported", any(x.startswith("LinkedIn:") for x in probs), probs)
    check("untouched template: no employer filled in", any(x.startswith("no employer filled in") for x in probs), probs)
    half = P.check_resume(P.parse(fill(TEMPLATE, [("Employer / client name", "Acme")])), TODAY)
    check("a started block reports its missing fields", any("Acme: Job title is missing" in x for x in half), half)
    txt = fill(TEMPLATE, personal("linkedin.com/in/test") + employer("Acme", "Engineer", "Jan 2024 - Present")
               + [("Domain guidance", "[left]"), ("Resume template", "fancy")])
    probs = P.check_resume(P.parse(clean_unused(txt)), TODAY)
    check("unknown template", any('Resume template "fancy"' in x for x in probs), probs)
    check("Education template text", any(x.startswith("Education:") for x in probs), probs)


def test_dates():
    def probs(*emps):
        pairs = personal()
        for e in emps:
            pairs += employer(*e)
        return P.check_resume(P.parse(clean_unused(fill(TEMPLATE, pairs + [("Resume template", ""), ("Education", ""), ("Certifications", "")]))), TODAY)
    check("good dates pass", probs(("A", "T", "Jan 2024 - Present"), ("B", "T", "Sept 2020 - Dec 2023")) == [],
          probs(("A", "T", "Jan 2024 - Present"), ("B", "T", "Sept 2020 - Dec 2023")))
    check("bad form", any("must look like" in x for x in probs(("A", "T", "2020 - 2023"),)))
    check("future", any("future" in x for x in probs(("A", "T", "Jan 2027 - Present"),)))
    check("start after end", any("start after" in x for x in probs(("A", "T", "Jan 2024 - Mar 2023"),)))
    check("order", any("most recent first" in x for x in probs(("A", "T", "Jan 2019 - Dec 2020"), ("B", "T", "Jan 2022 - Present"))))


def test_call_level():
    txt = clean_unused(fill(TEMPLATE, three_employers() + [
        ("Facts last confirmed on", "2026-10-01"), ("Education", "B.Tech, 2016"), ("Degree dates and place", "N/A"),
        ("Current employment status", "Contract ends 2026-11-28"), ("Availability / start date", "2026-11-01")]))
    blank, stale = P.check_call(P.parse(txt), TODAY)
    check("old confirmation date", any("51 days old" in x for x in stale), stale)
    check("date within 14 days", any("2026-11-28 is 7 day(s) away" in x for x in stale), stale)
    check("passed date", any("2026-11-01 has passed" in x for x in stale), stale)
    check("N/A is answered", not any(x.startswith("Degree dates and place") for x in blank), blank)
    check("filled is answered", not any(x.startswith("Education") for x in blank), blank)
    check("blank listed with its section", any(x.startswith("Work authorization (2. AUTHORIZATION)") for x in blank), blank)
    check("optional not listed", not any(x.startswith("Current compensation") for x in blank), blank)
    check("prefilled employer of record not listed", not any(x.startswith("Employer name") for x in blank), blank)
    blank, stale = P.check_call(P.parse(TEMPLATE), TODAY)
    check("missing confirmation date", any("missing" in x for x in stale), stale)


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
    if FAIL:
        print("FAILED:\n  " + "\n  ".join(FAIL)); sys.exit(1)
    print("all tests passed")
