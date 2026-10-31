# NOTE: application-workflow/Tools/contact_clean.py and publisher/Tools/contact_clean.py are separate copies so the two flows stay independent. If you change the cleaning rules or shared helpers in one, make the same change in the other.
"""One place that decides what a clean Recruiter Name / Recruiter Email / Recruiter Phone cell looks like.

Used by tracker_tool.py edit (refuses bad input) and tracker_tool.py check-contacts (finds bad cells already in the workbook). Pure functions, no file or network access.

  python contact_clean.py --selftest      runs the built-in checks

Rules (each one comes from a real bad cell found in the 2026-10-29 review):
  Email : exactly one lower-case address, nothing else in the cell. Words glued on by the page text are removed
          (name@co.comPhone, name@co.comWebsite, Specialistname@co.com, Taylortaylor@co.com). A bracketed note is not allowed in the cell.
          Anything that is not an address becomes "Not publicly verified".
  Name  : a person's name, optionally followed by bracketed notes. Description or signature text (pipes, "Regards,", "Email:", digits) is cut
          away; if no name is left the result is empty and the caller writes the standard "Not in posting text" / "Not in scraped data" text.
  Phone : one phone number (optionally "Ext: n"), never an email address or other text.
The standard placeholder texts ("Not publicly verified", "Not in posting text", "Not in scraped data - ...") are kept exactly as they are:
Module 04 looks for those words.
"""
import re

SENT_EMAIL = "Not publicly verified"
SENT_PHONE = "Not publicly verified"
PLACEHOLDER = re.compile(r"^\s*(not\s+(in|publicly|shown|stated|displayed|named|available)|unverified|n/?a\b|-+$)", re.I)

# ---------------------------------------------------------------- email
ADDR = re.compile(r"[A-Za-z0-9][A-Za-z0-9._%+\-]*@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}")
REAL_TLD = r"(?:com|net|org|us|io|ai|co|in|info|biz|edu|gov|dev|tech|uk|ca|de)"
GLUE_WORDS = r"(?:Phone|Website|Web|Email|E-?mail|Mobile|Direct|Tel|Fax|Cell|Contact|Linkedin|Address|Skype|Ph|Mob|Office|Www|Http)"
GLUE_SUFFIX = re.compile(r"^(.+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.%s)%s$" % (REAL_TLD, GLUE_WORDS), re.I)
# job-title words that the page text runs into the address ("Talent Acquisition Specialistchandan.s@..."). Only these words, capital first
# letter, and at least 4 characters must follow, so real names such as Leadbetter@ or Partnership@ are left alone.
TITLE_PREFIX = re.compile(r"^(Specialist|Recruiter|Sourcer|Coordinator|Executive|Consultant|Analyst)([a-z0-9][A-Za-z0-9._%+\-]{3,})$")


def clean_email(raw):
    """-> (email, note). email is "" when the cell holds no usable address. note says what was changed (empty when nothing was)."""
    if raw is None:
        return "", ""
    s = str(raw).strip()
    if not s or PLACEHOLDER.match(s):
        return "", ""
    notes = []
    m = ADDR.search(s)
    if not m:
        return "", "not an email address"
    e = m.group(0).strip(".,;:")
    if e != s:
        notes.append("extra text removed")
    g = GLUE_SUFFIX.match(e)
    if g:
        notes.append("glued word '%s' removed from the end" % e[len(g.group(1)):])
        e = g.group(1)
    local, _, dom = e.partition("@")
    t = TITLE_PREFIX.match(local)
    if t:
        notes.append("glued word '%s' removed from the start" % t.group(1))
        local = t.group(2)
    for i in range(3, len(local) // 2 + 1):  # signature first name run into the address: Vigneshvignesh@ -> vignesh@
        if local[0].isupper() and local[:i].isalpha() and local[i:] == local[:i].lower():
            notes.append("repeated name '%s' removed from the start" % local[:i])
            local = local[i:]
            break
    e = ("%s@%s" % (local, dom)).lower()
    if not re.match(r"^[a-z0-9][a-z0-9._%+\-]*@[a-z0-9\-]+(\.[a-z0-9\-]+)*\.[a-z]{2,}$", e):
        return "", "not an email address"
    return e, "; ".join(notes)


# ---------------------------------------------------------------- name
SIGN_OFF = re.compile(r"\b(?:warm |kind |best )?(?:regards|thanks|thank you|sincerely|cheers)\b[\s,:&\-]*", re.I)
CUT_AT = re.compile(r"\s*(?:\||,|;|:|\b(?:e-?mail|direct|phone|ph|mobile|cell|tel|fax|website|office|linkedin)\b).*$", re.I)
CREDENTIALS = re.compile(r",\s*(?:MBA|CPC|PMP|PHR|SPHR|CSM|CSP|MS|MSc|PhD|Jr\.?|Sr\.?|II|III)\b")
TITLE_CUT = re.compile(r"(?<=\S)\s+(?:account|talent|technical|senior|sr|recruiter|recruitment|manager|lead|director|executive|specialist|associate|sourcer|hr)\b.*$", re.I)
WORD = re.compile(r"^[A-Za-z][A-Za-z.'\-]*$")
NOT_NAME_WORDS = {"and", "the", "of", "for", "with", "to", "in", "on", "is", "are", "a", "an", "we", "you", "our", "your", "please", "contact",
                  "involving", "experience", "required", "skills", "services", "support", "cloud", "team"}
# characters outside the Latin letters (emoji, check marks) are never part of a name
NON_LATIN = re.compile("[^\u0000-ɏ’]")


def clean_name(raw, strict=False):
    """-> (name, note). name is "" when no person's name can be found.
    strict=True is for names picked out of free job-description text: every word must start with a capital letter.
    strict=False is for names taken from a structured poster field or typed by the user."""
    if raw is None:
        return "", ""
    s = re.sub(r"\s+", " ", NON_LATIN.sub("", str(raw))).strip()
    if not s:
        return "", ""
    if PLACEHOLDER.match(s):
        return s, ""
    parts = SIGN_OFF.split(s)
    t = parts[-1].strip() if len(parts) > 1 and parts[-1].strip() else s
    core, par, rest = t.partition("(")
    ann = ""
    if par:  # keep the bracketed notes exactly as they are (they may hold brackets inside); drop only loose text after the last ")"
        ann = "(" + rest
        if ")" in ann:
            ann = ann[: ann.rindex(")") + 1]
    cred = ""
    cm = CREDENTIALS.search(core)
    if cm:
        cred, core = cm.group(0), core[: cm.start()] + core[cm.end():]
    core = CUT_AT.sub("", core).strip()
    if strict:  # free text runs the job title on after the name ("Sandy G Account Manager")
        core = TITLE_CUT.sub("", core).strip()
    words = [w for w in re.split(r"\s+", core.replace(" / ", " ").replace(" - ", " ")) if w]
    ok = 1 <= len(words) <= 6 and all(WORD.match(w) for w in words) and len(core) <= 45
    if ok and strict and not all(w[0].isupper() for w in words):
        ok = False
    if ok and any(len(w.strip(".")) > 1 and w.lower().strip(".") in NOT_NAME_WORDS for w in words):
        ok = False
    if ok and len(words) == 1 and (core.isupper() or len(core) < 3):
        ok = False
    if not ok:
        return "", "no person's name found"
    letters = re.sub(r"[^A-Za-z]", "", core)
    if letters.isupper() or letters.islower():
        core = core.title()
    out = (core + cred + " " + ann).strip()
    return out, ("" if out == s else "tidied")


# ---------------------------------------------------------------- phone
PHONE = re.compile(r"(?:\+?1[\s.\-]?)?\(?\d{3}\)?[\s.\-]?\d{3}[\s.\-]?\d{4}")
INTL_PHONE = re.compile(r"\+\d[\d\s().\-]{7,}\d")  # non-US numbers (+91 98765 43210) are kept as written
EXT = re.compile(r"\s*(?:ext|extn|x)\.?\s*:?\s*\d{1,6}", re.I)


def clean_phone(raw):
    """-> (phone, note). phone is "" when no number is found. An extension is kept as written."""
    if raw is None:
        return "", ""
    s = str(raw).strip()
    if not s or PLACEHOLDER.match(s):
        return "", ""
    s = ADDR.sub(" ", s)  # digits inside an email address are not a phone number
    m = PHONE.search(s) or INTL_PHONE.search(s)
    if not m:
        return "", "not a phone number"
    p = m.group(0).strip()
    x = EXT.match(s[m.end():])
    if x:
        p += " " + x.group(0).strip()
    return p, ("" if p == s else "extra text removed")


# ---------------------------------------------------------------- one row
def name_email_mismatch(name, email):
    """True when both are real and no part of the name appears in the address, or the address is a shared mailbox
    (recruiter@, specialist@). A signal to look at the row, not an error."""
    if not name or not email or PLACEHOLDER.match(name) or "@" not in email:
        return False
    local = email.split("@")[0].lower()
    toks = [t for t in re.split(r"[^a-z]+", re.sub(r"\(.*", "", name.lower())) if len(t) > 1]
    if not toks:
        return False
    if local in ("recruiter", "specialist"):
        return True
    if local in ("info", "hr", "careers", "jobs", "talent"):
        return False
    return not any(t[:3] in local or local.startswith(t[0]) for t in toks)


def check_row(name, email, phone):
    """List of (column, problem, suggested value) for one row's contact cells (empty list = fine). Placeholders are fine.
    A cell is only reported when text would have to be removed or a case/spacing fix is needed."""
    out = []
    if email is not None and str(email).strip() and not PLACEHOLDER.match(str(email)):
        e, note = clean_email(email)
        if not e:
            out.append(("Recruiter Email", "not an address", SENT_EMAIL))
        elif e != str(email).strip() or note:
            out.append(("Recruiter Email", note or "needs tidying", e))
    if name is not None and str(name).strip() and not PLACEHOLDER.match(str(name)):
        n, note = clean_name(name)
        if not n:
            out.append(("Recruiter Name", "description/signature text, no name", "Not in posting text"))
        elif n != re.sub(r"\s+", " ", str(name)).strip():
            out.append(("Recruiter Name", "needs tidying", n))
    if phone is not None and str(phone).strip() and not PLACEHOLDER.match(str(phone)):
        p, note = clean_phone(phone)
        if not p:
            out.append(("Recruiter Phone", "not a phone number", SENT_PHONE))
        elif re.search(r"@|[A-Za-z]{4,}", re.sub(r"\(.*?\)|ext|extn", "", str(phone), flags=re.I)):
            out.append(("Recruiter Phone", "text mixed into the number", p))
    if name_email_mismatch(None if name is None else str(name), None if email is None else str(email).strip().lower()):
        out.append(("Name vs Email", "name does not match the address - check by hand", ""))
    return out


def _selftest():
    bad = []

    def eq(got, want, label):
        if got != want:
            bad.append("%s: got %r want %r" % (label, got, want))
    E = lambda x: clean_email(x)[0]
    eq(E("alex.morgan@examplecorp.comPhone"), "alex.morgan@examplecorp.com", "suffix Phone")
    eq(E("Pat.lee@examplecorp.comWebsite"), "pat.lee@examplecorp.com", "suffix Website")
    eq(E("Specialistsam.r@examplecorp.com"), "sam.r@examplecorp.com", "prefix Specialist")
    eq(E("Taylortaylor@example.com"), "taylor@example.com", "echo")
    eq(E("casey.nash@example.us (stated in posting; domain matches Example Technologies)"), "casey.nash@example.us", "annotation")
    eq(E("Joe@Example.com"), "joe@example.com", "case")
    eq(E("466-8260Link"), "", "phone fragment")
    eq(E("Eq"), "", "word")
    eq(E("Not publicly verified"), "", "placeholder")
    for ok in ("specialist@programmers.io", "recruiter@programmers.io", "Leadbetter@x.com", "Partnership@x.com", "Specialists@x.com", "john@company.co.in",
               "john@phoneco.com", "jane@website.com", "info@contact.com", "a@b.company.com", "x@y.communications.com", "bobbob@x.com", "Anna@x.com",
               "Hannah@x.com", "Mathew@x.com", "Annanna@x.com", "Bobbie@x.com"):
        eq(E(ok), ok.lower(), "unchanged " + ok)
    N = lambda x, s=False: clean_name(x, s)[0]
    eq(N("Kim Garner (LinkedIn poster, Accounting & Finance Recruiter at Example Staffing)"), "Kim Garner (LinkedIn poster, Accounting & Finance Recruiter at Example Staffing)", "annotated name kept")
    eq(N("Joe (Riley Brooks, Recruitment Manager, named in posting)"), "Joe (Riley Brooks, Recruitment Manager, named in posting)", "first name only kept")
    eq(N("Vic (Victor) (LinkedIn poster)"), "Vic (Victor) (LinkedIn poster)", "two groups kept")
    eq(N("JAMIE STONE (Sr Account Manager, LinkedIn poster)"), "Jamie Stone (Sr Account Manager, LinkedIn poster)", "caps")
    eq(N("Quinn  Reed (Recruiter, Dice poster card)"), "Quinn Reed (Recruiter, Dice poster card)", "double space")
    eq(N("taylor k (job poster, LinkedIn)"), "Taylor K (job poster, LinkedIn)", "lowercase")
    eq(N("Job Poster (Recruiter, Dice poster card)"), "Job Poster (Recruiter, Dice poster card)", "generic poster kept as is")
    eq(N("Drew Archer (Talent Acquisition (US Recruitment), LinkedIn poster)"), "Drew Archer (Talent Acquisition (US Recruitment), LinkedIn poster)", "nested brackets kept")
    eq(N("Sage A (LinkedIn poster, Recruitment Lead)"), "Sage A (LinkedIn poster, Recruitment Lead)", "single-letter initial")
    eq(N("Morgan M. / Morgan Smith (LinkedIn poster)"), "Morgan M. / Morgan Smith (LinkedIn poster)", "two spellings kept")
    eq(N("Kris - Dana Lane (job poster, LinkedIn)"), "Kris - Dana Lane (job poster, LinkedIn)", "dash kept")
    eq(N("Ada Parker, MBA (poster)"), "Ada Parker, MBA (poster)", "credentials kept")
    eq(N("Nolan Moore at Example Leaders (poster)"), "Nolan Moore at Example Leaders (poster)", "at company kept")
    eq(N("Robin G. (LinkedIn profile only)"), "Robin G. (LinkedIn profile only)", "initial with dot")
    eq(N("Mary-Anne O'Neil"), "Mary-Anne O'Neil", "hyphen apostrophe")
    eq(N("Mr.Sam Hartley ✅ (LinkedIn poster)"), "Mr.Sam Hartley (LinkedIn poster)", "emoji removed")
    eq(N("involving APIs| microservices| JavaSpring Boot| cloud-native integrat"), "", "junk pipes")
    eq(N("NS| DHCP| Group Policy| and related Windows infrastructure services.Troubleshoot (named in the posting)"), "", "junk NS")
    eq(N("loyment & Support | Cloud Services/APIs | AI Observability | Enterprise Se (named in the posting)"), "", "junk loyment")
    eq(N("esirable. Regards, Alan Jones (Alan) Direct: 5 5 5 0 1 0 0 1 2 3"), "Alan Jones (Alan)", "regards + direct")
    eq(N("ust. Warm Regards, Sandy G Email: Website: (named in the posting)"), "Sandy G (named in the posting)", "email label")
    eq(N(",Thanks & Regards,Casey Kent, Example International, INC. | 233 Eas"), "Casey Kent", "thanks & regards")
    eq(N("involving APIs", True), "", "strict lowercase start")
    eq(N("Not in posting text"), "Not in posting text", "placeholder kept")
    eq(N("Not in scraped data - check LinkedIn while logged in"), "Not in scraped data - check LinkedIn while logged in", "placeholder 2 kept")
    P = lambda x: clean_phone(x)[0]
    eq(P("+1 201-555-0142"), "+1 201-555-0142", "phone ok")
    eq(P("alex@example.com Ph No: 480-555-0123 Ext: 241"), "480-555-0123 Ext: 241", "phone with email in front")
    eq(P("302-555-0177 Ext. 135"), "302-555-0177 Ext. 135", "ext kept as written")
    eq(P("john5551234567@x.com"), "", "digits inside an email")
    eq(P("+91 98765 43210"), "+91 98765 43210", "international kept")
    eq(N("Sandy G Account Manager", True), "Sandy G", "title cut in free text")
    eq(N("Sandy G Account Manager"), "Sandy G Account Manager", "title kept when name is structured")
    eq(P("1 Great Neck, NY 11023Direct"), "", "address fragment")
    eq(P("Not publicly verified"), "", "phone placeholder")
    eq(check_row("Sana Dean (named in posting; Lead - US IT Recruitment)", "sana@example.com", "Not publicly verified"), [], "clean row has no issues")
    eq(check_row("Ada Parker, MBA (poster)", "ada.parker@x.com", "+1 (469) 555 0188 Ext 340 (from posting)"), [], "credentials and phone note are fine")
    eq([c for c, *_ in check_row("Rex Kim (Recruiter, Dice poster card)", "Not publicly verified", "x")], ["Recruiter Phone"], "check_row phone")
    eq([c for c, *_ in check_row("Neil Ward (job poster, LinkedIn)", "rhea.banks@example.com", None)], ["Name vs Email"], "mismatch flagged")
    eq([c for c, *_ in check_row("Dana Grant (job poster, LinkedIn)", "specialist@programmers.io", None)], ["Name vs Email"], "shared mailbox flagged")
    print("contact_clean self-test:", "FAILED" if bad else "all checks passed")
    for b in bad:
        print("  ", b)
    return 0 if not bad else 1


if __name__ == "__main__":
    import sys
    if "--selftest" in sys.argv:
        sys.exit(_selftest())
    print(__doc__)
