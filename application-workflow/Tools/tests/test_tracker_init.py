"""Tests for tracker_tool.py init. Offline, temporary folders; never touches the real tracker_path.txt or the template.

  python Tools/tests/test_tracker_init.py      prints "all tests passed"
"""
import contextlib, hashlib, io, os, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
TOOLS = os.path.dirname(HERE)
sys.path.insert(0, TOOLS)
import openpyxl
import tracker_tool as T
import setup_check as S

FAIL = []
TEMPLATE = os.path.join(T.APP_FOLDER, T.TEMPLATE_NAME)


def check(name, ok, detail=""):
    if not ok:
        FAIL.append(name + (": " + str(detail)[:400] if detail else ""))


def sha(p):
    return hashlib.sha256(open(p, "rb").read()).hexdigest()


def run(d, *args):
    saved = T.TRACKER_TXT
    T.TRACKER_TXT = os.path.join(d, "tracker_path.txt")
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            rc = T.main(["init", "--folder", d] + list(args))
    finally:
        T.TRACKER_TXT = saved
    return rc, buf.getvalue()


def test_names():
    check("spaces become underscores", T.file_safe_name("Asha  Rao") == "Asha_Rao")
    check("accents and punctuation", T.file_safe_name("José Rao-Kumar.") == "Jose_Rao_Kumar", T.file_safe_name("José Rao-Kumar."))
    check("nothing usable", T.file_safe_name("  --  ") == "")


def test_init_creates_a_clean_tracker():
    before = sha(TEMPLATE)
    with tempfile.TemporaryDirectory() as d:
        rc, out = run(d, "--name", "Asha Rao")
        target = os.path.join(d, "Master Job Tracker - Asha_Rao.xlsx")
        check("init succeeds", rc == 0 and os.path.isfile(target), out)
        check("template unchanged", sha(TEMPLATE) == before)
        check("tracker_path.txt points at it", open(os.path.join(d, "tracker_path.txt"), encoding="utf-8").read() == target + "\n")
        check("setup_check accepts it", S.check_tracker(target) == "", S.check_tracker(target))
        new, tpl = openpyxl.load_workbook(target), openpyxl.load_workbook(TEMPLATE)
        check("same sheets", new.sheetnames == tpl.sheetnames)
        for s in T.EXAMPLE_SHEETS:
            check("no EXAMPLE row on " + s, not T.example_rows(new[s]))
            check("headers kept on " + s, [c.value for c in new[s][1]] == [c.value for c in tpl[s][1]])
            check("filter kept on " + s, new[s].auto_filter.ref == tpl[s].auto_filter.ref, new[s].auto_filter.ref)
            check("freeze kept on " + s, new[s].freeze_panes == tpl[s].freeze_panes)
        dv = lambda wb: sorted(str(x.sqref) for x in wb[T.MASTER].data_validations.dataValidation)
        check("dropdown lists kept", dv(new) == dv(tpl), dv(new))
        f = lambda wb: [c.value for row in wb["Dashboard"].iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("=")]
        check("Dashboard formulas kept", f(new) == f(tpl) and f(new), len(f(new)))
        rc, out = run(d, "--name", "Asha Rao")
        check("second init refuses and changes nothing", rc == 2 and "already exists" in out, out)


def test_init_refusals():
    with tempfile.TemporaryDirectory() as d:
        rc, out = run(d, "--name", "!!")
        check("empty name refused", rc == 2 and "--name" in out, out)
        rc, out = run(d, "--name", "Live")
        check("reserved name refused", rc == 2 and "reserved" in out, out)
        check("nothing written on refusal", os.listdir(d) == [], os.listdir(d))


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
    if FAIL:
        print("FAILED:\n  " + "\n  ".join(FAIL)); sys.exit(1)
    print("all tests passed")
