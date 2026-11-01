"""Job selection (design section 6): resume-driven eligibility, exclusions, same recruiter, filters, limits and order."""
import datetime as dt
import unittest

from helpers import TempDir, cfg, job, make_resume
import dice_apply_select as SEL
import resume_find as RF

TODAY = dt.date(2026, 11, 2)


def blank(j):
    base = {c: "" for c in ("Job ID", "Priority", "Source Portal", "Run Timestamp", "Job Title", "Company", "Staffing Company", "Job URL",
                            "Recruiter Name", "Recruiter Email", "Pipeline Status", "Resume Filename", "Resume Path", "Applied Date",
                            "Notes", "Date Found", "Email Drafted Date")}
    base.update(j)
    return base


class Select(unittest.TestCase):
    def setUp(self):
        self.tmp = TempDir()
        self.root = self.tmp.__enter__() / "Resumes"
        self.cfg = cfg(max_per_round=None)

    def tearDown(self):
        self.tmp.__exit__()

    def res(self, n, tech="Workday", spec="HCM", name=None):
        name = name or f"Pat_{n}.docx"
        make_resume(self.root, tech, spec, f"JOB-{n}", name)
        return name

    def plan(self, jobs, ledger=(), queued=(), **kw):
        jobs = [blank(j) for j in jobs]
        return SEL.build_plan(jobs, kw.pop("filters", {}), self.cfg, set(ledger), set(queued), RF.ResumeIndex(self.root), today=TODAY, **kw)

    def codes(self, p):
        return {j["Job ID"]: c for j, c, _ in p.skipped}

    def picked(self, p):
        return [j["Job ID"] for j, _, _ in p.picked]

    def test_resume_driven_and_exclusions(self):
        jobs = [job(1, **{"Resume Filename": self.res(1)}),
                job(2),                                                            # no resume
                job(3, **{"Resume Filename": self.res(3), "Pipeline Status": "APPLIED"}),
                job(4, **{"Resume Filename": self.res(4), "Pipeline Status": "SKIPPED"}),
                job(5, **{"Resume Filename": self.res(5), "Applied Date": "2026-10-01"}),
                job(6, **{"Resume Filename": self.res(6), "Email Drafted Date": "2026-10-30 (resume attached)"}),
                job(7, **{"Resume Filename": self.res(7), "Email Drafted Date": "2026-10-30 (no resume)", "Pipeline Status": "EMAIL DRAFTED"}),
                job(8, **{"Resume Filename": self.res(8), "Source Portal": "LinkedIn"}),
                job(9, **{"Resume Filename": self.res(9)})]
        p = self.plan(jobs, ledger={job(9)["Job URL"].rsplit("/", 1)[1]})
        c = self.codes(p)
        self.assertEqual(c["JOB-2"], "NO_RESUME")
        self.assertEqual(c["JOB-3"], "STATUS_APPLIED")
        self.assertEqual(c["JOB-4"], "STATUS_SKIPPED")
        self.assertEqual(c["JOB-5"], "ALREADY_APPLIED")
        self.assertEqual(c["JOB-6"], "EMAILED_RESUME")
        self.assertEqual(c["JOB-9"], "IN_LEDGER")
        self.assertNotIn("JOB-8", c)                                               # not Dice: never considered
        self.assertEqual(sorted(self.picked(p)), ["JOB-1", "JOB-7"])

    def test_queue_counts_as_written(self):
        p = self.plan([job(1, **{"Resume Filename": self.res(1)})], queued={"JOB-1"})
        self.assertEqual(self.codes(p)["JOB-1"], "IN_LEDGER")

    def test_duplicate_title_company(self):
        jobs = [job(1, **{"Job Title": "Workday HCM", "Company": "Acme", "Pipeline Status": "APPLIED"}),
                job(2, **{"Job Title": "Workday HCM", "Company": "Acme", "Resume Filename": self.res(2)}),
                job(3, **{"Job Title": "Workday HCM Lead", "Company": "Acme", "Resume Filename": self.res(3)})]
        p = self.plan(jobs)
        self.assertEqual(self.codes(p)["JOB-2"], "DUPLICATE_OF JOB-1")
        self.assertIn("JOB-3", self.picked(p))
        self.assertIn("JOB-2", self.picked(self.plan(jobs, allow_duplicate=True)))

    def test_duplicate_posting_within_one_round(self):
        jobs = [job(1, **{"Job Title": "Java Developer", "Company": "Acme", "Resume Filename": self.res(1), "Date Found": "2026-10-01"}),
                job(2, **{"Job Title": "Java Developer", "Company": "Acme", "Resume Filename": self.res(2), "Date Found": "2026-10-02"})]
        p = self.plan(jobs)
        self.assertEqual(self.picked(p), ["JOB-1"])
        self.assertEqual(self.codes(p)["JOB-2"], "DUPLICATE_OF JOB-1")
        self.assertEqual(len(self.plan(jobs, allow_duplicate=True).picked), 2)

    def test_same_recruiter(self):
        rec = {"Recruiter Name": "Sam Recruit", "Recruiter Email": "sam@staff.example", "Staffing Company": "Staff Co"}
        jobs = [job(1, **rec, **{"Pipeline Status": "APPLIED", "Resume Filename": self.res(1, "Workday", "HCM")}),
                job(2, **rec, **{"Resume Filename": self.res(2, "Workday", "HCM")}),          # similar: allowed
                job(3, **rec, **{"Resume Filename": self.res(3, "Java", "")}),                 # different role
                job(4, **dict(rec, **{"Recruiter Email": "other@staff.example", "Recruiter Name": "Other Person"}),
                    **{"Resume Filename": self.res(4, "Java", "")})]                           # same agency, other recruiter: allowed
        p = self.plan(jobs)
        c = self.codes(p)
        self.assertEqual(c["JOB-3"], "SAME_RECRUITER JOB-1")
        self.assertIn("JOB-2", self.picked(p))
        self.assertIn("JOB-4", self.picked(p))
        self.assertIn("similar role", dict((j["Job ID"], n) for j, _, n in p.picked)["JOB-2"])
        self.assertIn("JOB-3", self.picked(self.plan(jobs, allow=["JOB-3"])))

    def test_same_recruiter_unclear_when_emailed(self):
        rec = {"Recruiter Email": "sam@staff.example"}
        jobs = [job(1, **rec, **{"Email Drafted Date": "2026-10-30 (resume attached)"}),
                job(2, **rec, **{"Resume Filename": self.res(2)})]
        p = self.plan(jobs)
        self.assertEqual(self.codes(p)["JOB-2"], "SAME_RECRUITER JOB-1")
        self.assertIn("role unclear", [d for j, _, d in p.skipped if j["Job ID"] == "JOB-2"][0])

    def test_same_recruiter_within_round(self):
        rec = {"Recruiter Email": "sam@staff.example"}
        jobs = [job(1, **rec, **{"Resume Filename": self.res(1, "Workday", "HCM"), "Date Found": "2026-10-01"}),
                job(2, **rec, **{"Resume Filename": self.res(2, "Java", ""), "Date Found": "2026-10-02"})]
        p = self.plan(jobs)
        self.assertEqual(self.picked(p), ["JOB-1"])
        self.assertEqual(self.codes(p)["JOB-2"], "SAME_RECRUITER JOB-1")

    def test_generic_inbox_is_not_a_recruiter(self):
        a, b = blank(job(1, **{"Recruiter Email": "jobs@staff.example", "Recruiter Name": ""})), blank(job(2, **{"Recruiter Email": "jobs@staff.example", "Recruiter Name": ""}))
        self.assertFalse(SEL.same_recruiter(a, b))

    def test_filters(self):
        jobs = [job(1, **{"Resume Filename": self.res(1), "Run Timestamp": "2026-11-02 07:00 MST", "Priority": "Tier 1", "Job Title": "Workday HCM"}),
                job(2, **{"Resume Filename": self.res(2), "Run Timestamp": "2026-11-01 11:30 MDT", "Priority": "Tier 4", "Job Title": "Java Dev"}),
                job(3, **{"Resume Filename": self.res(3), "Run Timestamp": "", "Date Found": "2026-11-02", "Priority": "Tier 5"})]
        f = lambda **k: self.picked(self.plan(jobs, filters=k))
        self.assertEqual(f(run_date="today"), ["JOB-1"])
        self.assertEqual(f(run_date="yesterday"), ["JOB-2"])
        self.assertEqual(f(run_latest=True), ["JOB-1"])
        self.assertEqual(f(priority="1,2"), ["JOB-1"])
        self.assertEqual(f(title_contains=["java"]), ["JOB-2"])
        self.assertEqual(f(found="today"), ["JOB-3"])
        self.assertEqual(f(jobs="JOB-2,JOB-3"), ["JOB-2", "JOB-3"])
        self.assertTrue(any("no Run Timestamp" in n for n in self.plan(jobs, filters={"run_date": "today"}).notes))

    def test_limits(self):
        jobs = [job(n, **{"Resume Filename": self.res(n)}) for n in range(1, 8)]
        p = self.plan(jobs, limit=5)
        self.assertEqual((len(p.picked), len(p.left)), (5, 2))
        self.cfg["max_per_round"] = 3
        p = self.plan(jobs, limit=5)
        self.assertEqual(len(p.picked), 3)
        self.assertIn("max_per_round", p.limit_note)
        p = self.plan(jobs, limit=5, remaining_today=2)
        self.assertEqual(len(p.picked), 2)

    def test_order_spreads_companies(self):
        jobs = [blank(job(n, **{"Company": "Acme" if n < 4 else "Beta", "Date Found": f"2026-10-0{n}"})) for n in range(1, 6)]
        ids = [j["Job ID"] for j in SEL.order(jobs)]
        self.assertEqual(sorted(ids), sorted(j["Job ID"] for j in jobs))
        comps = [next(j["Company"] for j in jobs if j["Job ID"] == i) for i in ids]
        self.assertLessEqual(sum(1 for a, b in zip(comps, comps[1:]) if a == b), 1)

    def test_estimate(self):
        self.assertGreater(SEL.estimate_minutes(25, self.cfg, "relaxed"), SEL.estimate_minutes(25, self.cfg, "brisk"))


if __name__ == "__main__":
    unittest.main()
