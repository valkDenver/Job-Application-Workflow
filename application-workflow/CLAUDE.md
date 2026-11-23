Read `README.md` in this folder first (its "For AI assistants" section holds the tracker rules).
The candidate usually types a slash command from `.claude/commands/` (/setup, /profile, /get-jobs, /check-setup, /call-prep, /build-resume, /email-draft, /email-follow-up, /email-sent, /my-progress); each command file says which prompt or tool to follow. Plain sentences work too:
For "set up the workflow" requests, follow `0 - Setup.txt` in this folder. For "build my profile" or "update my profile", follow `0b - Profile.txt`.
For "import my new jobs": run `python Tools/tracker_tool.py import-feed "<tracker from tracker_path.txt>"` (add `--feed "<file>"` and the starting `--after JOB-####` the first time; ask me for them), and report its result line. Never open the feed file in any other way.
For "check my sent emails" or "which emails did I send?", follow Section 10C of `4 - Recruiter Follow-up Email.txt`.
For "Run Module 02/03/04 ..." requests (vendor call, resume, recruiter email or follow-up), follow `2 - Vendor Call Assistant.txt`, `3 - Resume Generator.txt` or `4 - Recruiter Follow-up Email.txt` in this folder. For resumes for several jobs at once, follow `Run Notes/BATCH_RUN.md`.
My tracker is the path in `tracker_path.txt` and my Candidate Profile is the one "Candidate Profile - <name>.txt" here (not TEMPLATE or TestSample): use them without asking me for file names.
Run every `Tools/...` command from this folder.
