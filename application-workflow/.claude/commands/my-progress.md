---
description: Show my progress - jobs, calls, resumes, emails drafted and sent, submissions and applications - for the last days
argument-hint: "[number of days, default 7]"
---

Show my progress. My tracker is the path in `tracker_path.txt` (missing: tell me to run /setup and stop). Run from this folder (python3 on a Mac), changing nothing:
1. python Tools/tracker_tool.py summary "<my tracker>"
2. python Tools/tracker_tool.py report "<my tracker>" --days <"$ARGUMENTS" if I typed a number, else 7>
3. python Tools/tracker_tool.py unsent "<my tracker>" --limit 0 (only to count the drafts not recorded as sent)

Give me a short summary: totals by status, what I did in the period, how fast I reached first contact, and one line of next steps (for example: drafts not sent yet: run /email-sent; follow-ups due: run /email-follow-up; RESUME NEEDED jobs: run /build-resume all). Do not print the tracker or list every job.
