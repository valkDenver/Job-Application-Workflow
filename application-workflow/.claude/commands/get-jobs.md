---
description: Get my new jobs from the job feed into my tracker (with their full descriptions)
argument-hint: "[JOB-#### starting Job ID, first time only]"
---

Add the job feed's new jobs to my tracker. Use only `python Tools/tracker_tool.py import-feed` (python3 on a Mac) from this folder; never open the feed file any other way.
- My tracker is the path in `tracker_path.txt`. If that file is missing, tell me to run /setup first and stop.
- If `feed_path.txt` exists, run: python Tools/tracker_tool.py import-feed "<my tracker>"
- If it does not (first time): ask where my feed file is on my computer (the README's "Get jobs from the shared feed" says how to keep it synced) and use my starting Job ID: "$ARGUMENTS" if I typed one, else ask for it. Run: python Tools/tracker_tool.py import-feed "<my tracker>" --feed "<feed file>" --after <starting Job ID>
- If it says my tracker has no feed jobs yet and asks where to start (for example a new tracker), run it again with --after <starting Job ID>: "$ARGUMENTS" if I typed one, else ask me for the starting Job ID I was given.
- If it says the workbook is open in Excel, ask me to close it, then run it again.
Tell me its result line in one or two sentences (how many jobs were added and their Job ID range). Do not list the jobs.
