# Batch resume run (Module 03 for many jobs)

Instructions for the AI assistant when the candidate asks for resumes for several jobs at once, for example "build resumes for all my RESUME NEEDED jobs" or "run Module 03 for JOB-1234, JOB-1240 and JOB-1251". For one job, run Module 03 directly as usual; this file is not needed.

You are the main session. Helpers do the Module 03 work, one job each. Your job is to pick the jobs, launch the helpers, and record the results. Every call you make re-reads this whole conversation, so keep it small: never read Module 03, a job description or a resume yourself, and never paste helper output beyond the one result line.

## Before you start

1. The helper agent must exist: an agent named `resume-helper` in your list of agent types. It comes from `application-workflow/.claude/agents/resume-helper.md`. It loads when Claude Code is opened in the `application-workflow` folder; to use it from any folder, copy the file to `~/.claude/agents/`. Agents load when a session starts, so a newly copied file needs a new session. If `resume-helper` is not available, tell the candidate and stop. Do not fall back to a general-purpose helper or a fork: they carry a much larger context on every call and cost several times more.
2. You need the tracker file and the Candidate Profile file. If the candidate has not named them, ask.
3. Ask the candidate to close the tracker in Excel before the end of the run; the tracker update refuses to write while it is open.
4. Find a Python that can write Word files, as part of the first shell call in step 1: `python -c "import docx; print('docx ok')"` (`python3` on a Mac). If that fails, try the virtual environment in the `application-workflow` folder the same way: `venv/bin/python` on a Mac, `venv\Scripts\python.exe` on Windows. Use the first one that prints "docx ok" as the DOCX Python in every launch text. If none does, tell the candidate to run `pip install python-docx` (`pip3` on a Mac) and stop.

## 1. Pick the jobs

Steps 1 to 3 take at most two shell calls. If the candidate named Job IDs, run everything in one call:

`python Tools/tracker_tool.py list "<tracker>" --jobs JOB-####,JOB-####; python Tools/resume_path.py sync; python Tools/resume_path.py categories; python Tools/tracker_tool.py related "<tracker>" JOB-#### --limit 3; python Tools/tracker_tool.py related "<tracker>" JOB-#### --limit 3` (one `related` per job)

Otherwise run `python Tools/tracker_tool.py list "<tracker>"` (all RESUME NEEDED jobs) first, then `sync`, `categories` and every `related` together in a second call. (`sync` adds any resume folder missing from the technology list, so the library check can search it; pass its notes on to the candidate.)

(On a Mac use `python3` and forward slashes in every command in this file. Run every command from the `application-workflow` folder.)

The list shows Job ID, title, company, status, Resume Filename and the start of Why Niche. Before going on:
- A job whose status is SKIPPED or CLOSED: ask before including it.
- A job that already has a Resume Filename: ask whether to make another resume for it.
- Two jobs with the same title at the same company (or "Possible duplicate" in the tracker): ask in the plan (step 4) whether to build one resume and copy it to the other job. A second posting of the same job is often shorter, so Module 03's library check can score the first resume below 85% and build a second one for the same job.
  - Yes: launch a helper only for the first job. When it returns BUILT or REUSED, copy its resume yourself: `python Tools/resume_path.py reuse --source "<first job's folder>/<file name>" --tech "<technology>" [--specialty "<specialty>"] --job <second Job ID> --tracker "<tracker>"`, with the technology (and specialty, if any) taken from the first job's folder: the folder names right after `Resumes`. Append a result line for the second job to `results.txt`: REUSED, the file name, the FOLDER line the command printed, the first job's ATS and check results, and "same posting as <first Job ID>; copied at the candidate's request". If the first job stops at the gate or fails, run the second job normally.
  - No: keep both, but run them one after the other, never at the same time, so the second can still reuse the first's resume through the library check.

## 2. Same recruiter check (once for the whole run)

From the `related` output: Module 03's SAME RECRUITER OR AGENCY CHECK says what to tell the candidate. Put every heads-up into the plan message (step 4) instead of asking separately. Default: continue with a fresh resume. Leave out any job the candidate says no to.

## 3. Choose each job's folder

From the `categories` output, for each job choose one technology (and a specialty only when the title or Why Niche clearly points to one) from the job's title and Why Niche line, following Module 03's RESUME LIBRARY CHECK step 1. Do not read job descriptions for this; the helper corrects a clearly wrong choice and reports it.

## 4. Confirm the plan

Show the candidate, in one message: the number of jobs, and per job its Job ID, title and folder; the questions from step 1 and the heads-ups from step 2; and any run setting the candidate gave (for example "assume I have hands-on experience with X"). Wait for one answer that covers all of it.

## 5. Launch the helpers

- Make a run folder: `Run Notes/runs/<YYYY-MM-DD_HHMM>/` inside `application-workflow`. Helpers put their scratch files in `<run folder>/<Job ID>/` and append their result line to `<run folder>/results.txt`.
- One helper per job, with the Agent tool: subagent_type `resume-helper`, a fresh helper every time (never a fork), with the launch text below.
- Start four helpers at once, in one message (fewer if the run has fewer jobs, or if helpers fail on usage or rate limits). Each time a helper finishes, start the next job in that same turn, so no turn is spent only on launching.
- Do not read anything while helpers run, and do not reply between helpers beyond launching the next one. When a helper returns, keep only its result line.
- A helper that returns no result line, or ERROR: launch it once more for that job. If it fails again, report it and go on.

### Launch text

Replace the parts in angle brackets. Leave out the Run setting line when there is none.

```
Run Module 03 for one job: <Job ID>. Folder: technology "<technology>", <specialty "<specialty>" | no specialty> (if it clearly does not fit the job's required skills, use a better one and say so).

Files:
- Module 03: <full path of 3 - Resume Generator.txt> (read it fully once, in parts if needed)
- Tracker (read only for you): <full path of the tracker>
- Candidate Profile: <full path of the profile>
- Run every "python Tools/..." command from <full path of the application-workflow folder> (<Windows | Mac: use python3>). Build the Word file with: <the DOCX Python found before step 1> (it has python-docx).
- Scratch files: <run folder>/<Job ID>/
Run setting: <the candidate's run setting, word for word>

Run Module 03 in full for this job: Industry-Fit Gate, employer tool evidence check, library check, generation, saving with resume_path.py path, and resume_check.py until 0 FAIL and 0 WARN.
Batch rules (they override Module 03 where they differ): do not write the tracker (skip TRACKER UPDATE and the tracker step in gate rule 2); skip the SAME RECRUITER OR AGENCY CHECK; you cannot ask anything, so take the default that lets the job continue and note it; skip the RESUME PATH line.

When done, append your result line to <run folder>/results.txt, then reply with only that line:
JOB-#### | BUILT or REUSED or GATE-STOP or ERROR | <file name or -> | <full folder path, exactly as resume_path.py printed it on its FOLDER line, or -> | ATS <n>% | check FAIL <n> / WARN <n> | <notes: gate reason, shortfalls, omitted tools, assumptions; under 300 characters>
```

## 6. Record the results (one call)

When every helper is done, read `<run folder>/results.txt` once, then update the tracker for all jobs in one shell call, chaining one command per job:

- BUILT: `python Tools/tracker_tool.py set "<tracker>" JOB-#### --set "Resume Filename=<file>" --set "Resume Path=<folder>" --set "Pipeline Status=RESUME CREATED" --note "Resume generated"`
- REUSED: the same, with `--note "Resume reused from the library"`
- GATE-STOP: `python Tools/tracker_tool.py set "<tracker>" JOB-#### --note "Resume not generated: no employer domain supports <core requirement from the notes>"` (no status change)
- ERROR: no tracker change.

If a result line's folder is not a full path, leave out `--set "Resume Path=..."` for that job: given only the Resume Filename, the tool finds the file in that job's folder and stores its folder.

The tool never moves Pipeline Status backward and does not change a SKIPPED or CLOSED job; report what it prints. If it refuses because the tracker is open in Excel, ask the candidate to close it and run the same call again.

## 7. Report to the candidate

One table: Job ID, title, outcome, ATS score, file name and folder, and the notes (shortfalls, omitted tools, assumptions, folder changes). Then list the tracker changes made and any job that needs the candidate's attention (errors, gate stops, held-back jobs).
