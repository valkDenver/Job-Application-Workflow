# Batch resume run (Module 03 for many jobs)

Instructions for the AI assistant when the candidate asks for resumes for several jobs at once, for example "build resumes for all my RESUME NEEDED jobs" or "run Module 03 for JOB-1234, JOB-1240 and JOB-1251". For one job, run Module 03 directly as usual; this file is not needed.

You are the main session. Helpers do the Module 03 work, one job each. Your job is to pick the jobs, launch the helpers, and record the results. Every call you make re-reads this whole conversation, so keep it small: never read Module 03, a job description or a resume yourself (this file gives you the few Module 03 rules you need), and never paste helper output beyond the one result line.

## Before you start

1. The helper agent must exist: an agent named `resume-helper` in your list of agent types. It comes from `application-workflow/.claude/agents/resume-helper.md`. It loads when Claude Code is opened in the `application-workflow` folder; to use it from any folder, copy the file to `~/.claude/agents/`. Agents load when a session starts, so a newly copied file needs a new session. If `resume-helper` is not available, tell the candidate and stop. Do not fall back to a general-purpose helper or a fork: they carry a much larger context on every call and cost several times more.
2. You need the tracker file and the Candidate Profile file. If the candidate has not named them, ask.
3. Ask the candidate to close the tracker in Excel before the end of the run; the tracker update refuses to write while it is open.
4. Find a Python that can write Word files, as part of the first shell call in step 1: `python -c "import docx; print('docx ok')"` (`python3` on a Mac). If that fails and the candidate keeps a virtual environment named `venv` in the `application-workflow` folder, try it the same way: `venv/bin/python` on a Mac, `venv/Scripts/python.exe` on Windows (forward slashes). Use the first one that prints "docx ok" as the DOCX Python in every launch text. If none does, stop and tell the candidate to install it: `python -m pip install python-docx` on Windows, `python3 -m pip install --user python-docx` on a Mac.

## 1. Pick the jobs

Steps 1 to 3 take at most two shell calls. If the candidate named Job IDs, run everything in one call:

`python Tools/tracker_tool.py list "<tracker>" --jobs JOB-####,JOB-####; python Tools/resume_path.py sync; python Tools/resume_path.py categories; python Tools/tracker_tool.py related "<tracker>" JOB-#### --limit 3; python Tools/tracker_tool.py related "<tracker>" JOB-#### --limit 3` (one `related` per job)

Otherwise run `python Tools/tracker_tool.py list "<tracker>"` (all RESUME NEEDED jobs) first, then `sync`, `categories` and every `related` together in a second call. (`sync` adds any resume folder missing from the technology list, so the library check can search it; pass its notes on to the candidate.)

(On a Mac use `python3` and forward slashes in every command in this file. Run every command from the `application-workflow` folder. Chain commands with `;`, never `&&`, so one refusal does not skip the rest.)

The list shows Job ID, title, company, status, Resume Filename and the start of Why Niche. Before going on:
- A job whose status is SKIPPED or CLOSED: ask before including it.
- A job that already has a Resume Filename: ask whether to make another resume for it.
- Two jobs with the same title at the same company (or "Possible duplicate" in the tracker): ask in the plan (step 4) whether to build one resume and copy it to the other job. A second posting of the same job is often shorter, so Module 03's library check can score the first resume below 85% and build a second one for the same job.
  - Yes: launch a helper only for the first job. When it returns BUILT or REUSED, copy its resume yourself: `python Tools/resume_path.py reuse --source "<first job's folder>/<file name>" --tech "<technology>" [--specialty "<specialty>"] --job <second Job ID> --tracker "<tracker>"`. Take them from the first job's folder: the technology is the folder right after `Resumes`; the specialty is the folder between the technology and the `JOB-...` folder, only when there is one. Write the second job's result line to `<run folder>/<second Job ID>/result.txt`: REUSED, the file name, the FOLDER line the command printed, the first job's ATS and check results, and "same posting as <first Job ID>; copied at the candidate's request". If the first job stops at the gate or fails, run the second job normally.
  - No: keep both, but run them one after the other, never at the same time, so the second can still reuse the first's resume through the library check.

## 2. Same recruiter check (once for the whole run)

For each job, `related` prints the earlier jobs of the same recruiter or the same staffing company that already have a resume. Compare each one's title and Why Niche line with this job's (Module 03's SAME RECRUITER OR AGENCY CHECK, in short):
- Nothing listed: nothing to say.
- Same recruiter and a similar technology: tell the candidate, in two lines, which earlier job (title, status, resume file) already has a resume for a similar role. Module 03 would offer to update that resume for this job instead; a batch run cannot, so if the candidate wants that, hold the job back and tell them to run Module 03 for it on its own.
- A different technology: a one-line heads-up naming the recruiter, the earlier job(s), and that they may see two different backgrounds; for a different recruiter at the same agency, lighter wording that names the agency. Ask whether to go ahead with this job.

Put every heads-up into the plan message (step 4) instead of asking separately. Default: continue with a fresh resume. Leave out any job the candidate says no to.

## 3. Choose each job's folder

From the `categories` output, for each job choose the one technology whose name or hints best fit the job's main required skill, judged from its title and Why Niche line, and a specialty only when they clearly point to one. Do not read job descriptions for this; the helper corrects a clearly wrong choice and reports it.

## 4. Confirm the plan

Show the candidate, in one message: the number of jobs, and per job its Job ID, title and folder; the questions from step 1 and the heads-ups from step 2; and any run setting the candidate gave (for example "assume I have hands-on experience with X"). Wait for one answer that covers all of it.

## 5. Launch the helpers

- Make a run folder: `Run Notes/runs/<YYYY-MM-DD_HHMM>/` inside `application-workflow`. Each helper works in `<run folder>/<Job ID>/` and writes its result line there, in `result.txt`.
- One helper per job, with the Agent tool: subagent_type `resume-helper`, a fresh helper every time (never a fork), with the launch text below.
- Start five helpers at once, in one message (fewer if the run has fewer jobs, or if helpers fail on usage or rate limits). Each time a helper finishes, start the next job in that same turn, so no turn is spent only on launching.
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
- Work folder: <run folder>/<Job ID>/
Run setting: <the candidate's run setting, word for word>

Run Module 03 in full for this job: Industry-Fit Gate, employer tool evidence check, library check, generation, saving with resume_path.py path, its final validation, and resume_check.py until 0 FAIL, then every WARN fixed or, if it is a false alarm, kept and named in your notes (as Module 03 says).
Batch rules (they override Module 03 where they differ): do not write the tracker (skip TRACKER UPDATE and the tracker step in gate rule 2); skip the SAME RECRUITER OR AGENCY CHECK; you cannot ask anything, so take the default that lets the job continue and note it; skip the RESUME PATH line.

When done, write your result line to <run folder>/<Job ID>/result.txt, then reply with only that line:
JOB-#### | BUILT or REUSED or GATE-STOP or ERROR | <file name or -> | <full folder path, exactly as resume_path.py printed it on its FOLDER line, or -> | ATS <n>% | check FAIL <n> / WARN <n> | <notes: gate reason, WARNs kept as false alarms, shortfalls, omitted tools, assumptions; under 300 characters>
```

## 6. Record the results (one call)

When every helper is done, read all the `<run folder>/*/result.txt` files in one call. Check them against the result lines the helpers returned; where a file is missing or differs, use the returned line. Then update the tracker for all jobs in one shell call, one command per job, chained with `;`:

- BUILT: `python Tools/tracker_tool.py set "<tracker>" JOB-#### --set "Resume Filename=<file>" --set "Resume Path=<folder>" --set "Pipeline Status=RESUME CREATED" --note "Resume generated"`
- REUSED: the same, with `--note "Resume reused from the library"`
- GATE-STOP: `python Tools/tracker_tool.py set "<tracker>" JOB-#### --note "Resume not generated: no employer domain supports <core requirement from the notes>"` (no status change)
- ERROR: no tracker change.
- A job the candidate approved in step 1 although it was SKIPPED or CLOSED: add `--confirmed-closed`, or the tool refuses it. Its status stays as it is.
- A job that already had a Resume Filename: add the earlier file name to the note, for example `--note "Resume generated (replaces <earlier file name>)"`.

If a result line's folder is not a full path, leave out `--set "Resume Path=..."` for that job: given only the Resume Filename, the tool finds the file in that job's folder and stores its folder.

The tool never moves Pipeline Status backward; report anything it prints that is not a plain success. If it refuses because the tracker is open in Excel, ask the candidate to close it and run the same call again.

## 7. Report to the candidate

One table: Job ID, title, outcome, ATS score, file name and folder, and the notes (WARNs kept as false alarms, shortfalls, omitted tools, assumptions, folder changes). Then list the tracker changes made and any job that needs the candidate's attention (errors, gate stops, held-back jobs).
