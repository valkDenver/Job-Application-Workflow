# Job Application Workflow: candidate guide

For people and for AI assistants (Claude or ChatGPT). Read this file first.

Three prompts work in sequence on one Excel tracker, and the jobs come from a shared job feed that someone else keeps up to date:

| Step | What | File |
|---|---|---|
| Jobs | Copy new jobs from the shared feed into your tracker | `Master Job Tracker - SHARED.xlsx` (see [Get jobs](#get-jobs-from-the-shared-feed)) |
| Module 02 | Vendor call | `2 - Vendor Call Assistant.txt` |
| Module 03 | Resume | `3 - Resume Generator.txt` |
| Module 04 | Recruiter email and follow-ups | `4 - Recruiter Follow-up Email.txt` |

There is no Module 01 here: the job search is run by the person who publishes the feed. Nothing is sent or applied to automatically; you review and send everything. (Optional: to apply to Dice jobs automatically in rounds you approve, see [`../dice-auto-apply/README.md`](../dice-auto-apply/README.md).)

## Quick path

1. Get this folder (clone the repository or download the ZIP) and open a terminal in the `application-workflow` folder. Every `Tools/...` command in this guide runs from there. Install Python 3 ([python.org](https://www.python.org); on Windows tick "Add to PATH"), then once in that terminal: `pip install openpyxl` (Mac: `pip3`). Check with `python Tools/tracker_tool.py --help` (Mac: `python3`).
2. Rename `Master Job Tracker - TEMPLATE.xlsx` to `Master Job Tracker - <your name>.xlsx` and delete the grey EXAMPLE rows on every sheet. The Excluded Jobs sheet stays empty.
3. Copy `Candidate Profile - TEMPLATE.txt`, rename it with your name and fill it in: personal details, employers (most recent first), and the Vendor Call Facts at the end.
4. Connect Gmail (creates drafts only). Claude: Customize > Connectors. ChatGPT: Settings > Apps > Developer mode > Create app. Without it, Module 04 shows the email text to paste.
5. Get the link to the shared feed from the person who runs it, and copy the jobs you want (next section).
6. Optional: save each prompt as its own Claude Project or Custom GPT so you do not paste it every time.

Never share a filled-in tracker or profile; they contain personal data.

## Get jobs from the shared feed

The feed is "Master Job Tracker - SHARED.xlsx" in a shared Google Drive folder. It is view only, and nothing you do in it affects anyone else.

1. Open or download the file. Take a copy if you want to keep it open while it updates.
2. Find the jobs you do not have yet. Feed Job IDs start at JOB-1001 and go up, so they never clash with IDs in your own tracker (yours are below 1001). Anything higher than the last JOB number you copied is new.
3. Copy those rows from the "Master Tracker" sheet into your own Master Tracker (same column order), and the matching rows from the "JD Store" sheet into your JD Store, so each job keeps its full description.
4. Work the jobs in your own tracker. Set your own Pipeline Status, dates and notes there.

Good to know:
- New jobs are added after each search run, on weekdays (about 7:00 and 11:00 Mountain time). The feed is only as fresh as the last run.
- The feed only ever adds jobs. If a job closes or a recruiter email is added later, the feed does not show it. Check the job link before you spend time on it.
- Pipeline Status, Notes and similar columns show the publisher's values when the job was added, not your progress.
- Do not edit the shared file or save over it. If something looks wrong, tell the person who runs the feed.
- What the feed looks for: contract jobs posted in the last 24 hours, up to Senior level, niche technologies first, pay of $55/hr or more (a job whose stated hourly maximum is below $55 is excluded). Excluded: W2-only, no C2C, US-citizen or green-card or clearance-only, full-time or permanent only, Canada or outside the U.S., Architect/Lead/Manager levels, AI-training gig listings, postings no longer accepting applications, more than 13 years of experience, face-to-face interviews. Workday, Software Engineer, AI/ML, Java and front or back end roles are also kept when the hourly rate is $55 or more and C2C is not ruled out (they rank last).
- The Priority column shows the tier (Tier 1 is best). Tier 4 or 5 means something is unverified, usually rate or C2C: confirm it with the recruiter before spending time. The Tier Reason column says what is missing.

## How to run (copy-paste starters; attach the files named in brackets)

**Browse the jobs (optional).** Open `Tools/Job Tracker Viewer.html` in Chrome or Edge, press "Open workbook..." and pick your tracker. Filter by priority, niche, portal, pipeline status, C2C status, work arrangement, date found, minimum hourly rate, and whether the recruiter's name, email or phone is known. Click a job for every field and the full description; arrow keys (or j and k) move between jobs, "Pick" collects Job IDs, and the "Copy Module command" buttons copy a ready-to-paste line such as `Run Module 03 (Resume Generator) for JOB-1014`. After you add new jobs, press Reload. Nothing is uploaded. Opened on its own the viewer is read-only; for editing see [the tracker viewer](#the-tracker-viewer-so-you-never-open-excel).

**Choose.** Give the Job IDs you want to the next module. You can also mark jobs SHORTLIST (optional; no module requires it).

**Vendor call (Module 02)**: `Prepare JOB-1014 for vendor call.` [attach the tracker and your Candidate Profile]
- The briefing checks the tracker: if this recruiter already has one of your resumes it tells you, and for a similar role it offers to use that resume as the background for the call (the answers still follow this job's description).
- Practice or answer questions during the call. When it ends, type `The call is finished.`

**Resume (Module 03)**: `Run Module 03 for JOB-1014.` [attach the tracker and your Candidate Profile]; can be done before or after the call.
- It checks the matching technology folder first and reuses a resume scoring 85% or more by copying it into the job's folder; otherwise it builds a Word resume, runs `Tools/resume_check.py` on it until it passes, and shows the ATS score.
- Resumes are saved as `Documents/Job Application Workflow/Resumes/<Technology>/<Specialty>/<JOB-#### - Client (via Staffing company)>/<Name>_<Position title>.docx`, the same way on Windows and Mac.
- Industry-fit check: before writing anything it checks that at least one of your employers' industries can realistically use the job's core tools. If none can, it stops and records "Resume not generated" in the tracker Notes instead of inventing experience.
- Heads-up check: before it creates a new resume it looks for earlier resumes made for the same recruiter or staffing company. A different technology gets a one-time warning; a similar role offers to update that resume instead.
- It ends with a line `RESUME PATH: <full path or link>`. Copy it, open the job in the viewer, press Paste next to Resume path, then Save. The prompts never write the Resume Path column themselves.

**Email (Module 04)**: `Run Module 04 for JOB-1014.` [attach the tracker and your Candidate Profile]. It asks what happened with the recruiter and writes a different email for each case. The resume is attached only after you have actually reached the recruiter.
1. Called and they picked up: call follow-up email, resume attached.
2. Called and no answer: "tried reaching you" email, no resume.
3. No phone number: short introduction email, no resume.
4. They replied or asked for your resume: reply email, resume attached.

- Before drafting a new email it searches your Gmail sent mail for that recruiter's address. A similar earlier role is just mentioned; a different technology gets a warning and a question. Unsent drafts never count as sent.
- If a phone number is listed, call first, then run Module 04 and tell it what happened. Each run gives 3 differently worded versions; pick one and a Gmail draft is created. Attach the resume if the connector could not (cases 1 and 4), review, and send it yourself.
- Every first email has a short friendly intro line, states your years of experience matched to the job description, names 2 or 3 of the job's required tools, and gives one plain reason the role interests you, in 90 to 120 words. `Tools/email_check.py` checks each email before the draft is created.
- Emails end with your first name and email only. Say "include my phone" (or LinkedIn) to add them.

**Follow up (Module 04)**: `Follow up on JOB-1003.` or several Job IDs, or `Follow up on the emails from 2026-10-22.` or `What needs a follow-up?`
- It reads the Gmail conversation with that recruiter (read-only) and decides from who wrote last. No reply yet: a short reply in the same thread, no resume, asking the hourly rate range (and whether it is C2C) when the tracker does not have them, otherwise asking about visa limitations. The recruiter asked for your resume: a reply with the resume attached (or it tells you to run Module 03 first). The recruiter asked a question: a reply using only facts from your profile.
- It never drafts when your email was never sent, when the recruiter said the role is filled, or after 2 follow-ups with no reply (no limit once the recruiter has replied). Up to 10 jobs per run; one email per recruiter. Drafts only; you review and send.
- APPLIED means your resume was submitted by any route: a portal, or a case 1 or 4 email with the resume attached. An email without a resume (case 2 or 3) is not an application.

## The tracker viewer (so you never open Excel)

Double-click `Start Tracker Viewer.bat` (Mac: `Start Tracker Viewer.command`; if it will not open, run once in Terminal: `chmod +x "Start Tracker Viewer.command"`). The first time it asks for your tracker file (drag it onto the window, press Enter); after that it remembers it. Your browser opens the viewer. Keep the black window open while you work; close it (or press Ctrl+C) when done.

- **Editing:** open a job and use "Edit this job" to save straight into the workbook: mark SHORTLIST, mark APPLIED today, skip or reopen a job (with a reason), set the applied and follow-up dates, set the Resume Path, fix the recruiter name, email and phone, and add a dated note. It refreshes itself when a module updates the tracker.
- **Unseen jobs:** every job starts as unseen. A job counts as seen only when you act on it: change its status, skip it, mark it SHORTLIST or APPLIED, edit its dates, recruiter, resume path or notes, pick it, copy a module command for it, or press "Mark seen". Merely opening a job does not count. The "Only unseen" filter works as your inbox. The seen list is saved in a small file next to your tracker (for example `Master Job Tracker.seen.json`); it holds only Job IDs, but do not share it.
- **Progress chips:** each card shows Call, Resume and Email. Filled means done (hover for the date), dashed means not yet. They come from the Vendor Call Date, Resume Filename or Resume Path, and Email Drafted Date columns. "Drafted" means the Gmail draft was made, not that you sent it. The small x beside a done step clears it if it was added by mistake.
- **Why a tier:** hover the tier badge on a card to see what is missing, for example "Missing: rate partial; no recruiter phone".
- **Compact** shows two-line cards so about twice as many jobs fit on screen.
- **Good to know:** if the workbook is open in Excel, saves are refused (close Excel first). If a module saved the tracker after you opened the page, the viewer says so, reloads, and you repeat the edit; nothing is overwritten. Only your own computer can reach the helper. Without the helper (opening the file directly, or in Safari) the viewer is read-only.

## Skipping a job

`Skip JOB-1014: rate too low.` The job stays in the tracker (so you do not copy it in again), its Pipeline Status becomes SKIPPED, your reason goes in the Skip Reason column (as "Manual: ...") and a dated line goes in Notes. To undo it: `Reopen JOB-1014.` In the viewer use "Skip...". By hand in Excel: set Pipeline Status to SKIPPED, write "Manual: <reason>" in Skip Reason, add a dated line to Notes.

## For AI assistants

- The Master Tracker is the single source of truth. Every job has one permanent Job ID (JOB-0001...; feed jobs are JOB-1001 and up). Never change or reuse an ID.
- Pipeline Status moves forward only: REVIEW > SHORTLIST (optional) > CALL READY > VENDOR CALLED > RESUME NEEDED > RESUME CREATED > EMAIL DRAFTED > APPLIED > FOLLOW-UP > CLOSED (SKIPPED = dropped). Modules 02 to 04 never move it backward and ask before working on a SKIPPED or CLOSED job.
- Resume Path is filled in by the user only; no module writes to it.
- Module 02 reads the description and logs the call date and what was claimed on the call. Module 03 reads the description and the call notes and writes the Resume Filename. Module 04 reads the recruiter email and phone, and (for cases 1 and 4 only) the resume filename, asks which of its four situations applies, and creates the draft.
- If a prompt asks for something that is missing (Job ID, tracker, profile, a recruiter email), ask the user; never guess or invent it. Obey each module's "must NOT" list: never send email, never apply to jobs, never invent contact details, pay, work authorization or experience.
- If you cannot read or write .xlsx files, say so and use CSV files with the same column names as the Master Tracker sheet.
- Edits use `python Tools/tracker_tool.py edit "<tracker file>" JOB-#### --set "Field=Value" [--note ...] [--reason ...]`. Fields: Pipeline Status, Vendor Call Date, Resume Filename, Resume Path, Applied Date, Follow-Up Date, Recruiter Name, Recruiter Email, Recruiter Phone. It writes the dated Notes lines other modules rely on (for example "recruiter email ... supplied by me", which Module 04 uses to choose "Hi," or "Hi Name,"). It refuses a Recruiter Email that is not a plain address, a Recruiter Name that is description text, and a Recruiter Phone with no number. Put where a contact came from in `--note`, not in the cell.
- Skip: `python Tools/tracker_tool.py skip "<tracker file>" JOB-1014 --reason "<reason>"` (undo: `reopen`). A job already APPLIED or in FOLLOW-UP needs the user's OK first (`--force`). To count skipped jobs: `skipped`.
- Check contact cells across the tracker: `python Tools/tracker_tool.py check-contacts "<tracker file>"`. A "name does not match the address" line is not an error; look at that job by hand.

## Files

| File | Purpose |
|---|---|
| `2`–`4 - *.txt` | The three prompts |
| `Master Job Tracker - TEMPLATE.xlsx` | Your tracker (Master Tracker, JD Store, Excluded Jobs, Dashboard) |
| `Candidate Profile - TEMPLATE.txt` | Personal details, employers, vendor-call facts |
| `Start Tracker Viewer.bat` / `.command` | Open the viewer with editing |
| `Tools/Job Tracker Viewer.html` | Browse and filter your tracker (read-only on its own) |
| `Tools/tracker_server.py` | The small local helper behind the launchers (Python + openpyxl) |
| `Tools/tracker_tool.py` | Reads or changes ONE job without loading the whole workbook: get, summary, related, followups, drafted-notes, skipped, check-contacts, set/edit, skip/reopen, add-columns, extend-ranges, plus commands used by the job search |
| `Tools/contact_clean.py` | The one definition of a clean Recruiter Name, Email and Phone cell. `python Tools/contact_clean.py --selftest` checks it |
| `Tools/email_check.py` | Plans an email batch and checks each email (years claim, JD tools, word count, variety). Without it Module 04 checks by hand |
| `Tools/resume_path.py` | Chooses the resume folder, saves the resume, finds and copies reusable resumes, lets Module 04 find a job's resume file |
| `Tools/resume_check.py` | Checks the finished Word resume: layout, bullets, wording rules, tool limits, word count, ATS score. Without it Module 03 checks by hand |
| `Tools/resume_categories.json` | The technology folders for resumes; edit it to rename a folder |

## Troubleshooting

- **Viewer:** the tracker needs a sheet named "Master Tracker" with a "Job ID" column in row 1 (and, for descriptions, a "JD Store" sheet with "Job ID" and "Full Job Description"). Columns are matched by header name, so do not rename them. Extra columns appear under "Other" but get no filter. Chrome and Edge remember the last file; other browsers ask each time. If opening or Reload fails while the tracker is open in Excel, close it there or open a copy.
- **"I cannot access the tracker":** attach the .xlsx again, or paste its Master Tracker sheet as CSV.
- **Recruiter emails** are often missing from postings. The modules never guess an address; add it by hand if Module 04 reports it missing.
- **Gmail draft has no attachment:** expected for large files; the module tells you which file to attach.
- **Possible duplicates:** a job on more than one portal (often under different staffing firms) has separate rows marked "Possible duplicate of JOB-####: Review". Compare them yourself.
- **Resumes:** bullets are generated from the job description. Confirm that every statement is true for you before using it, and review every email before sending.

## Moving from an older tracker

Make a backup of your tracker and profile first, and keep your own files; the template is only for a new, blank tracker (never copy it over a working one).

1. Install Python and `pip install openpyxl` (above).
2. If your tracker lacks the Skip Reason, Tier Reason and Email Drafted Date columns, or its filters stop at row 1000: close the tracker in Excel, then run `python Tools/tracker_tool.py add-columns "<your tracker>.xlsx"` (adds Skip Reason, Tier Reason and Email Drafted Date) and `python Tools/tracker_tool.py extend-ranges "<your tracker>.xlsx"` (stretches filters, dropdowns and the Dashboard to row 100000). Both are safe to repeat.
3. If you type statuses in Excel, set the Pipeline Status dropdown (Data > Data Validation > List) to `REVIEW,SHORTLIST,CALL READY,VENDOR CALLED,RESUME NEEDED,RESUME CREATED,EMAIL DRAFTED,APPLIED,FOLLOW-UP,CLOSED,SKIPPED`. Old FOUND and SELECTED rows still work.
4. Fill in Work authorization and the other Vendor Call Facts in your profile. Follow-up emails use it only when a recruiter asks.
5. Resumes you made earlier are not in the current folder structure, so the first Module 03 run for each technology builds a new one. To reuse an older resume, copy it into `Documents/Job Application Workflow/Resumes/<Technology>/<any job folder>/`.
6. If you saved the prompts as Claude Projects or Custom GPTs, paste in the new text.

Known limits: the helper scripts were tested on Windows and in a Chromium-based browser. Not tested: Safari or Firefox with editing, the Mac launcher, and opening an edited workbook in Excel afterwards, so try those once on a copy. Gmail connectors differ in what they allow; if yours cannot search, create threaded drafts or attach files, the prompt says so and falls back to showing the text or telling you which file to attach.
