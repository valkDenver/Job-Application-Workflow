# Job Application Workflow: candidate guide

For people and for AI assistants. Read this file first.

**Where to run it: Claude Code**, in the Claude desktop app's Code tab (or the `claude` command line), opened in this `application-workflow` folder. It runs the workflow's checks and scripts on your computer, saves your resumes in the standard folders and updates your tracker. Chat apps (Claude Projects, ChatGPT, Custom GPTs) are limited: see [Chat apps (limited)](#chat-apps-limited).

Three prompts work in sequence on one Excel tracker, and the jobs come from a shared job feed that someone else keeps up to date:

| Step | What | File |
|---|---|---|
| Setup (once) | Check the computer, create your tracker | `0 - Setup.txt` |
| Profile | Build or update your Candidate Profile, from your resume or by answering questions | `0b - Profile.txt` |
| Jobs | Copy new jobs from the shared feed into your tracker | `Master Job Tracker - SHARED.xlsx` (see [Get jobs](#get-jobs-from-the-shared-feed)) |
| Module 02 | Vendor call | `2 - Vendor Call Assistant.txt` |
| Module 03 | Resume | `3 - Resume Generator.txt` |
| Module 04 | Recruiter email and follow-ups | `4 - Recruiter Follow-up Email.txt` |

There is no Module 01 here: the job search is run by the person who publishes the feed. Nothing is sent or applied to automatically; you review and send everything. (Optional: to apply to Dice jobs automatically in rounds you approve, see [`../dice-auto-apply/README.md`](../dice-auto-apply/README.md).)

## Quick path

1. Install the Claude desktop app ([claude.ai/download](https://claude.ai/download)) and sign in; Claude Code needs a Pro or Max plan. You use its **Code** tab for everything below.
2. Get this folder from GitHub: `git clone https://github.com/valkDenver/Job-Application-Workflow.git` (recommended: `git pull` then gets every update, see [Getting updates](#getting-updates)). Without git, use **Code > Download ZIP** on that page and unzip it; to update, download it again and move your tracker, profile and `Resumes` folder across. Use no other copy of the workflow.
3. Install Python 3 ([python.org](https://www.python.org); on Windows tick "Add to PATH") and Google Drive for desktop ([google.com/drive/download](https://www.google.com/drive/download/)), signed in with your Google account: your tracker lives in your Google Drive, shared with valktech.den@gmail.com as Viewer.
4. Open the `application-workflow` folder in the Code tab and type **`/setup`** (or say "Set up the workflow. Follow 0 - Setup.txt."). It checks your computer (`Tools/setup_check.py`), installs the two packages with your OK, offers to build your Candidate Profile, creates your tracker, checks Gmail, and finishes when the check says "All good." Run `python Tools/setup_check.py` (Mac: `python3`) any time something seems wrong.

Then use the commands in [How to run](#how-to-run) (for example `/get-jobs`, then `/build-resume JOB-1014`) in the Code tab: no pasting prompts or attaching files. Every `Tools/...` command in this guide runs from the `application-workflow` folder, and Claude Code runs them for you.

Never share a filled-in tracker or profile; they contain personal data.

### What the setup does (to do it by hand)

1. Packages: `pip install -r requirements.txt` (Mac: `python3 -m pip install --user -r requirements.txt`). For resumes fitted to exactly three pages you also need Microsoft Word (Windows) or LibreOffice (Mac); without it resumes are still saved at the template's normal size.
2. Tracker (required in Google Drive): install Google Drive for desktop and sign in, then `python Tools/tracker_tool.py init --name "<your name>" --folder "<your Google Drive>/Job Application Workflow"` creates `Master Job Tracker - <Your_Name>.xlsx` there from the template (which stays unchanged), without the grey EXAMPLE rows, and saves its location in `tracker_path.txt`. Then share that one file with **valktech.den@gmail.com as Viewer** (Google Drive on the web: right-click the file > Share), so your progress can be followed. You stay the only one who changes it; your resumes stay on your computer. `setup_check.py` reports a tracker outside Google Drive. Already have a tracker? Save its location with `python Tools/setup_check.py --tracker "<your tracker>" --fix`.
3. Profile: say **"Build my profile"** (it follows `0b - Profile.txt`): it fills the profile from your current resume if you have one, or asks you section by section, drafts each employer's Domain guidance and Seniority notes for you to confirm, and writes the file only after your yes. Say "Update my profile" later to change a part. By hand: copy `Candidate Profile - TEMPLATE.txt`, rename it `Candidate Profile - <Your_Name>.txt` and fill it in.
   - **For your first resume or email** you need only the top part: personal details, your employers (most recent first, dates like `Jan 2020 - Mar 2023` or `Jan 2025 - Present`, unused employer blocks left empty) and the resume layout. Check it with `python Tools/profile_check.py`.
   - **Before your first vendor call**, fill in the Vendor Call Facts (background and study overlap, authorization, rate, availability, standard answers, references; the employer of record comes filled in). Write N/A for anything that does not apply to you and leave a field blank only if it is truly unknown, then set "Facts last confirmed on". `python Tools/profile_check.py --level call` lists what is still blank or out of date, and Module 02 runs it before each call.
4. Gmail: connect it in Claude (Customize > Connectors); Module 04 uses it to create drafts and check your sent mail, never to send. Without it, Module 04 shows the email text to paste.
5. Jobs: get your feed file and starting Job ID from the person who runs the feed, and copy the jobs you want (next section).

## Getting updates

If you cloned the repository, get the latest version with `git pull`, run in the repository folder (the one that holds `application-workflow`).

If `git pull` stops with "Your local changes to the following files would be overwritten" and names `application-workflow/Tools/resume_categories.json`, keep your own copy of that file and pull again (Mac or Git Bash; on Windows PowerShell use `Copy-Item` for `cp`):

```
cp application-workflow/Tools/resume_categories.json ../resume_categories.backup.json
git checkout -- application-workflow/Tools/resume_categories.json
git pull
cp ../resume_categories.backup.json application-workflow/Tools/resume_categories.json
```

This is needed once: from this update on, `resume_categories.json` is your own file and no longer part of the repository, so it never blocks a pull again. For any other file named in that message, ask the person who maintains the workflow before discarding your change.

## Get jobs from the shared feed

The feed is "Master Job Tracker - SHARED.xlsx" in a shared Google Drive folder. It is view only, and nothing you do in it affects anyone else.

1. Have the feed file on your computer. Best: in Google Drive, open "Shared with me", right-click the feed file (or your feed's folder) and choose **Organize > Add shortcut**, then install Google Drive for desktop; the file then stays up to date at the same place on your computer. Otherwise download it after each update (a new download may get a new name, such as "(1)").
2. Close your tracker in Excel, then type **`/get-jobs`** in Claude Code (or say "Import my new jobs"), or run:
   `python Tools/tracker_tool.py import-feed "<your tracker>" --feed "<the feed file>" --after JOB-####` (first time: the starting Job ID you were given; the jobs above it are added)
   After that, `python Tools/tracker_tool.py import-feed "<your tracker>"` is enough: it remembers the feed and adds only the jobs newer than your newest feed job, with their descriptions. It is safe to run again; add `--dry-run` to see what it would add.
3. Work the jobs in your own tracker. Set your own Pipeline Status, dates and notes there. The import never changes a job already in your tracker.

By hand instead: copy the rows above the last Job ID you have from the feed's "Master Tracker" sheet into your Master Tracker (check the column headers line up), and the matching "JD Store" rows into your JD Store. Feed Job IDs start at JOB-1001 and go up, with gaps (those jobs went to other feeds); your own are below 1001.

Good to know:
- New jobs are added after each search run, on weekdays: the runs start at 9:00, 12:00 and 17:00 Mountain time, and their new jobs are usually in the feed by about 10:00, 13:00 and 18:00, sometimes later. Each run adds the jobs posted since the previous one; Monday morning's run also adds the weekend's jobs. The feed is only as fresh as the last run.
- The feed only ever adds jobs. If a job closes or a recruiter email is added later, the feed does not show it. Check the job link before you spend time on it.
- Pipeline Status, Notes and similar columns show the publisher's values when the job was added, not your progress.
- Do not edit the shared file or save over it. If something looks wrong, tell the person who runs the feed.
- What the feed looks for: contract jobs posted in the last 24 hours, up to Senior level, niche technologies first, pay of $55/hr or more (a job whose stated hourly maximum is below $55 is excluded). Excluded: W2-only, no C2C, US-citizen or green-card or clearance-only, full-time or permanent only, Canada or outside the U.S., Architect/Lead/Manager levels, AI-training gig listings, postings no longer accepting applications, more than 13 years of experience, face-to-face interviews. Workday, Software Engineer, AI/ML, Java and front or back end roles are also kept when the hourly rate is $55 or more and C2C is not ruled out (they rank last).
- The Priority column shows the tier (Tier 1 is best). Tier 4 or 5 means something is unverified, usually rate or C2C: confirm it with the recruiter before spending time. The Tier Reason column says what is missing.

## How to run

### Commands

Type these in Claude Code's Code tab, opened in this folder (type `/` to see them all). They find your tracker and Candidate Profile themselves; you never type a file name. A plain sentence works too ("draft the email for JOB-1234"): Claude picks the matching command.

| Command | What it does |
|---|---|
| `/setup` | One-time guided setup: checks this computer, your profile, your tracker in Google Drive, Gmail |
| `/profile` (`/profile update`) | Build or update your Candidate Profile, from your resume or by answering questions |
| `/get-jobs` (first time: `/get-jobs JOB-1234`) | Add the feed's new jobs to your tracker, with their full descriptions |
| `/check-setup` | Check your computer, tracker and profile, and say how to fix anything |
| `/call-prep JOB-1234` | Prepare for a vendor call (Module 02) |
| `/build-resume JOB-1234` (several IDs, or `all`) | Build resumes (Module 03; the batch run for several) |
| `/email-draft JOB-1234` (optional situation 1-4) | Draft the recruiter email in Gmail (Module 04); you review and send |
| `/email-follow-up` (Job IDs, a date, or nothing for what's due) | Draft follow-ups in the same Gmail thread |
| `/email-sent` | Check which drafts you actually sent (records the dates) and list the ones not sent |
| `/my-progress` (optional number of days) | Your progress: jobs, calls, resumes, emails drafted and sent, submissions, applications |

The sentences below still work too.

### Sentences

Type these starters in Claude Code's Code tab, opened in this folder. The files in brackets are the ones the module needs: name them the first time (for example "my tracker is Master Job Tracker - <your name>.xlsx"); there is nothing to attach.

**Browse the jobs (optional).** Open `Tools/Job Tracker Viewer.html` in Chrome or Edge, press "Open workbook..." and pick your tracker. Filter by priority, niche, portal, pipeline status, C2C status, work arrangement, date found, minimum hourly rate, and whether the recruiter's name, email or phone is known. Click a job for every field and the full description; arrow keys (or j and k) move between jobs, "Pick" collects Job IDs, and the "Copy Module command" buttons copy a ready-to-paste line such as `Run Module 03 (Resume Generator) for JOB-1014`. After you add new jobs, press Reload. Nothing is uploaded. Opened on its own the viewer is read-only; for editing see [the tracker viewer](#the-tracker-viewer-so-you-never-open-excel).

**Choose.** Give the Job IDs you want to the next module. You can also mark jobs SHORTLIST (optional; no module requires it).

**Vendor call (Module 02)**: `Prepare JOB-1014 for vendor call.` [your tracker and Candidate Profile]
- The briefing checks the tracker: if this recruiter already has one of your resumes it tells you, and for a similar role it offers to use that resume as the background for the call (the answers still follow this job's description).
- It starts with a fact check: it lists blank or stale Vendor Call Facts (education, where each job was worked, sponsorship, lowest rate, references and so on) so you can fill them before the call. Each answer carries a tag: PROFILE, JOB, CONFIRM or NEEDS JOB.
- Practice or answer questions during the call. Say `practice` for a mock screening, or give no Job ID for general answers without a tracker update. When the call ends, type `The call is finished.`; the Notes then record claims, rate and terms, right to represent, and what you owe the vendor.

**Resume (Module 03)**: `Run Module 03 for JOB-1014.` [your tracker and Candidate Profile]; can be done before or after the call.
- It checks the matching technology folder first and reuses a resume scoring 85% or more by copying it into the job's folder; otherwise it writes the resume content, renders a three-page Word resume in your chosen template with `Tools/resume_render.py`, runs `Tools/resume_check.py` on it until it passes, and shows the ATS score.
- Resume look: set `Resume template:` in your Candidate Profile to classic, modern, compact, executive, two_line, banner, hanging, rail, grid or spotlight (blank = classic). See them in the `Resume Templates` folder (start with `00 - All templates at a glance.pdf`); `python Tools/resume_render.py templates` lists them. All are ATS-safe except hanging (section titles in Word frames) and grid (skills in a two-column table). The renderer fits every resume onto exactly three pages; it needs Microsoft Word on Windows or LibreOffice on a Mac to measure pages, and without them it renders at the template's normal size and says so. Each resume gets a `<name>.json` content file beside it, so a reused resume can be re-rendered in your current template.
- Resumes are always Word files (.docx). Say `also make a PDF` when you run Module 03 if you want a PDF copy too; it is saved beside the Word file, and the tracker keeps pointing to the Word file.
- Resumes are saved as `Documents/Job Application Workflow/Resumes/<Technology>/<Specialty>/<JOB-#### - Client (via Staffing company)>/<Name>_<Position title>.docx`, the same way on Windows and Mac.
- Only technology folders in `Tools/resume_categories.json` are searched for reuse. If you made folders yourself, or the list was reset, run `python Tools/resume_path.py sync` once (`python3` on a Mac): it adds every folder already in your Resumes folder to the list, and never moves or deletes anything. A batch run does this for you.
- Industry-fit check: before writing anything it checks that at least one of your employers' industries can realistically use the job's core tools. If none can, it stops and records "Resume not generated" in the tracker Notes instead of inventing experience.
- Heads-up check: before it creates a new resume it looks for earlier resumes made for the same recruiter or staffing company. A different technology gets a one-time warning; a similar role offers to update that resume instead.
- It records the resume in the tracker: Resume Filename, Resume Path (the folder the resume is saved in, without the file name) and RESUME CREATED. It also ends with a line `RESUME PATH: <folder or link>`; if the tracker update could not run (for example the tracker was open in Excel), copy that line, open the job in the viewer, press Paste next to Resume path, then Save.

**Many resumes at once (batch run)**: `Follow "Run Notes/BATCH_RUN.md" for all my RESUME NEEDED jobs.` (or name the Job IDs) [your tracker and Candidate Profile]
- For Claude Code. It runs Module 03 for each job in a small helper of its own and records every result in the tracker at the end. You answer one question first: the plan, with any duplicate, same-recruiter or already-built job to decide.
- One-time setup: the helper is `.claude/agents/resume-helper.md`. It loads when Claude Code is opened in this folder; to use it from any folder, copy it to `~/.claude/agents/` (Windows: `C:\Users\<you>\.claude\agents\`). Start a new session after copying.
- Close the tracker in Excel before the run ends. In tests the helpers used about a third fewer tokens per resume than general-purpose helpers doing the same jobs, and more in large runs, where the old way's main conversation grows with every helper report.

**Email (Module 04)**: `Run Module 04 for JOB-1014.` [your tracker and Candidate Profile]. It asks what happened with the recruiter and writes a different email for each case. The resume is attached only after you have actually reached the recruiter.
1. Called and they picked up: call follow-up email, resume attached.
2. Called and no answer: "tried reaching you" email, no resume.
3. No phone number: short introduction email, no resume.
4. They replied or asked for your resume: reply email, resume attached.

- Before drafting a new email it searches your Gmail sent mail for that recruiter's address. A similar earlier role is just mentioned; a different technology gets a warning and a question. Unsent drafts never count as sent.
- When a run covers several jobs and some get no draft: a job that does not fit your profile is listed with the reason and marked SKIPPED only after you confirm. Fixable ones (no email, no resume, same recruiter already emailed) are never skipped; they get a dated Notes line with the reason.
- If a phone number is listed, call first, then run Module 04 and tell it what happened. Each run gives 3 differently worded versions; pick one and a Gmail draft is created. Attach the resume if the connector could not (cases 1 and 4), review, and send it yourself.
- Every first email has a short friendly intro line, states your years of experience matched to the job description, names 2 or 3 of the job's required tools, and gives one plain reason the role interests you, in 90 to 120 words. `Tools/email_check.py` checks each email before the draft is created.
- Emails end with your first name and email only. Say "include my phone" (or LinkedIn) to add them.

**Follow up (Module 04)**: `Follow up on JOB-1003.` or several Job IDs, or `Follow up on the emails from 2026-10-22.` or `What needs a follow-up?`
- It reads the Gmail conversation with that recruiter (read-only) and decides from who wrote last. No reply yet: a short reply in the same thread, no resume, asking the hourly rate range (and whether it is C2C) when the tracker does not have them, otherwise asking about visa limitations. The recruiter asked for your resume: a reply with the resume attached (or it tells you to run Module 03 first). The recruiter asked a question: a reply using only facts from your profile.
- It never drafts when your email was never sent, when the recruiter said the role is filled, or after 2 follow-ups with no reply (no limit once the recruiter has replied). Up to 10 jobs per run; one email per recruiter. Drafts only; you review and send.
- APPLIED means your resume was submitted by any route: a portal, or a case 1 or 4 email with the resume attached. An email without a resume (case 2 or 3) is not an application.
- It also records the day you actually sent each first email (Email Sent Date), when it finds it in your sent mail.
- When a recruiter's reply says they submitted you to the client, it asks "Mark JOB-#### as submitted?" and fills **Submitted Date** only after your yes. You can also fill it yourself in the tracker viewer when a recruiter tells you by phone. Module 02 then knows you were already submitted to that client (through any vendor), so you are never submitted twice.

**Check my sent emails (Module 04)**: `Check my sent emails.` Module 04 looks in your Gmail sent mail (read-only) for every job whose email was drafted but not yet recorded as sent, matching each job by its own email thread (not just the recruiter, who may have emails about several jobs), records the Email Sent Date, and lists the drafts you have not sent yet. You can also fill Email Sent Date yourself in the tracker viewer.

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

## Chat apps (limited)

Claude Projects, claude.ai chat, ChatGPT and Custom GPTs can run the module prompts if you paste the prompt and attach your tracker and profile, but they cannot run the workflow on your computer, so:
- the checks the modules run (`tracker_tool.py`, `email_check.py`, `resume_check.py`) are skipped or done by hand, which is less reliable;
- resumes are not rendered to three pages or saved in your `Resumes` folders, and your tracker is not updated: you download files and copy changes back yourself;
- the batch run does not work at all (it needs Claude Code).

Use them only when you cannot use Claude Code, and expect to do more by hand.

## For AI assistants

- The Master Tracker is the single source of truth. Every job has one permanent Job ID (JOB-0001...; feed jobs are JOB-1001 and up). Never change or reuse an ID.
- Pipeline Status moves forward only: REVIEW > SHORTLIST (optional) > CALL READY > VENDOR CALLED > RESUME NEEDED > RESUME CREATED > EMAIL DRAFTED > APPLIED > FOLLOW-UP > CLOSED (SKIPPED = dropped). Modules 02 to 04 never move it backward and ask before working on a SKIPPED or CLOSED job.
- Resume Path holds the folder of the job's resume, without the file name; the file is `<Resume Path>/<Resume Filename>`. Module 03 fills it when it records the resume (`tracker_tool.py set` checks the file is in that folder, or finds it in the job's folder); the user can change it in the viewer. No other module writes to it.
- Module 02 reads the description and your Vendor Call Facts, runs a fact check before the call, and afterwards logs the call date and the call notes under labels (Call claims, Rate and terms, Authorization and availability, Right to represent, Owed to vendor). Module 03 reads the description and only the "Call claims" line of those notes and writes the Resume Filename. Module 04 reads the recruiter email and phone, and (for cases 1 and 4 only) the resume filename, asks which of its four situations applies, and creates the draft.
- If a prompt asks for something that is missing (Job ID, tracker, profile, a recruiter email), ask the user; never guess or invent it. Obey each module's "must NOT" list: never send email, never apply to jobs, never invent contact details, pay, work authorization or experience.
- If you cannot read or write .xlsx files, say so and use CSV files with the same column names as the Master Tracker sheet.
- Edits use `python Tools/tracker_tool.py edit "<tracker file>" JOB-#### --set "Field=Value" [--note ...] [--reason ...]`. Fields: Pipeline Status, Vendor Call Date, Resume Filename, Resume Path, Applied Date, Follow-Up Date, Recruiter Name, Recruiter Email, Recruiter Phone. It writes the dated Notes lines other modules rely on (for example "recruiter email ... supplied by me", which Module 04 uses to choose "Hi," or "Hi Name,"). It refuses a Recruiter Email that is not a plain address, a Recruiter Name that is description text, and a Recruiter Phone with no number. Put where a contact came from in `--note`, not in the cell.
- Skip: `python Tools/tracker_tool.py skip "<tracker file>" JOB-1014 --reason "<reason>"` (undo: `reopen`). A job already APPLIED or in FOLLOW-UP needs the user's OK first (`--force`). To count skipped jobs: `skipped`.
- Check contact cells across the tracker: `python Tools/tracker_tool.py check-contacts "<tracker file>"`. A "name does not match the address" line is not an error; look at that job by hand.
- What was done in a period: `python Tools/tracker_tool.py report "<tracker file>" [--since YYYY-MM-DD] [--until YYYY-MM-DD]` (default: the last 7 days): jobs received, vendor calls, resumes, emails drafted with and without a resume, follow-ups, applied, skipped, and days from Date Found to first contact. It counts the date columns and the dated Notes lines, so every `set --note` starts with the date (added automatically when missing).

## Files

| File | Purpose |
|---|---|
| `0 - Setup.txt` | One-time guided setup |
| `0b - Profile.txt` | Builds or updates your Candidate Profile ("Build my profile") |
| `requirements.txt` | The Python packages the tools need (`pip install -r requirements.txt`) |
| `Tools/profile_check.py` | Checks your Candidate Profile: ready for resumes and emails, or (`--level call`) which Vendor Call Facts are blank or out of date |
| `Tools/setup_check.py` | Checks Python, packages, Word or LibreOffice, your tracker, your profile and the Resumes folder; `--fix` installs the packages and saves your tracker's location |
| `2`–`4 - *.txt` | The three prompts |
| `CLAUDE.md` | Points Claude Code at this README and the prompts when it opens this folder |
| `.claude/skills/<name>/SKILL.md` | The commands (`/setup`, `/get-jobs`, `/build-resume`...), as Claude Code skills; each one follows the prompt or tool it names |
| `.claude/settings.json` | Lets Claude Code run the workflow's checks and its routine, self-protecting updates (tracker updates that only move forward, importing new jobs, saving resumes) without asking each time. Editing a job by hand, skipping or reopening a job, deleting rows and one-time setup changes always ask. The tracker viewer is not affected |
| `Master Job Tracker - TEMPLATE.xlsx` | Your tracker (Master Tracker, JD Store, Excluded Jobs, Dashboard) |
| `Candidate Profile - TEMPLATE.txt` | Personal details, employers, vendor-call facts |
| `Start Tracker Viewer.bat` / `.command` | Open the viewer with editing |
| `Tools/Job Tracker Viewer.html` | Browse and filter your tracker (read-only on its own) |
| `Tools/tracker_server.py` | The small local helper behind the launchers (Python + openpyxl) |
| `Run Notes/BATCH_RUN.md` | Instructions for a batch resume run (many Module 03 jobs in one session) |
| `.claude/agents/resume-helper.md` | The small helper a batch run uses, one per job |
| `Run Notes/token_usage.py` | Totals the tokens of a session and its helpers from the Claude Code logs, to compare runs |
| `Tools/tracker_tool.py` | Reads or changes ONE job without loading the whole workbook: get, list, summary, related, followups, drafted-notes, skipped, check-contacts, set/edit, skip/reopen, add-columns, extend-ranges, plus commands used by the job search; `init --name` creates your own tracker from the template; `import-feed` adds the job feed's new jobs; `report` shows what was done in a period; `unsent` lists drafted emails not yet recorded as sent |
| `Tools/contact_clean.py` | The one definition of a clean Recruiter Name, Email and Phone cell. `python Tools/contact_clean.py --selftest` checks it |
| `Tools/email_check.py` | Plans an email batch and checks each email (years claim, JD tools, word count, variety). Without it Module 04 checks by hand |
| `Tools/resume_path.py` | Chooses the resume folder, saves the resume, finds and copies reusable resumes, lets Module 04 find a job's resume file |
| `Tools/resume_render.py` | Builds the Word resume from Module 03's content file in the candidate's template, fitted to three pages |
| `Resume Templates/` | Sample PDFs of every resume template with a made-up candidate, to help you choose. Rebuild them with `python Tools/resume_samples.py` after changing a template |
| `Tools/resume_templates.json` | The resume templates (fonts, sizes, colours, spacing, layout). Add or adjust a template here, no code change needed |
| `Tools/resume_check.py` | Checks the finished Word resume against its template: layout, bullets, wording rules, tool limits, word count, ATS score. Without it Module 03 checks by hand |
| `Tools/resume_categories.json` | Your technology folders for resumes. Created from a starting list the first time a resume tool runs, then grows as Module 03 adds technologies; it is yours and not part of the repository. Edit it to rename a folder |

## Troubleshooting

- **Viewer:** the tracker needs a sheet named "Master Tracker" with a "Job ID" column in row 1 (and, for descriptions, a "JD Store" sheet with "Job ID" and "Full Job Description"). Columns are matched by header name, so do not rename them. Extra columns appear under "Other" but get no filter. Chrome and Edge remember the last file; other browsers ask each time. If opening or Reload fails while the tracker is open in Excel, close it there or open a copy.
- **"I cannot access the tracker":** check that Claude Code is opened in the `application-workflow` folder and give the tracker's exact file name (or full path). If it says the file is in use, close it in Excel.
- **Recruiter emails** are often missing from postings. The modules never guess an address; add it by hand if Module 04 reports it missing.
- **Gmail draft has no attachment:** expected for large files; the module tells you which file to attach.
- **Possible duplicates:** a job on more than one portal (often under different staffing firms) has separate rows marked "Possible duplicate of JOB-####: Review". Compare them yourself.
- **Resumes:** bullets are generated from the job description. Confirm that every statement is true for you before using it, and review every email before sending.

## Moving from an older tracker

Make a backup of your tracker and profile first, and keep your own files; the template is only for a new, blank tracker (never copy it over a working one).

1. Install Python and `pip install openpyxl` (above).
2. If your tracker lacks the Skip Reason, Tier Reason and Email Drafted Date columns, or its filters stop at row 1000: close the tracker in Excel, then run `python Tools/tracker_tool.py add-columns "<your tracker>.xlsx"` (adds Skip Reason, Tier Reason and Email Drafted Date) and `python Tools/tracker_tool.py extend-ranges "<your tracker>.xlsx"` (stretches filters, dropdowns and the Dashboard to row 100000, and updates an older Dashboard: Applied then counts every job with an Applied Date, so it no longer drops when a job moves to FOLLOW-UP, and the always-empty Excluded Jobs row becomes Skipped). Both are safe to repeat; `python Tools/setup_check.py` tells you when your tracker needs them.
3. If you type statuses in Excel, set the Pipeline Status dropdown (Data > Data Validation > List) to `REVIEW,SHORTLIST,CALL READY,VENDOR CALLED,RESUME NEEDED,RESUME CREATED,EMAIL DRAFTED,APPLIED,FOLLOW-UP,CLOSED,SKIPPED`. Old FOUND and SELECTED rows still work.
4. Fill in Work authorization and the other Vendor Call Facts in your profile (the template now also asks for education, where each job was worked, sponsorship, lowest rate, assessment and in-person stance, references, and a "Facts last confirmed on" date; use real dates, not "in two weeks"). Follow-up emails use it only when a recruiter asks.
5. Resumes you made earlier are not in the current folder structure, so the first Module 03 run for each technology builds a new one. To reuse an older resume, copy it into `Documents/Job Application Workflow/Resumes/<Technology>/<any job folder>/`.
6. If you saved the prompts as Claude Projects or Custom GPTs, paste in the new text.

Known limits: the helper scripts were tested on Windows and in a Chromium-based browser. Not tested: Safari or Firefox with editing, the Mac launcher, and opening an edited workbook in Excel afterwards, so try those once on a copy. Gmail connectors differ in what they allow; if yours cannot search, create threaded drafts or attach files, the prompt says so and falls back to showing the text or telling you which file to attach.
