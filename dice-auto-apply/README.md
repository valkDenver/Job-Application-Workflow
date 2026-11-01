# Dice Auto Apply: candidate guide

Applies to Dice jobs from your Master Job Tracker for you. Each application gets that job's own resume, at a natural pace, only in rounds you approve. You ask your AI assistant in plain English; you never need to know the tools underneath.

This folder is separate from `application-workflow/`. The two share only your tracker file and your Resumes folder.

## What it does

- **It only applies to jobs that have a resume.** A job qualifies when its **Resume Path** (preferred) or **Resume Filename** is filled. Resume Path should be the full path of the file on your computer. When it isn't usable, the script searches your Resumes folder by the Resume Filename.
- **It never applies to:**
  - a job already APPLIED, FOLLOW-UP, CLOSED or SKIPPED
  - a job whose resume you already emailed to the recruiter (Module 04, "resume attached")
  - the same posting twice
  - a second job for **the same recruiter when it's a different role**. A similar role, meaning a resume in the same technology folder, is fine.
- **On Dice** it opens the job, clicks Apply, replaces the resume with the job's own file, answers screening questions you've saved, and submits.
- **In the tracker** it marks the job **APPLIED** with today's date and a note "by dice-auto-apply". It writes in batches of 5, so it waits if you have the tracker open in Excel.
- **Anything it can't do, it reports with the next step:** a company site outside Dice, a question with no saved answer, a job that has closed.

## One-time setup

You need Google Chrome and an AI assistant. Two ways to use it:

| Your assistant | How it works |
|---|---|
| **Claude Code** (Claude desktop app, Code tab) or **OpenAI Codex** (or any assistant that can run commands on your computer) | Open the `dice-auto-apply` folder in it. It runs everything for you. |
| **Claude chat** or **ChatGPT** in the browser (or any other chat) | It gives you one line. You paste that line into **Dice Auto Apply.bat** (Windows) or **Dice Auto Apply.command** (Mac). |

Then say: **"Set up Dice auto apply. Follow 5 - Dice Auto Apply.txt."** (Claude Code, Codex and Gemini CLI find that file on their own when opened in this folder.) The assistant then:
1. checks your computer and, with your OK, installs what's missing (including the right Python version),
2. opens Chrome so you can **log in to Dice yourself**. The script never sees your password.
3. shows the defaults (**200 applications a day**, every day, 08:00–16:00, normal pace) and changes any you want,
4. asks the common screening questions (work authorization, rate, relocation…) and saves your answers,
5. runs a 1–2 job **test** so you can watch the browser once.

**Mac:** if the `.command` file won't open, run this once in Terminal: `chmod +x "Dice Auto Apply.command"`.

## Everyday use: just ask

- *"Run all Dice jobs with a resume path, focus on today's runs, up to 25 in this round, leave some idle time."*
- *"Test 3 Workday jobs from the latest run, don't submit."*
- *"Tier 1 and 2 only, 10 now, a 10-minute break every 5 jobs, stop by noon."*
- *"How's it going?"* · *"Stop."* · *"What needs my attention?"*

Before anything is sent, the assistant shows the plan: how many jobs, which were skipped and why, how long it will take, and how many applications you'll have left today. It applies only after you say yes. The round keeps running if you close the assistant. Ask again later, or start a new conversation with **"continue Dice"**.

After a few rounds the assistant will suggest starting a fresh conversation, to keep it fast and cheap. Nothing is lost when you do.

## When something needs you

| You see | What to do |
|---|---|
| NEEDS_INPUT | Answer the questions; the assistant saves them and re-runs the job |
| EXTERNAL_APPLY | Apply on the company's own site (link given) |
| RESUME_NOT_FOUND / MISMATCH | Fix the Resume Path (or Resume Filename) in the tracker viewer, or run Module 03 |
| SAME_RECRUITER (role unclear) | Tell the assistant whether the two roles are similar |
| SUBMIT_UNCONFIRMED | Check Dice > Applied Jobs: it may have gone through |
| COOLDOWN | Dice showed a bot check. Don't use Dice from this tool until the time shown; then check Dice by hand and ask to clear the cooldown |
| tracker updates waiting | Close the tracker in Excel; the updates are written automatically |

## Where things are

- **Settings:** `Tools/dice_apply_config.json`, one file shared by everyone, with the defaults (200 applications a day, every day, 08:00–16:00, normal pace). To change a setting for yourself, ask your assistant or use `config set` (for example `config set max_submits_per_day=150`). Your change is saved in `settings.json` next to the Dice login (below), so the shared file never changes and `git pull` never conflicts. `config show` marks your own settings.
- **Your screening answers:** `screening_answers.json`, kept next to the Dice login (below), outside this folder, because they are personal.
- **The Dice login session, round results and screenshots:** `~/.job-workflow/dice-apply/` (Mac) or `%USERPROFILE%\.job-workflow\dice-apply\` (Windows). Deleting `browser-profile` there logs you out.
- **Tracker backups:** a `Backups` folder next to your tracker, once a day before the first update.

## For maintainers

- **Tests:** `python -m unittest discover -s Tools/tests` (no browser needed).
- **With the TestSample tracker:** `python Tools/tests/make_test_tracker.py "<TestSample.xlsx>" --check`
- **Before every commit:** `python Tools/tests/check_no_personal_data.py`
- **When Dice changes a page,** only `Tools/dice_apply_selectors.py` changes. Each locator is marked `verified <date>` or `UNVERIFIED`.
- **Design and plan** are kept outside the repo: `designs/Dice Auto Apply - Design.md` and `plans/Dice Auto Apply - Implementation Plan.md`.
