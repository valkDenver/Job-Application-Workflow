# Module 05 request test sheet (plan, Phase 6)

Run each request on Claude Code and OpenAI Codex (Tier A), and on a Claude chat and ChatGPT (Tier B). Use the tracker copy and test rounds only.

For each request, check:
- the options match the expected ones
- the plan is explained in plain words, with any interpretation stated
- nothing is submitted before approval

Where models disagree, fix the wording in `5 - Dice Auto Apply.txt`, not the model.

| # | Request | Expected plan options | Also check |
|---|---|---|---|
| 1 | Run all Dice jobs having resume path updated in excel tracker and focus on todays runs upto 25 in this round and leave some idle time | `--run-date today --limit 25 --pace relaxed` | asks once; says how "today's runs" was read |
| 2 | Just test 3 Workday jobs from the latest run | `--run-latest --title-contains Workday --limit 3`, then `--mode dry-run` | never submits |
| 3 | Tier 1 only, 10 now, a break every 5 | `--priority 1 --limit 10 --break-every 5` (asks or uses default break length) | |
| 4 | Apply to JOB-1014 and JOB-1020 | `--jobs JOB-1014,JOB-1020` | still shows the plan first |
| 5 | Apply to 20, stop by noon | `--limit 20 --until 12:00` | |
| 6 | Do yesterday's runs, go fast | `--run-date yesterday --pace brisk` (asks how many if no max_per_round) | |
| 7 | Only the ones I already called, 5 of them | `--status "VENDOR CALLED" --limit 5` | |
| 8 | Java roles found this week, 15 | `--found last:7d --title-contains Java --limit 15` | |
| 9 | Apply to 100 (with 30 left today) | trimmed to 30 | says how many are left; does not raise the limit |
| 10 | Run some Dice jobs | asks how many (no max_per_round) or uses max_per_round | does not guess "apply" without the plan |
| 11 | How's it going? (during a round) | `status` | one-sentence answer; no repeated checks |
| 12 | Stop | `stop` | says the current job finishes first |
| 13 | What needs my attention? (after a round) | `report` | explains each next step |
| 14 | A second round request while one runs | `status` only | does not start another round |
| 15 | Request while a COOLDOWN is active | `status` only | explains the cooldown; never touches Dice |
| 16 | "Raise my daily limit to 200" | `config set max_submits_per_day=200` | only because the user asked in words |
| 17 | The job asked "Expected hourly rate?", my answer is 75 | `config answer "Expected hourly rate?" "75"` then offers a re-run | |
| 18 | 4th round in the same conversation | suggests a fresh conversation once | "continue Dice" in a new chat runs status + report |
| 19 | Today's and yesterday's runs, Workday or ServiceNow | asks: two run dates are two rounds, or uses `--found last:2d`, and says so | ambiguity is stated, not guessed |
| 20 | Set up Dice auto apply | `doctor`, `doctor --fix` (with OK), `login`, `config set ...`, `config answer ...`, test round | |
