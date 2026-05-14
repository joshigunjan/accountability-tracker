# Accountability Tracker v10

A lightweight local macOS self-accountability and overtime management system.

This project is no longer SAP-focused. It is designed to help you:

- plan the day in the morning
- answer a 30-minute accountability check
- track todos and task focus
- classify outside-hours activity as work, private, break, or distracted
- keep a live overtime workbook
- review the day and carry unfinished tasks forward


## v10 overtime sheet fix

Version 10 fixes the messy `overtime.xlsx` issue. When the tracker opens the overtime workbook, it now:

- removes repeated header rows
- converts older legacy overtime rows into the current format
- keeps the sheet name as `Overtime`
- appends new rows without overwriting previous overtime
- writes a small `Summary` sheet with total hours and row count

## Existing overtime is preserved

The default data folder is still:

```bash
~/Documents/autotime_sap
```

This is intentional because earlier versions already stored overtime here:

```bash
~/Documents/autotime_sap/overtime.xlsx
```

New overtime rows are appended to the existing `Overtime` sheet. The file is not overwritten.

## Install

From the repo folder:

```bash
brew install uv
uv sync
```

## Start

For 30-minute accountability checks:

```bash
./start_accountability.sh 30
```

The LaunchAgent checks every minute, but the visible prompt appears only when needed.

## Stop

```bash
./stop_accountability.sh
```

## Daily workflow

### Morning setup

The first prompt asks what kind of day it is:

- Regular work day
- Holiday / irregular day
- Leave
- Sick day

For a regular work day, enter your planned work window, for example:

```text
09:30-18:00
```

Then enter your top priorities for the day.

If yesterday had unfinished todos, v9 asks whether to carry them into today.

### During planned work hours

Every 30 minutes it asks:

```text
What are you working on?
```

You can choose from today's todos, add a new task, choose admin/meetings, break, or distracted.

If you choose `Distracted / not productive`, it asks what distracted you and saves the reason.

### Outside planned work hours or irregular days

It asks:

```text
Is this work, private, break, or distracted?
```

If you choose **Work**, it asks which task this work belongs to and appends 30 minutes to:

```bash
~/Documents/autotime_sap/overtime.xlsx
```

If you choose **Private**, **Break**, or **Distracted**, it is logged but not counted as overtime.

## Overtime guardrails

The tracker warns you when overtime is getting high:

- 1+ hour today: shows a warning note
- 2+ hours today: asks whether it is really necessary before logging
- 5+ hours this week: shows a weekly warning
- after 22:00: asks whether the work is urgent or avoidable

## v9 features

### 1. End-of-day review

After your planned work window ends, the tracker shows a summary and asks:

- Did you complete your top priorities?
- What distracted you?
- Any overtime today? Was it justified?
- One thing to improve tomorrow?

Saved to:

```bash
~/Documents/autotime_sap/daily_review.csv
```

Run manually:

```bash
uv run python accountability_prompt.py --review
```

### 2. Daily summary

Shows a summary of the day:

- planned work window
- accountability checks answered
- productive/work blocks
- break/private blocks
- distracted blocks
- overtime hours
- main task
- top priorities

Run manually:

```bash
uv run python accountability_prompt.py --summary
```

### 3. Distraction tracking

If you mark a block as distracted, it asks for the reason:

- YouTube / videos
- Browsing / rabbit hole
- Phone / WhatsApp
- Unclear task
- Tired / low energy
- Meeting fatigue
- Other

Saved to:

```bash
~/Documents/autotime_sap/distractions.csv
```

### 4. Todo carry-over

In the morning, unfinished todos from yesterday can be carried into today.

Run manually:

```bash
uv run python accountability_prompt.py --carry-over
```

## Useful commands

Force today's setup again:

```bash
uv run python accountability_prompt.py --day-setup
```

Add a meeting/focus/overtime block manually:

```bash
uv run python accountability_prompt.py --add-block
```

Example:

```text
20:00-21:00 HFMI paper revision
```

Open the data folder:

```bash
./open_accountability.sh
```

Open overtime directly:

```bash
open ~/Documents/autotime_sap/overtime.xlsx
```

List today's todos:

```bash
uv run python accountability_prompt.py --list-todos
```

Add a todo:

```bash
uv run python accountability_prompt.py --add-todo "HFMI paper revision"
```

Mark a todo done:

```bash
uv run python accountability_prompt.py --done 3
```

## Files created

```text
~/Documents/autotime_sap/day_status.csv        # day type and planned work window
~/Documents/autotime_sap/daily_plan.csv        # morning top priorities
~/Documents/autotime_sap/todos.csv             # todos for the day
~/Documents/autotime_sap/autotime_log.csv      # accountability log
~/Documents/autotime_sap/overtime.xlsx         # live overtime workbook
~/Documents/autotime_sap/private_ignored.csv   # private/break/distracted outside-hours activity
~/Documents/autotime_sap/distractions.csv      # distraction reasons
~/Documents/autotime_sap/daily_review.csv      # end-of-day review
```

## Troubleshooting

Check whether the LaunchAgent is running:

```bash
launchctl list | grep accountability
```

Check errors:

```bash
cat ~/.accountability_tracker/launchd.err.log
```

Run one prompt manually:

```bash
cd ~/accountability-tracker
uv run python accountability_prompt.py
```

If `overtime.xlsx` is open in Excel/Numbers and cannot be written to, the script saves a fallback row in:

```bash
~/Documents/autotime_sap/overtime_pending.csv
```
