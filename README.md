# Accountability Tracker

A lightweight personal accountability and overtime-management tool for macOS.

The tracker asks short check-in questions during the day so you can see whether you are spending time on what you planned. It also records work done outside your planned work window as overtime and helps you confirm whether that overtime has already been added to your official time-management system.

This project is intended for personal use. All logs are stored locally on your computer.

---

## What it does

### Daily setup

At the start of the day, the tracker asks:

- Is today a regular work day, holiday, leave, sick day, or irregular day?
- What is your planned work window? For example: `09:30-18:00`
- What are your main tasks or priorities for today?

### During planned working hours

Every 30 minutes, it asks what you are working on.

The answer is saved as a regular accountability check.

### Outside planned working hours

If you are using the laptop before or after your planned work window, it asks whether this is:

- work
- private matter
- break
- distracted / not productive

If you choose work, it asks which task you are working on and adds a 30-minute entry to the overtime sheet.

### Meeting or focus blocks

You can add a time block, for example:

```text
14:00-15:30 meeting with Josh
```

During that block, the tracker will not interrupt you with check-ins.

### End-of-day review

At the end of the day, the tracker can show a summary and ask whether any overtime has already been added to your official time-management system.

The overtime sheet tracks both:

- overtime done
- overtime accounted for

---

## Files created by the tracker

The tracker stores your personal data locally in:

```text
~/Documents/autotime_sap/
```

Main files:

```text
autotime_log.csv       # regular accountability checks
overtime.xlsx          # overtime records and overtime accounting status
day_status.csv         # day type and planned work window
daily_review.csv       # end-of-day reflections
distractions.csv       # distraction notes, if any
```

The folder name still contains `autotime_sap` for backward compatibility with earlier versions.

Do not commit these files to GitHub.

---

## Installation

### 1. Install uv

The project uses `uv` to run Python commands.

On macOS:

```bash
brew install uv
```

Check that it works:

```bash
uv --version
```

### 2. Clone or download the repository

Using Git:

```bash
git clone https://github.com/YOUR_GITHUB_USERNAME/accountability-tracker.git
cd accountability-tracker
```

Or download the ZIP from GitHub, unzip it, and open Terminal inside the folder.

### 3. Install the environment

```bash
uv sync
```

### 4. Make shell scripts executable

```bash
chmod +x *.sh
```

---

## Start the tracker

Start check-ins every 30 minutes:

```bash
./start_accountability.sh 30
```

The tracker runs in the background using macOS LaunchAgent.

Check that it is active:

```bash
launchctl list | grep accountability
```

You should see something like:

```text
-    0    com.gunjan.accountability
```

---

## Stop the tracker

```bash
./stop_accountability.sh
```

---

## Useful commands

Run one accountability check manually:

```bash
uv run python accountability_prompt.py
```

Run daily setup again:

```bash
uv run python accountability_prompt.py --day-setup
```

Add a meeting or focus block:

```bash
uv run python accountability_prompt.py --meeting
```

Show today’s summary:

```bash
uv run python accountability_prompt.py --summary
```

Run end-of-day review:

```bash
uv run python accountability_prompt.py --review
```

Ask whether today’s overtime has been accounted for:

```bash
uv run python accountability_prompt.py --account-overtime
```

Check tracker status:

```bash
uv run python accountability_prompt.py --status
```

Run diagnostics:

```bash
uv run python accountability_prompt.py --doctor
```

---

## Overtime sheet

The overtime file is here:

```text
~/Documents/autotime_sap/overtime.xlsx
```

Open it with:

```bash
open ~/Documents/autotime_sap/overtime.xlsx
```

It contains an `Overtime` sheet and a `Summary` sheet.

The summary tracks:

```text
Today overtime done
Today overtime accounted for
Today overtime not yet accounted
Total overtime done
Total overtime accounted for
Total overtime not yet accounted
```

At the end of the day, the tracker can ask whether you already added today’s overtime to your official time-management system.

---

## Typical daily workflow

### Morning

When the first prompt appears:

1. Choose the day type.
2. Enter your planned work window, for example `09:30-18:00`.
3. Add your main tasks for the day.

### During the day

Answer the 30-minute check-ins honestly.

Examples:

```text
Instance segmentation
XAI course preparation
Admin / email / meetings
Break
Distracted / not productive
```

### During a meeting

Run:

```bash
uv run python accountability_prompt.py --meeting
```

Enter something like:

```text
14:00-15:30 project meeting
```

The tracker will pause prompts until the block ends.

### Evening

Run:

```bash
uv run python accountability_prompt.py --summary
uv run python accountability_prompt.py --review
uv run python accountability_prompt.py --account-overtime
```

---

## Troubleshooting

### The tracker does not start

Run:

```bash
chmod +x *.sh
./start_accountability.sh 30
```

If that still fails:

```bash
bash ./start_accountability.sh 30
```

### Check whether the background job is active

```bash
launchctl list | grep accountability
```

### Check which script macOS is running

```bash
grep -H "accountability" ~/Library/LaunchAgents/*.plist
```

### Run diagnostics

```bash
uv run python accountability_prompt.py --doctor
```

This checks the LaunchAgent, log files, and current tracker status.

### The summary looks wrong

Repair and diagnose the logs:

```bash
uv run python accountability_prompt.py --doctor
uv run python accountability_prompt.py --summary
```

The tracker includes automatic log-header repair for older CSV formats.

---

## Updating the tracker

After replacing files with a newer version:

```bash
cd ~/accountability-tracker
chmod +x *.sh
uv sync
./stop_accountability.sh
./start_accountability.sh 30
```

Then commit the update:

```bash
git add .
git commit -m "Update accountability tracker"
git push
```

---

## Privacy

Your personal logs are stored locally under:

```text
~/Documents/autotime_sap/
```

The repository should contain only the code.

Do not commit:

```text
*.csv
*.xlsx
*.log
```

The `.gitignore` file should exclude these by default.

---

## Project status

Current version: v13

Main features:

- daily setup
- 30-minute accountability checks
- meeting/focus blocks
- overtime tracking
- overtime accounting status
- daily summary
- end-of-day review
- diagnostics with `--doctor`
