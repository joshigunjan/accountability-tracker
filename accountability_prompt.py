#!/usr/bin/env python3
"""macOS self-accountability and overtime tracker.

Version: v13 overtime accounting check.

This script is intentionally small and local-only. It uses AppleScript popups,
CSV files for normal logs/todos, and an Excel workbook for overtime.

Default data directory is ~/Documents/autotime_sap to preserve older v5 data,
especially ~/Documents/autotime_sap/overtime.xlsx.
Set ACCOUNTABILITY_DATA_DIR to move it later.
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Iterable, Optional

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

DATA_DIR = Path(os.environ.get("ACCOUNTABILITY_DATA_DIR", "~/Documents/autotime_sap")).expanduser()
LOG_CSV = DATA_DIR / "autotime_log.csv"
DAY_STATUS_CSV = DATA_DIR / "day_status.csv"
DAILY_PLAN_CSV = DATA_DIR / "daily_plan.csv"
TODO_CSV = DATA_DIR / "todos.csv"
STATE_JSON = DATA_DIR / "state.json"
OVERTIME_XLSX = DATA_DIR / "overtime.xlsx"
PRIVATE_CSV = DATA_DIR / "private_ignored.csv"
DAILY_REVIEW_CSV = DATA_DIR / "daily_review.csv"
DISTRACTION_CSV = DATA_DIR / "distractions.csv"
EVENTS_CSV = DATA_DIR / "events.csv"
APP_VERSION = "v13.0.0"

LOG_FIELDS = [
    "timestamp",
    "date",
    "start",
    "end",
    "duration_min",
    "mode",
    "task",
    "note",
    "active_app",
    "window_title",
]
DAY_FIELDS = ["date", "day_type", "planned_start", "planned_end", "created_at"]
PLAN_FIELDS = ["date", "priority_no", "task", "created_at"]
TODO_FIELDS = ["id", "date", "task", "status", "created_at", "completed_at"]
PRIVATE_FIELDS = ["timestamp", "date", "start", "end", "duration_min", "note"]
REVIEW_FIELDS = [
    "date",
    "reviewed_at",
    "summary",
    "completed_priorities",
    "distractions",
    "overtime_justified",
    "overtime_done_hours",
    "overtime_accounted_hours",
    "overtime_accounted_status",
    "improve_tomorrow",
]
DISTRACTION_FIELDS = ["timestamp", "date", "start", "end", "duration_min", "context", "reason", "task", "note"]
EVENT_FIELDS = [
    "timestamp",
    "date",
    "start",
    "end",
    "duration_min",
    "type",
    "task",
    "note",
    "active_app",
    "window_title",
    "source",
]

OVERTIME_HEADERS = [
    "Date",
    "Start",
    "End",
    "Duration hours",
    "Task",
    "Reason",
    "Note",
    "Logged at",
    "Source",
    "Accounted in time management",
    "Accounted hours",
    "Accounted at",
    "Accounted note",
]


@dataclass
class DayStatus:
    date: str
    day_type: str
    planned_start: str = ""
    planned_end: str = ""
    created_at: str = ""


def ensure_data_dir() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    repair_log_csv(verbose=False)
    ensure_csv(LOG_CSV, LOG_FIELDS)
    ensure_csv(DAY_STATUS_CSV, DAY_FIELDS)
    ensure_csv(DAILY_PLAN_CSV, PLAN_FIELDS)
    ensure_csv(TODO_CSV, TODO_FIELDS)
    ensure_csv(PRIVATE_CSV, PRIVATE_FIELDS)
    ensure_csv(DAILY_REVIEW_CSV, REVIEW_FIELDS)
    ensure_csv(DISTRACTION_CSV, DISTRACTION_FIELDS)
    ensure_csv(EVENTS_CSV, EVENT_FIELDS)


def ensure_csv(path: Path, fields: list[str]) -> None:
    if not path.exists() or path.stat().st_size == 0:
        with path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fields)
            writer.writeheader()


def expected_header(fields: list[str]) -> str:
    return ",".join(fields)


def _looks_like_iso_date(text: str) -> bool:
    return bool(re.match(r"^\d{4}-\d{2}-\d{2}$", (text or "").strip()))


def _looks_like_timestamp(text: str) -> bool:
    return bool(re.match(r"^\d{4}-\d{2}-\d{2}[ T]\d{1,2}:\d{2}:\d{2}", (text or "").strip()))


def repair_log_csv(verbose: bool = True) -> bool:
    """Repair legacy/corrupt autotime_log.csv headers in-place.

    Older tracker versions sometimes reused a CSV with a different header or
    appended rows without the current header. The summary depends on named
    columns, so this function normalizes the file to LOG_FIELDS while keeping
    rows that look recoverable. A timestamped backup is created before repair.
    """
    if not LOG_CSV.exists() or LOG_CSV.stat().st_size == 0:
        return False
    raw = LOG_CSV.read_text(encoding="utf-8", errors="replace").splitlines()
    if not raw:
        return False
    current_header = raw[0].strip().lstrip("\ufeff")
    if current_header == expected_header(LOG_FIELDS):
        return False

    backup = LOG_CSV.with_name(f"autotime_log_backup_before_v12_{now_local().strftime('%Y%m%d_%H%M%S')}.csv")
    backup.write_text("\n".join(raw) + "\n", encoding="utf-8")

    repaired: list[dict[str, str]] = []
    for line in raw:
        if not line.strip():
            continue
        if line.lower().startswith("timestamp,"):
            continue
        try:
            parts = next(csv.reader([line]))
        except Exception:
            continue
        parts = [p.strip() for p in parts]

        # Current/new row without a proper file header:
        # timestamp,date,start,end,duration_min,mode,task,note,active_app,window_title
        if len(parts) >= 6 and _looks_like_timestamp(parts[0]) and _looks_like_iso_date(parts[1]):
            row = {field: "" for field in LOG_FIELDS}
            for field, value in zip(LOG_FIELDS, parts[:len(LOG_FIELDS)]):
                row[field] = value
            repaired.append(row)
            continue

        # Very old AutoTime row:
        # timestamp,entry,active_app,window_title
        if len(parts) >= 2 and _looks_like_timestamp(parts[0]):
            try:
                dt = datetime.fromisoformat(parts[0].replace(" ", "T"))
                day = dt.date().isoformat()
            except Exception:
                day = parts[0][:10]
            row = {field: "" for field in LOG_FIELDS}
            row.update({
                "timestamp": parts[0],
                "date": day,
                "start": "",
                "end": "",
                "duration_min": "0",
                "mode": "legacy",
                "task": parts[1] if len(parts) > 1 else "",
                "note": "imported from legacy log during v12 repair",
                "active_app": parts[2] if len(parts) > 2 else "",
                "window_title": parts[3] if len(parts) > 3 else "",
            })
            repaired.append(row)

    write_csv(LOG_CSV, LOG_FIELDS, repaired)
    if verbose:
        print(f"Repaired {LOG_CSV}")
        print(f"Backup saved to {backup}")
        print(f"Recovered rows: {len(repaired)}")
    return True


def log_header_status() -> str:
    if not LOG_CSV.exists():
        return "missing"
    first = LOG_CSV.read_text(encoding="utf-8", errors="replace").splitlines()[:1]
    if not first:
        return "empty"
    return "OK" if first[0].strip().lstrip("\ufeff") == expected_header(LOG_FIELDS) else "needs repair"

def read_csv(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open("r", newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def write_csv(path: Path, fields: list[str], rows: list[dict[str, str]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, "") for k in fields})


def append_csv(path: Path, fields: list[str], row: dict[str, str]) -> None:
    ensure_csv(path, fields)
    with path.open("a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writerow({k: row.get(k, "") for k in fields})


def now_local() -> datetime:
    return datetime.now().replace(microsecond=0)


def today_iso() -> str:
    return date.today().isoformat()


def applescript_escape(s: str) -> str:
    return s.replace("\\", "\\\\").replace('"', '\\"')


def run_osascript(script: str) -> str:
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=300,
        )
        if result.returncode != 0:
            return ""
        return result.stdout.strip()
    except Exception:
        return ""


def input_dialog(title: str, prompt: str, default: str = "") -> str:
    script = f'''
try
  set answer to text returned of (display dialog "{applescript_escape(prompt)}" default answer "{applescript_escape(default)}" buttons {{"Cancel", "Save"}} default button "Save" with title "{applescript_escape(title)}")
  return answer
on error
  return ""
end try
'''
    return run_osascript(script)


def choose_dialog(title: str, prompt: str, choices: list[str]) -> str:
    if not choices:
        return ""
    items = ", ".join(f'"{applescript_escape(c)}"' for c in choices)
    script = f'''
try
  set chosen to choose from list {{{items}}} with title "{applescript_escape(title)}" with prompt "{applescript_escape(prompt)}" OK button name "Select" cancel button name "Cancel"
  if chosen is false then
    return ""
  else
    return item 1 of chosen
  end if
on error
  return ""
end try
'''
    return run_osascript(script)



def choose_multiple_dialog(title: str, prompt: str, choices: list[str]) -> list[str]:
    if not choices:
        return []
    items = ", ".join(f'"{applescript_escape(c)}"' for c in choices)
    script = f"""
try
  set chosen to choose from list {{{items}}} with title "{applescript_escape(title)}" with prompt "{applescript_escape(prompt)}" OK button name "Select" cancel button name "Cancel" with multiple selections allowed
  if chosen is false then
    return ""
  else
    set AppleScript's text item delimiters to "||"
    set output to chosen as text
    set AppleScript's text item delimiters to ""
    return output
  end if
on error
  return ""
end try
"""
    result = run_osascript(script)
    if not result:
        return []
    return [x.strip() for x in result.split("||") if x.strip()]


def message_dialog(title: str, message: str) -> None:
    script = f'''
try
  display dialog "{applescript_escape(message)}" buttons {{"OK"}} default button "OK" with title "{applescript_escape(title)}"
end try
'''
    run_osascript(script)


def active_app_info() -> tuple[str, str]:
    app = run_osascript('tell application "System Events" to get name of first application process whose frontmost is true')
    title = ""
    if app:
        title = run_osascript(f'tell application "System Events" to tell process "{applescript_escape(app)}" to get name of front window')
    return app, title


def load_state() -> dict:
    if not STATE_JSON.exists():
        return {}
    try:
        return json.loads(STATE_JSON.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_state(state: dict) -> None:
    STATE_JSON.write_text(json.dumps(state, indent=2, sort_keys=True), encoding="utf-8")


def get_day_status(day: str | None = None) -> Optional[DayStatus]:
    day = day or today_iso()
    rows = read_csv(DAY_STATUS_CSV)
    for row in reversed(rows):
        if row.get("date") == day:
            return DayStatus(
                date=row.get("date", day),
                day_type=row.get("day_type", ""),
                planned_start=row.get("planned_start", ""),
                planned_end=row.get("planned_end", ""),
                created_at=row.get("created_at", ""),
            )
    return None


def set_day_status(status: DayStatus) -> None:
    rows = [r for r in read_csv(DAY_STATUS_CSV) if r.get("date") != status.date]
    rows.append({
        "date": status.date,
        "day_type": status.day_type,
        "planned_start": status.planned_start,
        "planned_end": status.planned_end,
        "created_at": status.created_at or now_local().isoformat(sep=" "),
    })
    write_csv(DAY_STATUS_CSV, DAY_FIELDS, rows)


def parse_time_window(text: str) -> tuple[str, str] | None:
    text = text.strip().replace("–", "-").replace("—", "-")
    m = re.search(r"(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})", text)
    if not m:
        return None
    start, end = m.group(1), m.group(2)
    try:
        parse_hhmm(start)
        parse_hhmm(end)
    except ValueError:
        return None
    return normalize_hhmm(start), normalize_hhmm(end)


def parse_hhmm(s: str) -> time:
    h, m = [int(x) for x in s.split(":")]
    if not (0 <= h <= 23 and 0 <= m <= 59):
        raise ValueError(s)
    return time(h, m)


def normalize_hhmm(s: str) -> str:
    t = parse_hhmm(s)
    return f"{t.hour:02d}:{t.minute:02d}"


def is_inside_planned(status: DayStatus, moment: datetime | None = None) -> bool:
    moment = moment or now_local()
    if status.day_type != "regular":
        return False
    if not status.planned_start or not status.planned_end:
        return False
    start = parse_hhmm(status.planned_start)
    end = parse_hhmm(status.planned_end)
    current = moment.time().replace(second=0, microsecond=0)
    if start <= end:
        return start <= current <= end
    # rare overnight window
    return current >= start or current <= end


def parse_plan_text(plan_text: str) -> list[str]:
    tasks: list[str] = []
    for line in plan_text.splitlines():
        clean = re.sub(r"^\s*\d+[.)-]?\s*", "", line).strip(" -\t")
        if clean:
            tasks.append(clean)
    if not tasks and plan_text.strip():
        tasks.append(plan_text.strip())
    return tasks[:10]


def next_todo_id() -> str:
    rows = read_csv(TODO_CSV)
    max_id = 0
    for r in rows:
        try:
            max_id = max(max_id, int(r.get("id", "0")))
        except ValueError:
            pass
    return str(max_id + 1)


def add_todo(task: str, day: str | None = None) -> str:
    day = day or today_iso()
    task = task.strip()
    if not task:
        return ""
    # avoid duplicate active tasks for same day
    for r in read_csv(TODO_CSV):
        if r.get("date") == day and r.get("task", "").strip().lower() == task.lower() and r.get("status") != "done":
            return r.get("id", "")
    tid = next_todo_id()
    append_csv(TODO_CSV, TODO_FIELDS, {
        "id": tid,
        "date": day,
        "task": task,
        "status": "todo",
        "created_at": now_local().isoformat(sep=" "),
        "completed_at": "",
    })
    return tid


def active_todos(day: str | None = None) -> list[dict[str, str]]:
    day = day or today_iso()
    return [r for r in read_csv(TODO_CSV) if r.get("date") == day and r.get("status") in {"todo", "doing"}]



def unfinished_todos(day: str) -> list[dict[str, str]]:
    return [
        r for r in read_csv(TODO_CSV)
        if r.get("date") == day and r.get("status") in {"todo", "doing"}
    ]


def update_todo_status(todo_ids: set[str], status: str) -> None:
    rows = read_csv(TODO_CSV)
    changed = False
    for r in rows:
        if r.get("id") in todo_ids:
            r["status"] = status
            if status == "done":
                r["completed_at"] = now_local().isoformat(sep=" ")
            changed = True
    if changed:
        write_csv(TODO_CSV, TODO_FIELDS, rows)


def maybe_offer_todo_carryover(force: bool = False) -> None:
    """Offer to carry yesterday's unfinished todos into today."""
    state = load_state()
    key = f"carryover_checked_{today_iso()}"
    if state.get(key) and not force:
        return

    yesterday = (date.today() - timedelta(days=1)).isoformat()
    unfinished = unfinished_todos(yesterday)
    if not unfinished:
        state[key] = True
        save_state(state)
        return

    preview = "\n".join(f"- {r.get('task','')}" for r in unfinished[:8])
    choice = choose_dialog(
        "Todo carry-over",
        f"Yesterday you did not finish:\n{preview}\n\nCarry these into today?",
        ["Carry all unfinished", "Choose items", "Drop all unfinished", "Decide later"],
    )
    if choice == "Carry all unfinished":
        for r in unfinished:
            add_todo(r.get("task", ""), today_iso())
    elif choice == "Choose items":
        labels = [f"#{r['id']} {r['task']}" for r in unfinished]
        selected = choose_multiple_dialog("Carry over todos", "Select tasks to carry into today:", labels)
        selected_ids = {re.match(r"#(\d+)", item).group(1) for item in selected if re.match(r"#(\d+)", item)}
        for r in unfinished:
            if r.get("id") in selected_ids:
                add_todo(r.get("task", ""), today_iso())
    elif choice == "Drop all unfinished":
        update_todo_status({r.get("id", "") for r in unfinished}, "dropped")

    if choice != "Decide later":
        state[key] = True
        save_state(state)


def mark_todo_done(todo_id: str) -> bool:
    rows = read_csv(TODO_CSV)
    changed = False
    for r in rows:
        if r.get("id") == str(todo_id):
            r["status"] = "done"
            r["completed_at"] = now_local().isoformat(sep=" ")
            changed = True
    if changed:
        write_csv(TODO_CSV, TODO_FIELDS, rows)
    return changed


def day_setup(force: bool = False) -> None:
    ensure_data_dir()
    existing = get_day_status()
    if existing and not force:
        return

    choice = choose_dialog(
        "Daily setup",
        "What kind of day is today?",
        ["Regular work day", "Holiday / irregular day", "Leave", "Sick day", "Skip for now"],
    )
    if not choice or choice == "Skip for now":
        return

    day_type_map = {
        "Regular work day": "regular",
        "Holiday / irregular day": "irregular",
        "Leave": "leave",
        "Sick day": "sick",
    }
    day_type = day_type_map.get(choice, "irregular")

    # Offer carry-over early so unfinished tasks can appear in today's choices.
    maybe_offer_todo_carryover()

    planned_start = ""
    planned_end = ""
    tasks: list[str] = []

    if day_type == "regular":
        while True:
            window_text = input_dialog(
                "Planned work window",
                "What is your planned work window today?\nExample: 09:30-18:00",
                "09:30-18:00",
            )
            if not window_text:
                return
            parsed = parse_time_window(window_text)
            if parsed:
                planned_start, planned_end = parsed
                break
            message_dialog("Invalid time", "Please enter time like 09:30-18:00")

        plan_text = input_dialog(
            "Today's priorities",
            "What are your top priorities today?\nPut one task per line.",
            "1. \n2. \n3. ",
        )
        tasks = parse_plan_text(plan_text)

    elif day_type == "irregular":
        plan_text = input_dialog(
            "Irregular day tasks",
            "If you do work today, what should it be for?\nPut one task per line. Leave blank if you want to add tasks later.",
            "1. \n2. \n3. ",
        )
        tasks = parse_plan_text(plan_text)

    elif day_type in {"leave", "sick"}:
        plan_text = input_dialog(
            "Optional emergency tasks",
            "You marked today as leave/sick. If you absolutely must work, what tasks are allowed?\nLeave blank if none.",
            "",
        )
        tasks = parse_plan_text(plan_text)

    # Replace today's plan rows, but keep todos already created manually.
    plan_rows = [r for r in read_csv(DAILY_PLAN_CSV) if r.get("date") != today_iso()]
    for i, task in enumerate(tasks, start=1):
        plan_rows.append({
            "date": today_iso(),
            "priority_no": str(i),
            "task": task,
            "created_at": now_local().isoformat(sep=" "),
        })
        add_todo(task)
    write_csv(DAILY_PLAN_CSV, PLAN_FIELDS, plan_rows)

    set_day_status(DayStatus(today_iso(), day_type, planned_start, planned_end, now_local().isoformat(sep=" ")))


def choose_task(
    title: str = "Accountability",
    prompt: str = "What are you working on?",
    *,
    include_nonwork: bool = True,
) -> str:
    todos = active_todos()
    choices = [f"#{r['id']} {r['task']}" for r in todos]
    choices.extend(["Add new task", "Other / type manually", "Admin / email / meetings", "Add meeting/focus block"])
    if include_nonwork:
        choices.extend(["Break", "Distracted / not productive"])
    choices.append("Skip")
    choice = choose_dialog(title, prompt, choices)
    if not choice or choice == "Skip":
        return ""
    if choice == "Add new task":
        task = input_dialog("Add task", "Add a new task:", "")
        if task:
            add_todo(task)
        return task.strip()
    if choice == "Other / type manually":
        return input_dialog("Activity", "What are you working on?", "").strip()
    if choice == "Add meeting/focus block":
        return "__ADD_BLOCK__"
    if choice.startswith("#"):
        return re.sub(r"^#\d+\s+", "", choice).strip()
    return choice.strip()


def log_event(event_type: str, task: str, start_dt: datetime, end_dt: datetime, note: str = "", active_app: str = "", window_title: str = "", source: str = "tracker") -> None:
    duration_min = max(0, int((end_dt - start_dt).total_seconds() // 60))
    append_csv(EVENTS_CSV, EVENT_FIELDS, {
        "timestamp": now_local().isoformat(sep=" "),
        "date": start_dt.date().isoformat(),
        "start": start_dt.strftime("%H:%M"),
        "end": end_dt.strftime("%H:%M"),
        "duration_min": str(duration_min),
        "type": event_type,
        "task": task,
        "note": note,
        "active_app": active_app,
        "window_title": window_title,
        "source": source,
    })


def log_activity(mode: str, task: str, start_dt: datetime, end_dt: datetime, note: str = "") -> None:
    active_app, window_title = active_app_info()
    duration_min = max(0, int((end_dt - start_dt).total_seconds() // 60))
    row = {
        "timestamp": now_local().isoformat(sep=" "),
        "date": start_dt.date().isoformat(),
        "start": start_dt.strftime("%H:%M"),
        "end": end_dt.strftime("%H:%M"),
        "duration_min": str(duration_min),
        "mode": mode,
        "task": task,
        "note": note,
        "active_app": active_app,
        "window_title": window_title,
    }
    append_csv(LOG_CSV, LOG_FIELDS, row)
    log_event(mode, task, start_dt, end_dt, note=note, active_app=active_app, window_title=window_title, source="autotime_log")


def _cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ", timespec="seconds")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, time):
        return value.strftime("%H:%M:%S")
    return str(value).strip()


def _is_header_like(values: list) -> bool:
    cleaned = [_cell_text(v).lower() for v in values]
    joined = "|".join(cleaned)
    return (
        joined.startswith("date|start|end|duration hours")
        or joined.startswith("timestamp|date|weekday|start|end|duration min")
        or cleaned[0] in {"date", "timestamp"}
    )


def _parse_date_cell(value) -> Optional[date]:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = _cell_text(value)
    for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(text[:10], fmt).date()
        except ValueError:
            pass
    return None


def _parse_hours(value) -> float:
    text = _cell_text(value)
    if not text:
        return 0.0
    try:
        return round(float(text), 2)
    except ValueError:
        pass
    # Legacy duration such as 00:30.
    m = re.match(r"^(\d{1,2}):(\d{2})$", text)
    if m:
        return round((int(m.group(1)) * 60 + int(m.group(2))) / 60, 2)
    return 0.0


def _normalize_overtime_rows(ws) -> list[list]:
    """Return clean overtime rows from mixed v5/v9/cleaned workbook formats.

    This removes repeated headers, title rows, total rows, and blank rows.
    It also preserves older legacy rows that used Timestamp/Duration min columns.
    """
    clean_rows: list[list] = []
    seen: set[tuple] = set()
    for row in ws.iter_rows(values_only=True):
        values = list(row)
        if not any(_cell_text(v) for v in values):
            continue
        first = _cell_text(values[0]).lower()
        if first in {"cleaned overtime log", "total overtime hours"}:
            continue
        if _is_header_like(values):
            continue

        normalized = None

        # Current format: Date | Start | End | Duration hours | Task | Reason | Note | Logged at | Source
        current_date = _parse_date_cell(values[0] if len(values) > 0 else None)
        if current_date and len(values) >= 4:
            duration_hours = _parse_hours(values[3] if len(values) > 3 else 0)
            accounted_text = _cell_text(values[9] if len(values) > 9 else "")
            accounted_hours = _parse_hours(values[10] if len(values) > 10 else 0)
            # Older rows had a yes/no account flag but no accounted-hours column.
            if accounted_hours == 0 and accounted_text.lower() in {"yes", "y", "true", "accounted", "all"}:
                accounted_hours = duration_hours
            normalized = [
                current_date.isoformat(),
                _cell_text(values[1] if len(values) > 1 else ""),
                _cell_text(values[2] if len(values) > 2 else ""),
                duration_hours,
                _cell_text(values[4] if len(values) > 4 else ""),
                _cell_text(values[5] if len(values) > 5 else ""),
                _cell_text(values[6] if len(values) > 6 else ""),
                _cell_text(values[7] if len(values) > 7 else ""),
                _cell_text(values[8] if len(values) > 8 else "tracker"),
                accounted_text,
                accounted_hours,
                _cell_text(values[11] if len(values) > 11 else ""),
                _cell_text(values[12] if len(values) > 12 else ""),
            ]

        # Legacy format: Timestamp | Date | Weekday | Start | End | Duration min | Duration | Activity | Source | Notes
        legacy_date = _parse_date_cell(values[1] if len(values) > 1 else None)
        if legacy_date and len(values) >= 10 and _cell_text(values[3]) and _cell_text(values[4]):
            try:
                hours = round(float(_cell_text(values[5])) / 60, 2)
            except ValueError:
                hours = _parse_hours(values[6])
            normalized = [
                legacy_date.isoformat(),
                _cell_text(values[3]),
                _cell_text(values[4]),
                hours,
                _cell_text(values[7]),
                _cell_text(values[8]),
                _cell_text(values[9]),
                _cell_text(values[0]),
                "legacy format",
                "",
                0.0,
                "",
                "",
            ]

        if not normalized:
            continue

        # Skip accidental repeated data-less rows.
        if not normalized[0] or normalized[0].lower() in {"date", "timestamp"}:
            continue
        key = tuple(str(x) for x in normalized[:8])
        if key in seen:
            continue
        seen.add(key)
        clean_rows.append(normalized)
    return clean_rows


def _write_overtime_workbook(wb: Workbook, rows: list[list]) -> None:
    if "Overtime" in wb.sheetnames:
        ws = wb["Overtime"]
        # Remove old contents completely so repeated headers / totals disappear.
        if ws.max_row:
            ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet("Overtime")

    ws.append(OVERTIME_HEADERS)
    for row in rows:
        ws.append(row)
    style_overtime_sheet(ws)

    # Keep a small summary sheet that is safe to regenerate.
    if "Summary" in wb.sheetnames:
        del wb["Summary"]
    summary = wb.create_sheet("Summary")
    today = today_iso()
    total_done = round(sum(float(r[3] or 0) for r in rows), 2)
    total_accounted = round(sum(float(r[10] or 0) for r in rows), 2)
    today_done = round(sum(float(r[3] or 0) for r in rows if str(r[0]) == today), 2)
    today_accounted = round(sum(float(r[10] or 0) for r in rows if str(r[0]) == today), 2)
    summary.append(["Metric", "Value"])
    summary.append(["Today overtime done", today_done])
    summary.append(["Today overtime accounted for", today_accounted])
    summary.append(["Today overtime not yet accounted", max(0, round(today_done - today_accounted, 2))])
    summary.append(["Total overtime done", total_done])
    summary.append(["Total overtime accounted for", total_accounted])
    summary.append(["Total overtime not yet accounted", max(0, round(total_done - total_accounted, 2))])
    summary.append(["Rows", len(rows)])
    summary.append(["Last cleaned", now_local().isoformat(sep=" ")])
    for cell in summary[1]:
        cell.font = Font(bold=True)
    summary.column_dimensions["A"].width = 34
    summary.column_dimensions["B"].width = 24


def ensure_overtime_workbook() -> None:
    if OVERTIME_XLSX.exists():
        try:
            wb = load_workbook(OVERTIME_XLSX)
        except Exception:
            # Do not destroy existing file. Create sidecar if current workbook is unreadable.
            sidecar = DATA_DIR / f"overtime_recovery_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Overtime"
            ws.append(OVERTIME_HEADERS)
            wb.save(sidecar)
            return
        if "Overtime" not in wb.sheetnames:
            ws = wb.create_sheet("Overtime")
            rows: list[list] = []
        else:
            ws = wb["Overtime"]
            rows = _normalize_overtime_rows(ws)
        _write_overtime_workbook(wb, rows)
        try:
            wb.save(OVERTIME_XLSX)
        except PermissionError:
            # If the workbook is open/locked, do not interrupt the prompt flow.
            pass
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Overtime"
    ws.append(OVERTIME_HEADERS)
    style_overtime_sheet(ws)
    wb.save(OVERTIME_XLSX)


def style_overtime_sheet(ws) -> None:
    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
    widths = [14, 10, 10, 15, 32, 24, 38, 22, 18, 28, 16, 22, 38]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width


def append_overtime(start_dt: datetime, end_dt: datetime, task: str, reason: str = "outside planned time", note: str = "") -> None:
    ensure_overtime_workbook()
    try:
        wb = load_workbook(OVERTIME_XLSX)
        ws = wb["Overtime"]
        duration_hours = round((end_dt - start_dt).total_seconds() / 3600, 2)
        # Re-read normalized rows, append exactly one data row, then rewrite.
        rows = _normalize_overtime_rows(ws)
        rows.append([
            start_dt.date().isoformat(),
            start_dt.strftime("%H:%M"),
            end_dt.strftime("%H:%M"),
            duration_hours,
            task,
            reason,
            note,
            now_local().isoformat(sep=" "),
            "tracker",
            "no",
            0.0,
            "",
            "",
        ])
        _write_overtime_workbook(wb, rows)
        wb.save(OVERTIME_XLSX)
    except PermissionError:
        # Excel/Numbers may lock the file. Preserve data in CSV fallback.
        append_csv(DATA_DIR / "overtime_pending.csv", OVERTIME_HEADERS, {
            "Date": start_dt.date().isoformat(),
            "Start": start_dt.strftime("%H:%M"),
            "End": end_dt.strftime("%H:%M"),
            "Duration hours": str(round((end_dt - start_dt).total_seconds() / 3600, 2)),
            "Task": task,
            "Reason": reason,
            "Note": note,
            "Logged at": now_local().isoformat(sep=" "),
            "Source": "tracker_pending_csv",
            "Accounted in time management": "no",
            "Accounted hours": "0",
            "Accounted at": "",
            "Accounted note": "",
        })
        message_dialog("Overtime saved to fallback", "overtime.xlsx seems to be open/locked. I saved this row to overtime_pending.csv instead.")


def should_prompt(interval_minutes: int, mode: str) -> bool:
    state = load_state()
    key = f"last_prompt_at_{mode}"
    last = state.get(key) or state.get("last_prompt_at")
    if not last:
        return True
    try:
        last_dt = datetime.fromisoformat(last)
    except ValueError:
        return True
    return now_local() - last_dt >= timedelta(minutes=interval_minutes)


def mark_prompt(mode: str) -> None:
    state = load_state()
    stamp = now_local().isoformat()
    state["last_prompt_at"] = stamp
    state[f"last_prompt_at_{mode}"] = stamp
    save_state(state)


def in_quiet_period() -> bool:
    state = load_state()
    quiet_until = state.get("quiet_until")
    if not quiet_until:
        return False
    try:
        until = datetime.fromisoformat(quiet_until)
    except ValueError:
        return False
    if now_local() < until:
        return True
    state.pop("quiet_until", None)
    save_state(state)
    return False


def quiet_until_text() -> str:
    state = load_state()
    quiet_until = state.get("quiet_until")
    if not quiet_until:
        return "none"
    try:
        until = datetime.fromisoformat(quiet_until)
    except ValueError:
        return f"invalid value: {quiet_until}"
    if now_local() >= until:
        return "expired"
    return until.strftime("%Y-%m-%d %H:%M")


def next_prompt_due_text(interval_minutes: int = 30) -> str:
    state = load_state()
    last_values = [v for k, v in state.items() if k.startswith("last_prompt_at") and v]
    if not last_values:
        return "now / no previous prompt recorded"
    try:
        last_dt = max(datetime.fromisoformat(v) for v in last_values)
    except Exception:
        return "unknown"
    due = last_dt + timedelta(minutes=interval_minutes)
    if now_local() >= due:
        return "now"
    return due.strftime("%Y-%m-%d %H:%M")



def ask_distraction_reason(context: str = "", task: str = "") -> str:
    choice = choose_dialog(
        "Distraction check",
        "What distracted you?",
        [
            "YouTube / videos",
            "Browsing / rabbit hole",
            "Phone / WhatsApp",
            "Unclear task",
            "Tired / low energy",
            "Meeting fatigue",
            "Other",
            "Skip",
        ],
    )
    if not choice or choice == "Skip":
        return ""
    if choice == "Other":
        return input_dialog("Distraction reason", "What distracted you?", "").strip()
    return choice


def log_distraction(start_dt: datetime, end_dt: datetime, context: str, reason: str, task: str = "", note: str = "") -> None:
    append_csv(DISTRACTION_CSV, DISTRACTION_FIELDS, {
        "timestamp": now_local().isoformat(sep=" "),
        "date": start_dt.date().isoformat(),
        "start": start_dt.strftime("%H:%M"),
        "end": end_dt.strftime("%H:%M"),
        "duration_min": str(max(0, int((end_dt - start_dt).total_seconds() // 60))),
        "context": context,
        "reason": reason,
        "task": task,
        "note": note,
    })


def regular_prompt(interval_minutes: int) -> None:
    task = choose_task("Accountability check", "What are you working on?")
    if task == "__ADD_BLOCK__":
        add_block()
        return
    if not task:
        return
    end_dt = now_local()
    start_dt = end_dt - timedelta(minutes=interval_minutes)
    if task == "Break":
        log_activity("break", task, start_dt, end_dt, note="inside planned work window")
    elif task == "Distracted / not productive":
        reason = ask_distraction_reason("planned work time")
        log_distraction(start_dt, end_dt, "planned work time", reason, task=task)
        log_activity("distracted", task, start_dt, end_dt, note=reason)
    else:
        log_activity("regular", task, start_dt, end_dt)


def overtime_totals() -> tuple[float, float]:
    """Return overtime totals for today and the current Monday-Sunday week."""
    if not OVERTIME_XLSX.exists():
        return 0.0, 0.0
    try:
        ensure_overtime_workbook()
        wb = load_workbook(OVERTIME_XLSX, data_only=True)
        ws = wb["Overtime"]
    except Exception:
        return 0.0, 0.0

    today = today_iso()
    week_start = date.today() - timedelta(days=date.today().weekday())
    total_today = 0.0
    total_week = 0.0

    for row in ws.iter_rows(min_row=2, values_only=True):
        row_date, _, _, hours, *_ = row
        try:
            d = date.fromisoformat(str(row_date))
            h = float(hours or 0)
        except Exception:
            continue
        if d.isoformat() == today:
            total_today += h
        if week_start <= d <= date.today():
            total_week += h
    return total_today, total_week


def confirm_overtime_thresholds(projected_today: float, projected_week: float) -> bool:
    """Return True if the user still wants to log overtime after threshold checks."""
    if projected_today >= 2:
        choice = choose_dialog(
            "Overtime check",
            f"This will bring today's overtime to {projected_today:.1f} hours. Is this really necessary?",
            ["Yes, log overtime", "No, skip this block"],
        )
        if choice != "Yes, log overtime":
            return False
    elif projected_today >= 1:
        message_dialog(
            "Overtime note",
            f"You will have logged {projected_today:.1f} hour(s) overtime today after this block.",
        )

    if projected_week >= 5:
        message_dialog(
            "Weekly overtime warning",
            f"You will have logged {projected_week:.1f} hours overtime this week after this block.",
        )
    return True


def ask_after_22() -> str:
    """Ask why the user is working late. Returns a note string, or empty string to skip logging."""
    if now_local().time() < time(22, 0):
        return ""
    choice = choose_dialog(
        "Late work check",
        "It is after 22:00. Is this urgent work or avoidable work?",
        ["Urgent / necessary", "Avoidable / I should stop", "Continue, but mark as avoidable", "Skip"],
    )
    if choice == "Urgent / necessary":
        return "after 22:00; urgent/necessary"
    if choice == "Continue, but mark as avoidable":
        return "after 22:00; avoidable but continued"
    return "SKIP_OVERTIME"


def outside_hours_prompt(interval_minutes: int, status: DayStatus | None) -> None:
    label = "outside planned time"
    if status and status.day_type in {"irregular", "leave", "sick"}:
        label = status.day_type
    prompt = "This is outside your planned work window. Is this work or private?"
    if label in {"irregular", "leave", "sick"}:
        prompt = f"Today is marked as {label}. Is this work or private?"
    choice = choose_dialog(
        "Outside planned time",
        prompt,
        ["Work", "Private matter", "Break", "Distracted / not productive", "Add meeting/focus block", "Skip"],
    )
    if not choice or choice == "Skip":
        return
    if choice == "Add meeting/focus block":
        add_block()
        return
    end_dt = now_local()
    start_dt = end_dt - timedelta(minutes=interval_minutes)
    if choice in {"Private matter", "Break", "Distracted / not productive"}:
        note = choice.lower()
        mode = "private"
        if choice == "Break":
            mode = "break"
        elif choice == "Distracted / not productive":
            mode = "distracted"
            reason = ask_distraction_reason(label)
            note = f"distracted: {reason}" if reason else "distracted"
            log_distraction(start_dt, end_dt, label, reason, task=choice, note="outside planned/irregular")
        append_csv(PRIVATE_CSV, PRIVATE_FIELDS, {
            "timestamp": now_local().isoformat(sep=" "),
            "date": start_dt.date().isoformat(),
            "start": start_dt.strftime("%H:%M"),
            "end": end_dt.strftime("%H:%M"),
            "duration_min": str(interval_minutes),
            "note": note,
        })
        log_activity(mode, choice, start_dt, end_dt, note=f"ignored for overtime; {label}; {note}")
        return

    late_note = ask_after_22()
    if late_note == "SKIP_OVERTIME":
        log_activity("overtime_skipped", "avoidable late work", start_dt, end_dt, note="after 22:00; skipped")
        return

    task = choose_task("Overtime / irregular work", "Which task is this work for?", include_nonwork=False)
    if not task:
        task = input_dialog("Overtime", "What is this overtime work for?", "")
    if not task:
        return

    duration_hours = round((end_dt - start_dt).total_seconds() / 3600, 2)
    total_today, total_week = overtime_totals()
    projected_today = total_today + duration_hours
    projected_week = total_week + duration_hours
    if not confirm_overtime_thresholds(projected_today, projected_week):
        log_activity("overtime_skipped", task, start_dt, end_dt, note=f"threshold check skipped; {label}")
        return

    note = late_note if late_note else ""
    append_overtime(start_dt, end_dt, task, reason=label, note=note)
    log_activity("overtime", task, start_dt, end_dt, note="; ".join(x for x in [label, note] if x))
    overtime_warning()


def overtime_warning() -> None:
    total_today, total_week = overtime_totals()
    if total_today >= 2:
        message_dialog("Overtime warning", f"You have logged {total_today:.1f} hours overtime today. Please check if this is really necessary.")
    elif total_today >= 1:
        message_dialog("Overtime note", f"You have already logged {total_today:.1f} hour(s) overtime today.")
    if total_week >= 5:
        message_dialog("Weekly overtime warning", f"You have logged {total_week:.1f} hours overtime this week.")



def overtime_total_for_day(day: str) -> float:
    if not OVERTIME_XLSX.exists():
        return 0.0
    try:
        ensure_overtime_workbook()
        wb = load_workbook(OVERTIME_XLSX, data_only=True)
        ws = wb["Overtime"]
    except Exception:
        return 0.0
    total = 0.0
    for row in ws.iter_rows(min_row=2, values_only=True):
        row_date, _, _, hours, *_ = row
        try:
            if date.fromisoformat(str(row_date)).isoformat() == day:
                total += float(hours or 0)
        except Exception:
            continue
    return total


def overtime_rows_for_day(day: str) -> list[dict[str, str]]:
    """Return normalized overtime workbook rows for a date.

    This lets the daily summary count overtime even if an older version wrote
    overtime.xlsx but failed to mirror the same block into autotime_log.csv.
    """
    if not OVERTIME_XLSX.exists():
        return []
    try:
        ensure_overtime_workbook()
        wb = load_workbook(OVERTIME_XLSX, data_only=True)
        ws = wb["Overtime"]
    except Exception:
        return []
    rows: list[dict[str, str]] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        values = list(row) + [""] * 13
        row_date, start, end, hours, task, reason, note, logged_at, source, accounted, accounted_hours, accounted_at, accounted_note = values[:13]
        try:
            if date.fromisoformat(str(row_date)).isoformat() != day:
                continue
        except Exception:
            continue
        rows.append({
            "date": str(row_date),
            "start": _cell_text(start),
            "end": _cell_text(end),
            "hours": str(hours or 0),
            "task": _cell_text(task),
            "reason": _cell_text(reason),
            "note": _cell_text(note),
            "logged_at": _cell_text(logged_at),
            "source": _cell_text(source),
            "accounted": _cell_text(accounted),
            "accounted_hours": str(accounted_hours or 0),
            "accounted_at": _cell_text(accounted_at),
            "accounted_note": _cell_text(accounted_note),
        })
    return rows



def overtime_done_accounted_for_day(day: str | None = None) -> tuple[float, float]:
    day = day or today_iso()
    rows = overtime_rows_for_day(day)
    done = 0.0
    accounted = 0.0
    for r in rows:
        try:
            done += float(r.get("hours") or 0)
        except ValueError:
            pass
        try:
            accounted += float(r.get("accounted_hours") or 0)
        except ValueError:
            # Older boolean-style accounted rows count as fully accounted.
            if (r.get("accounted") or "").lower() in {"yes", "y", "true", "accounted", "all"}:
                try:
                    accounted += float(r.get("hours") or 0)
                except ValueError:
                    pass
    return round(done, 2), round(accounted, 2)


def mark_overtime_accounted(day: str | None = None, accounted_hours: float | None = None, note: str = "") -> tuple[float, float]:
    """Mark today's overtime as accounted in the time-management system.

    If accounted_hours is None, all overtime rows for the day are marked as accounted.
    If a number is provided, rows are marked sequentially until that many hours are accounted.
    """
    day = day or today_iso()
    ensure_overtime_workbook()
    wb = load_workbook(OVERTIME_XLSX)
    ws = wb["Overtime"]
    rows = _normalize_overtime_rows(ws)
    total = round(sum(float(r[3] or 0) for r in rows if str(r[0]) == day), 2)
    if total <= 0:
        return 0.0, 0.0

    target = total if accounted_hours is None else max(0.0, min(float(accounted_hours), total))
    remaining = target
    stamp = now_local().isoformat(sep=" ")
    account_note = note.strip() or "Overtime accounted for in time management"

    # Reset today's accounting status first, then re-apply the current answer.
    for r in rows:
        if str(r[0]) == day:
            r[9] = "no"
            r[10] = 0.0
            r[11] = ""
            r[12] = ""

    for r in rows:
        if str(r[0]) != day or remaining <= 0:
            continue
        duration = float(r[3] or 0)
        accounted = round(min(duration, remaining), 2)
        if accounted <= 0:
            continue
        r[9] = "yes" if accounted >= duration else "partial"
        r[10] = accounted
        r[11] = stamp
        r[12] = account_note
        remaining = round(remaining - accounted, 2)

    _write_overtime_workbook(wb, rows)
    wb.save(OVERTIME_XLSX)
    return total, round(target, 2)


def ask_overtime_accounting(day: str | None = None) -> tuple[str, float, float]:
    day = day or today_iso()
    done, already_accounted = overtime_done_accounted_for_day(day)
    if done <= 0:
        return "no overtime", 0.0, 0.0

    outstanding = max(0.0, round(done - already_accounted, 2))
    choice = choose_dialog(
        "Overtime accounting check",
        (
            f"Today overtime done: {done:.1f} h\n"
            f"Already marked accounted: {already_accounted:.1f} h\n"
            f"Not yet accounted: {outstanding:.1f} h\n\n"
            "Have you already added this overtime to your time-management system?"
        ),
        ["Yes, all added", "Partly added", "No, not yet", "Skip"],
    )
    if choice == "Yes, all added":
        note = input_dialog("Overtime accounting note", "Optional note", "Overtime accounted for in time management")
        _, accounted = mark_overtime_accounted(day, None, note)
        return "all accounted", done, accounted
    if choice == "Partly added":
        raw = input_dialog("Overtime accounting", "How many overtime hours did you already add?", f"{already_accounted:.1f}")
        try:
            hours = float(raw.replace(",", "."))
        except ValueError:
            hours = already_accounted
        note = input_dialog("Overtime accounting note", "Optional note", "Overtime partly accounted for in time management")
        _, accounted = mark_overtime_accounted(day, hours, note)
        return "partly accounted", done, accounted
    if choice == "No, not yet":
        return "not yet accounted", done, already_accounted
    return "skipped", done, already_accounted

def build_daily_summary(day: str | None = None) -> str:
    day = day or today_iso()
    rows = [r for r in read_csv(LOG_CSV) if r.get("date") == day]
    overtime_rows = overtime_rows_for_day(day)

    priorities = [r.get("task", "") for r in read_csv(DAILY_PLAN_CSV) if r.get("date") == day]
    status = get_day_status(day)
    planned = ""
    if status and status.day_type == "regular":
        planned = f"Planned work: {status.planned_start}-{status.planned_end}"
    elif status:
        planned = f"Day type: {status.day_type}"

    # Some older versions wrote overtime.xlsx but did not also mirror the row
    # into autotime_log.csv. Count those overtime rows as answered checks too,
    # but avoid double-counting overtime rows that already exist in the log.
    logged_overtime = sum(1 for r in rows if r.get("mode") == "overtime")
    overtime_missing_from_log = max(0, len(overtime_rows) - logged_overtime)

    checks = len(rows) + overtime_missing_from_log
    productive_modes = {"regular", "regular_block", "overtime"}
    productive = sum(1 for r in rows if r.get("mode") in productive_modes) + overtime_missing_from_log
    distracted = sum(1 for r in rows if r.get("mode") == "distracted")
    breaks_private = sum(1 for r in rows if r.get("mode") in {"private", "break", "private_block"})
    overtime_hours, overtime_accounted = overtime_done_accounted_for_day(day)
    overtime_unaccounted = max(0.0, round(overtime_hours - overtime_accounted, 2))

    task_minutes: dict[str, int] = {}
    for r in rows:
        task = (r.get("task") or "").strip()
        if not task or task in {"Break", "Distracted / not productive", "Private matter"}:
            continue
        try:
            minutes = int(float(r.get("duration_min") or 0))
        except ValueError:
            minutes = 0
        task_minutes[task] = task_minutes.get(task, 0) + minutes

    # Add overtime workbook tasks that are missing from the CSV log.
    if overtime_missing_from_log:
        for r in overtime_rows:
            task = (r.get("task") or "").strip()
            if not task:
                continue
            try:
                minutes = int(round(float(r.get("hours") or 0) * 60))
            except ValueError:
                minutes = 0
            task_minutes[task] = task_minutes.get(task, 0) + minutes

    main_task = "None yet"
    if task_minutes:
        task, mins = max(task_minutes.items(), key=lambda item: item[1])
        main_task = f"{task} ({mins / 60:.1f} h)"

    priority_text = ""
    if priorities:
        priority_text = "\nTop priorities:\n" + "\n".join(f"- {p}" for p in priorities[:5])

    return (
        f"Today: {day}\n"
        f"{planned}\n"
        f"Accountability checks answered: {checks}\n"
        f"Productive/work blocks: {productive}\n"
        f"Break/private blocks: {breaks_private}\n"
        f"Distracted blocks: {distracted}\n"
        f"Overtime done: {overtime_hours:.1f} h\n"
        f"Overtime accounted for: {overtime_accounted:.1f} h\n"
        f"Overtime not yet accounted: {overtime_unaccounted:.1f} h\n"
        f"Main task: {main_task}"
        f"{priority_text}"
    )

def review_already_done(day: str | None = None) -> bool:
    day = day or today_iso()
    state = load_state()
    if state.get(f"daily_review_done_{day}"):
        return True
    return any(r.get("date") == day for r in read_csv(DAILY_REVIEW_CSV))


def should_run_daily_review(status: DayStatus | None) -> bool:
    if not status or review_already_done(status.date):
        return False
    now = now_local()
    if status.day_type == "regular" and status.planned_end:
        try:
            end_dt = datetime.combine(date.today(), parse_hhmm(status.planned_end)) + timedelta(minutes=15)
            if status.planned_start and parse_hhmm(status.planned_end) < parse_hhmm(status.planned_start):
                end_dt += timedelta(days=1)
            return now >= end_dt
        except Exception:
            return now.time() >= time(18, 30)
    return now.time() >= time(21, 0)


def run_daily_review(force: bool = False) -> None:
    ensure_data_dir()
    day = today_iso()
    if review_already_done(day) and not force:
        return
    summary = build_daily_summary(day)
    message_dialog("Daily summary", summary)

    # v13: before the subjective review, ask whether today's overtime has
    # already been added to the user's official/time-management system.
    overtime_accounting_status, overtime_done, overtime_accounted = ask_overtime_accounting(day)
    summary = build_daily_summary(day)

    completed = input_dialog(
        "End-of-day review",
        "Did you complete your top priorities today?\nWrite a short note.",
        "",
    )
    distractions = input_dialog(
        "End-of-day review",
        "What distracted you today, if anything?",
        "",
    )
    overtime_note = input_dialog(
        "End-of-day review",
        "Optional overtime note / reason for record.",
        overtime_accounting_status,
    )
    improve = input_dialog(
        "End-of-day review",
        "One thing to improve tomorrow?",
        "",
    )
    append_csv(DAILY_REVIEW_CSV, REVIEW_FIELDS, {
        "date": day,
        "reviewed_at": now_local().isoformat(sep=" "),
        "summary": summary.replace("\n", " | "),
        "completed_priorities": completed,
        "distractions": distractions,
        "overtime_justified": overtime_note,
        "overtime_done_hours": f"{overtime_done:.2f}",
        "overtime_accounted_hours": f"{overtime_accounted:.2f}",
        "overtime_accounted_status": overtime_accounting_status,
        "improve_tomorrow": improve,
    })
    state = load_state()
    state[f"daily_review_done_{day}"] = True
    save_state(state)

def show_daily_summary() -> None:
    message_dialog("Daily summary", build_daily_summary())


def scheduled_run(interval_minutes: int) -> None:
    ensure_data_dir()
    if in_quiet_period():
        return
    status = get_day_status()
    if not status:
        day_setup(force=True)
        status = get_day_status()
        if not status:
            return
    if should_run_daily_review(status):
        run_daily_review()

    now = now_local()
    inside = bool(status and is_inside_planned(status, now))
    mode = "regular" if inside else "outside"
    if should_prompt(interval_minutes, mode):
        if inside:
            regular_prompt(interval_minutes)
        else:
            outside_hours_prompt(interval_minutes, status)
        mark_prompt(mode)


def parse_block(text: str) -> tuple[datetime, datetime, str] | None:
    text = text.strip().replace("–", "-").replace("—", "-")
    m = re.search(r"(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})\s*(.*)", text)
    if not m:
        return None
    start_s, end_s, task = normalize_hhmm(m.group(1)), normalize_hhmm(m.group(2)), m.group(3).strip()
    today = date.today()
    start_dt = datetime.combine(today, parse_hhmm(start_s))
    end_dt = datetime.combine(today, parse_hhmm(end_s))
    if end_dt <= start_dt:
        end_dt += timedelta(days=1)
    return start_dt, end_dt, task


def add_block() -> None:
    ensure_data_dir()
    text = input_dialog("Add time block", "Enter time block and task.\nExample: 20:00-21:00 HFMI paper", "")
    if not text:
        return
    parsed = parse_block(text)
    if not parsed:
        message_dialog("Invalid block", "Use format like 20:00-21:00 HFMI paper")
        return
    start_dt, end_dt, task = parsed
    if not task:
        task = input_dialog("Task", "What was this block for?", "")
    if not task:
        return
    status = get_day_status()
    is_regular_block = status and status.day_type == "regular" and is_inside_planned(status, start_dt) and is_inside_planned(status, end_dt)
    if is_regular_block:
        log_activity("regular_block", task, start_dt, end_dt, note="manual block")
    else:
        choice = choose_dialog("Manual block", "Is this block work overtime or private?", ["Work overtime", "Private matter", "Regular/accountability only"])
        if choice == "Work overtime":
            append_overtime(start_dt, end_dt, task, reason="manual block")
            log_activity("overtime", task, start_dt, end_dt, note="manual block")
        elif choice == "Private matter":
            append_csv(PRIVATE_CSV, PRIVATE_FIELDS, {
                "timestamp": now_local().isoformat(sep=" "),
                "date": start_dt.date().isoformat(),
                "start": start_dt.strftime("%H:%M"),
                "end": end_dt.strftime("%H:%M"),
                "duration_min": str(int((end_dt - start_dt).total_seconds() // 60)),
                "note": task,
            })
            log_activity("private_block", task, start_dt, end_dt, note="manual block")
        elif choice == "Regular/accountability only":
            log_activity("regular_block", task, start_dt, end_dt, note="manual block")
        else:
            return
    # Silence prompts until the end of a meeting/block if it is in the future.
    if end_dt > now_local():
        state = load_state()
        state["quiet_until"] = end_dt.isoformat()
        save_state(state)


def list_todos() -> None:
    ensure_data_dir()
    rows = active_todos()
    if not rows:
        print("No active todos for today.")
        return
    for r in rows:
        print(f"#{r['id']} [{r['status']}] {r['task']}")


def add_todo_cli(task: str) -> None:
    ensure_data_dir()
    tid = add_todo(task)
    print(f"Added todo #{tid}: {task}")


def done_todo_cli(todo_id: str) -> None:
    ensure_data_dir()
    if mark_todo_done(todo_id):
        print(f"Marked todo #{todo_id} done.")
    else:
        print(f"Todo #{todo_id} not found.")


def today_counts(day: str | None = None) -> dict[str, float | int]:
    day = day or today_iso()
    rows = [r for r in read_csv(LOG_CSV) if r.get("date") == day]
    logged_overtime = sum(1 for r in rows if r.get("mode") == "overtime")
    overtime_rows = overtime_rows_for_day(day)
    overtime_missing = max(0, len(overtime_rows) - logged_overtime)
    return {
        "log_rows": len(rows),
        "checks": len(rows) + overtime_missing,
        "regular": sum(1 for r in rows if r.get("mode") in {"regular", "regular_block"}),
        "overtime_rows": len(overtime_rows),
        "overtime_hours": overtime_done_accounted_for_day(day)[0],
        "overtime_accounted_hours": overtime_done_accounted_for_day(day)[1],
        "private_break_distracted": sum(1 for r in rows if r.get("mode") in {"private", "private_block", "break", "distracted"}),
    }


def launchagent_status() -> tuple[str, str]:
    plist = Path.home() / "Library/LaunchAgents/com.gunjan.accountability.plist"
    runner = Path.home() / ".accountability_tracker/run_accountability.sh"
    status = "not checked"
    if sys.platform == "darwin":
        try:
            result = subprocess.run(["launchctl", "list"], capture_output=True, text=True, timeout=10)
            line = next((ln for ln in result.stdout.splitlines() if "com.gunjan.accountability" in ln), "")
            status = line.strip() or "not loaded"
        except Exception as exc:
            status = f"could not query launchctl: {exc}"
    runner_text = ""
    if runner.exists():
        try:
            runner_text = runner.read_text(encoding="utf-8", errors="replace").strip()
        except Exception:
            runner_text = "exists but could not read"
    else:
        runner_text = "runner missing"
    return status, f"plist={plist} | runner={runner} | runner_content={runner_text}"


def show_status(interval_minutes: int = 30) -> None:
    ensure_data_dir()
    counts = today_counts()
    status = get_day_status()
    planned = "not set"
    if status:
        planned = status.day_type
        if status.planned_start or status.planned_end:
            planned += f" {status.planned_start}-{status.planned_end}"
    text = (
        f"Accountability Tracker {APP_VERSION}\n"
        f"Today: {today_iso()}\n"
        f"Day status: {planned}\n"
        f"Data folder: {DATA_DIR}\n"
        f"Log header: {log_header_status()}\n"
        f"Today checks: {counts['checks']}\n"
        f"Today log rows: {counts['log_rows']}\n"
        f"Today overtime done: {counts['overtime_hours']:.1f} h\n"
        f"Today overtime accounted: {counts['overtime_accounted_hours']:.1f} h\n"
        f"Quiet until: {quiet_until_text()}\n"
        f"Next prompt due: {next_prompt_due_text(interval_minutes)}"
    )
    print(text)
    message_dialog("Tracker status", text)


def run_doctor(interval_minutes: int = 30) -> None:
    ensure_data_dir()
    repaired = repair_log_csv(verbose=True)
    ensure_overtime_workbook()
    status_line, runner_details = launchagent_status()
    counts = today_counts()
    repo = Path(__file__).resolve().parent
    problems: list[str] = []
    if log_header_status() != "OK":
        problems.append("autotime_log.csv header still needs repair")
    if sys.platform == "darwin" and "com.gunjan.accountability" not in status_line:
        problems.append("LaunchAgent is not loaded. Run ./start_accountability.sh 30")
    runner_path_hint = str(repo)
    if "runner_content=" in runner_details and runner_path_hint not in runner_details:
        problems.append("LaunchAgent runner may point to an older folder. Run ./start_accountability.sh 30 from this repo.")

    text = (
        f"Accountability Tracker doctor ({APP_VERSION})\n"
        f"Script folder: {repo}\n"
        f"Data folder: {DATA_DIR}\n"
        f"Log header: {log_header_status()}\n"
        f"Log repair run: {'yes' if repaired else 'not needed'}\n"
        f"Today checks: {counts['checks']}\n"
        f"Today overtime done: {counts['overtime_hours']:.1f} h\n"
        f"Today overtime accounted: {counts['overtime_accounted_hours']:.1f} h\n"
        f"Quiet until: {quiet_until_text()}\n"
        f"Next prompt due: {next_prompt_due_text(interval_minutes)}\n"
        f"LaunchAgent: {status_line}\n"
        f"Runner: {runner_details}\n"
        f"Problems: {', '.join(problems) if problems else 'none detected'}"
    )
    print(text)
    message_dialog("Tracker doctor", text)


def open_paths() -> None:
    ensure_data_dir()
    ensure_overtime_workbook()
    if sys.platform == "darwin":
        subprocess.run(["open", str(DATA_DIR)])
    elif sys.platform.startswith("win"):
        subprocess.run(["explorer", str(DATA_DIR)])
    else:
        subprocess.run(["xdg-open", str(DATA_DIR)])


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Self-accountability and overtime tracker")
    parser.add_argument("--scheduled", action="store_true", help="Run as scheduled launchd check")
    parser.add_argument("--interval-minutes", type=int, default=30, help="Visible prompt interval")
    parser.add_argument("--day-setup", action="store_true", help="Force daily setup popup")
    parser.add_argument("--add-block", action="store_true", help="Manually add a time block")
    parser.add_argument("--meeting", action="store_true", help="Add a meeting/focus block and pause prompts until it ends")
    parser.add_argument("--open", action="store_true", help="Open data folder")
    parser.add_argument("--list-todos", action="store_true", help="List today's active todos")
    parser.add_argument("--add-todo", type=str, help="Add a todo for today")
    parser.add_argument("--done", type=str, help="Mark todo id done")
    parser.add_argument("--summary", action="store_true", help="Show today's daily summary popup")
    parser.add_argument("--review", action="store_true", help="Run end-of-day review now")
    parser.add_argument("--carry-over", action="store_true", help="Offer todo carry-over now")
    parser.add_argument("--status", action="store_true", help="Show tracker status and next prompt information")
    parser.add_argument("--doctor", action="store_true", help="Run health check and repair common log/header problems")
    parser.add_argument("--repair-logs", action="store_true", help="Repair autotime_log.csv header/legacy rows")
    parser.add_argument("--account-overtime", action="store_true", help="Mark today's overtime as already added to time management")
    args = parser.parse_args(argv)

    ensure_data_dir()
    ensure_overtime_workbook()

    if args.day_setup:
        day_setup(force=True)
    elif args.add_block or args.meeting:
        add_block()
    elif args.open:
        open_paths()
    elif args.list_todos:
        list_todos()
    elif args.add_todo:
        add_todo_cli(args.add_todo)
    elif args.done:
        done_todo_cli(args.done)
    elif args.summary:
        show_daily_summary()
    elif args.review:
        run_daily_review(force=True)
    elif args.carry_over:
        maybe_offer_todo_carryover(force=True)
    elif args.status:
        show_status(max(1, args.interval_minutes))
    elif args.doctor:
        run_doctor(max(1, args.interval_minutes))
    elif args.repair_logs:
        repair_log_csv(verbose=True)
    elif args.account_overtime:
        ask_overtime_accounting(today_iso())
    elif args.scheduled:
        scheduled_run(max(1, args.interval_minutes))
    else:
        # Manual normal run: no rate limit.
        status = get_day_status()
        if not status:
            day_setup(force=True)
            status = get_day_status()
        if status and is_inside_planned(status):
            regular_prompt(max(1, args.interval_minutes))
        else:
            outside_hours_prompt(max(1, args.interval_minutes), status)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
