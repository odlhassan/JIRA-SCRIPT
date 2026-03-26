from __future__ import annotations

import json
import re
import sqlite3
from pathlib import Path


def normalize_issue_key(value: object) -> str:
    return str(value or "").strip().upper()


def normalize_text(value: object) -> str:
    return str(value or "").strip()


def normalize_date_text(value: object) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    match = re.match(r"^(\d{4}-\d{2}-\d{2})", text)
    if match:
        return match.group(1)
    return text[:10]


def parse_json_object(value: object) -> dict:
    text = normalize_text(value)
    if not text:
        return {}
    try:
        parsed = json.loads(text)
    except Exception:
        return {}
    return parsed if isinstance(parsed, dict) else {}


def parse_planner_hours_from_man_days(value: object) -> float | None:
    text = normalize_text(value)
    if not text:
        return None
    try:
        man_days = float(text)
    except (TypeError, ValueError):
        return None
    if man_days != man_days or man_days < 0:
        return None
    return round(man_days * 8.0, 4)


def _sqlite_table_exists(conn: sqlite3.Connection, table_name: str) -> bool:
    row = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (normalize_text(table_name),),
    ).fetchone()
    return bool(row)


def load_epics_planner_epic_plan_by_key(db_path: Path) -> dict[str, dict[str, object]]:
    if not db_path.exists():
        return {}
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
    except Exception:
        return {}
    try:
        if not _sqlite_table_exists(conn, "epics_management"):
            return {}
        rows = conn.execute(
            "SELECT epic_key, epic_plan_json FROM epics_management"
        ).fetchall()
    except Exception:
        return {}
    finally:
        conn.close()

    lookup: dict[str, dict[str, object]] = {}
    for row in rows:
        epic_key = normalize_issue_key(row["epic_key"])
        if not epic_key:
            continue
        epic_plan = parse_json_object(row["epic_plan_json"])
        lookup[epic_key] = {
            "planner_start_date": normalize_date_text(epic_plan.get("start_date")),
            "planner_end_date": normalize_date_text(epic_plan.get("due_date")),
            "planner_planned_hours": parse_planner_hours_from_man_days(epic_plan.get("man_days")),
        }
    return lookup


def load_epics_planner_story_dates_by_key(db_path: Path) -> dict[str, dict[str, object]]:
    if not db_path.exists():
        return {}
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
    except Exception:
        return {}
    try:
        if not _sqlite_table_exists(conn, "epics_management_story_sync"):
            return {}
        columns = conn.execute("PRAGMA table_info(epics_management_story_sync)").fetchall()
        column_names = {normalize_text(col[1]) for col in columns}
        has_estimate_hours = "estimate_hours" in column_names
        if has_estimate_hours:
            rows = conn.execute(
                "SELECT story_key, start_date, due_date, estimate_hours FROM epics_management_story_sync"
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT story_key, start_date, due_date FROM epics_management_story_sync"
            ).fetchall()
    except Exception:
        return {}
    finally:
        conn.close()

    lookup: dict[str, dict[str, object]] = {}
    for row in rows:
        story_key = normalize_issue_key(row["story_key"])
        if not story_key:
            continue
        estimate_hours = None
        try:
            raw_hours = row["estimate_hours"] if "estimate_hours" in row.keys() else None
            estimate_hours = float(raw_hours) if raw_hours is not None else None
        except Exception:
            estimate_hours = None
        start_date = normalize_date_text(row["start_date"])
        due_date = normalize_date_text(row["due_date"])
        if not start_date and not due_date and estimate_hours is None:
            continue
        lookup[story_key] = {
            "start_date": start_date,
            "due_date": due_date,
            "estimate_hours": estimate_hours,
        }
    return lookup


def load_epics_management_dashboard_meta(db_path: Path) -> dict[str, dict[str, str]]:
    if not db_path.exists():
        return {}
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
    except Exception:
        return {}
    try:
        if not _sqlite_table_exists(conn, "epics_management"):
            return {}
        rows = conn.execute(
            """
            SELECT epic_key, ipp_meeting_planned, actual_production_date, remarks, epic_plan_json
            FROM epics_management
            """
        ).fetchall()
    except Exception:
        return {}
    finally:
        conn.close()

    lookup: dict[str, dict[str, str]] = {}
    for row in rows:
        epic_key = normalize_issue_key(row["epic_key"])
        if not epic_key:
            continue
        epic_plan = parse_json_object(row["epic_plan_json"])
        lookup[epic_key] = {
            "latest_ipp_meeting": "Yes" if normalize_text(row["ipp_meeting_planned"]).lower() == "yes" else "No",
            "ipp_planned_start_date": normalize_date_text(epic_plan.get("start_date")),
            "ipp_planned_end_date": normalize_date_text(epic_plan.get("due_date")),
            "ipp_actual_date": normalize_date_text(row["actual_production_date"]),
            "ipp_remarks": normalize_text(row["remarks"]),
        }
    return lookup


def apply_epics_management_ipp_fields(
    item: dict[str, object],
    epic_key: str,
    epic_meta_by_key: dict[str, dict[str, str]],
    *,
    jira_start_date: object = "",
    jira_end_date: object = "",
) -> None:
    normalized_epic_key = normalize_issue_key(epic_key)
    meta = epic_meta_by_key.get(normalized_epic_key) or {}
    planner_start = normalize_date_text(meta.get("ipp_planned_start_date"))
    planner_end = normalize_date_text(meta.get("ipp_planned_end_date"))
    jira_start = normalize_date_text(jira_start_date or item.get("jira_start_date"))
    jira_end = normalize_date_text(jira_end_date or item.get("jira_end_date"))
    actual_date = normalize_date_text(meta.get("ipp_actual_date"))

    item["latest_ipp_meeting"] = normalize_text(meta.get("latest_ipp_meeting")) or "No"
    item["ipp_planned_start_date"] = planner_start
    item["ipp_planned_end_date"] = planner_end
    item["ipp_actual_date"] = actual_date
    item["ipp_remarks"] = normalize_text(meta.get("ipp_remarks"))

    planner_has_dates = bool(planner_start or planner_end)
    if planner_has_dates and (planner_start != jira_start or planner_end != jira_end):
        item["jira_ipp_rmi_dates_altered"] = "Yes"
    else:
        item["jira_ipp_rmi_dates_altered"] = "No"

    item["ipp_actual_matches_jira_end_date"] = (
        "Yes" if actual_date and jira_end and actual_date == jira_end else "No"
    )
