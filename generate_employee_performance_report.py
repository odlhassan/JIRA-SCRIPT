from __future__ import annotations

import argparse
import json
import os
import re
import sqlite3
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from generate_assignee_hours_report import _list_capacity_profiles
from manage_fields_registry import load_manage_fields
from report_entity_registry import load_report_entities

DEFAULT_WORKLOG_INPUT_XLSX = "2_jira_subtask_worklogs.xlsx"
DEFAULT_WORK_ITEMS_INPUT_XLSX = "1_jira_work_items_export.xlsx"
DEFAULT_LEAVE_REPORT_INPUT_XLSX = "rlt_leave_report.xlsx"
DEFAULT_HTML_OUTPUT = "employee_performance_report.html"
DEFAULT_CAPACITY_DB = "assignee_hours_capacity.db"
LEAVE_HOURS_PER_DAY = 8.0

DEFAULT_PERFORMANCE_SETTINGS: dict[str, float] = {
    "base_score": 100.0,
    "min_score": 0.0,
    "max_score": 100.0,
    "points_per_bug_hour": 0.5,
    "points_per_bug_late_hour": 1.5,
    "points_per_unplanned_leave_hour": 0.75,
    "points_per_subtask_late_hour": 1.0,
    "points_per_estimate_overrun_hour": 1.25,
    "points_per_missed_due_date": 2.0,
}


def _to_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _to_float(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _resolve_path(value: str, base_dir: Path) -> Path:
    path = Path(value)
    return path if path.is_absolute() else base_dir / path


def _parse_iso_date(value: Any) -> str:
    text = _to_text(value)
    if not text:
        return ""
    if len(text) >= 10:
        try:
            date.fromisoformat(text[:10])
            return text[:10]
        except ValueError:
            pass
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f%z", "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return ""


def _normalize_header_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "_", _to_text(value).strip().lower()).strip("_")


def _extract_project_key(issue_id: str) -> str:
    text = _to_text(issue_id).upper()
    if not text or "-" not in text:
        return "UNKNOWN"
    project = text.split("-", 1)[0].strip()
    return project if re.match(r"^[A-Z0-9]+$", project) else "UNKNOWN"


def _is_bug_type(text: str) -> bool:
    low = _to_text(text).lower()
    return "bug" in low and ("subtask" in low or "sub-task" in low or "task" in low)


def _normalize_performance_settings(payload: dict, require_all_fields: bool = True) -> dict[str, float]:
    source = payload or {}
    out: dict[str, float] = {}
    for key, default in DEFAULT_PERFORMANCE_SETTINGS.items():
        if key not in source and require_all_fields:
            raise ValueError(f"Missing field: {key}")
        value = _to_float(source.get(key, default))
        if key.startswith("points_per_") and value < 0:
            raise ValueError(f"{key} must be >= 0")
        out[key] = round(value, 4)
    if out["min_score"] > out["base_score"] or out["base_score"] > out["max_score"]:
        raise ValueError("Score bounds must satisfy: max_score >= base_score >= min_score")
    return out


def _init_performance_settings_db(db_path: Path) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS performance_point_settings (
                id INTEGER PRIMARY KEY CHECK(id = 1),
                base_score REAL NOT NULL,
                min_score REAL NOT NULL,
                max_score REAL NOT NULL,
                points_per_bug_hour REAL NOT NULL,
                points_per_bug_late_hour REAL NOT NULL,
                points_per_unplanned_leave_hour REAL NOT NULL,
                points_per_subtask_late_hour REAL NOT NULL,
                points_per_estimate_overrun_hour REAL NOT NULL,
                points_per_missed_due_date REAL NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS performance_teams (
                team_name TEXT PRIMARY KEY,
                team_leader TEXT NOT NULL DEFAULT '',
                assignees_json TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        team_cols = [str(row[1]).lower() for row in conn.execute("PRAGMA table_info(performance_teams)").fetchall()]
        if "team_leader" not in team_cols:
            conn.execute("ALTER TABLE performance_teams ADD COLUMN team_leader TEXT NOT NULL DEFAULT ''")
        settings_cols = [str(row[1]).lower() for row in conn.execute("PRAGMA table_info(performance_point_settings)").fetchall()]
        if "points_per_missed_due_date" not in settings_cols:
            conn.execute("ALTER TABLE performance_point_settings ADD COLUMN points_per_missed_due_date REAL NOT NULL DEFAULT 2.0")
        row = conn.execute("SELECT id FROM performance_point_settings WHERE id = 1").fetchone()
        if not row:
            defaults = _normalize_performance_settings(DEFAULT_PERFORMANCE_SETTINGS, require_all_fields=True)
            conn.execute(
                """
                INSERT INTO performance_point_settings (
                    id, base_score, min_score, max_score,
                    points_per_bug_hour, points_per_bug_late_hour, points_per_unplanned_leave_hour,
                    points_per_subtask_late_hour, points_per_estimate_overrun_hour, points_per_missed_due_date, updated_at
                ) VALUES (1, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    defaults["base_score"],
                    defaults["min_score"],
                    defaults["max_score"],
                    defaults["points_per_bug_hour"],
                    defaults["points_per_bug_late_hour"],
                    defaults["points_per_unplanned_leave_hour"],
                    defaults["points_per_subtask_late_hour"],
                    defaults["points_per_estimate_overrun_hour"],
                    defaults["points_per_missed_due_date"],
                    datetime.now(timezone.utc).isoformat(),
                ),
            )
        conn.commit()


def _load_performance_settings(db_path: Path) -> dict[str, float]:
    _init_performance_settings_db(db_path)
    with sqlite3.connect(db_path) as conn:
        row = conn.execute(
            """
            SELECT base_score, min_score, max_score, points_per_bug_hour, points_per_bug_late_hour,
                   points_per_unplanned_leave_hour, points_per_subtask_late_hour, points_per_estimate_overrun_hour, points_per_missed_due_date
            FROM performance_point_settings WHERE id = 1
            """
        ).fetchone()
    if not row:
        return dict(DEFAULT_PERFORMANCE_SETTINGS)
    return _normalize_performance_settings(
        {
            "base_score": row[0],
            "min_score": row[1],
            "max_score": row[2],
            "points_per_bug_hour": row[3],
            "points_per_bug_late_hour": row[4],
            "points_per_unplanned_leave_hour": row[5],
            "points_per_subtask_late_hour": row[6],
            "points_per_estimate_overrun_hour": row[7],
            "points_per_missed_due_date": row[8],
        },
        require_all_fields=True,
    )


def _save_performance_settings(db_path: Path, payload: dict) -> dict[str, float]:
    _init_performance_settings_db(db_path)
    normalized = _normalize_performance_settings(payload, require_all_fields=True)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            UPDATE performance_point_settings
            SET base_score=?, min_score=?, max_score=?, points_per_bug_hour=?, points_per_bug_late_hour=?,
                points_per_unplanned_leave_hour=?, points_per_subtask_late_hour=?, points_per_estimate_overrun_hour=?, points_per_missed_due_date=?, updated_at=?
            WHERE id=1
            """,
            (
                normalized["base_score"],
                normalized["min_score"],
                normalized["max_score"],
                normalized["points_per_bug_hour"],
                normalized["points_per_bug_late_hour"],
                normalized["points_per_unplanned_leave_hour"],
                normalized["points_per_subtask_late_hour"],
                normalized["points_per_estimate_overrun_hour"],
                normalized["points_per_missed_due_date"],
                datetime.now(timezone.utc).isoformat(),
            ),
        )
        conn.commit()
    return normalized


def _normalize_team_name(value: Any) -> str:
    name = _to_text(value)
    if not name:
        raise ValueError("Team name is required.")
    if len(name) > 80:
        raise ValueError("Team name is too long (max 80 chars).")
    return name


def _normalize_assignees(values: list[Any]) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for raw in values or []:
        value = _to_text(raw)
        if not value:
            continue
        key = value.casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(value)
    if not out:
        raise ValueError("Select at least one assignee.")
    out.sort(key=lambda s: s.casefold())
    return out


def _normalize_team_leader(value: Any, assignees: list[str]) -> str:
    leader = _to_text(value)
    if not leader:
        raise ValueError("Team leader is required.")
    if leader.casefold() not in {a.casefold() for a in assignees}:
        raise ValueError("Team leader must be one of selected assignees.")
    for a in assignees:
        if a.casefold() == leader.casefold():
            return a
    return leader


def _list_performance_teams(db_path: Path) -> list[dict]:
    _init_performance_settings_db(db_path)
    with sqlite3.connect(db_path) as conn:
        rows = conn.execute(
            """
            SELECT team_name, team_leader, assignees_json, updated_at
            FROM performance_teams
            ORDER BY lower(team_name) ASC
            """
        ).fetchall()
    out: list[dict] = []
    for row in rows:
        try:
            assignees = json.loads(_to_text(row[2]))
        except json.JSONDecodeError:
            assignees = []
        out.append(
            {
                "team_name": _to_text(row[0]),
                "team_leader": _to_text(row[1]),
                "assignees": [str(a) for a in (assignees or []) if _to_text(a)],
                "updated_at": _to_text(row[3]),
            }
        )
    return out


def _save_performance_team(db_path: Path, team_name: Any, assignees: list[Any], team_leader: Any) -> dict:
    _init_performance_settings_db(db_path)
    normalized_name = _normalize_team_name(team_name)
    normalized_assignees = _normalize_assignees(assignees)
    normalized_leader = _normalize_team_leader(team_leader, normalized_assignees)
    updated_at = datetime.now(timezone.utc).isoformat()
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO performance_teams (team_name, team_leader, assignees_json, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(team_name) DO UPDATE SET
              team_leader=excluded.team_leader,
              assignees_json=excluded.assignees_json,
              updated_at=excluded.updated_at
            """,
            (normalized_name, normalized_leader, json.dumps(normalized_assignees, ensure_ascii=True), updated_at),
        )
        conn.commit()
    return {
        "team_name": normalized_name,
        "team_leader": normalized_leader,
        "assignees": normalized_assignees,
        "updated_at": updated_at,
    }


def _delete_performance_team(db_path: Path, team_name: Any) -> bool:
    _init_performance_settings_db(db_path)
    normalized_name = _normalize_team_name(team_name)
    with sqlite3.connect(db_path) as conn:
        cur = conn.execute("DELETE FROM performance_teams WHERE team_name = ?", (normalized_name,))
        conn.commit()
        return cur.rowcount > 0


def _load_work_items(path: Path) -> dict[str, dict]:
    if not path.exists():
        return {}
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not header:
            return {}
        headers = [_to_text(h) for h in header]
        idx = {name: i for i, name in enumerate(headers)}
        if "issue_key" not in idx:
            return {}
        issue_type_col = "jira_issue_type" if "jira_issue_type" in idx else ("work_item_type" if "work_item_type" in idx else "")
        out: dict[str, dict] = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            issue_key = _to_text(row[idx["issue_key"]]).upper()
            if not issue_key:
                continue
            out[issue_key] = {
                "issue_key": issue_key,
                "project_key": _to_text(row[idx.get("project_key", -1)]).upper() if "project_key" in idx else _extract_project_key(issue_key),
                "issue_type": _to_text(row[idx[issue_type_col]]) if issue_type_col else "",
                "fix_type": _to_text(row[idx["fix_type"]]).lower() if "fix_type" in idx else "",
                "summary": _to_text(row[idx["summary"]]) if "summary" in idx else "",
                "status": _to_text(row[idx["status"]]) if "status" in idx else "",
                "assignee": _to_text(row[idx["assignee"]]) if "assignee" in idx else "",
                "start_date": _parse_iso_date(row[idx["start_date"]]) if "start_date" in idx else "",
                "due_date": _parse_iso_date(row[idx["end_date"]]) if "end_date" in idx else "",
                "resolved_stable_since_date": _parse_iso_date(row[idx["resolved_stable_since_date"]]) if "resolved_stable_since_date" in idx else "",
                "original_estimate_hours": round(_to_float(row[idx["original_estimate_hours"]]), 2) if "original_estimate_hours" in idx else 0.0,
                "parent_issue_key": _to_text(row[idx["parent_issue_key"]]).upper() if "parent_issue_key" in idx else "",
            }
        return out
    finally:
        wb.close()


def _load_worklogs(path: Path, work_items: dict[str, dict]) -> list[dict]:
    if not path.exists():
        return []
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not header:
            return []
        headers = [_to_text(h) for h in header]
        idx = {name: i for i, name in enumerate(headers)}
        required = {"issue_id", "issue_assignee", "worklog_started", "hours_logged"}
        if not required.issubset(set(idx)):
            return []
        out: list[dict] = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            issue_id = _to_text(row[idx["issue_id"]]).upper()
            if not issue_id:
                continue
            worklog_date = _parse_iso_date(row[idx["worklog_started"]])
            hours = round(_to_float(row[idx["hours_logged"]]), 2)
            if not worklog_date or hours <= 0:
                continue
            assignee = _to_text(row[idx["issue_assignee"]]) or "Unassigned"
            item = work_items.get(issue_id, {})
            parent_story_id = _to_text(row[idx["parent_story_id"]]).upper() if "parent_story_id" in idx else ""
            parent_story_id = parent_story_id or _to_text(item.get("parent_issue_key")).upper()
            story_item = work_items.get(parent_story_id, {}) if parent_story_id else {}
            raw_type = _to_text(row[idx["issue_type"]]) if "issue_type" in idx else ""
            type_hint = raw_type or _to_text(item.get("issue_type"))
            out.append(
                {
                    "issue_id": issue_id,
                    "issue_assignee": assignee,
                    "worklog_date": worklog_date,
                    "hours_logged": hours,
                    "project_key": _to_text(item.get("project_key")) or _extract_project_key(issue_id),
                    "is_bug": _is_bug_type(type_hint),
                    "fix_type": _to_text(item.get("fix_type")).lower(),
                    "item_summary": _to_text(item.get("summary")),
                    "item_status": _to_text(item.get("status")),
                    "item_issue_type": _to_text(item.get("issue_type")),
                    "item_assignee": _to_text(item.get("assignee")),
                    "item_parent_issue_key": _to_text(item.get("parent_issue_key")).upper(),
                    "item_start_date": _to_text(item.get("start_date")),
                    "item_due_date": _to_text(item.get("due_date")),
                    "item_resolved_stable_since_date": _to_text(item.get("resolved_stable_since_date")),
                    "story_due_date": _to_text(story_item.get("due_date")),
                    "original_estimate_hours": round(_to_float(item.get("original_estimate_hours")), 2),
                }
            )
        return out
    finally:
        wb.close()


def _load_unplanned_leave_rows(path: Path) -> list[dict]:
    if not path.exists():
        return []
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        preferred_sheets = ["Daily_Assignee", "Leaves", "Leave_Daily", "Daily_Leaves"]
        sheet_names = [name for name in preferred_sheets if name in wb.sheetnames]
        sheet_names.extend([name for name in wb.sheetnames if name not in sheet_names])
        for sheet_name in sheet_names:
            ws = wb[sheet_name]
            header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
            if not header:
                continue
            header_keys = [_normalize_header_key(h) for h in header]
            index_by_header = {key: i for i, key in enumerate(header_keys) if key}

            def idx(*aliases: str) -> int | None:
                for alias in aliases:
                    key = _normalize_header_key(alias)
                    if key in index_by_header:
                        return index_by_header[key]
                return None

            assignee_idx = idx("assignee", "employee", "resource_name")
            day_idx = idx("period_day", "day", "date", "leave_date")
            planned_idx = idx("planned_taken_hours", "planned_leave_hours", "planned_hours")
            unplanned_idx = idx("unplanned_taken_hours", "unplanned_leave_hours", "unplanned_hours")
            leave_type_idx = idx("leave_type", "type")
            leave_hours_idx = idx("leave_hours", "hours", "taken_hours")
            leave_days_idx = idx("leave_days", "days")

            has_wide_schema = planned_idx is not None and unplanned_idx is not None
            has_long_schema = leave_type_idx is not None and (leave_hours_idx is not None or leave_days_idx is not None)
            if assignee_idx is None or day_idx is None or (not has_wide_schema and not has_long_schema):
                continue

            out: list[dict] = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                iso_day = _parse_iso_date(row[day_idx])
                if not iso_day:
                    continue
                assignee = _to_text(row[assignee_idx]) or "Unassigned"
                planned_hours = 0.0
                unplanned_hours = 0.0
                if has_wide_schema:
                    planned_hours = round(_to_float(row[planned_idx]), 2)
                    unplanned_hours = round(_to_float(row[unplanned_idx]), 2)
                else:
                    leave_type = _to_text(row[leave_type_idx]).lower()
                    hours = _to_float(row[leave_hours_idx]) if leave_hours_idx is not None else 0.0
                    if leave_days_idx is not None:
                        hours = hours or (_to_float(row[leave_days_idx]) * LEAVE_HOURS_PER_DAY)
                    hours = round(hours, 2)
                    if "planned" in leave_type and "unplanned" not in leave_type:
                        planned_hours = hours
                    elif "unplanned" in leave_type:
                        unplanned_hours = hours
                    else:
                        continue
                out.append(
                    {
                        "assignee": assignee,
                        "period_day": iso_day,
                        "unplanned_taken_hours": unplanned_hours,
                        "planned_taken_hours": planned_hours,
                    }
                )
            return out
        return []
    finally:
        wb.close()


def _default_range(rows: list[dict]) -> tuple[str, str]:
    today = datetime.now(timezone.utc).date()
    month_start = date(today.year, today.month, 1)
    next_month_start = date(today.year + (1 if today.month == 12 else 0), 1 if today.month == 12 else today.month + 1, 1)
    month_end = next_month_start - timedelta(days=1)
    return month_start.isoformat(), month_end.isoformat()


def _build_payload(
    worklogs: list[dict],
    work_items: list[dict],
    leave_rows: list[dict],
    settings: dict[str, float],
    teams: list[dict],
    entities_catalog: list[dict],
    managed_fields: list[dict],
    capacity_profiles: list[dict],
) -> dict:
    projects = sorted(
        {
            *({_to_text(r.get("project_key")) or "UNKNOWN" for r in worklogs}),
            *({_to_text(r.get("project_key")) or "UNKNOWN" for r in work_items}),
        }
    )
    default_from, default_to = _default_range(worklogs)
    return {
        "worklogs": worklogs,
        "work_items": work_items,
        "leave_rows": leave_rows,
        "teams": teams or [],
        "projects": projects,
        "default_from": default_from,
        "default_to": default_to,
        "leave_hours_per_day": LEAVE_HOURS_PER_DAY,
        "settings": settings,
        "entities_catalog": entities_catalog or [],
        "managed_fields": managed_fields or [],
        "capacity_profiles": capacity_profiles or [],
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
    }


def _build_html(payload: dict) -> str:
    data = json.dumps(payload, ensure_ascii=True)
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Employee Performance Dashboard</title>
  <style>
    :root {{ --bg:#0a1325; --panel:#121f39; --line:#2a3f65; --ink:#dce8ff; --muted:#9db1d8; --bad:#fb7185; --good:#34d399; }}
    * {{ box-sizing:border-box; }} body {{ margin:0; font-family:"Trebuchet MS","Segoe UI",sans-serif; color:var(--ink); background:radial-gradient(circle at 5% 0%,#1a3356 0%,#0a1325 60%); }}
    .wrap {{ max-width:1800px; margin:0 auto; padding:14px; }} .hero {{ background:#101e37; border:1px solid #34507e; border-radius:12px; padding:12px; }}
    .meta {{ color:#c4d4ef; font-size:.8rem; margin-top:4px; }} .toolbar {{ display:grid; gap:8px; grid-template-columns:145px 145px minmax(180px,1fr) minmax(210px,1fr) auto auto auto; margin-top:8px; }}
    .shortcut-bar {{ display:flex; gap:8px; flex-wrap:wrap; margin-top:8px; }}
    .shortcut-btn {{ border:1px solid #3f5f93; background:#0f2342; color:#dce8ff; border-radius:999px; font-size:.74rem; padding:5px 10px; cursor:pointer; }}
    .shortcut-btn:hover {{ background:#17325a; }}
    .f label {{ display:block; font-size:.7rem; color:var(--muted); margin-bottom:3px; text-transform:uppercase; font-weight:700; }} .f input,.f select {{ width:100%; border:1px solid #3a5c91; border-radius:8px; background:#0d1830; color:var(--ink); padding:7px; }}
    #projects option {{ color:#0f172a; background:#ffffff; }}
    .btn {{ border:1px solid #4a6ea9; background:#1b325a; color:#eef4ff; border-radius:8px; font-weight:700; padding:7px 10px; cursor:pointer; text-decoration:none; display:inline-flex; align-items:center; }}
    .guide {{ margin-top:8px; border:1px solid #2f517e; border-radius:10px; background:#0f2140; padding:8px; }}
    .guide h2 {{ margin:0; font-size:.86rem; }}
    .guide p {{ margin:5px 0 0; font-size:.77rem; color:#c7d7f3; line-height:1.35; }}
    .discover {{ display:flex; gap:8px; flex-wrap:wrap; margin-top:8px; }}
    .discover .pill {{ border:1px solid #365c8d; border-radius:999px; padding:4px 10px; font-size:.74rem; color:#dce8ff; background:#132949; }}
    .leader-controls {{ display:flex; gap:8px; flex-wrap:wrap; align-items:end; padding:8px 10px 0; }}
    .leader-controls .f {{ min-width:160px; }}
    .section-head {{ margin:10px 0 4px; font-size:.8rem; color:var(--muted); text-transform:uppercase; letter-spacing:.04em; font-weight:800; }}
    .section-head.collapse-toggle {{ cursor:pointer; user-select:none; }}
    .section-head.collapse-toggle .hint {{ font-size:.68rem; color:#7fa3d6; margin-left:8px; text-transform:none; letter-spacing:0; }}
    .is-collapsed {{ display:none; }}
    .kpis {{ display:grid; gap:8px; grid-template-columns:repeat(4,minmax(0,1fr)); margin-top:8px; }} .kpi {{ border:1px solid var(--line); border-radius:10px; background:var(--panel); padding:8px; }} .kpi .k {{ font-size:.72rem; color:var(--muted); text-transform:uppercase; }} .kpi .v {{ margin-top:4px; font-size:1.1rem; font-weight:800; }}
    .planning-kpis .kpi {{ border-color:#315a86; background:#10223f; }}
    .top3-wrap {{ display:grid; gap:8px; grid-template-columns:repeat(2,minmax(0,1fr)); margin-top:8px; }}
    .top3-card {{ border:1px solid var(--line); border-radius:10px; background:var(--panel); padding:8px; }}
    .top3-title {{ font-size:.74rem; color:var(--muted); text-transform:uppercase; font-weight:700; margin-bottom:5px; }}
    .top3-item {{ display:flex; justify-content:space-between; gap:8px; padding:3px 0; font-size:.82rem; }}
    .top3-item .nm {{ color:var(--ink); font-weight:700; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
    .top3-item .sc {{ font-weight:800; }}
    .top3-item.high .sc {{ color:var(--good); }}
    .top3-item.low .sc {{ color:var(--bad); }}
    .teams-wrap {{ margin-top:8px; border:1px solid var(--line); border-radius:10px; background:var(--panel); padding:8px; }}
    .teams-wrap.is-collapsed .team-score-layout {{ display:none; }}
    .teams-title {{ font-size:.74rem; color:var(--muted); text-transform:uppercase; font-weight:700; margin-bottom:6px; }}
    .team-score-layout {{ display:grid; gap:8px; grid-template-columns:7fr 3fr; align-items:stretch; }}
    .team-card {{ border:1px solid #314d7a; border-radius:8px; background:#111e37; padding:8px; cursor:pointer; transition:box-shadow .12s ease, border-color .12s ease, transform .12s ease; }}
    .team-card:hover {{ border-color:#5f86bf; transform:translateY(-1px); }}
    .team-card.sel {{ border-color:#6ea8ff; box-shadow:inset 0 0 0 1px #6ea8ff; }}
    .team-head {{ display:flex; justify-content:space-between; gap:8px; align-items:flex-start; }}
    .team-name {{ font-size:.86rem; font-weight:800; }}
    .team-sub {{ font-size:.74rem; color:var(--muted); margin-top:2px; }}
    .team-score {{ font-size:1.1rem; font-weight:900; }}
    .team-metrics {{ margin-top:6px; display:grid; gap:3px; font-size:.76rem; }}
    .team-chart-shell {{ border:1px solid #2b446e; border-radius:10px; background:#0f1b32; padding:8px; }}
    .team-chart-svg {{ width:100%; height:220px; display:block; border:1px solid #223a61; border-radius:8px; background:#0d172b; }}
    .team-detail-shell {{ border:1px solid #2b446e; border-radius:10px; background:#0f1b32; padding:8px; }}
    .team-detail-body {{ margin-top:6px; }}
    .arena {{ display:grid; gap:10px; grid-template-columns:minmax(320px,38%) minmax(0,62%); margin-top:10px; }} .panel {{ border:1px solid var(--line); border-radius:12px; background:var(--panel); overflow:hidden; }} .panel h2 {{ margin:0; padding:9px 10px; font-size:.9rem; border-bottom:1px solid var(--line); }}
    .leaderboard {{ max-height:72vh; overflow:auto; }} .row {{ display:grid; gap:6px; grid-template-columns:24px 1fr auto; align-items:center; padding:8px 10px; border-bottom:1px solid #243b61; cursor:pointer; }} .row.sel {{ background:#193766; box-shadow:inset 0 0 0 1px #5d89cf; }} .rank {{ color:#5eead4; font-weight:800; }} .sub {{ color:var(--muted); font-size:.72rem; }} .score {{ border:1px solid #3f5f93; border-radius:999px; padding:2px 8px; font-weight:800; display:inline-flex; gap:4px; align-items:center; }}
    .leader-metrics {{ display:flex; gap:6px; flex-wrap:wrap; margin-top:2px; }}
    .metric-chip {{ display:inline-flex; align-items:center; gap:4px; border:1px solid #36598a; border-radius:999px; padding:2px 8px; background:#112546; }}
    .metric-chip .metric-value {{ color:#e2ecff; font-weight:900; font-size:.74rem; }}
    .metric-chip .metric-value.warn {{ color:#f59e0b; }}
    .metric-chip .material-symbols-outlined {{ font-size:15px; color:#93c5fd; font-variation-settings:"FILL" 1, "wght" 500, "GRAD" 0, "opsz" 20; }}
    .detail {{ padding:10px; }} .card {{ border:1px solid #314e7f; border-radius:10px; background:#12213d; padding:10px; }} .big {{ font-size:2rem; font-weight:900; line-height:1; }}
    .tabs {{ display:flex; gap:6px; flex-wrap:wrap; margin-top:8px; }}
    .tab-btn {{ border:1px solid #3b5f91; background:#12284b; color:#e6efff; border-radius:999px; padding:4px 10px; font-size:.74rem; cursor:pointer; }}
    .tab-btn.active {{ border-color:#7cb2ff; box-shadow:inset 0 0 0 1px #7cb2ff; background:#173866; }}
    .tab-pane {{ display:none; }}
    .tab-pane.active {{ display:block; }}
    .grid2 {{ display:grid; gap:8px; grid-template-columns:repeat(2,minmax(0,1fr)); margin-top:8px; }} .mini {{ border:1px solid #2b446e; border-radius:10px; background:#0f1b32; padding:8px; }} .mini h3 {{ margin:0 0 6px; font-size:.8rem; }} .mini .l {{ display:flex; justify-content:space-between; font-size:.78rem; padding:2px 0; }}
    .mini .l.actionable {{ cursor:pointer; border-radius:6px; transition:background .15s ease; }}
    .mini .l.actionable:hover {{ background:#173158; }}
    .metric-link-btn {{ border:1px solid #36598a; border-radius:999px; background:#12284b; color:#dce8ff; font-size:.72rem; padding:2px 8px; cursor:pointer; }}
    .metric-link-btn:hover {{ background:#1a3a67; border-color:#5f88c0; }}
    .ts-card {{ margin-top:8px; border:1px solid #2b446e; border-radius:10px; background:#0f1b32; padding:8px; }}
    .ts-svg {{ width:100%; height:220px; display:block; border:1px solid #223a61; border-radius:8px; background:#0d172b; }}
    .kpi-charts {{ display:grid; gap:8px; grid-template-columns:repeat(2,minmax(0,1fr)); margin-top:8px; }}
    .mini-chart {{ border:1px solid #2b446e; border-radius:10px; background:#0f1b32; padding:8px; }}
    .mini-svg {{ width:100%; height:160px; display:block; border:1px solid #223a61; border-radius:8px; background:#0d172b; }}
    .tbl-wrap {{ margin-top:8px; border:1px solid #2b446e; border-radius:10px; background:#0f1b32; overflow:auto; }}
    .tbl-title {{ margin:0; padding:8px; font-size:.82rem; border-bottom:1px solid #21385e; }}
    .tbl {{ width:100%; border-collapse:collapse; min-width:720px; }}
    .tbl th,.tbl td {{ border-bottom:1px solid #21385e; padding:6px 8px; font-size:.76rem; vertical-align:top; }}
    .tbl th {{ color:#9db1d8; text-transform:uppercase; font-size:.68rem; letter-spacing:.04em; position:sticky; top:0; background:#10223f; }}
    .tbl tr.exec-negative-subtask td {{ background:rgba(244,63,94,.12); }}
    .tbl tr.exec-negative-subtask .issue-id,.tbl tr.exec-negative-subtask .issue-title {{ color:#fecdd3; }}
    .tbl tr.due-missed td {{ background:rgba(249,115,22,.12); }}
    .due-bucket {{ display:inline-flex; align-items:center; border-radius:999px; padding:2px 8px; border:1px solid transparent; font-weight:700; }}
    .due-before {{ background:rgba(34,197,94,.18); border-color:#22c55e; color:#bbf7d0; }}
    .due-on {{ background:rgba(56,189,248,.18); border-color:#38bdf8; color:#bae6fd; }}
    .due-after {{ background:rgba(249,115,22,.2); border-color:#f97316; color:#fed7aa; }}
    .exec-accordion {{ padding:8px; display:grid; gap:8px; }}
    .exec-epic {{ border:1px solid #2f4f7f; border-radius:10px; background:#10203c; overflow:hidden; }}
    .exec-epic > summary {{ list-style:none; cursor:pointer; display:flex; justify-content:space-between; gap:8px; align-items:flex-start; padding:8px; }}
    .exec-epic > summary::-webkit-details-marker {{ display:none; }}
    .exec-epic-left {{ min-width:0; }}
    .exec-epic-metrics {{ display:flex; gap:6px; flex-wrap:wrap; justify-content:flex-end; }}
    .exec-epic-body {{ border-top:1px solid #29476f; padding:8px; display:grid; gap:8px; }}
    .exec-story-block {{ border:1px solid #29476f; border-radius:10px; background:#0e1c35; overflow:hidden; }}
    .exec-story-head {{ display:flex; justify-content:space-between; gap:8px; align-items:flex-start; padding:8px; border-bottom:1px solid #21385e; }}
    .exec-story-metrics {{ display:flex; gap:6px; flex-wrap:wrap; justify-content:flex-end; }}
    .exec-subtask-table {{ min-width:0; border-top:0; }}
    .exec-neg-pill {{ border-color:#fb7185; color:#fecdd3; background:#3b1020; }}
    .focus-pulse {{ animation:focusPulse 1.2s ease; }}
    @keyframes focusPulse {{ 0% {{ box-shadow:0 0 0 0 rgba(56,189,248,.65); }} 100% {{ box-shadow:0 0 0 12px rgba(56,189,248,0); }} }}
    .issue-id {{ font-size:.67rem; color:#8fb1e8; font-family:Consolas,monospace; }}
    .issue-title {{ font-size:.83rem; color:#e2ebff; line-height:1.25; }}
    .tree-wrap {{ border:1px solid #2b446e; border-radius:10px; background:#0f1b32; padding:8px; margin-top:8px; }}
    .tree-node {{ margin:4px 0; }}
    .tree-node details {{ border:1px solid #223a61; border-radius:8px; background:#112241; padding:6px; }}
    .tree-node summary {{ cursor:pointer; list-style:none; display:flex; justify-content:space-between; gap:8px; align-items:flex-start; }}
    .tree-node summary::-webkit-details-marker {{ display:none; }}
    .tree-left {{ min-width:0; }}
    .tree-metrics {{ display:flex; gap:6px; flex-wrap:wrap; justify-content:flex-end; }}
    .metric-pill {{ border:1px solid #375f95; border-radius:999px; padding:2px 8px; font-size:.68rem; color:#d7e7ff; background:#17325a; white-space:nowrap; }}
    .tree-children {{ margin-left:14px; margin-top:6px; border-left:1px dashed #2e4f7d; padding-left:8px; }}
    .exec-metrics {{ display:grid; gap:10px; }}
    .exec-metric {{ border:1px solid #29486f; border-radius:8px; padding:8px; background:#10223f; }}
    .exec-metric.actionable {{ cursor:pointer; }}
    .exec-metric.actionable:hover {{ border-color:#5f88c0; background:#12284b; }}
    .exec-m-head {{ display:flex; justify-content:space-between; gap:8px; align-items:center; }}
    .exec-m-name {{ font-size:.86rem; font-weight:800; color:#e6f0ff; }}
    .exec-m-value {{ font-size:.88rem; font-weight:900; color:#7dd3fc; }}
    .exec-m-meaning {{ margin-top:2px; font-size:.74rem; color:#b7c9e8; line-height:1.3; }}
    .exec-bar-track {{ margin-top:6px; width:100%; height:12px; border-radius:999px; border:1px solid #2f4e7d; background:#0a1a33; overflow:hidden; }}
    .exec-bar-fill {{ height:100%; background:linear-gradient(90deg,#38bdf8,#60a5fa); }}
    .exec-scale-note {{ font-size:.72rem; color:#93acd2; margin-top:2px; }}
    .availability-breakdown {{ margin-top:8px; border-top:1px dashed #2f4e7d; padding-top:8px; display:grid; gap:4px; }}
    .availability-line {{ display:flex; justify-content:space-between; gap:8px; font-size:.75rem; }}
    .availability-name {{ color:#dce8ff; }}
    .availability-num {{ color:#f8fafc; font-weight:800; font-family:Consolas,monospace; }}
    .availability-note {{ color:#93acd2; font-size:.72rem; }}
    .neg {{ color:var(--bad); font-weight:700; }} .feed {{ margin-top:8px; border:1px solid #2b446e; border-radius:10px; max-height:230px; overflow:auto; background:#0e182d; }} .feed .i {{ padding:7px 8px; border-bottom:1px solid #21385e; font-size:.76rem; }} .empty {{ color:var(--muted); font-style:italic; }}
    @media (max-width:1200px) {{ .toolbar{{grid-template-columns:1fr 1fr;}} .kpis{{grid-template-columns:1fr 1fr;}} .top3-wrap{{grid-template-columns:1fr;}} .team-score-layout{{grid-template-columns:1fr;}} .arena{{grid-template-columns:1fr;}} }}
  </style>
  <link rel="stylesheet" href="shared-nav.css">
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:FILL,wght,GRAD,opsz@1,500,0,20">
</head>
<body>
<div class="wrap">
  <section class="hero">
    <h1>Employee Performance Dashboard</h1><div class="meta" id="meta"></div>
    <section class="guide">
      <h2>Executive View Guide</h2>
      <p>Start with Planning & Start-Adherence KPIs to check workload realism, then use Performance Score KPIs for risk posture. In leaderboard, sort by the lens you want and click a person for full diagnostic detail.</p>
      <div class="discover" id="discover-insights"></div>
    </section>
    <div class="toolbar">
      <div class="f"><label for="from">From</label><input id="from" type="date"></div>
      <div class="f"><label for="to">To</label><input id="to" type="date"></div>
      <div class="f"><label for="projects">Project</label><select id="projects" multiple size="1"></select></div>
      <div class="f"><label for="capacity-profile">Capacity Profile</label><select id="capacity-profile"></select></div>
      <div class="f"><label for="search">Search Assignee</label><input id="search" type="text"></div>
      <button id="apply" class="btn" type="button">Apply Filters</button>
      <button id="reset" class="btn" type="button">Reset</button>
      <a href="/settings/performance" class="btn">Performance Settings</a>
    </div>
    <div class="shortcut-bar">
      <button id="shortcut-current-month" class="shortcut-btn" type="button">Current Month</button>
      <button id="shortcut-previous-month" class="shortcut-btn" type="button">Previous Month</button>
      <button id="shortcut-last-30-days" class="shortcut-btn" type="button">Last 30 Days</button>
      <button id="shortcut-quarter-to-date" class="shortcut-btn" type="button">Quarter To Date</button>
      <button id="shortcut-reset" class="shortcut-btn" type="button">Reset</button>
    </div>
    <div class="meta" id="capacity-profile-meta"></div>
    <div class="section-head collapse-toggle" id="toggle-performance-kpis">Performance Score KPIs<span class="hint">click to expand</span></div>
    <div id="performance-kpis-wrap" class="is-collapsed">
    <div class="kpis" id="performance-kpis">
      <div class="kpi"><div class="k">Team Avg Score</div><div class="v" id="kpi-avg">0</div></div>
      <div class="kpi"><div class="k">Top Performer</div><div class="v" id="kpi-top">-</div></div>
      <div class="kpi"><div class="k">At Risk (&lt;60)</div><div class="v" id="kpi-risk">0</div></div>
      <div class="kpi"><div class="k">Total Penalty</div><div class="v" id="kpi-pen">0</div></div>
      <div class="kpi"><div class="k">Total Rework Hours</div><div class="v" id="kpi-rework">0h</div></div>
    </div>
    </div>
    <div class="top3-wrap">
      <div class="top3-card">
        <div class="top3-title">Top 3 High Performing</div>
        <div id="top3-high"></div>
      </div>
      <div class="top3-card">
        <div class="top3-title">Top 3 Lowest Performing</div>
        <div id="top3-low"></div>
      </div>
    </div>
    <div class="teams-wrap is-collapsed" id="team-performance-section">
      <div class="teams-title collapse-toggle" id="toggle-team-performance">Team Performance <span class="hint">click to expand</span></div>
      <div class="team-score-layout">
        <div class="team-chart-shell">
          <div class="teams-title">Team Score Chart</div>
          <div id="team-performance-chart"></div>
        </div>
        <div class="team-detail-shell">
          <div class="teams-title">Team Performance Card</div>
          <div id="selected-team-performance" class="team-detail-body"><div class="empty">Select a team from chart.</div></div>
        </div>
      </div>
    </div>
  </section>
  <section class="arena">
    <article class="panel"><h2 id="leaderboard-title">Leaderboard</h2><div class="leader-controls"><div class="f"><label for="leader-sort">Sort By</label><select id="leader-sort"><option value="rmis">RMIs In Range (Desc)</option><option value="score">Performance Score</option><option value="missed">Missed Start Ratio</option><option value="capacity_gap">Capacity Gap (Cap - Planned)</option></select></div><div class="f"><label for="filter-risk">At-Risk View</label><select id="filter-risk"><option value="all">All Assignees</option><option value="risk">Only At-Risk (&lt;60)</option></select></div><div class="f"><label for="filter-missed">Start Discipline</label><select id="filter-missed"><option value="all">All</option><option value="missed">Only Missed Starts</option></select></div></div><div id="leaderboard-filter" class="sub" style="padding:0 10px 8px;"></div><div id="leaderboard" class="leaderboard"></div></article>
    <article class="panel"><h2>Assignee Drilldown</h2><div id="detail" class="detail"><div class="empty">Select an assignee.</div></div></article>
  </section>
</div>
<script>
const payload = {data};
const worklogs = Array.isArray(payload.worklogs) ? payload.worklogs : [];
const workItems = Array.isArray(payload.work_items) ? payload.work_items : [];
const leaveRows = Array.isArray(payload.leave_rows) ? payload.leave_rows : [];
const teams = Array.isArray(payload.teams) ? payload.teams : [];
const projects = Array.isArray(payload.projects) ? payload.projects : [];
const entitiesCatalog = Array.isArray(payload.entities_catalog) ? payload.entities_catalog : [];
const managedFields = Array.isArray(payload.managed_fields) ? payload.managed_fields : [];
const capacityProfiles = Array.isArray(payload.capacity_profiles) ? payload.capacity_profiles : [];
const capacityProfileSelectEl = document.getElementById("capacity-profile");
const capacityProfileMetaEl = document.getElementById("capacity-profile-meta");
const workItemsByKey = new Map(workItems.map((row) => [String(row && row.issue_key || "").toUpperCase(), row || {{}}]));
const defaultFrom = payload.default_from || "";
const defaultTo = payload.default_to || "";
const leaveHoursPerDay = n(payload.leave_hours_per_day) > 0 ? n(payload.leave_hours_per_day) : 8;
const settings = payload.settings || {{}};
let selectedName = "";
let selectedTeam = "";
let availabilityBreakdownForAssignee = "";
function n(v) {{ const x = Number(v); return Number.isFinite(x) ? x : 0; }}
function e(t) {{ return String(t ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }}
function clamp(v, minv, maxv) {{ return Math.max(minv, Math.min(maxv, v)); }}
function inRange(day, from, to) {{ if (!day) return false; if (from && day < from) return false; if (to && day > to) return false; return true; }}
function matchesPlannedRange(row, from, to) {{
  const start = String(row && row.item_start_date || "");
  const end = String(row && row.item_due_date || "");
  if (!start && !end) return false;
  return inRange(start, from, to) || inRange(end, from, to);
}}
function selectedProjects() {{ return new Set(Array.from(document.getElementById("projects").selectedOptions).map(o => o.value)); }}
function addDayPenalty(rec, day, points) {{
  const d = String(day || "");
  const p = n(points);
  if (!d || p <= 0) return;
  rec.daily_penalty_by_day[d] = n(rec.daily_penalty_by_day[d]) + p;
}}
function dateRangeDays(from, to) {{
  const out = [];
  if (!from || !to || to < from) return out;
  let cur = new Date(from + "T00:00:00");
  const end = new Date(to + "T00:00:00");
  while (cur <= end) {{
    const y = cur.getFullYear();
    const m = String(cur.getMonth() + 1).padStart(2, "0");
    const d = String(cur.getDate()).padStart(2, "0");
    out.push(`${{y}}-${{m}}-${{d}}`);
    cur.setDate(cur.getDate() + 1);
  }}
  return out;
}}
function ensure(map, name) {{
  if (!map.has(name)) map.set(name, {{assignee:name, bug_hours:0, bug_late_hours:0, subtask_late_hours:0, estimate_overrun_hours:0, rework_hours:0, unplanned_leave_hours:0, planned_leave_hours:0, unplanned_leave_count:0, planned_leave_count:0, unplanned_leave_days:0, planned_leave_days:0, missing_story_due_count:0, missing_due_count:0, missing_estimate_issue_count:0, total_hours:0, planned_hours_assigned:0, employee_capacity_hours:0, assigned_counts:{{epic:0,story:0,subtask:0}}, total_assigned_count:0, due_dated_assigned_count:0, missed_start_count:0, missed_start_ratio:0, missed_due_date_count:0, missed_due_date_ratio:0, active_rmi_count:0, assigned_hierarchy:[], missed_start_items:[], due_compliance_items:[], start_day_activity:[], last_log_by_issue:{{}}, issue_logged_hours_by_issue:{{}}, subtask_late_by_issue:{{}}, entity_values:{{}}, managed_values:{{}}, managed_scope:{{}}, feed:[], daily_penalty_by_day:{{}}, daily_series:[]}});
  return map.get(name);
}}
function normalizeType(t) {{
  const low = String(t || "").toLowerCase();
  if (low.includes("epic")) return "epic";
  if (low.includes("story")) return "story";
  return "subtask";
}}
function issueTypeLabel(t) {{
  return String(t || "").trim().toLowerCase();
}}
function _isoDateLocal(dt) {{
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, "0");
  const d = String(dt.getDate()).padStart(2, "0");
  return `${{y}}-${{m}}-${{d}}`;
}}
function capacityProfileLabel(profile) {{
  const fromDate = String(profile?.from_date || "");
  const toDate = String(profile?.to_date || "");
  const std = n(profile?.standard_hours_per_day) > 0 ? n(profile.standard_hours_per_day) : 8;
  const ramadan = n(profile?.ramadan_hours_per_day) > 0 ? n(profile.ramadan_hours_per_day) : std;
  return `${{fromDate}} -> ${{toDate}} | Std ${{std.toFixed(1)}}h | Ramadan ${{ramadan.toFixed(1)}}h`;
}}
function refreshCapacityProfileOptions() {{
  if (!capacityProfileSelectEl) return;
  const options = [`<option value="auto">Auto (Match selected date range)</option>`];
  capacityProfiles.forEach((profile, idx) => {{
    options.push(`<option value="${{idx}}">${{e(capacityProfileLabel(profile))}}</option>`);
  }});
  capacityProfileSelectEl.innerHTML = options.join("");
  capacityProfileSelectEl.value = "auto";
}}
function resolveActiveCapacityProfile(fromIso, toIso) {{
  if (!capacityProfileSelectEl) return null;
  const selected = String(capacityProfileSelectEl.value || "auto");
  if (selected === "auto") {{
    return capacityProfiles.find((p) => String(p.from_date || "") === String(fromIso || "") && String(p.to_date || "") === String(toIso || "")) || null;
  }}
  const idx = Number(selected);
  if (Number.isInteger(idx) && idx >= 0 && idx < capacityProfiles.length) return capacityProfiles[idx];
  return null;
}}
function updateCapacityProfileMeta(fromIso, toIso, profile) {{
  if (!capacityProfileMetaEl) return;
  const selected = capacityProfileSelectEl ? String(capacityProfileSelectEl.value || "auto") : "auto";
  if (!profile) {{
    capacityProfileMetaEl.textContent = "Capacity profile: Default weekdays (8h/day)";
    return;
  }}
  const mode = selected === "auto" ? "Auto profile" : "Applied profile";
  capacityProfileMetaEl.textContent = `${{mode}}: ${{capacityProfileLabel(profile)}} | Active range: ${{String(fromIso || "")}} -> ${{String(toIso || "")}}`;
}}
function computePerAssigneeCapacity(fromIso, toIso, profile) {{
  const fromDate = String(fromIso || "");
  const toDate = String(toIso || "");
  if (!fromDate || !toDate || toDate < fromDate) return 0;
  const start = new Date(fromDate + "T00:00:00");
  const end = new Date(toDate + "T00:00:00");
  if (!Number.isFinite(start.getTime()) || !Number.isFinite(end.getTime())) return 0;
  const standardHours = n(profile?.standard_hours_per_day) > 0 ? n(profile.standard_hours_per_day) : 8;
  const ramadanHours = n(profile?.ramadan_hours_per_day) > 0 ? n(profile.ramadan_hours_per_day) : standardHours;
  const ramadanStart = String(profile?.ramadan_start_date || "");
  const ramadanEnd = String(profile?.ramadan_end_date || "");
  const holidaySet = new Set((Array.isArray(profile?.holiday_dates) ? profile.holiday_dates : []).map((day) => String(day || "")));
  let total = 0;
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {{
    const dow = d.getDay();
    if (dow === 0 || dow === 6) continue;
    const dayIso = _isoDateLocal(d);
    if (holidaySet.has(dayIso)) continue;
    const inRamadan = ramadanStart && ramadanEnd && dayIso >= ramadanStart && dayIso <= ramadanEnd;
    total += inRamadan ? ramadanHours : standardHours;
  }}
  return total;
}}
function computeBusinessDays(fromIso, toIso, profile) {{
  const fromDate = String(fromIso || "");
  const toDate = String(toIso || "");
  if (!fromDate || !toDate || toDate < fromDate) return 0;
  const start = new Date(fromDate + "T00:00:00");
  const end = new Date(toDate + "T00:00:00");
  if (!Number.isFinite(start.getTime()) || !Number.isFinite(end.getTime())) return 0;
  const holidaySet = new Set((Array.isArray(profile?.holiday_dates) ? profile.holiday_dates : []).map((day) => String(day || "")));
  let days = 0;
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {{
    const dow = d.getDay();
    if (dow === 0 || dow === 6) continue;
    const dayIso = _isoDateLocal(d);
    if (holidaySet.has(dayIso)) continue;
    days += 1;
  }}
  return days;
}}
function resolveParentEpicKey(wi) {{
  const parentKey = String(wi && wi.parent_issue_key || "").toUpperCase();
  if (!parentKey) return "";
  const parentRow = workItemsByKey.get(parentKey) || null;
  if (!parentRow) return "";
  const parentType = issueTypeLabel(parentRow.jira_issue_type || parentRow.issue_type || parentRow.work_item_type);
  if (parentType.includes("epic")) return parentKey;
  return String(parentRow.parent_issue_key || "").toUpperCase();
}}
function resolveEpicKeyFromWorklogRow(row) {{
  const issueKey = String(row && row.issue_id || "").toUpperCase();
  const issueRow = workItemsByKey.get(issueKey) || null;
  if (issueRow) {{
    return resolveParentEpicKey(issueRow);
  }}
  const parentKey = String(row && row.item_parent_issue_key || "").toUpperCase();
  if (!parentKey) return "";
  const parentRow = workItemsByKey.get(parentKey) || null;
  if (!parentRow) return "";
  const parentType = issueTypeLabel(parentRow.jira_issue_type || parentRow.issue_type || parentRow.work_item_type);
  if (parentType.includes("epic")) return parentKey;
  return String(parentRow.parent_issue_key || "").toUpperCase();
}}
function setDateFilterRange(from, to) {{
  document.getElementById("from").value = String(from || "");
  document.getElementById("to").value = String(to || "");
}}
function applyDateShortcut(kind) {{
  const now = new Date();
  const today = _isoDateLocal(now);
  if (kind === "current_month") {{
    const first = new Date(now.getFullYear(), now.getMonth(), 1);
    const last = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    setDateFilterRange(_isoDateLocal(first), _isoDateLocal(last));
    return;
  }}
  if (kind === "previous_month") {{
    const first = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const last = new Date(now.getFullYear(), now.getMonth(), 0);
    setDateFilterRange(_isoDateLocal(first), _isoDateLocal(last));
    return;
  }}
  if (kind === "last_30_days") {{
    const start = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 29);
    setDateFilterRange(_isoDateLocal(start), today);
    return;
  }}
  if (kind === "quarter_to_date") {{
    const qStartMonth = Math.floor(now.getMonth() / 3) * 3;
    const qStart = new Date(now.getFullYear(), qStartMonth, 1);
    setDateFilterRange(_isoDateLocal(qStart), today);
    return;
  }}
  setDateFilterRange(defaultFrom, defaultTo);
}}
function isSubtaskPerformanceType(t) {{
  const label = issueTypeLabel(t);
  if (!label) return false;
  if (label.includes("sub-task") || label.includes("subtask")) return true;
  return label.includes("bug") && (label.includes("sub-task") || label.includes("subtask"));
}}
function buildCapacityMap(from, to, assigneeNames, activeProfile) {{
  const out = new Map();
  const names = Array.from(assigneeNames || []);
  const fromIso = String(from || "");
  const toIso = String(to || "");
  const perAssignee = computePerAssigneeCapacity(fromIso, toIso, activeProfile || null);
  names.forEach((name) => out.set(String(name || ""), perAssignee));
  return out;
}}
function evalFormula(expr, scope) {{
  const text = String(expr || "").trim();
  if (!text) return 0;
  const src = text
    .replace(/\\baverage\\(/gi, "AVG(")
    .replace(/\\bsum\\(/gi, "SUM(")
    .replace(/\\bcount\\(/gi, "COUNT(")
    .replace(/\\bmin\\(/gi, "MIN(")
    .replace(/\\bmax\\(/gi, "MAX(");
  const api = {{
    SUM:(v)=>n(v),
    COUNT:(v)=>n(v)>0?1:0,
    MIN:(v)=>n(v),
    MAX:(v)=>n(v),
    AVG:(v)=>n(v)
  }};
  const replaced = src.replace(/[A-Za-z_][A-Za-z0-9_]*/g, (token) => {{
    if (token in api) return token;
    return String(n(scope[token.toLowerCase()]));
  }});
  if (!/^[0-9+\\-*/().,\\sA-Z]+$/.test(replaced)) return 0;
  try {{
    const fn = Function("SUM","COUNT","MIN","MAX","AVG", `return (${{replaced}});`);
    return n(fn(api.SUM, api.COUNT, api.MIN, api.MAX, api.AVG));
  }} catch (_err) {{
    return 0;
  }}
}}
function compute() {{
  const from = document.getElementById("from").value || defaultFrom;
  const to = document.getElementById("to").value || defaultTo;
  const activeCapacityProfile = resolveActiveCapacityProfile(from, to);
  updateCapacityProfileMeta(from, to, activeCapacityProfile);
  const s = String(document.getElementById("search").value || "").trim().toLowerCase();
  const pset = selectedProjects();
  const useP = pset.size > 0;
  const logs = worklogs.filter((r) => {{
    if (useP && !pset.has(String(r.project_key || "UNKNOWN"))) return false;
    if (!inRange(String(r.worklog_date || ""), from, to)) return false;
    if (!matchesPlannedRange(r, from, to)) return false;
    if (s && !String(r.issue_assignee || "").toLowerCase().includes(s)) return false;
    const issueType = String(r.item_issue_type || r.issue_type || "");
    return isSubtaskPerformanceType(issueType);
  }});
  const assignedItems = workItems.filter((r) => {{
    const assignee = String(r.assignee || "");
    if (!assignee) return false;
    const project = String(r.project_key || "UNKNOWN");
    if (useP && !pset.has(project)) return false;
    if (s && !assignee.toLowerCase().includes(s)) return false;
    const issueType = String(r.jira_issue_type || r.issue_type || r.work_item_type || "");
    if (!isSubtaskPerformanceType(issueType)) return false;
    const start = String(r.start_date || "");
    const due = String(r.due_date || "");
    if (!start && !due) return false;
    return inRange(start, from, to) || inRange(due, from, to);
  }});
  const leaves = leaveRows.filter((r) => inRange(String(r.period_day || ""), from, to));
  const byA = new Map();
  const epicKeysByAssignee = new Map();
  function addAssigneeEpic(name, epicKey) {{
    const assignee = String(name || "Unassigned");
    const epic = String(epicKey || "").toUpperCase();
    if (!epic) return;
    if (!epicKeysByAssignee.has(assignee)) epicKeysByAssignee.set(assignee, new Set());
    epicKeysByAssignee.get(assignee).add(epic);
  }}
  const issueAgg = new Map();
  for (const wi of assignedItems) ensure(byA, String(wi.assignee || "Unassigned"));
  for (const l of logs) {{
    const a = ensure(byA, String(l.issue_assignee || "Unassigned"));
    addAssigneeEpic(a.assignee, resolveEpicKeyFromWorklogRow(l));
    const hrs = n(l.hours_logged);
    a.total_hours += hrs;
    const issueKey = String(l.issue_id || "");
    const lastLogged = String(a.last_log_by_issue[issueKey] || "");
    if (!lastLogged || String(l.worklog_date || "") > lastLogged) a.last_log_by_issue[issueKey] = String(l.worklog_date || "");
    a.issue_logged_hours_by_issue[issueKey] = n(a.issue_logged_hours_by_issue[issueKey]) + hrs;
    const issue = String(l.issue_id || "");
    if (!issueAgg.has(issue)) issueAgg.set(issue, {{ total:0, estimate:n(l.original_estimate_hours), shares:new Map(), dayShares:new Map() }});
    const agg = issueAgg.get(issue);
    agg.total += hrs;
    agg.shares.set(a.assignee, n(agg.shares.get(a.assignee)) + hrs);
    if (!agg.dayShares.has(a.assignee)) agg.dayShares.set(a.assignee, new Map());
    const byDay = agg.dayShares.get(a.assignee);
    const dayKey = String(l.worklog_date || "");
    byDay.set(dayKey, n(byDay.get(dayKey)) + hrs);
    if (l.is_bug) {{
      a.bug_hours += hrs;
      addDayPenalty(a, dayKey, hrs * n(settings.points_per_bug_hour));
      const due = String(l.story_due_date || "");
      if (due && String(l.worklog_date || "") > due) {{
        a.bug_late_hours += hrs;
        addDayPenalty(a, dayKey, hrs * n(settings.points_per_bug_late_hour));
      }}
      else if (!due) a.missing_story_due_count += 1;
    }} else {{
      const due = String(l.item_due_date || "");
      if (due && String(l.worklog_date || "") > due) {{
        a.subtask_late_hours += hrs;
        a.subtask_late_by_issue[issueKey] = n(a.subtask_late_by_issue[issueKey]) + hrs;
        addDayPenalty(a, dayKey, hrs * n(settings.points_per_subtask_late_hour));
      }}
      else if (!due) a.missing_due_count += 1;
    }}
    if (String(l.fix_type || "").toLowerCase() === "rework") {{
      a.rework_hours += hrs;
    }}
  }}
  for (const agg of issueAgg.values()) {{
    if (n(agg.estimate) <= 0) {{
      for (const name of agg.shares.keys()) ensure(byA, name).missing_estimate_issue_count += 1;
      continue;
    }}
    const overrun = Math.max(0, n(agg.total) - n(agg.estimate));
    if (overrun <= 0 || n(agg.total) <= 0) continue;
    for (const [name, hrs] of agg.shares.entries()) {{
      const rec = ensure(byA, name);
      const assigneeOverrun = overrun * (n(hrs) / n(agg.total));
      rec.estimate_overrun_hours += assigneeOverrun;
      const dayMap = agg.dayShares.get(name) || new Map();
      if (n(hrs) > 0) {{
        for (const [day, dayHours] of dayMap.entries()) {{
          const dayShare = n(dayHours) / n(hrs);
          const dayOverrun = assigneeOverrun * dayShare;
          addDayPenalty(rec, day, dayOverrun * n(settings.points_per_estimate_overrun_hour));
        }}
      }}
    }}
  }}
  for (const r of leaves) {{
    const a = ensure(byA, String(r.assignee || "Unassigned"));
    const unplannedHours = n(r.unplanned_taken_hours);
    const plannedHours = n(r.planned_taken_hours);
    a.unplanned_leave_hours += unplannedHours;
    a.planned_leave_hours += plannedHours;
    if (unplannedHours > 0) {{
      a.unplanned_leave_count += 1;
      a.unplanned_leave_days += unplannedHours / leaveHoursPerDay;
    }}
    if (plannedHours > 0) {{
      a.planned_leave_count += 1;
      a.planned_leave_days += plannedHours / leaveHoursPerDay;
    }}
    addDayPenalty(a, String(r.period_day || ""), unplannedHours * n(settings.points_per_unplanned_leave_hour));
  }}
  const assigneeNames = new Set(Array.from(byA.keys()));
  const capacityByAssignee = buildCapacityMap(from, to, assigneeNames, activeCapacityProfile);
  for (const wi of assignedItems) {{
    const a = ensure(byA, String(wi.assignee || "Unassigned"));
    addAssigneeEpic(a.assignee, resolveParentEpicKey(wi));
    const issueKey = String(wi.issue_key || "");
    const startDate = String(wi.start_date || "");
    const dueDate = String(wi.due_date || "");
    const issueType = normalizeType(wi.issue_type || wi.work_item_type || wi.jira_issue_type);
    a.planned_hours_assigned += n(wi.original_estimate_hours);
    a.assigned_counts[issueType] = n(a.assigned_counts[issueType]) + 1;
    a.total_assigned_count += 1;
    if (dueDate) a.due_dated_assigned_count += 1;
    const lastLogDate = String(a.last_log_by_issue[issueKey] || "");
    a.assigned_hierarchy.push({{
      issue_key: issueKey,
      summary: String(wi.summary || ""),
      issue_type: issueType,
      parent_issue_key: String(wi.parent_issue_key || ""),
      parent_epic_key: resolveParentEpicKey(wi),
      due_date: dueDate,
      start_date: startDate,
      original_estimate_hours: n(wi.original_estimate_hours),
      actual_hours: n(a.issue_logged_hours_by_issue[issueKey]),
      negative_hours: n(a.subtask_late_by_issue[issueKey]),
      completion_date: lastLogDate,
      resolved_stable_since_date: String(wi.resolved_stable_since_date || ""),
      status: String(wi.status || "")
    }});
    let dueStatus = "Not completed";
    let missedDueDate = false;
    if (lastLogDate && dueDate) {{
      if (lastLogDate < dueDate) dueStatus = "Before due";
      else if (lastLogDate === dueDate) dueStatus = "On due";
      else {{
        dueStatus = "After due";
        missedDueDate = true;
      }}
    }} else if (lastLogDate) {{
      dueStatus = "No due date";
    }} else if (dueDate && to >= dueDate) {{
      missedDueDate = true;
    }}
    if (missedDueDate) {{
      a.missed_due_date_count += 1;
      addDayPenalty(a, dueDate || lastLogDate || to, n(settings.points_per_missed_due_date));
    }}
    a.due_compliance_items.push({{
      issue_key: issueKey,
      summary: String(wi.summary || ""),
      due_date: dueDate,
      completion_date: lastLogDate,
      resolved_stable_since_date: String(wi.resolved_stable_since_date || ""),
      status_bucket: dueStatus,
      is_missed_due_date: missedDueDate
    }});
    if (startDate && lastLogDate !== startDate) {{
      const loggedOnStart = logs.some((x) => String(x.issue_assignee || "") === a.assignee && String(x.issue_id || "") === issueKey && String(x.worklog_date || "") === startDate);
      if (!loggedOnStart) {{
        a.missed_start_count += 1;
        const otherWork = logs
          .filter((x) => String(x.issue_assignee || "") === a.assignee && String(x.worklog_date || "") === startDate && String(x.issue_id || "") !== issueKey)
          .map((x) => ({{issue_id:String(x.issue_id || ""), summary:String(x.item_summary || ""), hours:n(x.hours_logged)}}));
        const leaveOnDay = leaves
          .filter((x) => String(x.assignee || "") === a.assignee && String(x.period_day || "") === startDate)
          .reduce((acc, x) => acc + n(x.planned_taken_hours) + n(x.unplanned_taken_hours), 0);
        const idle = !otherWork.length && leaveOnDay <= 0;
        const entry = {{
          issue_key: issueKey,
          summary: String(wi.summary || ""),
          start_date: startDate,
          other_work: otherWork,
          leave_hours: leaveOnDay,
          idle: idle
        }};
        a.missed_start_items.push(entry);
        a.start_day_activity.push(entry);
      }}
    }}
  }}
  const items = Array.from(byA.values());
  for (const it of items) {{
    it.active_rmi_count = (epicKeysByAssignee.get(String(it.assignee || "")) || new Set()).size;
  }}
  const seriesDays = dateRangeDays(from, to);
  for (const it of items) {{
    it.penalties = {{ bug: it.bug_hours * n(settings.points_per_bug_hour), bug_late: it.bug_late_hours * n(settings.points_per_bug_late_hour), leave: it.unplanned_leave_hours * n(settings.points_per_unplanned_leave_hour), subtask_late: it.subtask_late_hours * n(settings.points_per_subtask_late_hour), estimate: it.estimate_overrun_hours * n(settings.points_per_estimate_overrun_hour), missed_due_date: it.missed_due_date_count * n(settings.points_per_missed_due_date) }};
    it.rework_ratio_pct = n(it.total_hours) > 0 ? (it.rework_hours / n(it.total_hours)) * 100 : 0;
    it.total_penalty = Object.values(it.penalties).reduce((a,b)=>a+n(b),0);
    it.raw_score = n(settings.base_score) - it.total_penalty;
    it.final_score = clamp(it.raw_score, n(settings.min_score), n(settings.max_score));
    it.missed_due_date_ratio = n(it.due_dated_assigned_count) > 0 ? (n(it.missed_due_date_count) / n(it.due_dated_assigned_count)) * 100 : 0;
    if (it.penalties.bug) it.feed.push({{label:"Bug hours", points:it.penalties.bug, hours:it.bug_hours}});
    if (it.penalties.bug_late) it.feed.push({{label:"Bug hours after story due", points:it.penalties.bug_late, hours:it.bug_late_hours}});
    if (it.penalties.subtask_late) it.feed.push({{label:"Subtask due overrun", points:it.penalties.subtask_late, hours:it.subtask_late_hours}});
    if (it.penalties.estimate) it.feed.push({{label:"Estimate overrun", points:it.penalties.estimate, hours:it.estimate_overrun_hours}});
    if (it.penalties.leave) it.feed.push({{label:"Unplanned leave", points:it.penalties.leave, hours:it.unplanned_leave_hours}});
    if (it.penalties.missed_due_date) it.feed.push({{label:"Missed due dates", points:it.penalties.missed_due_date, hours:it.missed_due_date_count}});
    it.feed.sort((a,b)=>n(b.points)-n(a.points));
    it.employee_capacity_hours = Math.max(0, n(capacityByAssignee.get(it.assignee)) - n(it.planned_leave_hours) - n(it.unplanned_leave_hours));
    it.missed_start_ratio = n(it.total_assigned_count) > 0 ? (n(it.missed_start_count) / n(it.total_assigned_count)) * 100 : 0;
    it.entity_values = {{
      capacity: n(it.employee_capacity_hours),
      planned_hours: n(it.planned_hours_assigned),
      actual_hours: n(it.total_hours),
      planned_leaves: n(it.planned_leave_hours),
      unplanned_leaves: n(it.unplanned_leave_hours),
      planned_dates: n(it.total_assigned_count),
      status: n(it.final_score),
      activity: n(it.total_hours)
    }};
    const scope = {{}};
    for (const e of entitiesCatalog) {{
      const key = String(e.entity_key || "").toLowerCase();
      if (!key) continue;
      scope[key] = n(it.entity_values[key]);
    }}
    for (const m of managedFields) {{
      const key = String(m.field_key || "").toLowerCase();
      if (!key) continue;
      const formula = String(m.formula_expression || "");
      const value = evalFormula(formula, scope);
      it.managed_values[key] = value;
      scope[key] = value;
    }}
    it.managed_scope = {{...scope}};
    const days = seriesDays.length ? seriesDays : Object.keys(it.daily_penalty_by_day || {{}}).sort();
    let cumulative = 0;
    it.daily_series = days.map((d) => {{
      cumulative += n(it.daily_penalty_by_day[d]);
      return {{
        day: d,
        penalty: n(it.daily_penalty_by_day[d]),
        score: clamp(n(settings.base_score) - cumulative, n(settings.min_score), n(settings.max_score))
      }};
    }});
  }}
  items.sort((a,b)=>n(b.final_score)-n(a.final_score) || a.assignee.localeCompare(b.assignee));
  return items;
}}
function renderSeriesSvg(series) {{
  const data = Array.isArray(series) ? series : [];
  if (!data.length) return '<div class="empty">No daily performance series in selected range.</div>';
  const w = 920;
  const h = 220;
  const padL = 42, padR = 12, padT = 12, padB = 28;
  const minS = n(settings.min_score), maxS = n(settings.max_score);
  const plotW = Math.max(1, w - padL - padR), plotH = Math.max(1, h - padT - padB);
  function x(i) {{ return padL + ((data.length <= 1 ? 0 : i / (data.length - 1)) * plotW); }}
  function y(score) {{
    const clamped = clamp(n(score), minS, maxS);
    const ratio = (maxS - minS) > 0 ? ((clamped - minS) / (maxS - minS)) : 0;
    return padT + (1 - ratio) * plotH;
  }}
  const pts = data.map((d, i) => `${{x(i).toFixed(2)}},${{y(d.score).toFixed(2)}}`).join(" ");
  const ticks = [minS, minS + (maxS - minS) * 0.5, maxS];
  const yTicks = ticks.map((v) => {{
    const yy = y(v).toFixed(2);
    return `<line x1="${{padL}}" y1="${{yy}}" x2="${{w - padR}}" y2="${{yy}}" stroke="#27406a" stroke-width="1"></line><text x="${{padL - 6}}" y="${{Number(yy) + 4}}" fill="#9db1d8" font-size="10" text-anchor="end">${{n(v).toFixed(0)}}</text>`;
  }}).join("");
  const first = data[0], last = data[data.length - 1];
  const xLabels = `<text x="${{padL}}" y="${{h - 8}}" fill="#9db1d8" font-size="10">${{e(first.day)}}</text><text x="${{w - padR}}" y="${{h - 8}}" fill="#9db1d8" font-size="10" text-anchor="end">${{e(last.day)}}</text>`;
  return `<svg class="ts-svg" viewBox="0 0 ${{w}} ${{h}}" preserveAspectRatio="none"><rect x="0" y="0" width="${{w}}" height="${{h}}" fill="#0d172b"></rect>${{yTicks}}<polyline fill="none" stroke="#60a5fa" stroke-width="2.5" points="${{pts}}"></polyline>${{xLabels}}</svg>`;
}}
function renderTeamChartSvg(rows, selectedTeamName) {{
  const data = Array.isArray(rows) ? rows.slice(0, 12) : [];
  if (!data.length) return '<div class="empty">No team data to chart.</div>';
  const w = 920;
  const barH = 18;
  const gap = 10;
  const h = 18 + (data.length * (barH + gap)) + 16;
  const padL = 180, padR = 24, padT = 14, padB = 12;
  const minS = n(settings.min_score), maxS = n(settings.max_score);
  const plotW = Math.max(1, w - padL - padR);
  function x(score) {{
    const clamped = clamp(n(score), minS, maxS);
    const ratio = (maxS - minS) > 0 ? ((clamped - minS) / (maxS - minS)) : 0;
    return padL + (ratio * plotW);
  }}
  const grid = [minS, minS + (maxS - minS) * 0.5, maxS].map((v) => {{
    const xx = x(v).toFixed(2);
    return `<line x1="${{xx}}" y1="${{padT - 2}}" x2="${{xx}}" y2="${{h - padB}}" stroke="#27406a" stroke-width="1"></line><text x="${{xx}}" y="${{padT - 4}}" fill="#9db1d8" font-size="10" text-anchor="middle">${{n(v).toFixed(0)}}</text>`;
  }}).join("");
  const bars = data.map((tr, i) => {{
    const y = padT + i * (barH + gap);
    const xx = x(tr.avg_score);
    const width = Math.max(2, xx - padL);
    const selected = String(tr.team_name || "") === String(selectedTeamName || "");
    const fill = selected ? "#10b981" : "#22c55e";
    const stroke = selected ? "#93c5fd" : "none";
    const strokeW = selected ? "1.2" : "0";
    return `<g><text x="${{padL - 8}}" y="${{y + 12}}" fill="#dce8ff" font-size="11" text-anchor="end">${{e(tr.team_name)}}</text><rect x="${{padL}}" y="${{y}}" width="${{plotW}}" height="${{barH}}" rx="6" ry="6" fill="#122746"></rect><rect x="${{padL}}" y="${{y}}" width="${{width.toFixed(2)}}" height="${{barH}}" rx="6" ry="6" fill="${{fill}}" stroke="${{stroke}}" stroke-width="${{strokeW}}"></rect><rect class="team-bar-hit" data-team-name="${{e(tr.team_name)}}" x="${{padL}}" y="${{y}}" width="${{plotW}}" height="${{barH}}" rx="6" ry="6" fill="transparent" style="cursor:pointer;"></rect><text x="${{(xx + 6).toFixed(2)}}" y="${{y + 12}}" fill="#dce8ff" font-size="11">${{n(tr.avg_score).toFixed(1)}}</text></g>`;
  }}).join("");
  return `<svg class="team-chart-svg" viewBox="0 0 ${{w}} ${{h}}" preserveAspectRatio="none"><rect x="0" y="0" width="${{w}}" height="${{h}}" fill="#0d172b"></rect>${{grid}}${{bars}}</svg>`;
}}
function renderSimpleList(rows, formatter) {{
  const data = Array.isArray(rows) ? rows : [];
  if (!data.length) return '<div class="empty">No items.</div>';
  return data.map(formatter).join("");
}}
function renderAssignedMixChart(item) {{
  const vals = [n(item?.assigned_counts?.epic), n(item?.assigned_counts?.story), n(item?.assigned_counts?.subtask)];
  const labels = ["Epic", "Story", "Subtask"];
  const colors = ["#38bdf8", "#22c55e", "#f59e0b"];
  const maxv = Math.max(1, ...vals);
  const w = 520, h = 160, padL = 85, padR = 10, padT = 10, padB = 16;
  const rowH = 34;
  const bars = vals.map((v, i) => {{
    const y = padT + i * rowH;
    const ww = ((w - padL - padR) * (v / maxv));
    return `<g><text x="${{padL - 8}}" y="${{y + 15}}" fill="#dce8ff" font-size="11" text-anchor="end">${{labels[i]}}</text><rect x="${{padL}}" y="${{y}}" width="${{(w-padL-padR)}}" height="18" fill="#132846" rx="6"></rect><rect x="${{padL}}" y="${{y}}" width="${{ww.toFixed(2)}}" height="18" fill="${{colors[i]}}" rx="6"></rect><text x="${{(padL + ww + 6).toFixed(2)}}" y="${{y + 14}}" fill="#cfe3ff" font-size="11">${{v.toFixed(0)}}</text></g>`;
  }}).join("");
  return `<svg class="mini-svg" viewBox="0 0 ${{w}} ${{h}}" preserveAspectRatio="none"><rect x="0" y="0" width="${{w}}" height="${{h}}" fill="#0d172b"></rect>${{bars}}</svg>`;
}}
function renderDueComplianceChart(rows) {{
  const data = Array.isArray(rows) ? rows : [];
  const buckets = {{
    "Before due": 0,
    "On due": 0,
    "After due": 0,
    "Not completed": 0
  }};
  data.forEach((r) => {{
    const k = String(r?.status_bucket || "Not completed");
    if (!(k in buckets)) buckets["Not completed"] += 1;
    else buckets[k] += 1;
  }});
  const vals = [buckets["Before due"], buckets["On due"], buckets["After due"], buckets["Not completed"]];
  const labels = ["Before", "On", "After", "Open"];
  const colors = ["#22c55e", "#38bdf8", "#f43f5e", "#94a3b8"];
  const total = Math.max(1, vals.reduce((a, b) => a + b, 0));
  const w = 520, h = 160, cx = 115, cy = 80, r = 52, ir = 28;
  let angle = -Math.PI / 2;
  const segs = vals.map((v, i) => {{
    const frac = v / total;
    const next = angle + frac * Math.PI * 2;
    const large = frac > 0.5 ? 1 : 0;
    const x1 = cx + Math.cos(angle) * r, y1 = cy + Math.sin(angle) * r;
    const x2 = cx + Math.cos(next) * r, y2 = cy + Math.sin(next) * r;
    const x3 = cx + Math.cos(next) * ir, y3 = cy + Math.sin(next) * ir;
    const x4 = cx + Math.cos(angle) * ir, y4 = cy + Math.sin(angle) * ir;
    const d = `M ${{x1}} ${{y1}} A ${{r}} ${{r}} 0 ${{large}} 1 ${{x2}} ${{y2}} L ${{x3}} ${{y3}} A ${{ir}} ${{ir}} 0 ${{large}} 0 ${{x4}} ${{y4}} Z`;
    angle = next;
    return `<path d="${{d}}" fill="${{colors[i]}}"></path>`;
  }}).join("");
  const legend = vals.map((v, i) => `<text x="210" y="${{26 + i*24}}" fill="#dce8ff" font-size="11">${{labels[i]}}: ${{v}}</text><rect x="188" y="${{18 + i*24}}" width="14" height="14" rx="3" fill="${{colors[i]}}"></rect>`).join("");
  return `<svg class="mini-svg" viewBox="0 0 ${{w}} ${{h}}" preserveAspectRatio="none"><rect x="0" y="0" width="${{w}}" height="${{h}}" fill="#0d172b"></rect>${{segs}}<text x="${{cx}}" y="${{cy + 4}}" fill="#cfe3ff" font-size="12" text-anchor="middle">${{total}}</text>${{legend}}</svg>`;
}}
function renderHierarchyTable(rows, dueRows, missedRows) {{
  const data = Array.isArray(rows) ? rows : [];
  if (!data.length) return '<div class="empty" style="padding:8px;">No assigned items.</div>';
  const dueMap = new Map((Array.isArray(dueRows) ? dueRows : []).map((r) => [String(r.issue_key || ""), String(r.status_bucket || "")]));
  const missedSet = new Set((Array.isArray(missedRows) ? missedRows : []).map((r) => String(r.issue_key || "")));
  const byKey = new Map(data.map((r) => [String(r.issue_key || ""), r]));
  const childMap = new Map();
  for (const r of data) {{
    const p = String(r.parent_issue_key || "");
    if (!childMap.has(p)) childMap.set(p, []);
    childMap.get(p).push(r);
  }}
  function typeRank(t) {{
    const x = String(t || "");
    if (x === "epic") return 0;
    if (x === "story") return 1;
    return 2;
  }}
  function sorted(list) {{
    return (list || []).slice().sort((a,b) => typeRank(a.issue_type) - typeRank(b.issue_type) || String(a.issue_key).localeCompare(String(b.issue_key)));
  }}
  function renderNode(node) {{
    const key = String(node.issue_key || "");
    const kids = sorted(childMap.get(key) || []);
    const dueBucket = dueMap.get(key) || "-";
    const missed = missedSet.has(key) ? "Yes" : "No";
    const openAttr = node.issue_type === "epic" ? " open" : "";
    const summary = `<summary><div class="tree-left"><div class="issue-id">${{e(key)}}</div><div class="issue-title">${{e(node.summary || "")}}</div></div><div class="tree-metrics"><span class="metric-pill">${{e(String(node.issue_type || "").toUpperCase())}}</span><span class="metric-pill">Est: ${{n(node.original_estimate_hours).toFixed(1)}}h</span><span class="metric-pill">Start: ${{e(node.start_date || "-")}}</span><span class="metric-pill">Due: ${{e(node.due_date || "-")}}</span><span class="metric-pill">Done: ${{e(node.completion_date || "-")}}</span><span class="metric-pill">Due Status: ${{e(dueBucket)}}</span><span class="metric-pill">Missed Start: ${{missed}}</span></div></summary>`;
    if (!kids.length) return `<div class="tree-node"><details${{openAttr}}>${{summary}}</details></div>`;
    return `<div class="tree-node"><details${{openAttr}}>${{summary}}<div class="tree-children">${{kids.map(renderNode).join("")}}</div></details></div>`;
  }}
  const roots = sorted(data.filter((r) => {{
    const p = String(r.parent_issue_key || "");
    return !p || !byKey.has(p);
  }}));
  return `<div class="tree-wrap">${{roots.map(renderNode).join("")}}</div>`;
}}
function renderDueTable(rows) {{
  const data = Array.isArray(rows) ? rows : [];
  if (!data.length) return '<div class="empty" style="padding:8px;">No logged items.</div>';
  function dueBucketClass(bucket) {{
    const b = String(bucket || "");
    if (b === "Before due") return "due-bucket due-before";
    if (b === "On due") return "due-bucket due-on";
    if (b === "After due") return "due-bucket due-after";
    return "due-bucket";
  }}
  return `<table class="tbl"><thead><tr><th>Issue</th><th>Due</th><th>Completion</th><th>Stable Resolved</th><th>Bucket</th></tr></thead><tbody>${{data.map((r)=>`<tr class="${{r.is_missed_due_date ? "due-missed" : ""}}"><td><div class="issue-id">${{e(r.issue_key)}}</div><div class="issue-title">${{e(r.summary)}}</div></td><td>${{e(r.due_date || "-")}}</td><td>${{e(r.completion_date || "-")}}</td><td>${{e(r.resolved_stable_since_date || "-")}}</td><td><span class="${{dueBucketClass(r.status_bucket)}}">${{e(r.status_bucket)}}</span></td></tr>`).join("")}}</tbody></table>`;
}}
function renderMissedTable(rows) {{
  const data = Array.isArray(rows) ? rows : [];
  if (!data.length) return '<div class="empty" style="padding:8px;">No missed starts.</div>';
  return `<table class="tbl"><thead><tr><th>Issue</th><th>Planned Start</th><th>Other Work</th><th>Leave (h)</th><th>State</th></tr></thead><tbody>${{data.map((r)=>{{ const other=(r.other_work||[]).map((x)=>`${{e(x.issue_id)}} (${{n(x.hours).toFixed(1)}}h)`).join(", "); const state=r.idle ? "Idle" : (n(r.leave_hours)>0 ? "On Leave" : "Working on other items"); return `<tr><td><div class="issue-id">${{e(r.issue_key)}}</div><div class="issue-title">${{e(r.summary)}}</div></td><td>${{e(r.start_date || "-")}}</td><td>${{other || "-"}}</td><td>${{n(r.leave_hours).toFixed(1)}}</td><td>${{state}}</td></tr>`; }}).join("")}}</tbody></table>`;
}}
function renderExecutionHierarchyTable(item) {{
  const nodes = Array.isArray(item?.assigned_hierarchy) ? item.assigned_hierarchy : [];
  const subtasks = nodes.filter((n0) => String(n0?.issue_type || "").toLowerCase() === "subtask");
  if (!subtasks.length) return '<div class="empty" style="padding:8px;">No subtasks in current scope.</div>';

  function wiByKey(key) {{
    return workItemsByKey.get(String(key || "").toUpperCase()) || null;
  }}
  function itemStart(x) {{ return String(x?.start_date || x?.item_start_date || ""); }}
  function itemDue(x) {{ return String(x?.due_date || x?.end_date || x?.item_due_date || ""); }}
  function itemSummary(x, fallback) {{ return String(x?.summary || fallback || ""); }}
  function maxDate(values) {{
    const vals = (Array.isArray(values) ? values : []).map((v) => String(v || "")).filter(Boolean).sort();
    return vals.length ? vals[vals.length - 1] : "";
  }}
  function minDate(values) {{
    const vals = (Array.isArray(values) ? values : []).map((v) => String(v || "")).filter(Boolean).sort();
    return vals.length ? vals[0] : "";
  }}

  const epicMap = new Map();
  for (const st of subtasks) {{
    const stKey = String(st.issue_key || "").toUpperCase();
    const storyKey = String(st.parent_issue_key || "").toUpperCase();
    const storyWi = wiByKey(storyKey);
    const epicKey = String(st.parent_epic_key || storyWi?.parent_issue_key || "").toUpperCase();
    const epicWi = wiByKey(epicKey);
    const storyActual = n(st.actual_hours);
    const storyCompletion = String(st.completion_date || "");

    if (!epicMap.has(epicKey || "NO_EPIC")) {{
      epicMap.set(epicKey || "NO_EPIC", {{
        epic_key: epicKey || "",
        epic_summary: itemSummary(epicWi, epicKey || "No Epic"),
        planned_start: itemStart(epicWi),
        planned_due: itemDue(epicWi),
        planned_hours: n(epicWi?.original_estimate_hours),
        actual_hours: 0,
        completion_dates: [],
        resolved_stable: String(epicWi?.resolved_stable_since_date || ""),
        stories: new Map(),
      }});
    }}
    const epicNode = epicMap.get(epicKey || "NO_EPIC");
    epicNode.actual_hours += storyActual;
    if (storyCompletion) epicNode.completion_dates.push(storyCompletion);

    if (!epicNode.stories.has(storyKey || "NO_STORY")) {{
      epicNode.stories.set(storyKey || "NO_STORY", {{
        story_key: storyKey || "",
        story_summary: itemSummary(storyWi, storyKey || "No Story"),
        planned_start: itemStart(storyWi),
        planned_due: itemDue(storyWi),
        planned_hours: n(storyWi?.original_estimate_hours),
        actual_hours: 0,
        completion_dates: [],
        resolved_stable: String(storyWi?.resolved_stable_since_date || ""),
        subtasks: new Map(),
      }});
    }}
    const storyNode = epicNode.stories.get(storyKey || "NO_STORY");
    storyNode.actual_hours += storyActual;
    if (storyCompletion) storyNode.completion_dates.push(storyCompletion);
    const stNodeKey = stKey || "NO_SUBTASK";
    if (!storyNode.subtasks.has(stNodeKey)) {{
      storyNode.subtasks.set(stNodeKey, {{
        issue_key: stKey,
        summary: String(st.summary || ""),
        planned_start: String(st.start_date || ""),
        planned_due: String(st.due_date || ""),
        planned_hours: n(st.original_estimate_hours),
        actual_hours: 0,
        negative_hours: 0,
        completion_date: storyCompletion,
        resolved_stable: String(st.resolved_stable_since_date || ""),
        status: String(st.status || ""),
      }});
    }}
    const stNode = storyNode.subtasks.get(stNodeKey);
    stNode.actual_hours += n(st.actual_hours);
    stNode.negative_hours += n(st.negative_hours);
    if (storyCompletion && (!stNode.completion_date || storyCompletion > stNode.completion_date)) stNode.completion_date = storyCompletion;
  }}

  const epicNodes = Array.from(epicMap.values()).sort((a, b) => a.epic_summary.localeCompare(b.epic_summary) || a.epic_key.localeCompare(b.epic_key));
  return `<div class="exec-accordion">${{epicNodes.map((epic, epicIndex) => {{
    const storyNodes = Array.from(epic.stories.values()).sort((a, b) => a.story_summary.localeCompare(b.story_summary) || a.story_key.localeCompare(b.story_key));
    const epicPlannedHours = epic.planned_hours > 0 ? epic.planned_hours : storyNodes.reduce((acc, s0) => acc + n(s0.planned_hours), 0);
    const epicStart = epic.planned_start || minDate(storyNodes.map((s0) => s0.planned_start));
    const epicDue = epic.planned_due || maxDate(storyNodes.map((s0) => s0.planned_due));
    const epicDone = maxDate(epic.completion_dates);
    const epicNegativeCount = storyNodes.reduce((acc, story) => acc + Array.from(story.subtasks.values()).filter((st0) => n(st0.negative_hours) > 0).length, 0);
    const storyHtml = storyNodes.map((story) => {{
      const subRows = Array.from(story.subtasks.values()).sort((a, b) => a.summary.localeCompare(b.summary) || a.issue_key.localeCompare(b.issue_key));
      const storyPlannedHours = story.planned_hours > 0 ? story.planned_hours : subRows.reduce((acc, st0) => acc + n(st0.planned_hours), 0);
      const storyStart = story.planned_start || minDate(subRows.map((st0) => st0.planned_start));
      const storyDue = story.planned_due || maxDate(subRows.map((st0) => st0.planned_due));
      return `<section class="exec-story-block"><div class="exec-story-head"><div><span class="metric-pill">STORY</span><div class="issue-id">${{e(story.story_key || "-")}}</div><div class="issue-title">${{e(story.story_summary || "-")}}</div></div><div class="exec-story-metrics"><span class="metric-pill">Start: ${{e(storyStart || "-")}}</span><span class="metric-pill">Due: ${{e(storyDue || "-")}}</span><span class="metric-pill">Planned: ${{n(storyPlannedHours).toFixed(2)}}h</span><span class="metric-pill">Actual: ${{n(story.actual_hours).toFixed(2)}}h</span><span class="metric-pill">Done: ${{e(maxDate(story.completion_dates) || "-")}}</span></div></div><table class="tbl exec-subtask-table"><thead><tr><th>Subtask</th><th>Planned Start</th><th>Planned Due</th><th>Planned Hours</th><th>Actual Hours</th><th>Last Activity</th><th>Stable Resolved</th><th>Status</th></tr></thead><tbody>${{subRows.map((st) => {{
        const negHrs = n(st.negative_hours);
        const status = negHrs > 0 ? `Penalty hit: late by ${{negHrs.toFixed(2)}}h` : (st.status || "");
        return `<tr class="${{negHrs > 0 ? "exec-negative-subtask" : ""}}"><td><div><span class="metric-pill">SUBTASK</span><div class="issue-id">${{e(st.issue_key || "-")}}</div><div class="issue-title">${{e(st.summary || "-")}}</div></div></td><td>${{e(st.planned_start || "-")}}</td><td>${{e(st.planned_due || "-")}}</td><td>${{n(st.planned_hours).toFixed(2)}}h</td><td>${{n(st.actual_hours).toFixed(2)}}h</td><td>${{e(st.completion_date || "-")}}</td><td>${{e(st.resolved_stable || "-")}}</td><td>${{e(status)}}</td></tr>`;
      }}).join("")}}</tbody></table></section>`;
    }}).join("");
    return `<details class="exec-epic"${{epicIndex === 0 ? " open" : ""}}><summary><div class="exec-epic-left"><span class="metric-pill">EPIC</span><div class="issue-id">${{e(epic.epic_key || "-")}}</div><div class="issue-title">${{e(epic.epic_summary || "-")}}</div></div><div class="exec-epic-metrics"><span class="metric-pill">Stories: ${{storyNodes.length}}</span><span class="metric-pill">Subtasks: ${{storyNodes.reduce((acc, s0) => acc + s0.subtasks.size, 0)}}</span><span class="metric-pill">Start: ${{e(epicStart || "-")}}</span><span class="metric-pill">Due: ${{e(epicDue || "-")}}</span><span class="metric-pill">Planned: ${{n(epicPlannedHours).toFixed(2)}}h</span><span class="metric-pill">Actual: ${{n(epic.actual_hours).toFixed(2)}}h</span><span class="metric-pill">Done: ${{e(epicDone || "-")}}</span>${{epicNegativeCount > 0 ? `<span class="metric-pill exec-neg-pill">Penalty Subtasks: ${{epicNegativeCount}}</span>` : ""}}</div></summary><div class="exec-epic-body">${{storyHtml}}</div></details>`;
  }}).join("")}}</div>`;
}}
function toTitleCaseKey(key) {{
  const txt = String(key || "").replace(/_/g, " ").trim();
  if (!txt) return "Metric";
  return txt.replace(/\\w\\S*/g, (w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
}}
function getManagedValue(item, candidates) {{
  const values = item && item.managed_values ? item.managed_values : {{}};
  const byKey = new Map(Object.entries(values).map(([k, v]) => [String(k || "").toLowerCase(), n(v)]));
  for (const raw of (Array.isArray(candidates) ? candidates : [])) {{
    const key = String(raw || "").toLowerCase();
    if (byKey.has(key)) return n(byKey.get(key));
  }}
  return NaN;
}}
function render(items) {{
  document.getElementById("kpi-avg").textContent = (items.length ? (items.reduce((a,b)=>a+n(b.final_score),0)/items.length) : 0).toFixed(1);
  document.getElementById("kpi-top").textContent = items[0]?.assignee || "-";
  document.getElementById("kpi-risk").textContent = String(items.filter(i => n(i.final_score) < 60).length);
  document.getElementById("kpi-pen").textContent = items.reduce((a,b)=>a+n(b.total_penalty),0).toFixed(1);
  document.getElementById("kpi-rework").textContent = items.reduce((a,b)=>a+n(b.rework_hours),0).toFixed(1) + "h";
  const totalAssignees = items.length;
  const atRiskCount = items.filter((i) => n(i.final_score) < 60).length;
  const highMissed = items.filter((i) => n(i.missed_start_ratio) >= 30).length;
  const overloaded = items.filter((i) => (n(i.planned_hours_assigned) - n(i.employee_capacity_hours)) > 0).length;
  document.getElementById("discover-insights").innerHTML = [
    `<span class="pill">Assignees: ${{totalAssignees}}</span>`,
    `<span class="pill">At-Risk: ${{atRiskCount}}</span>`,
    `<span class="pill">High Missed-Start (>=30%): ${{highMissed}}</span>`,
    `<span class="pill">Over Capacity: ${{overloaded}}</span>`
  ].join("");
  const topHigh = items.slice(0, 3);
  const topLow = items.slice(-3).reverse();
  document.getElementById("top3-high").innerHTML = topHigh.length
    ? topHigh.map((it, idx) => `<div class="top3-item high"><span class="nm">#${{idx + 1}} ${{e(it.assignee)}}</span><span class="sc">${{n(it.final_score).toFixed(1)}}</span></div>`).join("")
    : '<div class="empty">No data.</div>';
  document.getElementById("top3-low").innerHTML = topLow.length
    ? topLow.map((it, idx) => `<div class="top3-item low"><span class="nm">#${{idx + 1}} ${{e(it.assignee)}}</span><span class="sc">${{n(it.final_score).toFixed(1)}}</span></div>`).join("")
    : '<div class="empty">No data.</div>';
  const teamChartHost = document.getElementById("team-performance-chart");
  const teamDetailHost = document.getElementById("selected-team-performance");
  let teamRows = [];
  if (!teams.length) {{
    teamChartHost.innerHTML = '<div class="empty">No teams configured in settings.</div>';
    teamDetailHost.innerHTML = '<div class="empty">No teams configured in settings.</div>';
    selectedTeam = "";
  }} else {{
    const byName = new Map(items.map((it) => [String(it.assignee || "").toLowerCase(), it]));
    teamRows = teams.map((t) => {{
      const members = Array.isArray(t.assignees) ? t.assignees : [];
      const matched = members.map((m) => byName.get(String(m || "").toLowerCase())).filter(Boolean);
      const avgScore = matched.length ? (matched.reduce((a, b) => a + n(b.final_score), 0) / matched.length) : 0;
      const totalPenalty = matched.reduce((a, b) => a + n(b.total_penalty), 0);
      const atRisk = matched.filter((m) => n(m.final_score) < 60).length;
      return {{
        team_name: String(t.team_name || ""),
        team_leader: String(t.team_leader || "-"),
        members: members,
        total_members: members.length,
        active_members: matched.length,
        avg_score: avgScore,
        total_penalty: totalPenalty,
        total_rework_hours: matched.reduce((a, b) => a + n(b.rework_hours), 0),
        at_risk: atRisk,
        planned_leave_count: matched.reduce((a, b) => a + n(b.planned_leave_count), 0),
        planned_leave_hours: matched.reduce((a, b) => a + n(b.planned_leave_hours), 0),
        planned_leave_days: matched.reduce((a, b) => a + n(b.planned_leave_days), 0),
        unplanned_leave_count: matched.reduce((a, b) => a + n(b.unplanned_leave_count), 0),
        unplanned_leave_hours: matched.reduce((a, b) => a + n(b.unplanned_leave_hours), 0),
        unplanned_leave_days: matched.reduce((a, b) => a + n(b.unplanned_leave_days), 0)
      }};
    }}).sort((a, b) => n(b.avg_score) - n(a.avg_score) || a.team_name.localeCompare(b.team_name));
    if (selectedTeam && !teamRows.some((tr) => tr.team_name === selectedTeam)) selectedTeam = "";
    teamChartHost.innerHTML = renderTeamChartSvg(teamRows, selectedTeam);
    Array.from(teamChartHost.querySelectorAll(".team-bar-hit")).forEach((node) => {{
      node.addEventListener("click", () => {{
        const clicked = String(node.getAttribute("data-team-name") || "");
        selectedTeam = selectedTeam === clicked ? "" : clicked;
        render(compute());
      }});
    }});
  }}
  const selectedTeamRow = selectedTeam ? teamRows.find((tr) => tr.team_name === selectedTeam) : null;
  if (selectedTeamRow) {{
    teamDetailHost.innerHTML = `
      <div class="team-card sel" style="cursor:default;transform:none;">
        <div class="team-head">
          <div>
            <div class="team-name">${{e(selectedTeamRow.team_name || "-")}}</div>
            <div class="team-sub">Lead: ${{e(selectedTeamRow.team_leader)}} | Members: ${{selectedTeamRow.active_members}}/${{selectedTeamRow.total_members}}</div>
          </div>
          <div class="team-score">${{n(selectedTeamRow.avg_score).toFixed(1)}}</div>
        </div>
        <div class="team-metrics">
          <div>Total Penalty: ${{n(selectedTeamRow.total_penalty).toFixed(1)}}</div>
          <div>Total Rework: ${{n(selectedTeamRow.total_rework_hours).toFixed(1)}}h</div>
          <div>At Risk (&lt;60): ${{selectedTeamRow.at_risk}}</div>
          <div>Planned Leaves: ${{n(selectedTeamRow.planned_leave_count).toFixed(0)}} | ${{n(selectedTeamRow.planned_leave_hours).toFixed(2)}}h / ${{n(selectedTeamRow.planned_leave_days).toFixed(2)}}d</div>
          <div>Unplanned Leaves: ${{n(selectedTeamRow.unplanned_leave_count).toFixed(0)}} | ${{n(selectedTeamRow.unplanned_leave_hours).toFixed(2)}}h / ${{n(selectedTeamRow.unplanned_leave_days).toFixed(2)}}d</div>
        </div>
      </div>
    `;
  }} else if (teams.length) {{
    teamDetailHost.innerHTML = '<div class="empty">Select a team from chart.</div>';
  }}
  let viewItems = items.map((it) => ({{...it, capacity_gap_hours: n(it.employee_capacity_hours) - n(it.planned_hours_assigned)}}));
  const riskMode = String(document.getElementById("filter-risk")?.value || "all");
  const missedMode = String(document.getElementById("filter-missed")?.value || "all");
  const sortMode = String(document.getElementById("leader-sort")?.value || "rmis");
  if (riskMode === "risk") viewItems = viewItems.filter((it) => n(it.final_score) < 60);
  if (missedMode === "missed") viewItems = viewItems.filter((it) => n(it.missed_start_count) > 0);
  if (sortMode === "rmis") viewItems.sort((a,b)=>n(b.active_rmi_count)-n(a.active_rmi_count)||n(b.final_score)-n(a.final_score)||a.assignee.localeCompare(b.assignee));
  else if (sortMode === "missed") viewItems.sort((a,b)=>n(b.missed_start_ratio)-n(a.missed_start_ratio)||n(b.final_score)-n(a.final_score));
  else if (sortMode === "capacity_gap") viewItems.sort((a,b)=>n(a.capacity_gap_hours)-n(b.capacity_gap_hours)||n(b.final_score)-n(a.final_score));
  else viewItems.sort((a,b)=>n(b.final_score)-n(a.final_score)||a.assignee.localeCompare(b.assignee));
  if (selectedTeamRow) {{
    const allowed = new Set((selectedTeamRow.members || []).map((m) => String(m || "").toLowerCase()));
    viewItems = items.filter((it) => allowed.has(String(it.assignee || "").toLowerCase()));
  }}
  document.getElementById("leaderboard-title").textContent = selectedTeam ? `Leaderboard - ${{selectedTeam}}` : "Leaderboard";
  document.getElementById("leaderboard-filter").textContent = selectedTeam
    ? `Filtered by team "${{selectedTeam}}" (${{viewItems.length}} assignee${{viewItems.length === 1 ? "" : "s"}})`
    : "";
  const lb = document.getElementById("leaderboard");
  if (!viewItems.length) {{ lb.innerHTML = '<div class="empty" style="padding:10px;">No assignee activity for current filter.</div>'; document.getElementById("detail").innerHTML = '<div class="empty">No assignee activity for current filter.</div>'; return; }}
  lb.innerHTML = viewItems.map((it, i) => {{
    const capMoreManaged = getManagedValue(it, ["capacity_available_for_more_work", "capacityavailableformorework", "capacity_available_more_work"]);
    const capMore = Number.isFinite(capMoreManaged) ? capMoreManaged : n(it.capacity_gap_hours);
    return `<div class="row${{it.assignee===selectedName?' sel':''}}" data-name="${{e(it.assignee)}}"><div class="rank">#${{i+1}}</div><div><div>${{e(it.assignee)}}</div><div class="leader-metrics"><span class="metric-chip"><span class="material-symbols-outlined">deployed_code</span><span class="metric-value">${{n(it.active_rmi_count).toFixed(0)}}</span></span><span class="metric-chip"><span class="material-symbols-outlined">sliders</span><span class="metric-value${{capMore < 0 ? " warn" : ""}}">${{capMore.toFixed(1)}}h</span></span><span class="metric-chip"><span class="material-symbols-outlined">award_star</span><span class="metric-value">${{n(it.final_score).toFixed(1)}}</span></span></div><div class="sub">${{n(it.total_hours).toFixed(1)}}h logged | Missed: ${{n(it.missed_start_ratio).toFixed(1)}}% | Cap Gap: ${{n(it.capacity_gap_hours).toFixed(1)}}h</div></div><div class="score"><span class="material-symbols-outlined">award_star</span>${{n(it.final_score).toFixed(1)}}</div></div>`;
  }}).join("");
  Array.from(lb.querySelectorAll(".row")).forEach((el)=>el.addEventListener("click", ()=>{{ selectedName = String(el.getAttribute("data-name") || ""); availabilityBreakdownForAssignee = ""; render(compute()); }}));
  let item = viewItems.find(x => x.assignee === selectedName); if (!item) {{ item = viewItems[0]; selectedName = item.assignee; }}
  const feed = (item.feed || []).map((v) => `<div class="i"><strong>${{e(v.label)}}</strong><br>${{n(v.hours).toFixed(2)}}h | <span class="neg">-${{n(v.points).toFixed(2)}}</span></div>`).join("") || '<div class="i empty">No violations.</div>';
  const hierarchyTable = renderHierarchyTable(item.assigned_hierarchy, item.due_compliance_items, item.missed_start_items);
  const executionHierarchyTable = renderExecutionHierarchyTable(item);
  const dueTable = renderDueTable(item.due_compliance_items);
  const missedTable = renderMissedTable(item.missed_start_items);
  const assignedMixChart = renderAssignedMixChart(item);
  const dueMixChart = renderDueComplianceChart(item.due_compliance_items);
  const managedCatalog = new Map((managedFields || []).map((mf) => [String(mf?.field_key || "").toLowerCase(), mf]));
  const managedEntries = Object.entries(item.managed_values || {{}})
    .map(([k, v]) => {{
      const meta = managedCatalog.get(String(k || "").toLowerCase()) || {{}};
      const label = String(meta?.label || "").trim() || toTitleCaseKey(k);
      const meaning = String(meta?.description || "").trim() || "Derived leadership metric from configured managed fields.";
      return {{ key: k, label, meaning, value: n(v) }};
    }})
    .sort((a, b) => a.label.localeCompare(b.label));
  const plannedSubtaskHours = (item.assigned_hierarchy || [])
    .filter((row) => String(row?.issue_type || "").toLowerCase() === "subtask")
    .reduce((acc, row) => acc + n(row?.original_estimate_hours), 0);
  const plannedAssignedEntry = {{
    key: "planned_hours_assigned_static",
    label: "Planned Hours Assigned",
    meaning: "Total planned estimate hours for assigned subtasks within current filters.",
    value: n(plannedSubtaskHours),
  }};
  const availabilityIndex = managedEntries.findIndex((entry) => {{
    const key = String(entry?.key || "").trim().toLowerCase();
    const label = String(entry?.label || "").trim().toLowerCase();
    return key === "availability" || label === "availability";
  }});
  if (availabilityIndex >= 0) managedEntries.splice(availabilityIndex + 1, 0, plannedAssignedEntry);
  else managedEntries.push(plannedAssignedEntry);
  function normMetricToken(text) {{
    return String(text || "").toLowerCase().replace(/[^a-z0-9]+/g, "");
  }}
  function extractFormulaTokens(expr) {{
    const src = String(expr || "");
    const found = src.match(/[A-Za-z_][A-Za-z0-9_]*/g) || [];
    const skip = new Set(["SUM", "COUNT", "MIN", "MAX", "AVG"]);
    const out = [];
    const seen = new Set();
    for (const token of found) {{
      const upper = String(token || "").toUpperCase();
      const norm = String(token || "").toLowerCase();
      if (!norm || skip.has(upper) || seen.has(norm)) continue;
      seen.add(norm);
      out.push(norm);
    }}
    return out;
  }}
  const availabilityEntry = managedEntries.find((entry) => {{
    const k = normMetricToken(entry?.key);
    const l = normMetricToken(entry?.label);
    return k === "availability" || l === "availability";
  }});
  const availabilityFormulaField = availabilityEntry
    ? (managedCatalog.get(String(availabilityEntry.key || "").toLowerCase())
      || Array.from(managedCatalog.values()).find((meta) => normMetricToken(meta?.label) === "availability")
      || null)
    : null;
  const availabilityFormula = String(availabilityFormulaField?.formula_expression || "").trim();
  const availabilityFormulaTokens = extractFormulaTokens(availabilityFormula);
  const availabilityScope = item?.managed_scope && typeof item.managed_scope === "object" ? item.managed_scope : {{}};
  const availabilityIngredients = availabilityFormulaTokens.map((token) => {{
    const key = String(token || "").toLowerCase();
    let source = "scope";
    let value = availabilityScope[key];
    if (!Number.isFinite(Number(value))) {{
      value = item?.entity_values?.[key];
      source = "entity";
    }}
    if (!Number.isFinite(Number(value))) {{
      value = item?.managed_values?.[key];
      source = "managed";
    }}
    const missing = !Number.isFinite(Number(value));
    return {{
      key,
      value: missing ? 0 : n(value),
      missing,
      source: missing ? "default" : source,
    }};
  }});
  const availabilityRawValue = n(availabilityEntry?.value);
  const otherMetricMax = Math.max(1, ...managedEntries
    .filter((entry) => {{
      const k = normMetricToken(entry?.key);
      const l = normMetricToken(entry?.label);
      return !(k === "availability" || l === "availability");
    }})
    .map((entry) => n(entry?.value)));
  function isCapacityAvailableForMoreWork(entry) {{
    const k = normMetricToken(entry?.key);
    const l = normMetricToken(entry?.label);
    return k === "capacityavailableformorework" || l === "capacityavailableformorework";
  }}
  function isPlannedHoursAssignedMetric(entry) {{
    const k = normMetricToken(entry?.key);
    const l = normMetricToken(entry?.label);
    return k === "plannedhoursassignedstatic" || l === "plannedhoursassigned";
  }}
  function barMeta(entry) {{
    const key = normMetricToken(entry?.key);
    const label = normMetricToken(entry?.label);
    const isAvailability = key === "availability" || label === "availability";
    const isPlannedAssigned = isPlannedHoursAssignedMetric(entry);
    const isCapacityAvail = isCapacityAvailableForMoreWork(entry);
    const maxValue = isAvailability
      ? 178
      : (isPlannedAssigned
        ? 178
        : (isCapacityAvail
          ? Math.max(1, availabilityRawValue)
          : otherMetricMax));
    const rawValue = n(entry?.value);
    const clampedValue = Math.max(0, Math.min(rawValue, maxValue));
    const pct = maxValue > 0 ? (clampedValue / maxValue) * 100 : 0;
    const overflow = rawValue > maxValue;
    let fill = "linear-gradient(90deg,#38bdf8,#60a5fa)";
    if (overflow && isPlannedAssigned) {{
      fill = "linear-gradient(90deg,#38bdf8 0%, #60a5fa 42%, #f97316 72%, #ef4444 100%)";
    }}
    return {{
      isAvailability,
      isPlannedAssigned,
      maxValue,
      rawValue,
      pct: Math.max(0, Math.min(100, pct)),
      fill,
      note: isAvailability
        ? "Scale: 0 to 178h"
        : (isPlannedAssigned
          ? (overflow ? `Scale: 0 to 178h (Exceeded by ${{(rawValue - 178).toFixed(2)}}h)` : "Scale: 0 to 178h")
          : `Scale: 0 to ${{maxValue.toFixed(2)}}`)
    }};
  }}
  const activeFrom = document.getElementById("from").value || defaultFrom;
  const activeTo = document.getElementById("to").value || defaultTo;
  const activeProfileForBreakdown = resolveActiveCapacityProfile(activeFrom, activeTo);
  const businessDaysInRange = computeBusinessDays(activeFrom, activeTo, activeProfileForBreakdown);
  const managedHtml = managedEntries.length
    ? `<div class="exec-metrics">${{managedEntries.map((m) => {{ const b = barMeta(m); const toggleAttr = b.isAvailability ? ' data-action="toggle-availability-breakdown"' : ""; const cardClass = b.isAvailability ? "exec-metric actionable" : "exec-metric"; const isOpen = b.isAvailability && availabilityBreakdownForAssignee === String(item.assignee || ""); const ingredientRows = availabilityIngredients.length ? availabilityIngredients.map((part) => `<div class="availability-line"><span class="availability-name">${{e(part.key)}}${{part.missing ? " (default)" : ""}}</span><span class="availability-num">${{n(part.value).toFixed(2)}}${{b.isAvailability ? "h" : ""}}</span></div>`).join("") : '<div class="availability-note">No ingredients detected from formula.</div>'; const formulaBlock = b.isAvailability && isOpen ? `<div class="availability-breakdown"><div class="availability-line"><span class="availability-name">Formula</span><span class="availability-num">${{availabilityFormula ? e(availabilityFormula) : "Formula not configured"}}</span></div><div class="availability-line"><span class="availability-name">Business Days</span><span class="availability-num">${{n(businessDaysInRange).toFixed(0)}}d</span></div><div class="availability-note">Ingredients</div>${{ingredientRows}}<div class="availability-line"><span class="availability-name"><strong>Result</strong></span><span class="availability-num"><strong>${{b.rawValue.toFixed(2)}}h</strong></span></div></div>` : ""; return `<div class="${{cardClass}}"${{toggleAttr}}><div class="exec-m-head"><div class="exec-m-name">${{e(m.label)}}</div><div class="exec-m-value">${{b.rawValue.toFixed(2)}}${{(b.isAvailability || b.isPlannedAssigned) ? "h" : ""}}</div></div><div class="exec-m-meaning">${{e(m.meaning)}}</div><div class="exec-bar-track"><div class="exec-bar-fill" style="width:${{b.pct.toFixed(2)}}%;background:${{b.fill}};"></div></div><div class="exec-scale-note">${{e(b.note)}}${{b.isAvailability ? " | Click to view formula" : ""}}</div>${{formulaBlock}}</div>`; }}).join("")}}</div>`
    : '<div class="empty">No managed metrics configured.</div>';
  const activeProjectList = Array.from(selectedProjects()).sort();
  const activeProjectsText = activeProjectList.length ? activeProjectList.join(", ") : "All";
  const healthTag = n(item.final_score) < 60 ? "High Risk" : (n(item.missed_start_ratio) >= 30 ? "Start Discipline Risk" : "Stable");
  const healthColor = n(item.final_score) < 60 ? "#f43f5e" : (n(item.missed_start_ratio) >= 30 ? "#f59e0b" : "#22c55e");
  const summaryCapacityManaged = getManagedValue(item, ["capacity_available_for_more_work", "capacityavailableformorework", "capacity_available_more_work"]);
  const summaryCapacity = Number.isFinite(summaryCapacityManaged) ? summaryCapacityManaged : (n(item.employee_capacity_hours) - n(item.planned_hours_assigned));
  const summaryRmis = n(item.active_rmi_count);
  const summaryScore = n(item.final_score);
  const summaryMetricsHtml = `<div class="kpis" style="margin-top:8px;"><div class="kpi"><div class="k"><span class="material-symbols-outlined" style="font-size:15px;vertical-align:middle;margin-right:4px;">deployed_code</span>RMIs</div><div class="v">${{summaryRmis.toFixed(0)}}</div></div><div class="kpi"><div class="k"><span class="material-symbols-outlined" style="font-size:15px;vertical-align:middle;margin-right:4px;">sliders</span>Capacity</div><div class="v">${{summaryCapacity.toFixed(1)}}h</div></div><div class="kpi"><div class="k"><span class="material-symbols-outlined" style="font-size:15px;vertical-align:middle;margin-right:4px;">award_star</span>Score</div><div class="v">${{summaryScore.toFixed(1)}}</div></div></div>`;
  const summaryHtml = `<div class="card"><div style="display:flex;justify-content:space-between;align-items:end;"><div><div class="sub">Assignee</div><div style="font-size:1.1rem;font-weight:800;">${{e(item.assignee)}}</div></div><div class="big">${{summaryScore.toFixed(1)}}</div></div>${{summaryMetricsHtml}}<div class="sub">Raw ${{n(item.raw_score).toFixed(2)}} | Total Penalty -${{n(item.total_penalty).toFixed(2)}} | Base ${{n(settings.base_score).toFixed(0)}}</div><div class="discover"><span class="pill" style="border-color:${{healthColor}};">Health: ${{healthTag}}</span><span class="pill">Capacity Gap: ${{(n(item.employee_capacity_hours)-n(item.planned_hours_assigned)).toFixed(1)}}h</span><span class="pill">Missed Starts: ${{n(item.missed_start_ratio).toFixed(1)}}%</span><span class="pill">Missed Due Dates: ${{n(item.missed_due_date_ratio).toFixed(1)}}%</span></div></div>`;
  const managedSectionHtml = `<div class="mini" style="margin:8px 0;"><h3>Managed Field Metrics - ${{e(item.assignee)}}</h3><div class="sub">Employee-only metrics within filters | Date: ${{e(activeFrom)}} to ${{e(activeTo)}} | Projects: ${{e(activeProjectsText)}}</div>${{managedHtml}}</div>`;
  const tabsHtml = `<div class="tabs"><button class="tab-btn active" data-tab="overview">Overview</button><button class="tab-btn" data-tab="execution">Execution</button><button class="tab-btn" data-tab="planning">Planning</button></div><div class="tab-pane active" data-pane="overview"><div class="grid2"><section class="mini"><h3>Score Breakdown</h3><div class="l"><span>Bug Hours</span><span class="neg">-${{n(item.penalties.bug).toFixed(2)}}</span></div><div class="l"><span>Bug Late Hours</span><span class="neg">-${{n(item.penalties.bug_late).toFixed(2)}}</span></div><div class="l"><span>Unplanned Leaves</span><span class="neg">-${{n(item.penalties.leave).toFixed(2)}}</span></div><div class="l"><span>Subtask Late Hours</span><span class="neg">-${{n(item.penalties.subtask_late).toFixed(2)}}</span></div><div class="l"><span>Missed Due Dates</span><span class="neg">-${{n(item.penalties.missed_due_date).toFixed(2)}}</span></div><div class="l"><span>Estimate Overrun</span><span class="neg">-${{n(item.penalties.estimate).toFixed(2)}}</span></div></section><section class="mini"><h3>Planning Scorecards</h3><div class="l"><span>Employee Capacity</span><span>${{n(item.employee_capacity_hours).toFixed(2)}}h</span></div><div class="l"><span>Planned Assigned</span><span>${{n(item.planned_hours_assigned).toFixed(2)}}h</span></div><div class="l"><span>Assigned (E/S/ST)</span><span>${{n(item.assigned_counts.epic).toFixed(0)}}/${{n(item.assigned_counts.story).toFixed(0)}}/${{n(item.assigned_counts.subtask).toFixed(0)}}</span></div><div class="l actionable" data-action="open-missed-starts"><span>Missed Starts</span><span><button type="button" class="metric-link-btn">View subtasks</button> ${{n(item.missed_start_count).toFixed(0)}} / ${{n(item.total_assigned_count).toFixed(0)}} (${{n(item.missed_start_ratio).toFixed(1)}}%)</span></div><div class="l actionable" data-action="open-missed-due"><span>Missed Due Dates</span><span><button type="button" class="metric-link-btn">View subtasks</button> ${{n(item.missed_due_date_count).toFixed(0)}} / ${{n(item.due_dated_assigned_count).toFixed(0)}} (${{n(item.missed_due_date_ratio).toFixed(1)}}%)</span></div><div class="l"><span>Planned Leaves</span><span>${{n(item.planned_leave_count).toFixed(0)}} | ${{n(item.planned_leave_hours).toFixed(2)}}h / ${{n(item.planned_leave_days).toFixed(2)}}d</span></div><div class="l"><span>Unplanned Leaves</span><span>${{n(item.unplanned_leave_count).toFixed(0)}} | ${{n(item.unplanned_leave_hours).toFixed(2)}}h / ${{n(item.unplanned_leave_days).toFixed(2)}}d</span></div></section></div><div class="kpi-charts"><div class="mini-chart"><h3 style="margin:0 0 6px;font-size:.82rem;">Assigned Mix Chart</h3>${{assignedMixChart}}</div><div class="mini-chart"><h3 style="margin:0 0 6px;font-size:.82rem;">Due Compliance Chart</h3>${{dueMixChart}}</div></div><div class="ts-card"><h3 style="margin:0 0 6px;font-size:.82rem;">Performance Over Days</h3>${{renderSeriesSvg(item.daily_series)}}</div></div><div class="tab-pane" data-pane="execution"><div class="tbl-wrap"><h3 class="tbl-title">Execution Nested View (Epic -> Story -> Subtask) | Planned vs Actual</h3>${{executionHierarchyTable}}</div><div class="tbl-wrap" id="due-compliance-context"><h3 class="tbl-title">Due Compliance Table (Logged Items)</h3>${{dueTable}}</div><div class="tbl-wrap" id="missed-start-context"><h3 class="tbl-title">Missed Start Context Table</h3>${{missedTable}}</div><div class="feed">${{feed}}</div></div><div class="tab-pane" data-pane="planning"><div class="tbl-wrap"><h3 class="tbl-title">Interactive Hierarchy Breakdown (Epic -> Story -> Subtask)</h3>${{hierarchyTable}}</div></div>`;
  document.getElementById("detail").innerHTML = `${{summaryHtml}}${{managedSectionHtml}}${{tabsHtml}}`;
  const detailHost = document.getElementById("detail");
  function activateDetailTab(tab) {{
    Array.from(detailHost.querySelectorAll(".tab-btn")).forEach((b) => b.classList.toggle("active", String(b.getAttribute("data-tab") || "") === tab));
    Array.from(detailHost.querySelectorAll(".tab-pane")).forEach((p) => p.classList.toggle("active", String(p.getAttribute("data-pane") || "") === tab));
  }}
  Array.from(detailHost.querySelectorAll(".tab-btn")).forEach((btn) => {{
    btn.addEventListener("click", () => {{
      const tab = String(btn.getAttribute("data-tab") || "overview");
      activateDetailTab(tab);
    }});
  }});
  function focusContext(targetId) {{
    activateDetailTab("execution");
    const target = detailHost.querySelector(targetId);
    if (!target) return;
    target.classList.add("focus-pulse");
    target.scrollIntoView({{ behavior: "smooth", block: "start" }});
    setTimeout(() => target.classList.remove("focus-pulse"), 1300);
  }}
  const missedStartsTrigger = detailHost.querySelector('[data-action="open-missed-starts"]');
  if (missedStartsTrigger) {{
    missedStartsTrigger.addEventListener("click", () => {{
      focusContext("#missed-start-context");
    }});
  }}
  const missedDueTrigger = detailHost.querySelector('[data-action="open-missed-due"]');
  if (missedDueTrigger) {{
    missedDueTrigger.addEventListener("click", () => {{
      focusContext("#due-compliance-context");
    }});
  }}
  const availabilityTrigger = detailHost.querySelector('[data-action="toggle-availability-breakdown"]');
  if (availabilityTrigger) {{
    availabilityTrigger.addEventListener("click", () => {{
      const assigneeName = String(item.assignee || "");
      availabilityBreakdownForAssignee = availabilityBreakdownForAssignee === assigneeName ? "" : assigneeName;
      render(items);
    }});
  }}
}}
function renderAll() {{ availabilityBreakdownForAssignee = ""; render(compute()); }}
document.getElementById("projects").innerHTML = projects.map((p) => `<option value="${{e(p)}}" selected>${{e(p)}}</option>`).join("");
refreshCapacityProfileOptions();
document.getElementById("from").value = defaultFrom; document.getElementById("to").value = defaultTo;
document.getElementById("meta").textContent = `Generated: ${{payload.generated_at || "-"}} | Data window: ${{defaultFrom}} to ${{defaultTo}}`;
document.getElementById("toggle-performance-kpis").addEventListener("click", () => {{
  const wrap = document.getElementById("performance-kpis-wrap");
  const collapsed = wrap.classList.toggle("is-collapsed");
  document.querySelector("#toggle-performance-kpis .hint").textContent = collapsed ? "click to expand" : "click to collapse";
}});
document.getElementById("toggle-team-performance").addEventListener("click", () => {{
  const sec = document.getElementById("team-performance-section");
  const collapsed = sec.classList.toggle("is-collapsed");
  document.querySelector("#toggle-team-performance .hint").textContent = collapsed ? "click to expand" : "click to collapse";
}});
document.getElementById("apply").addEventListener("click", renderAll);
if (capacityProfileSelectEl) capacityProfileSelectEl.addEventListener("change", renderAll);
document.getElementById("reset").addEventListener("click", ()=>{{ document.getElementById("from").value=defaultFrom; document.getElementById("to").value=defaultTo; document.getElementById("search").value=\"\"; document.getElementById("leader-sort").value=\"rmis\"; document.getElementById("filter-risk").value=\"all\"; document.getElementById("filter-missed").value=\"all\"; if (capacityProfileSelectEl) capacityProfileSelectEl.value=\"auto\"; selectedTeam = \"\"; Array.from(document.getElementById("projects").options).forEach(o => o.selected=true); renderAll(); }});
document.getElementById("shortcut-current-month").addEventListener("click", ()=>{{ applyDateShortcut("current_month"); renderAll(); }});
document.getElementById("shortcut-previous-month").addEventListener("click", ()=>{{ applyDateShortcut("previous_month"); renderAll(); }});
document.getElementById("shortcut-last-30-days").addEventListener("click", ()=>{{ applyDateShortcut("last_30_days"); renderAll(); }});
document.getElementById("shortcut-quarter-to-date").addEventListener("click", ()=>{{ applyDateShortcut("quarter_to_date"); renderAll(); }});
document.getElementById("shortcut-reset").addEventListener("click", ()=>{{ applyDateShortcut("reset"); renderAll(); }});
renderAll();
</script>
<script src="shared-nav.js"></script>
</body>
</html>"""


def _resolve_runtime_paths(base_dir: Path) -> dict[str, Path]:
    worklog_name = os.getenv("JIRA_WORKLOG_XLSX_PATH", DEFAULT_WORKLOG_INPUT_XLSX).strip() or DEFAULT_WORKLOG_INPUT_XLSX
    work_items_name = os.getenv("JIRA_EXPORT_XLSX_PATH", DEFAULT_WORK_ITEMS_INPUT_XLSX).strip() or DEFAULT_WORK_ITEMS_INPUT_XLSX
    leave_name = os.getenv("JIRA_LEAVE_REPORT_XLSX_PATH", DEFAULT_LEAVE_REPORT_INPUT_XLSX).strip() or DEFAULT_LEAVE_REPORT_INPUT_XLSX
    html_name = os.getenv("JIRA_EMPLOYEE_PERFORMANCE_HTML_PATH", DEFAULT_HTML_OUTPUT).strip() or DEFAULT_HTML_OUTPUT
    db_name = os.getenv("JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH", DEFAULT_CAPACITY_DB).strip() or DEFAULT_CAPACITY_DB
    return {
        "worklog_path": _resolve_path(worklog_name, base_dir),
        "work_items_path": _resolve_path(work_items_name, base_dir),
        "leave_report_path": _resolve_path(leave_name, base_dir),
        "html_path": _resolve_path(html_name, base_dir),
        "db_path": _resolve_path(db_name, base_dir),
    }


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    paths = _resolve_runtime_paths(base_dir)
    _init_performance_settings_db(paths["db_path"])
    settings = _load_performance_settings(paths["db_path"])
    teams = _list_performance_teams(paths["db_path"])
    entities_catalog = load_report_entities(paths["db_path"])
    managed_fields = load_manage_fields(paths["db_path"], include_inactive=False)
    capacity_profiles = _list_capacity_profiles(paths["db_path"])
    work_items = _load_work_items(paths["work_items_path"])
    worklogs = _load_worklogs(paths["worklog_path"], work_items)
    leave_rows = _load_unplanned_leave_rows(paths["leave_report_path"])
    payload = _build_payload(
        worklogs,
        list(work_items.values()),
        leave_rows,
        settings,
        teams,
        entities_catalog=entities_catalog,
        managed_fields=managed_fields,
        capacity_profiles=capacity_profiles,
    )
    paths["html_path"].write_text(_build_html(payload), encoding="utf-8")
    print(f"Employee performance logs: {len(worklogs)}")
    print(f"Wrote HTML report: {paths['html_path']}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate employee performance report.")
    parser.parse_args()
    main()

