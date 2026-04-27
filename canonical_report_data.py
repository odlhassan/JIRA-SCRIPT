from __future__ import annotations

import json
import sqlite3
from collections import defaultdict
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from jira_client import BASE_URL


def _to_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _to_float(value: Any) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _normalize_date_text(value: Any) -> str:
    text = _to_text(value)
    if not text:
        return ""
    if len(text) >= 10:
        try:
            return date.fromisoformat(text[:10]).isoformat()
        except ValueError:
            pass
    for fmt in ("%d-%b-%Y", "%d-%B-%Y", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date().isoformat()
    except ValueError:
        return ""


def resolve_canonical_run_id(db_path: Path, run_id: str = "") -> str:
    requested = _to_text(run_id)
    if requested:
        return requested
    if not db_path.exists():
        return ""
    try:
        with sqlite3.connect(db_path) as conn:
            try:
                row = conn.execute(
                    "SELECT last_success_run_id FROM canonical_refresh_state WHERE id = 1"
                ).fetchone()
                state_run_id = _to_text(row[0] if row else "")
                if state_run_id:
                    return state_run_id
            except sqlite3.Error:
                pass

            # Fallback: if refresh state is stale/missing, use latest run_id present in canonical tables.
            for table in ("canonical_issues", "canonical_worklogs"):
                try:
                    fallback_row = conn.execute(
                        f"SELECT run_id FROM {table} WHERE run_id IS NOT NULL AND trim(run_id) <> '' ORDER BY rowid DESC LIMIT 1"
                    ).fetchone()
                except sqlite3.Error:
                    continue
                fallback_run_id = _to_text(fallback_row[0] if fallback_row else "")
                if fallback_run_id:
                    return fallback_run_id
    except sqlite3.Error:
        return ""
    return ""


def load_canonical_issues(db_path: Path, run_id: str = "") -> list[dict[str, Any]]:
    effective_run_id = resolve_canonical_run_id(db_path, run_id)
    if not effective_run_id or not db_path.exists():
        return []
    try:
        with sqlite3.connect(db_path) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                """
                SELECT run_id, issue_id, issue_key, project_key, issue_type, summary, status, assignee,
                       start_date, due_date, created_utc, updated_utc, resolved_stable_since_date,
                       original_estimate_hours, total_hours_logged, fix_type, parent_issue_key,
                       story_key, epic_key, raw_payload_json
                FROM canonical_issues
                WHERE run_id = ?
                ORDER BY project_key ASC, issue_key ASC
                """,
                (effective_run_id,),
            ).fetchall()
    except sqlite3.Error:
        return []
    return [dict(row) for row in rows]


def load_canonical_worklogs(db_path: Path, run_id: str = "") -> list[dict[str, Any]]:
    effective_run_id = resolve_canonical_run_id(db_path, run_id)
    if not effective_run_id or not db_path.exists():
        return []
    try:
        with sqlite3.connect(db_path) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                """
                SELECT run_id, worklog_id, issue_key, project_key, worklog_author, issue_assignee,
                       started_utc, started_date, updated_utc, hours_logged
                FROM canonical_worklogs
                WHERE run_id = ?
                ORDER BY started_date ASC, worklog_id ASC
                """,
                (effective_run_id,),
            ).fetchall()
    except sqlite3.Error:
        return []
    return [dict(row) for row in rows]


def load_canonical_actuals_by_issue(db_path: Path, run_id: str = "") -> dict[str, dict[str, Any]]:
    effective_run_id = resolve_canonical_run_id(db_path, run_id)
    if not effective_run_id or not db_path.exists():
        return {}
    try:
        with sqlite3.connect(db_path) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                """
                SELECT issue_key, first_worklog_date, last_worklog_date, actual_complete_date,
                       total_worklog_hours, worklog_count
                FROM canonical_issue_actuals
                WHERE run_id = ?
                """,
                (effective_run_id,),
            ).fetchall()
    except sqlite3.Error:
        return {}
    out: dict[str, dict[str, Any]] = {}
    for row in rows:
        issue_key = _to_text(row["issue_key"]).upper()
        if issue_key:
            out[issue_key] = dict(row)
    return out


def build_canonical_dashboard_source_rows(
    db_path: Path,
    run_id: str = "",
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], str]:
    effective_run_id = resolve_canonical_run_id(db_path, run_id)
    issues = load_canonical_issues(db_path, effective_run_id)
    worklogs = load_canonical_worklogs(db_path, effective_run_id)
    actuals = load_canonical_actuals_by_issue(db_path, effective_run_id)
    if not effective_run_id or not issues:
        return [], [], [], ""

    issue_by_key = {_to_text(item.get("issue_key")).upper(): item for item in issues if _to_text(item.get("issue_key"))}
    rows1: list[dict[str, Any]] = []
    for item in issues:
        issue_key = _to_text(item.get("issue_key")).upper()
        actual = actuals.get(issue_key, {})
        rows1.append(
            {
                "project_key": _to_text(item.get("project_key")).upper(),
                "issue_key": issue_key,
                "parent_issue_key": _to_text(item.get("parent_issue_key")).upper(),
                "jira_issue_type": _to_text(item.get("issue_type")),
                "summary": _to_text(item.get("summary")),
                "status": _to_text(item.get("status")),
                "start_date": _normalize_date_text(item.get("start_date")),
                "end_date": _normalize_date_text(item.get("due_date")),
                "actual_start_date": _normalize_date_text(actual.get("first_worklog_date")),
                "actual_end_date": _normalize_date_text(actual.get("actual_complete_date") or actual.get("last_worklog_date")),
                "original_estimate": round(_to_float(item.get("original_estimate_hours")), 2),
                "original_estimate_hours": round(_to_float(item.get("original_estimate_hours")), 2),
                "total_hours_logged": round(_to_float(item.get("total_hours_logged")), 2),
                "assignee": _to_text(item.get("assignee")),
                "jira_url": f"{BASE_URL.rstrip('/')}/browse/{issue_key}" if issue_key else "",
                "created": _to_text(item.get("created_utc")),
                "updated": _to_text(item.get("updated_utc")),
                "fix_type": _to_text(item.get("fix_type")),
            }
        )

    rows2: list[dict[str, Any]] = []
    for worklog in worklogs:
        issue_key = _to_text(worklog.get("issue_key")).upper()
        issue = issue_by_key.get(issue_key, {})
        epic_key = _to_text(issue.get("epic_key")).upper()
        if not epic_key:
            continue
        rows2.append(
            {
                "issue_id": issue_key,
                "parent_story_id": _to_text(issue.get("story_key")).upper() or _to_text(issue.get("parent_issue_key")).upper(),
                "parent_epic_id": epic_key,
                "hours_logged": round(_to_float(worklog.get("hours_logged")), 2),
                "Latest IPP Meeting": "No",
                "Jira IPP RMI Dates Altered": "No",
                "IPP Actual Date (Production Date)": "",
                "IPP Remarks": "",
                "IPP Actual Date Matches Jira End Date": "No",
            }
        )

    rows3: list[dict[str, Any]] = []
    for issue in issues:
        issue_key = _to_text(issue.get("issue_key")).upper()
        issue_type = _to_text(issue.get("issue_type")).lower()
        if "story" not in issue_type and "sub" not in issue_type:
            continue
        epic_key = _to_text(issue.get("epic_key")).upper()
        if not epic_key:
            continue
        rows3.append(
            {
                "issue_id": issue_key,
                "parent_story_id": _to_text(issue.get("story_key")).upper() or issue_key,
                "parent_epic_id": epic_key,
                "planned start date": _normalize_date_text(issue.get("start_date")),
                "Latest IPP Meeting": "No",
                "Jira IPP RMI Dates Altered": "No",
                "IPP Actual Date (Production Date)": "",
                "IPP Remarks": "",
                "IPP Actual Date Matches Jira End Date": "No",
            }
        )
    return rows1, rows2, rows3, effective_run_id


def load_nested_rows_from_canonical(db_path: Path, run_id: str = "") -> list[dict[str, Any]]:
    from generate_nested_view_html import load_nested_view_tree_for_api

    effective_run_id = resolve_canonical_run_id(db_path, run_id)
    if not effective_run_id:
        return []
    try:
        return list(load_nested_view_tree_for_api(db_path, run_id=effective_run_id))
    except Exception:
        return []


def build_rlt_leave_snapshot(
    db_path: Path,
    run_id: str = "",
    from_date: str = "1900-01-01",
    to_date: str = "2999-12-31",
) -> dict[str, Any]:
    from generate_rlt_leave_report import (
        _compute_aggregates,
        _day_hours_profile_from_env,
        _load_canonical_project_rows,
        _redistribute_continuous_leave_subtasks,
        _redistribute_continuous_leave_worklogs,
    )

    effective_run_id = resolve_canonical_run_id(db_path, run_id)
    empty = {
        "raw_subtasks": [],
        "distributed_subtasks": [],
        "worklogs_normalized": [],
        "assignee_summary": [],
        "daily": [],
        "weekly": [],
        "monthly": [],
        "defective": [],
        "clubbed": [],
    }
    if not effective_run_id:
        return empty
    try:
        subtasks, raw_worklogs = _load_canonical_project_rows(db_path, effective_run_id, "RLT", "")
    except Exception:
        return empty
    day_profile = _day_hours_profile_from_env()
    distributed = _redistribute_continuous_leave_subtasks(subtasks, day_profile)
    redistributed = _redistribute_continuous_leave_worklogs(subtasks, raw_worklogs, day_profile)
    aggregates = _compute_aggregates(subtasks, redistributed, from_date, to_date, day_profile)
    return {
        "raw_subtasks": [
            {
                "issue_key": item.issue_key,
                "issue_id": item.issue_id,
                "summary": item.summary,
                "status": item.status,
                "assignee": item.assignee,
                "parent_task_key": item.parent_task_key,
                "parent_task_assignee": item.parent_task_assignee,
                "created": item.created,
                "updated": item.updated,
                "start_date": item.start_date,
                "due_date": item.due_date,
                "original_estimate_hours": round(item.original_estimate_hours, 2),
                "timespent_hours": round(item.timespent_hours, 2),
                "leave_type_raw": item.leave_type_raw,
                "leave_classification": item.leave_classification,
                "classification_source": item.classification_source,
                "total_worklog_hours": round(item.total_worklog_hours, 2),
                "planned_date_for_bucket": item.planned_date_for_bucket,
                "clubbed_leave": item.clubbed_leave,
                "no_entry_flag": item.no_entry_flag,
                "verification_reference_date": item.verification_reference_date,
                "created_after_leave_date_flag": item.created_after_leave_date_flag,
                "created_after_leave_days": item.created_after_leave_days,
                "verification_note": item.verification_note,
            }
            for item in subtasks
        ],
        "distributed_subtasks": [
            {
                "issue_key": item.issue_key,
                "issue_id": item.issue_id,
                "summary": item.summary,
                "status": item.status,
                "assignee": item.assignee,
                "parent_task_key": item.parent_task_key,
                "parent_task_assignee": item.parent_task_assignee,
                "created": item.created,
                "updated": item.updated,
                "start_date": item.start_date,
                "due_date": item.due_date,
                "original_estimate_hours": round(item.original_estimate_hours, 2),
                "timespent_hours": round(item.timespent_hours, 2),
                "leave_type_raw": item.leave_type_raw,
                "leave_classification": item.leave_classification,
                "classification_source": item.classification_source,
                "total_worklog_hours": round(item.total_worklog_hours, 2),
                "planned_date_for_bucket": item.planned_date_for_bucket,
                "clubbed_leave": item.clubbed_leave,
                "no_entry_flag": item.no_entry_flag,
                "verification_reference_date": item.verification_reference_date,
                "created_after_leave_date_flag": item.created_after_leave_date_flag,
                "created_after_leave_days": item.created_after_leave_days,
                "verification_note": item.verification_note,
            }
            for item in distributed
        ],
        "worklogs_normalized": [
            {
                "issue_key": item.issue_key,
                "started_raw": item.started_raw,
                "started_date": item.started_date,
                "author": item.author,
                "hours_logged": round(item.hours_logged, 2),
            }
            for item in redistributed
        ],
        **aggregates,
    }


def build_assignee_hours_rows_from_canonical(db_path: Path, run_id: str = "") -> list[dict[str, Any]]:
    effective_run_id = resolve_canonical_run_id(db_path, run_id)
    if not effective_run_id:
        return []
    worklogs = load_canonical_worklogs(db_path, effective_run_id)
    out: list[dict[str, Any]] = []
    for row in worklogs:
        started_date = _normalize_date_text(row.get("started_date") or row.get("started_utc"))
        hours_logged = round(_to_float(row.get("hours_logged")), 2)
        if not started_date or hours_logged <= 0:
            continue
        issue_key = _to_text(row.get("issue_key")).upper()
        project_key = _to_text(row.get("project_key")).upper() or (issue_key.split("-", 1)[0] if "-" in issue_key else "UNKNOWN")
        out.append(
            {
                "issue_id": issue_key,
                "project_key": project_key,
                "issue_assignee": _to_text(row.get("issue_assignee")) or "Unassigned",
                "worklog_author": _to_text(row.get("worklog_author")) or "Unassigned",
                "worklog_date": started_date,
                "period_day": started_date,
                "period_week": date.fromisoformat(started_date).isocalendar(),
                "period_month": started_date[:7],
                "hours_logged": hours_logged,
            }
        )
    for item in out:
        iso_year, iso_week, _ = item["period_week"]
        item["period_week"] = f"{iso_year:04d}-W{iso_week:02d}"
    return out


def build_planned_work_items_from_canonical(db_path: Path, run_id: str = "") -> list[dict[str, Any]]:
    effective_run_id = resolve_canonical_run_id(db_path, run_id)
    if not effective_run_id:
        return []
    items = load_canonical_issues(db_path, effective_run_id)
    out: list[dict[str, Any]] = []
    for row in items:
        issue_key = _to_text(row.get("issue_key")).upper()
        issue_type = _to_text(row.get("issue_type")).strip().lower()
        if "epic" in issue_type:
            normalized_type = "epic"
        elif "story" in issue_type:
            normalized_type = "story"
        elif "sub" in issue_type:
            normalized_type = "subtask"
        else:
            continue
        estimate_hours = round(_to_float(row.get("original_estimate_hours")), 2)
        if estimate_hours <= 0:
            continue
        out.append(
            {
                "issue_key": issue_key,
                "project_key": _to_text(row.get("project_key")).upper(),
                "issue_type": normalized_type,
                "planned_start": _normalize_date_text(row.get("start_date")),
                "planned_end": _normalize_date_text(row.get("due_date")),
                "original_estimate_hours": estimate_hours,
                "summary": _to_text(row.get("summary")),
                "status": _to_text(row.get("status")),
                "assignee": _to_text(row.get("assignee")),
                "epic_key": _to_text(row.get("epic_key")).upper(),
                "story_key": _to_text(row.get("story_key")).upper(),
            }
        )
    return out


def build_project_planned_hours_from_canonical(db_path: Path, run_id: str = "", excluded_project_key: str = "RLT") -> float:
    excluded = _to_text(excluded_project_key).upper()
    total = 0.0
    for row in build_planned_work_items_from_canonical(db_path, run_id):
        if row["issue_type"] != "epic":
            continue
        if _to_text(row.get("project_key")).upper() == excluded:
            continue
        total += _to_float(row.get("original_estimate_hours"))
    return round(total, 2)


def build_rlt_leaves_planned_rows_from_canonical(db_path: Path, run_id: str = "") -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for row in build_planned_work_items_from_canonical(db_path, run_id):
        if row["issue_type"] != "epic":
            continue
        if _to_text(row.get("project_key")).upper() != "RLT":
            continue
        out.append(
            {
                "issue_key": _to_text(row.get("issue_key")).upper(),
                "jira_start_date": _to_text(row.get("planned_start")),
                "jira_end_date": _to_text(row.get("planned_end")),
                "original_estimate_hours": round(_to_float(row.get("original_estimate_hours")), 2),
            }
        )
    return out


def build_epic_logged_hours_by_key_from_canonical(db_path: Path, run_id: str = "") -> dict[str, float]:
    out: dict[str, float] = defaultdict(float)
    for row in load_canonical_issues(db_path, run_id):
        if "epic" not in _to_text(row.get("issue_type")).lower():
            continue
        issue_key = _to_text(row.get("issue_key")).upper()
        if issue_key:
            out[issue_key] += round(_to_float(row.get("total_hours_logged")), 2)
    return {key: round(value, 2) for key, value in out.items()}


def build_rnd_epics_from_canonical(db_path: Path, run_id: str = "") -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for row in load_canonical_issues(db_path, run_id):
        if "epic" not in _to_text(row.get("issue_type")).lower():
            continue
        issue_key = _to_text(row.get("issue_key")).upper()
        out.append(
            {
                "issue_key": issue_key,
                "project_key": _to_text(row.get("project_key")).upper() or (issue_key.split("-", 1)[0] if "-" in issue_key else ""),
                "summary": _to_text(row.get("summary")),
                "status": _to_text(row.get("status")),
                "start_date": _normalize_date_text(row.get("start_date")),
                "end_date": _normalize_date_text(row.get("due_date")),
                "original_estimate_hours": round(_to_float(row.get("original_estimate_hours")), 2),
            }
        )
    return out
