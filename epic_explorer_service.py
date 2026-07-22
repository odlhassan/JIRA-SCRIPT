from __future__ import annotations

import calendar
import json
import sqlite3
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path
from typing import Any

from canonical_report_data import build_rlt_leave_snapshot, resolve_canonical_run_id
from generate_employee_performance_report import _list_performance_teams
from monthly_epic_plan_progress_service import (
    HOURS_PER_DAY,
    build_workforce_month_payload,
    load_support_team,
)


def _to_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _to_float(value: Any) -> float:
    try:
        return float(value or 0.0)
    except (TypeError, ValueError):
        return 0.0


def _round_hours(value: Any) -> float:
    return round(_to_float(value) + 1e-9, 2)


def _round_pct(value: Any) -> float:
    return round(_to_float(value) + 1e-9, 1)


def _parse_iso_date(value: Any) -> date | None:
    text = _to_text(value)
    if not text:
        return None
    try:
        return date.fromisoformat(text[:10])
    except ValueError:
        return None


def _date_range_overlaps(start_text: Any, due_text: Any, range_start: date, range_end: date) -> bool:
    start_day = _parse_iso_date(start_text)
    due_day = _parse_iso_date(due_text)
    return bool(start_day and due_day and start_day <= range_end and due_day >= range_start)


def _normalize_issue_type(value: Any) -> str:
    text = _to_text(value)
    lowered = text.casefold().replace("_", " ").replace("-", " ")
    if "bug" in lowered and "subtask" in lowered:
        return "Bug Subtask"
    if "subtask" in lowered or "sub task" in lowered:
        return "Sub-task"
    if "story" in lowered:
        return "Story"
    if "epic" in lowered:
        return "Epic"
    return text


def _is_drilldown_parent_type(value: Any) -> bool:
    text = _normalize_issue_type(value)
    return text in {"Story", "Task", "Bug"}


def _is_resolved_status_text(value: Any) -> bool:
    text = _to_text(value).casefold()
    return text in {"resolved", "resolved!", "done", "closed", "complete", "completed"}


def _issue_url(jira_base_url: str, issue_key: str) -> str:
    base = _to_text(jira_base_url).rstrip("/")
    key = _to_text(issue_key).upper()
    return f"{base}/browse/{key}" if base and key else ""


def _hours_from_man_days(value: Any) -> float | None:
    text = _to_text(value)
    if not text:
        return None
    try:
        parsed = float(text)
    except (TypeError, ValueError):
        return None
    if parsed < 0:
        return None
    return _round_hours(parsed * HOURS_PER_DAY)


def _planner_tk_budget_hours(planner_row: dict[str, Any] | None) -> float | None:
    if not planner_row:
        return None
    plans = planner_row.get("plans") if isinstance(planner_row.get("plans"), dict) else {}
    epic_plan = plans.get("epic_plan") if isinstance(plans.get("epic_plan"), dict) else {}
    return _hours_from_man_days(
        epic_plan.get("tk_budgeted_man_days")
        if epic_plan.get("tk_budgeted_man_days") not in (None, "")
        else epic_plan.get("man_days")
    )


def _project_name_from_issue(issue: dict[str, Any]) -> str:
    raw = _to_text(issue.get("raw_payload_json"))
    if not raw:
        return ""
    try:
        payload = json.loads(raw)
    except Exception:
        return ""
    if not isinstance(payload, dict):
        return ""
    fields = payload.get("fields")
    if not isinstance(fields, dict):
        return ""
    project = fields.get("project")
    if not isinstance(project, dict):
        return ""
    return _to_text(project.get("name"))


def _derive_actual_completion(
    planned_due_date: Any,
    last_logged_date: Any,
    resolved_stable_since_date: Any = None,
) -> dict[str, str]:
    due_date = _parse_iso_date(planned_due_date)
    last_log = _parse_iso_date(last_logged_date)
    resolved_stable = _parse_iso_date(resolved_stable_since_date)
    candidates = [d for d in (last_log, resolved_stable) if d]
    actual_complete_date = ""
    actual_complete_source = "none"
    if candidates:
        actual_complete_date = max(candidates).isoformat()
        if last_log and resolved_stable:
            actual_complete_source = "max_last_logged_resolved_stable"
        elif last_log:
            actual_complete_source = "last_logged_date"
        else:
            actual_complete_source = "resolved_stable_since_date"

    if not due_date:
        completion_bucket = "no_due_date"
    elif not actual_complete_date:
        completion_bucket = "not_completed"
    elif actual_complete_date < due_date.isoformat():
        completion_bucket = "before_due"
    elif actual_complete_date == due_date.isoformat():
        completion_bucket = "on_due"
    else:
        completion_bucket = "after_due"

    return {
        "actual_complete_date": actual_complete_date,
        "actual_complete_source": actual_complete_source,
        "completion_bucket": completion_bucket,
    }


def _planned_total_hours(
    jira_original_estimate_hours: float,
    tk_budget_hours: float | None,
    story_estimate_hours: float,
    subtask_estimate_hours: float,
) -> float:
    return _planned_estimate_basis(
        jira_original_estimate_hours,
        tk_budget_hours,
        story_estimate_hours,
        subtask_estimate_hours,
    )["hours"]


def _planned_estimate_basis(
    jira_original_estimate_hours: float,
    tk_budget_hours: float | None,
    story_estimate_hours: float,
    subtask_estimate_hours: float,
) -> dict[str, Any]:
    candidates = [
        ("jira_original_estimate", "Jira Original Estimate", jira_original_estimate_hours),
        ("tk_budget", "TK Budget", tk_budget_hours if tk_budget_hours is not None else 0.0),
        ("story_estimates", "Story Estimates", story_estimate_hours),
        ("subtask_estimates", "Subtask Estimates", subtask_estimate_hours),
    ]
    for source, label, value in candidates:
        rounded = _round_hours(value)
        if rounded > 0:
            return {"source": source, "label": label, "hours": rounded}
    return {"source": "none", "label": "No estimate available", "hours": 0.0}


def _schedule_proration_days(planned_start: Any, planned_due: Any, as_of_day: date | None) -> dict[str, int | None]:
    start_day = _parse_iso_date(planned_start)
    due_day = _parse_iso_date(planned_due)
    if not start_day or not due_day or due_day < start_day or not as_of_day:
        return {"total_calendar_days": None, "elapsed_calendar_days": None}
    total_days = (due_day - start_day).days + 1
    elapsed_days = min(total_days, max(0, (as_of_day - start_day).days + 1))
    return {"total_calendar_days": total_days, "elapsed_calendar_days": elapsed_days}


def _planned_to_date_hours(planned_total_hours: float, planned_start: Any, planned_due: Any, as_of_day: date | None) -> float:
    total = _round_hours(planned_total_hours)
    if total <= 0:
        return 0.0
    start_day = _parse_iso_date(planned_start)
    due_day = _parse_iso_date(planned_due)
    if not start_day or not due_day or due_day < start_day or not as_of_day:
        return total
    if as_of_day < start_day:
        return 0.0
    if as_of_day >= due_day:
        return total
    total_days = (due_day - start_day).days + 1
    elapsed_days = (as_of_day - start_day).days + 1
    return _round_hours(total * (elapsed_days / total_days))


def _position_for_signed_delta(value: float | None, *, neutral: str = "on_track") -> str:
    if value is None:
        return "insufficient_data"
    rounded = _round_hours(value)
    if rounded < -0.01:
        return "behind"
    if rounded > 0.01:
        return "ahead"
    return neutral


def _estimation_accuracy(planned_total_hours: float, total_actual_hours: float) -> dict[str, Any]:
    planned = _round_hours(planned_total_hours)
    actual = _round_hours(total_actual_hours)
    if planned <= 0 or actual <= 0:
        return {"pct": None, "status": "no_actuals"}
    pct = _round_hours((planned / actual) * 100)
    if 85 <= pct <= 115:
        status = "ideal"
    elif pct < 70:
        status = "broken"
    else:
        status = "outside_ideal"
    return {"pct": pct, "status": status}


def _month_code(day: date) -> str:
    return f"{day.year:04d}-{day.month:02d}"


def _month_bounds(month_code: str) -> tuple[date, date]:
    year = int(month_code[:4])
    month = int(month_code[5:7])
    return date(year, month, 1), date(year, month, calendar.monthrange(year, month)[1])


def _month_end(month_code: str) -> date:
    return _month_bounds(month_code)[1]


def _iter_months(start_day: date, end_day: date) -> list[str]:
    months: list[str] = []
    y, m = start_day.year, start_day.month
    while (y, m) <= (end_day.year, end_day.month):
        months.append(f"{y:04d}-{m:02d}")
        if m == 12:
            y += 1
            m = 1
        else:
            m += 1
    return months


def _planned_effort_hours(*values: Any) -> float:
    for value in values:
        if value is None:
            continue
        hours = _round_hours(value)
        if hours > 0:
            return hours
    return 0.0


def _schedule_position(value: float) -> str:
    if value < -0.01:
        return "behind"
    if value > 0.01:
        return "ahead"
    return "on_track"


def _estimate_accuracy_status(value: float | None) -> str:
    if value is None:
        return "no_actuals"
    if value < 70:
        return "broken"
    if 85 <= value <= 115:
        return "ideal"
    return "outside_ideal"


def _load_canonical_rows(db_path: Path, run_id: str) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        issue_rows = [
            dict(row)
            for row in conn.execute(
                """
                SELECT run_id, issue_id, issue_key, project_key, issue_type, summary, status, assignee,
                       start_date, due_date, created_utc, updated_utc, resolved_stable_since_date,
                       original_estimate_hours, total_hours_logged, fix_type, parent_issue_key,
                       story_key, epic_key, raw_payload_json
                FROM canonical_issues
                WHERE run_id = ?
                ORDER BY project_key ASC, issue_key ASC
                """,
                (run_id,),
            ).fetchall()
        ]
        worklog_rows = [
            dict(row)
            for row in conn.execute(
                """
                SELECT run_id, worklog_id, issue_key, project_key, worklog_author, issue_assignee,
                       started_utc, started_date, updated_utc, hours_logged
                FROM canonical_worklogs
                WHERE run_id = ?
                ORDER BY started_date ASC, worklog_id ASC
                """,
                (run_id,),
            ).fetchall()
        ]
    return issue_rows, worklog_rows


def _team_membership(db_path: Path) -> dict[str, str]:
    out: dict[str, str] = {}
    try:
        teams = _list_performance_teams(db_path)
    except Exception:
        teams = []
    for team in teams or []:
        if not isinstance(team, dict):
            continue
        team_name = _to_text(team.get("team_name"))
        if not team_name:
            continue
        for member in team.get("assignees") or []:
            name = _to_text(member)
            if name:
                out[name.casefold()] = team_name
    return out


def _flatten_leaves(snapshot: dict[str, Any], range_start: date, range_end: date) -> list[dict[str, Any]]:
    leaves: list[dict[str, Any]] = []
    distributed = [row for row in snapshot.get("distributed_subtasks") or [] if isinstance(row, dict)]
    if distributed:
        for row in distributed:
            day = _parse_iso_date(row.get("planned_date_for_bucket") or row.get("start_date"))
            if day is None or not (range_start <= day <= range_end):
                continue
            hours = _round_hours(row.get("original_estimate_hours"))
            if hours <= 0:
                continue
            classification = _to_text(row.get("leave_classification")) or "Planned"
            lowered = classification.casefold()
            leave_class = "unplanned" if "unplanned" in lowered else "planned"
            leaves.append(
                {
                    "assignee": _to_text(row.get("assignee")) or "Unassigned",
                    "date": day.isoformat(),
                    "hours": hours,
                    "leave_class": leave_class,
                    "leave_type": _to_text(row.get("leave_type_raw")),
                    "summary": _to_text(row.get("summary")),
                    "issue_key": _to_text(row.get("issue_key")).upper(),
                }
            )
        return leaves

    for row in snapshot.get("daily") or []:
        if not isinstance(row, dict):
            continue
        day = _parse_iso_date(row.get("period_day"))
        if day is None or not (range_start <= day <= range_end):
            continue
        assignee = _to_text(row.get("assignee")) or "Unassigned"
        planned = _round_hours(_to_float(row.get("planned_taken_hours")) + _to_float(row.get("planned_not_taken_hours")))
        unplanned = _round_hours(row.get("unplanned_taken_hours"))
        if planned > 0:
            leaves.append({"assignee": assignee, "date": day.isoformat(), "hours": planned, "leave_class": "planned", "leave_type": "", "summary": "", "issue_key": ""})
        if unplanned > 0:
            leaves.append({"assignee": assignee, "date": day.isoformat(), "hours": unplanned, "leave_class": "unplanned", "leave_type": "", "summary": "", "issue_key": ""})
    return leaves


def _capacity_lookup(
    db_path: Path,
    run_id: str,
    months: list[str],
    jira_base_url: str,
) -> dict[str, dict[str, dict[str, float]]]:
    lookup: dict[str, dict[str, dict[str, float]]] = {}
    for month in months[:48]:
        month_start, month_end = _month_bounds(month)
        try:
            workforce = build_workforce_month_payload(
                db_path,
                month_start,
                month_end,
                run_id,
                jira_base=jira_base_url,
            )
        except Exception:
            continue
        month_lookup: dict[str, dict[str, float]] = {}
        default_available = _round_hours(workforce.get("availability_hours"))
        for row in workforce.get("assignees") or []:
            if not isinstance(row, dict):
                continue
            name = _to_text(row.get("name"))
            if not name:
                continue
            month_lookup[name.casefold()] = {
                "available_hours": _round_hours(row.get("per_person_availability_hours")),
                "capacity_hours": _round_hours(row.get("per_person_capacity_hours")),
            }
        month_lookup[""] = {"available_hours": default_available, "capacity_hours": _round_hours(workforce.get("capacity_hours"))}
        lookup[month] = month_lookup
    return lookup


EPIC_EXPLORER_CAPACITY_BASIS_VALUES = {
    "assignee_capacity_after_leaves",
    "standard_workdays",
}
DEFAULT_EPIC_EXPLORER_CAPACITY_BASIS = "assignee_capacity_after_leaves"


def normalize_capacity_basis(value: Any) -> str:
    text = _to_text(value).lower()
    if text in EPIC_EXPLORER_CAPACITY_BASIS_VALUES:
        return text
    return DEFAULT_EPIC_EXPLORER_CAPACITY_BASIS


def _load_capacity_profiles_for_distribution(db_path: Path) -> list[dict[str, Any]]:
    try:
        with sqlite3.connect(db_path) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                """
                SELECT from_date, to_date, standard_hours_per_day,
                       ramadan_start_date, ramadan_end_date, ramadan_hours_per_day,
                       holiday_dates_json
                FROM assignee_capacity_settings
                ORDER BY from_date ASC, to_date ASC
                """
            ).fetchall()
    except sqlite3.Error:
        return []

    profiles: list[dict[str, Any]] = []
    for row in rows:
        start_day = _parse_iso_date(row["from_date"])
        end_day = _parse_iso_date(row["to_date"])
        if not start_day or not end_day:
            continue
        holidays: set[date] = set()
        try:
            decoded = json.loads(_to_text(row["holiday_dates_json"]) or "[]")
            if isinstance(decoded, list):
                holidays = {day for day in (_parse_iso_date(item) for item in decoded) if day}
        except json.JSONDecodeError:
            holidays = set()
        profiles.append(
            {
                "from_date": start_day,
                "to_date": end_day,
                "standard_hours_per_day": _to_float(row["standard_hours_per_day"]) or HOURS_PER_DAY,
                "ramadan_start_date": _parse_iso_date(row["ramadan_start_date"]),
                "ramadan_end_date": _parse_iso_date(row["ramadan_end_date"]),
                "ramadan_hours_per_day": _to_float(row["ramadan_hours_per_day"]) or 6.5,
                "holiday_dates": holidays,
            }
        )
    return profiles


def _profile_for_day(day: date, profiles: list[dict[str, Any]]) -> dict[str, Any] | None:
    for profile in profiles:
        if profile["from_date"] <= day <= profile["to_date"]:
            return profile
    return None


def _base_capacity_weight_for_day(day: date, profiles: list[dict[str, Any]], basis: str) -> float:
    if day.weekday() >= 5:
        return 0.0
    if basis == "standard_workdays":
        return HOURS_PER_DAY
    profile = _profile_for_day(day, profiles)
    if not profile:
        return HOURS_PER_DAY
    if day in profile.get("holiday_dates", set()):
        return 0.0
    ramadan_start = profile.get("ramadan_start_date")
    ramadan_end = profile.get("ramadan_end_date")
    if ramadan_start and ramadan_end and ramadan_start <= day <= ramadan_end:
        return _to_float(profile.get("ramadan_hours_per_day")) or 6.5
    return _to_float(profile.get("standard_hours_per_day")) or HOURS_PER_DAY


def _story_distribution_weights(
    story: dict[str, Any],
    profiles: list[dict[str, Any]],
    leave_by_assignee_day: dict[str, dict[date, float]],
    capacity_basis: str,
) -> list[tuple[date, float]]:
    start_day = _parse_iso_date(story.get("start_date"))
    due_day = _parse_iso_date(story.get("due_date"))
    if not start_day and due_day:
        start_day = due_day
    if start_day and not due_day:
        due_day = start_day
    if not start_day or not due_day:
        return []
    if due_day < start_day:
        start_day, due_day = due_day, start_day

    assignee_key = _to_text(story.get("assignee")).casefold()
    rows: list[tuple[date, float]] = []
    cursor = start_day
    while cursor <= due_day:
        weight = _base_capacity_weight_for_day(cursor, profiles, capacity_basis)
        if capacity_basis == "assignee_capacity_after_leaves" and assignee_key:
            weight = max(0.0, weight - leave_by_assignee_day.get(assignee_key, {}).get(cursor, 0.0))
        rows.append((cursor, weight))
        cursor += timedelta(days=1)

    if any(weight > 0 for _day, weight in rows):
        return rows

    cursor = start_day
    fallback: list[tuple[date, float]] = []
    while cursor <= due_day:
        fallback.append((cursor, 1.0))
        cursor += timedelta(days=1)
    return fallback


def _distribute_story_estimates_by_month(
    stories: list[dict[str, Any]],
    profiles: list[dict[str, Any]],
    leave_by_assignee_day: dict[str, dict[date, float]],
    capacity_basis: str,
) -> tuple[dict[str, float], dict[str, dict[str, float]], list[dict[str, Any]]]:
    planned_by_month: dict[str, float] = defaultdict(float)
    planned_by_assignee_month: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    distribution_details: list[dict[str, Any]] = []
    for story in stories:
        estimate = _to_float(story.get("original_estimate_hours"))
        if estimate <= 0:
            continue
        weights = _story_distribution_weights(story, profiles, leave_by_assignee_day, capacity_basis)
        if not weights:
            distribution_details.append(
                {
                    "issue_key": _to_text(story.get("issue_key")).upper(),
                    "summary": _to_text(story.get("summary")) or _to_text(story.get("issue_key")).upper(),
                    "assignee": _to_text(story.get("assignee")) or "Unassigned",
                    "original_estimate_hours": _round_hours(estimate),
                    "start_date": _to_text(story.get("start_date")),
                    "due_date": _to_text(story.get("due_date")),
                    "allocation_status": "missing_dates",
                    "monthly_hours": [],
                }
            )
            continue
        total_weight = sum(weight for _day, weight in weights)
        if total_weight <= 0:
            continue
        story_months: dict[str, float] = defaultdict(float)
        assignee = _to_text(story.get("assignee")) or "Unassigned"
        for day, weight in weights:
            allocated = estimate * (weight / total_weight)
            month = _month_code(day)
            planned_by_month[month] += allocated
            planned_by_assignee_month[assignee][month] += allocated
            story_months[month] += allocated
        distribution_details.append(
            {
                "issue_key": _to_text(story.get("issue_key")).upper(),
                "summary": _to_text(story.get("summary")) or _to_text(story.get("issue_key")).upper(),
                "assignee": assignee,
                "original_estimate_hours": _round_hours(estimate),
                "start_date": _to_text(story.get("start_date")),
                "due_date": _to_text(story.get("due_date")),
                "allocation_status": "allocated",
                "monthly_hours": [
                    {"month": month, "hours": _round_hours(hours)}
                    for month, hours in sorted(story_months.items())
                ],
            }
        )
    return planned_by_month, planned_by_assignee_month, distribution_details


def build_epic_explorer_payload(
    db_path: Path,
    planner_rows: list[dict[str, Any]],
    canonical_run_id: str = "",
    *,
    from_date: str | None = None,
    to_date: str | None = None,
    selected_projects: set[str] | None = None,
    jira_base_url: str = "",
    capacity_basis: str = DEFAULT_EPIC_EXPLORER_CAPACITY_BASIS,
) -> dict[str, Any]:
    capacity_basis = normalize_capacity_basis(capacity_basis)
    run_id = resolve_canonical_run_id(db_path, canonical_run_id)
    if not run_id:
        raise ValueError("No successful canonical refresh found. Run the canonical refresh first.")

    range_start = _parse_iso_date(from_date)
    range_end = _parse_iso_date(to_date)
    if range_start and not range_end:
        range_end = range_start
    if range_end and not range_start:
        range_start = range_end
    if range_start and range_end and range_end < range_start:
        range_end = range_start
    date_filter_active = bool(range_start and range_end)
    selected_project_keys = {
        _to_text(project).upper()
        for project in (selected_projects or set())
        if _to_text(project)
    }

    issues, worklogs = _load_canonical_rows(db_path, run_id)
    planner_by_epic = {_to_text(row.get("epic_key")).upper(): row for row in planner_rows if _to_text(row.get("epic_key"))}
    project_name_by_key: dict[str, str] = {}
    for planner_row in planner_rows:
        project_key_text = _to_text(planner_row.get("project_key")).upper()
        project_name_text = _to_text(planner_row.get("project_name")) or project_key_text
        if project_key_text:
            project_name_by_key.setdefault(project_key_text, project_name_text)
    epic_rows: dict[str, dict[str, Any]] = {}
    story_rows_by_epic: dict[str, list[dict[str, Any]]] = defaultdict(list)
    story_to_epic: dict[str, str] = {}
    subtask_rows_by_epic: dict[str, list[dict[str, Any]]] = defaultdict(list)
    subtask_to_epic: dict[str, str] = {}
    subtask_to_story: dict[str, str] = {}
    issue_by_key: dict[str, dict[str, Any]] = {}

    for issue in issues:
        key = _to_text(issue.get("issue_key")).upper()
        if not key:
            continue
        issue["issue_key"] = key
        issue["issue_type"] = _normalize_issue_type(issue.get("issue_type"))
        issue_by_key[key] = issue
        issue_type = issue["issue_type"]
        if issue_type == "Epic":
            epic_rows[key] = issue

    for issue in issues:
        key = _to_text(issue.get("issue_key")).upper()
        issue_type = _normalize_issue_type(issue.get("issue_type"))
        if _is_drilldown_parent_type(issue_type):
            epic_key = _to_text(issue.get("epic_key")).upper() or _to_text(issue.get("parent_issue_key")).upper()
            if epic_key:
                story_to_epic[key] = epic_key
                story_rows_by_epic[epic_key].append(issue)

    for issue in issues:
        key = _to_text(issue.get("issue_key")).upper()
        issue_type = _normalize_issue_type(issue.get("issue_type"))
        if issue_type not in {"Sub-task", "Bug Subtask"}:
            continue
        story_key = _to_text(issue.get("story_key")).upper() or _to_text(issue.get("parent_issue_key")).upper()
        epic_key = _to_text(issue.get("epic_key")).upper() or story_to_epic.get(story_key, "")
        if epic_key:
            subtask_to_epic[key] = epic_key
            subtask_to_story[key] = story_key
            subtask_rows_by_epic[epic_key].append(issue)

    worklogs_by_issue: dict[str, list[dict[str, Any]]] = defaultdict(list)
    headcount_by_epic: dict[str, set[str]] = defaultdict(set)
    last_log_by_epic: dict[str, str] = {}
    actual_by_epic_month: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    actual_by_epic_author: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    actual_by_epic_story: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    actual_by_epic_team: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    support_by_epic_author: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    support_members = {name.casefold() for name in load_support_team(db_path)}
    team_by_author = _team_membership(db_path)

    for raw_wl in worklogs:
        issue_key = _to_text(raw_wl.get("issue_key")).upper()
        epic_key = subtask_to_epic.get(issue_key)
        if not epic_key:
            continue
        started_date = _to_text(raw_wl.get("started_date"))[:10]
        hours = _round_hours(raw_wl.get("hours_logged"))
        author = _to_text(raw_wl.get("worklog_author")) or _to_text(raw_wl.get("issue_assignee")) or "Unassigned"
        story_key = subtask_to_story.get(issue_key, "")
        worklog = {
            "worklog_id": _to_text(raw_wl.get("worklog_id")),
            "issue_key": issue_key,
            "date": started_date,
            "author": author,
            "hours": hours,
        }
        worklogs_by_issue[issue_key].append(worklog)
        actual_by_epic_author[epic_key][author] += hours
        if started_date:
            month = started_date[:7]
            actual_by_epic_month[epic_key][month] += hours
            if not last_log_by_epic.get(epic_key) or started_date > last_log_by_epic[epic_key]:
                last_log_by_epic[epic_key] = started_date
        if author:
            headcount_by_epic[epic_key].add(author.casefold())
        team_name = team_by_author.get(author.casefold(), "Unmapped")
        actual_by_epic_team[epic_key][team_name] += hours
        if author.casefold() in support_members:
            support_by_epic_author[epic_key][author] += hours

    actual_by_subtask: dict[str, float] = {}
    actual_by_epic: dict[str, float] = defaultdict(float)
    actual_by_epic_story: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    for epic_key, subtasks in subtask_rows_by_epic.items():
        for subtask in subtasks:
            subtask_key = _to_text(subtask.get("issue_key")).upper()
            if not subtask_key:
                continue
            worklog_total = _round_hours(sum(_to_float(log.get("hours")) for log in worklogs_by_issue.get(subtask_key, [])))
            issue_total = _round_hours(subtask.get("total_hours_logged"))
            actual_total = max(worklog_total, issue_total)
            actual_by_subtask[subtask_key] = actual_total
            actual_by_epic[epic_key] += actual_total
            story_key = _to_text(subtask.get("story_key")).upper() or _to_text(subtask.get("parent_issue_key")).upper()
            if story_key:
                actual_by_epic_story[epic_key][story_key] += actual_total
    actual_by_epic = defaultdict(float, {key: _round_hours(value) for key, value in actual_by_epic.items()})
    actual_by_epic_story = defaultdict(
        lambda: defaultdict(float),
        {
            epic_key: defaultdict(float, {story_key: _round_hours(hours) for story_key, hours in story_hours.items()})
            for epic_key, story_hours in actual_by_epic_story.items()
        },
    )

    all_date_values = [
        _parse_iso_date(issue.get(field))
        for issue in issues
        for field in ("start_date", "due_date")
    ]
    all_date_values += [_parse_iso_date(wl.get("started_date")) for wl in worklogs]
    valid_dates = [d for d in all_date_values if d]
    global_start = min(valid_dates) if valid_dates else date.today()
    global_end = max(valid_dates) if valid_dates else date.today()
    try:
        leave_rows = _flatten_leaves(build_rlt_leave_snapshot(db_path, run_id, global_start.isoformat(), global_end.isoformat()), global_start, global_end)
    except Exception:
        leave_rows = []
    leave_by_assignee_day: dict[str, dict[date, float]] = defaultdict(lambda: defaultdict(float))
    for leaf in leave_rows:
        leaf_day = _parse_iso_date(leaf.get("date"))
        if not leaf_day:
            continue
        leave_by_assignee_day[_to_text(leaf.get("assignee")).casefold()][leaf_day] += _to_float(leaf.get("hours"))

    capacity_by_month = _capacity_lookup(db_path, run_id, _iter_months(global_start, global_end), jira_base_url)
    capacity_profiles = _load_capacity_profiles_for_distribution(db_path)

    project_options: dict[str, str] = {}
    rows: list[dict[str, Any]] = []
    for epic_key, epic in epic_rows.items():
        project_key = _to_text(epic.get("project_key")).upper()
        canonical_project_name = _project_name_from_issue(epic)
        if canonical_project_name:
            project_name_by_key.setdefault(project_key, canonical_project_name)
        project_options[project_key] = project_name_by_key.get(project_key, canonical_project_name or project_key)
        if selected_project_keys and project_key not in selected_project_keys:
            continue
        if date_filter_active and not _date_range_overlaps(epic.get("start_date"), epic.get("due_date"), range_start, range_end):
            continue

        planner = planner_by_epic.get(epic_key)
        stories = sorted(story_rows_by_epic.get(epic_key, []), key=lambda item: _to_text(item.get("issue_key")))
        subtasks = sorted(subtask_rows_by_epic.get(epic_key, []), key=lambda item: (_to_text(item.get("story_key")), _to_text(item.get("issue_key"))))
        subtasks_by_story: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for subtask in subtasks:
            story_key = _to_text(subtask.get("story_key")).upper() or _to_text(subtask.get("parent_issue_key")).upper()
            subtasks_by_story[story_key].append(subtask)
        assigned_subtask_assignees = {
            _to_text(subtask.get("assignee")).casefold()
            for subtask in subtasks
            if _to_text(subtask.get("assignee"))
        }

        story_estimate_hours = _round_hours(sum(_to_float(story.get("original_estimate_hours")) for story in stories))
        subtask_estimate_hours = _round_hours(sum(_to_float(subtask.get("original_estimate_hours")) for subtask in subtasks))
        total_actual_hours = _round_hours(actual_by_epic.get(epic_key))
        if _is_resolved_status_text(epic.get("status")):
            completion = _derive_actual_completion(
                epic.get("due_date"),
                last_log_by_epic.get(epic_key, ""),
                epic.get("resolved_stable_since_date"),
            )
        else:
            # A worklog date is evidence of activity, not completion. Open and
            # reopened epics must keep aging against today's reporting date.
            completion = _derive_actual_completion(epic.get("due_date"), "", "")
        tk_budget_hours = _planner_tk_budget_hours(planner)
        jira_original_estimate_hours = _round_hours(epic.get("original_estimate_hours"))
        planned_estimate_basis = _planned_estimate_basis(
            jira_original_estimate_hours,
            tk_budget_hours,
            story_estimate_hours,
            subtask_estimate_hours,
        )
        planned_total_hours = planned_estimate_basis["hours"]
        schedule_as_of_day = _parse_iso_date(completion["actual_complete_date"]) or date.today()
        planned_to_date_hours = _planned_to_date_hours(
            planned_total_hours,
            epic.get("start_date"),
            epic.get("due_date"),
            schedule_as_of_day,
        )
        actual_to_date_hours = _round_hours(
            sum(
                actual_by_subtask.get(_to_text(subtask.get("issue_key")).upper(), 0.0)
                for subtask in subtasks
            )
        )
        schedule_variance_hours = _round_hours(actual_to_date_hours - planned_to_date_hours)
        schedule_variance_pct = _round_hours((schedule_variance_hours / planned_to_date_hours) * 100) if planned_to_date_hours > 0 else None
        schedule_proration_days = _schedule_proration_days(
            epic.get("start_date"),
            epic.get("due_date"),
            schedule_as_of_day,
        )
        planned_due_day = _parse_iso_date(epic.get("due_date"))
        schedule_variance_days = (planned_due_day - schedule_as_of_day).days if planned_due_day and schedule_as_of_day else None
        accuracy = _estimation_accuracy(planned_total_hours, total_actual_hours)
        jira_url = _to_text((planner or {}).get("jira_url")) or _issue_url(jira_base_url, epic_key)

        nested_stories: list[dict[str, Any]] = []
        for story in stories:
            story_key = _to_text(story.get("issue_key")).upper()
            nested_subtasks: list[dict[str, Any]] = []
            for subtask in subtasks_by_story.get(story_key, []):
                subtask_key = _to_text(subtask.get("issue_key")).upper()
                subtask_logs = list(worklogs_by_issue.get(subtask_key, []))
                subtask_actual = _round_hours(actual_by_subtask.get(subtask_key, 0.0))
                subtask_last_log = max([_to_text(log.get("date")) for log in subtask_logs if _to_text(log.get("date"))] or [""])
                subtask_completion = _derive_actual_completion(
                    subtask.get("due_date"),
                    subtask_last_log,
                    subtask.get("resolved_stable_since_date"),
                )
                nested_subtasks.append(
                    {
                        "issue_key": subtask_key,
                        "summary": _to_text(subtask.get("summary")) or subtask_key,
                        "issue_type": _normalize_issue_type(subtask.get("issue_type")),
                        "assignee": _to_text(subtask.get("assignee")),
                        "status": _to_text(subtask.get("status")),
                        "start_date": _to_text(subtask.get("start_date")),
                        "due_date": _to_text(subtask.get("due_date")),
                        "original_estimate_hours": _round_hours(subtask.get("original_estimate_hours")),
                        "actual_hours": subtask_actual,
                        "actual_complete_date": subtask_completion["actual_complete_date"],
                        "completion_bucket": subtask_completion["completion_bucket"],
                        "jira_url": _issue_url(jira_base_url, subtask_key),
                        "worklogs": subtask_logs,
                    }
                )
            story_last_log = max(
                [_to_text(log.get("date")) for sub in nested_subtasks for log in sub.get("worklogs", []) if _to_text(log.get("date"))]
                or [""]
            )
            story_completion = _derive_actual_completion(story.get("due_date"), story_last_log, story.get("resolved_stable_since_date"))
            nested_stories.append(
                {
                    "issue_key": story_key,
                    "summary": _to_text(story.get("summary")) or story_key,
                    "issue_type": _normalize_issue_type(story.get("issue_type")),
                    "assignee": _to_text(story.get("assignee")),
                    "status": _to_text(story.get("status")),
                    "start_date": _to_text(story.get("start_date")),
                    "due_date": _to_text(story.get("due_date")),
                    "original_estimate_hours": _round_hours(story.get("original_estimate_hours")),
                    "actual_hours": _round_hours(sum(_to_float(sub.get("actual_hours")) for sub in nested_subtasks)),
                    "actual_complete_date": story_completion["actual_complete_date"],
                    "completion_bucket": story_completion["completion_bucket"],
                    "jira_url": _issue_url(jira_base_url, story_key),
                    "subtasks": nested_subtasks,
                }
            )

        epic_dates = [
            _parse_iso_date(epic.get("start_date")),
            _parse_iso_date(epic.get("due_date")),
            _parse_iso_date(last_log_by_epic.get(epic_key, "")),
        ]
        epic_dates.extend(
            _parse_iso_date(item.get(field))
            for item in stories
            for field in ("start_date", "due_date")
        )
        epic_dates = [d for d in epic_dates if d]
        epic_start = min(epic_dates) if epic_dates else global_start
        epic_end = max(epic_dates) if epic_dates else global_end
        month_codes = _iter_months(epic_start, epic_end)
        planned_by_month, planned_by_assignee_month, story_planning_distribution = _distribute_story_estimates_by_month(
            stories,
            capacity_profiles,
            leave_by_assignee_day,
            capacity_basis,
        )
        monthly_plan_actual = [
            {
                "month": month,
                "planned_hours": _round_hours(planned_by_month.get(month, 0.0)),
                "actual_hours": _round_hours(actual_by_epic_month.get(epic_key, {}).get(month, 0.0)),
            }
            for month in month_codes
        ]
        for item in monthly_plan_actual:
            planned = _to_float(item.get("planned_hours"))
            actual = _to_float(item.get("actual_hours"))
            variance = _round_hours(actual - planned)
            item["schedule_variance_hours"] = variance
            item["schedule_variance_pct"] = _round_pct((variance / planned) * 100) if planned else None
            item["schedule_position"] = _schedule_position(variance)

        status_day = _parse_iso_date(completion["actual_complete_date"]) or date.today()
        status_month = _month_code(status_day)
        planned_total_hours = planned_estimate_basis["hours"]
        planned_to_date_hours = _planned_to_date_hours(planned_total_hours, epic.get("start_date"), epic.get("due_date"), status_day)
        actual_to_date_hours = _round_hours(
            sum(
                actual_by_subtask.get(_to_text(subtask.get("issue_key")).upper(), 0.0)
                for subtask in subtasks
            )
        )
        schedule_variance_hours = _round_hours(actual_to_date_hours - planned_to_date_hours)
        schedule_variance_pct = _round_pct((schedule_variance_hours / planned_to_date_hours) * 100) if planned_to_date_hours else None
        planned_due_day = _parse_iso_date(epic.get("due_date"))
        schedule_variance_days = (planned_due_day - status_day).days if planned_due_day else None
        estimation_accuracy_pct = _round_pct((planned_total_hours / total_actual_hours) * 100) if total_actual_hours else None
        recent_sv_trend = [item for item in monthly_plan_actual if _month_end(item["month"]) <= _month_end(status_month)][-3:]
        if len(recent_sv_trend) >= 2:
            first_sv = _to_float(recent_sv_trend[0].get("schedule_variance_hours"))
            last_sv = _to_float(recent_sv_trend[-1].get("schedule_variance_hours"))
            trend_delta = _round_hours(last_sv - first_sv)
            trend_direction = "improving" if trend_delta > 0.01 else ("declining" if trend_delta < -0.01 else "flat")
        else:
            trend_delta = 0.0
            trend_direction = "insufficient_data"
        schedule_variance_trend = {
            "months": recent_sv_trend,
            "delta_hours": trend_delta,
            "direction": trend_direction,
        }

        leaves_for_epic = []
        for leaf in leave_rows:
            leaf_day = _parse_iso_date(leaf.get("date"))
            leaf_assignee_key = _to_text(leaf.get("assignee")).casefold()
            if not leaf_day or not (epic_start <= leaf_day <= epic_end):
                continue
            if assigned_subtask_assignees and leaf_assignee_key not in assigned_subtask_assignees:
                continue
            leaves_for_epic.append(leaf)
        leave_summary: dict[str, dict[str, float]] = defaultdict(lambda: {"planned_hours": 0.0, "unplanned_hours": 0.0})
        for leaf in leaves_for_epic:
            assignee = _to_text(leaf.get("assignee")) or "Unassigned"
            if leaf.get("leave_class") == "unplanned":
                leave_summary[assignee]["unplanned_hours"] += _to_float(leaf.get("hours"))
            else:
                leave_summary[assignee]["planned_hours"] += _to_float(leaf.get("hours"))

        story_estimate_by_key = { _to_text(story.get("issue_key")).upper(): _to_float(story.get("original_estimate_hours")) for story in stories }
        story_summary_by_key = { _to_text(story.get("issue_key")).upper(): _to_text(story.get("summary")) for story in stories }
        unassigned_planning_stories = [
            {
                "issue_key": _to_text(story.get("issue_key")).upper(),
                "summary": _to_text(story.get("summary")) or _to_text(story.get("issue_key")).upper(),
                "issue_type": _normalize_issue_type(story.get("issue_type")),
                "start_date": _to_text(story.get("start_date")),
                "due_date": _to_text(story.get("due_date")),
                "original_estimate_hours": _round_hours(story.get("original_estimate_hours")),
            }
            for story in stories
            if _to_float(story.get("original_estimate_hours")) > 0 and not _to_text(story.get("assignee"))
        ]
        equal_story_estimate_count = 0
        over_original_estimate_count = 0
        estimate_quality_details: list[dict[str, Any]] = []
        for subtask in subtasks:
            subtask_key = _to_text(subtask.get("issue_key")).upper()
            story_key = _to_text(subtask.get("story_key")).upper() or _to_text(subtask.get("parent_issue_key")).upper()
            sub_est = _to_float(subtask.get("original_estimate_hours"))
            sub_actual = actual_by_subtask.get(subtask_key, 0.0)
            matches_story_estimate = bool(story_key and abs(sub_est - story_estimate_by_key.get(story_key, -1)) < 0.01)
            over_original_estimate = sub_actual > sub_est
            if matches_story_estimate:
                equal_story_estimate_count += 1
            if over_original_estimate:
                over_original_estimate_count += 1
            estimate_quality_details.append(
                {
                    "issue_key": subtask_key,
                    "summary": _to_text(subtask.get("summary")) or subtask_key,
                    "story_key": story_key,
                    "story_name": story_summary_by_key.get(story_key, story_key),
                    "assignee": _to_text(subtask.get("assignee")) or "Unassigned",
                    "issue_type": _normalize_issue_type(subtask.get("issue_type")),
                    "original_estimate_hours": _round_hours(sub_est),
                    "actual_hours": _round_hours(sub_actual),
                    "matches_story_estimate": matches_story_estimate,
                    "over_original_estimate": over_original_estimate,
                    "jira_url": _issue_url(jira_base_url, subtask_key),
                }
            )

        resource_utilization: list[dict[str, Any]] = []
        assignee_schedule_variance: list[dict[str, Any]] = []
        assignee_names = set(planned_by_assignee_month.keys()) | set(actual_by_epic_author.get(epic_key, {}).keys())
        for assignee in sorted(assignee_names, key=lambda value: value.lower()):
            planned = _round_hours(
                sum(hours for month, hours in planned_by_assignee_month.get(assignee, {}).items() if _month_end(month) <= _month_end(status_month))
            )
            actual = _round_hours(actual_by_epic_author.get(epic_key, {}).get(assignee, 0.0))
            variance = _round_hours(actual - planned)
            assignee_schedule_variance.append(
                {
                    "assignee": assignee,
                    "planned_to_date_hours": planned,
                    "actual_to_date_hours": actual,
                    "schedule_variance_hours": variance,
                    "schedule_variance_pct": _round_pct((variance / planned) * 100) if planned else None,
                    "schedule_position": _schedule_position(variance),
                }
            )

        for author, hours in sorted(actual_by_epic_author.get(epic_key, {}).items(), key=lambda item: (-item[1], item[0].lower())):
            for month in month_codes:
                epic_hours = _round_hours(sum(_to_float(log.get("hours")) for sub in subtasks for log in worklogs_by_issue.get(_to_text(sub.get("issue_key")).upper(), []) if _to_text(log.get("author")) == author and _to_text(log.get("date")).startswith(month)))
                if epic_hours <= 0:
                    continue
                available = _round_hours(capacity_by_month.get(month, {}).get(author.casefold(), capacity_by_month.get(month, {}).get("", {})).get("available_hours"))
                resource_utilization.append(
                    {
                        "month": month,
                        "assignee": author,
                        "epic_hours": epic_hours,
                        "available_hours": available,
                        "utilization_pct": _round_hours((epic_hours / available) * 100) if available else 0.0,
                    }
                )

        row = {
            "epic_key": epic_key,
            "epic_name": _to_text(epic.get("summary")) or _to_text((planner or {}).get("epic_name")) or epic_key,
            "jira_url": jira_url,
            "assignee": _to_text(epic.get("assignee")),
            "project_key": project_key,
            "project_name": _to_text((planner or {}).get("project_name")) or project_name_by_key.get(project_key, "") or canonical_project_name or project_key,
            "product": _to_text((planner or {}).get("product_category")),
            "tk_budget_hours": tk_budget_hours,
            "tk_budget_days": None if tk_budget_hours is None else _round_hours(tk_budget_hours / HOURS_PER_DAY),
            "jira_original_estimate_hours": jira_original_estimate_hours,
            "story_estimate_hours": story_estimate_hours,
            "subtask_estimate_hours": subtask_estimate_hours,
            "planned_total_hours": planned_total_hours,
            "planned_to_date_hours": planned_to_date_hours,
            "planned_start": _to_text(epic.get("start_date")),
            "planned_due": _to_text(epic.get("due_date")),
            "total_actual_hours": total_actual_hours,
            "total_actual_days": _round_hours(total_actual_hours / HOURS_PER_DAY),
            "actual_to_date_hours": actual_to_date_hours,
            "actual_complete_date": completion["actual_complete_date"],
            "actual_complete_source": completion["actual_complete_source"],
            "completion_bucket": completion["completion_bucket"],
            "schedule_variance_days": schedule_variance_days,
            "schedule_variance_date_position": _position_for_signed_delta(schedule_variance_days),
            "schedule_variance_hours": schedule_variance_hours,
            "schedule_variance_pct": schedule_variance_pct,
            "schedule_variance_hours_position": _position_for_signed_delta(schedule_variance_hours),
            "estimation_accuracy_pct": accuracy["pct"],
            "estimation_accuracy_status": accuracy["status"],
            "epic_status": _to_text(epic.get("status")),
            "headcount": len(headcount_by_epic.get(epic_key, set())),
            "story_count": len(stories),
            "subtask_count": len(subtasks),
            "planned_to_date_hours": planned_to_date_hours,
            "actual_to_date_hours": actual_to_date_hours,
            "schedule_variance_days": schedule_variance_days,
            "schedule_variance_date_basis": completion["actual_complete_date"] or status_day.isoformat(),
            "schedule_variance_date_basis_type": "actual_complete_date" if completion["actual_complete_date"] else "current_date",
            "schedule_variance_date_position": _schedule_position(float(schedule_variance_days or 0)),
            "schedule_variance_hours": schedule_variance_hours,
            "schedule_variance_pct": schedule_variance_pct,
            "schedule_variance_hours_position": _schedule_position(schedule_variance_hours),
            "schedule_variance_breakdown": {
                "estimate_source": planned_estimate_basis["source"],
                "estimate_label": planned_estimate_basis["label"],
                "estimate_hours": planned_total_hours,
                "planned_start": _to_text(epic.get("start_date")),
                "planned_due": _to_text(epic.get("due_date")),
                "as_of_date": status_day.isoformat(),
                "as_of_basis": "actual_complete_date" if completion["actual_complete_date"] else "current_date",
                **schedule_proration_days,
                "planned_to_date_hours": planned_to_date_hours,
                "actual_to_date_hours": actual_to_date_hours,
                "schedule_variance_hours": schedule_variance_hours,
                "schedule_variance_pct": schedule_variance_pct,
                "position": _schedule_position(schedule_variance_hours),
            },
            "estimation_accuracy_pct": estimation_accuracy_pct,
            "estimation_accuracy_status": _estimate_accuracy_status(estimation_accuracy_pct),
            "stories": nested_stories,
            "analytics": {
                "monthly_plan_actual": monthly_plan_actual,
                "monthly_plan_basis": {
                    "estimate_source": "story_original_estimate",
                    "capacity_basis": capacity_basis,
                    "distribution": "capacity_weighted_daily_proration",
                },
                "story_planning_distribution": story_planning_distribution,
                "unassigned_planning_stories": unassigned_planning_stories,
                "schedule_variance": {
                    "planned_to_date_hours": planned_to_date_hours,
                    "actual_to_date_hours": actual_to_date_hours,
                    "schedule_variance_hours": schedule_variance_hours,
                    "schedule_variance_pct": schedule_variance_pct,
                    "schedule_variance_hours_position": _schedule_position(schedule_variance_hours),
                    "schedule_variance_days": schedule_variance_days,
                    "schedule_variance_date_position": _schedule_position(float(schedule_variance_days or 0)),
                    "schedule_variance_date_basis": completion["actual_complete_date"] or status_day.isoformat(),
                    "schedule_variance_date_basis_type": "actual_complete_date" if completion["actual_complete_date"] else "current_date",
                    "planned_total_hours": planned_total_hours,
                    "actual_total_hours": total_actual_hours,
                    "planned_due": _to_text(epic.get("due_date")),
                    "actual_complete_date": completion["actual_complete_date"],
                    "estimation_accuracy_pct": estimation_accuracy_pct,
                    "estimation_accuracy_status": _estimate_accuracy_status(estimation_accuracy_pct),
                    "trend_3_months": schedule_variance_trend,
                    "assignees": assignee_schedule_variance,
                },
                "gantt": {
                    "start_date": epic_start.isoformat(),
                    "end_date": epic_end.isoformat(),
                    "worklogs": [log for sub in subtasks for log in worklogs_by_issue.get(_to_text(sub.get("issue_key")).upper(), [])],
                    "leaves": leaves_for_epic,
                },
                "completion_counts": {
                    bucket: sum(1 for story in nested_stories for sub in story.get("subtasks", []) if sub.get("completion_bucket") == bucket)
                    for bucket in ("before_due", "on_due", "after_due", "not_completed", "no_due_date")
                },
                "completion_stats": [
                    {
                        "bucket": bucket,
                        "label": {
                            "before_due": "Completed before due date",
                            "on_due": "Completed on due date",
                            "after_due": "Completed after due date",
                            "not_completed": "Not completed",
                            "no_due_date": "No due date",
                        }[bucket],
                        "count": sum(1 for story in nested_stories for sub in story.get("subtasks", []) if sub.get("completion_bucket") == bucket),
                        "items": [
                            {
                                "issue_key": sub.get("issue_key"),
                                "summary": sub.get("summary"),
                                "story_key": story.get("issue_key"),
                                "story_name": story.get("summary"),
                                "assignee": sub.get("assignee"),
                                "status": sub.get("status"),
                                "due_date": sub.get("due_date"),
                                "actual_complete_date": sub.get("actual_complete_date"),
                                "jira_url": sub.get("jira_url"),
                            }
                            for story in nested_stories
                            for sub in story.get("subtasks", [])
                            if sub.get("completion_bucket") == bucket
                        ],
                    }
                    for bucket in ("before_due", "on_due", "after_due", "not_completed", "no_due_date")
                ],
                "resource_effort": [
                    {"assignee": author, "hours": _round_hours(hours), "pct": _round_hours((hours / total_actual_hours) * 100) if total_actual_hours else 0.0}
                    for author, hours in sorted(actual_by_epic_author.get(epic_key, {}).items(), key=lambda item: (-item[1], item[0].lower()))
                ],
                "story_effort": [
                    {
                        "story_key": story.get("issue_key"),
                        "story_name": story.get("summary"),
                        "hours": _round_hours(actual_by_epic_story.get(epic_key, {}).get(story.get("issue_key"), 0.0)),
                        "pct": _round_hours((actual_by_epic_story.get(epic_key, {}).get(story.get("issue_key"), 0.0) / total_actual_hours) * 100) if total_actual_hours else 0.0,
                    }
                    for story in nested_stories
                ],
                "team_effort": [
                    {"team": team, "hours": _round_hours(hours), "pct": _round_hours((hours / total_actual_hours) * 100) if total_actual_hours else 0.0}
                    for team, hours in sorted(actual_by_epic_team.get(epic_key, {}).items(), key=lambda item: (-item[1], item[0].lower()))
                ],
                "resource_utilization": resource_utilization,
                "support_effort": [
                    {"assignee": author, "hours": _round_hours(hours)}
                    for author, hours in sorted(support_by_epic_author.get(epic_key, {}).items(), key=lambda item: (-item[1], item[0].lower()))
                ],
                "estimate_quality": {
                    "subtasks_equal_story_estimate_count": equal_story_estimate_count,
                    "subtasks_over_original_estimate_count": over_original_estimate_count,
                    "subtask_count": len(subtasks),
                    "within_original_estimate_count": max(0, len(subtasks) - over_original_estimate_count),
                    "equal_story_estimate_pct": _round_hours((equal_story_estimate_count / len(subtasks)) * 100) if subtasks else 0.0,
                    "over_original_estimate_pct": _round_hours((over_original_estimate_count / len(subtasks)) * 100) if subtasks else 0.0,
                    "details": estimate_quality_details,
                },
                "leave_summary": [
                    {
                        "assignee": assignee,
                        "planned_hours": _round_hours(values.get("planned_hours")),
                        "unplanned_hours": _round_hours(values.get("unplanned_hours")),
                    }
                    for assignee, values in sorted(leave_summary.items(), key=lambda item: item[0].lower())
                ],
            },
        }
        rows.append(row)

    rows.sort(key=lambda item: (_to_text(item.get("project_name")).lower(), _to_text(item.get("planned_start")), _to_text(item.get("epic_key"))))
    total_planned_hours = _round_hours(sum(_to_float(row.get("planned_total_hours")) for row in rows))
    total_actual_hours = _round_hours(sum(_to_float(row.get("total_actual_hours")) for row in rows))
    total_planned_to_date_hours = _round_hours(sum(_to_float(row.get("planned_to_date_hours")) for row in rows))
    total_schedule_variance_hours = _round_hours(sum(_to_float(row.get("schedule_variance_hours")) for row in rows))
    totals = {
        "epic_count": len(rows),
        "story_count": sum(int(row.get("story_count") or 0) for row in rows),
        "subtask_count": sum(int(row.get("subtask_count") or 0) for row in rows),
        "tk_budget_hours": _round_hours(sum(_to_float(row.get("tk_budget_hours")) for row in rows)),
        "jira_original_estimate_hours": _round_hours(sum(_to_float(row.get("jira_original_estimate_hours")) for row in rows)),
        "planned_total_hours": _round_hours(sum(_to_float(row.get("planned_total_hours")) for row in rows)),
        "story_estimate_hours": _round_hours(sum(_to_float(row.get("story_estimate_hours")) for row in rows)),
        "subtask_estimate_hours": _round_hours(sum(_to_float(row.get("subtask_estimate_hours")) for row in rows)),
        "planned_total_hours": total_planned_hours,
        "planned_to_date_hours": total_planned_to_date_hours,
        "total_actual_hours": total_actual_hours,
        "actual_to_date_hours": _round_hours(sum(_to_float(row.get("actual_to_date_hours")) for row in rows)),
        "schedule_variance_hours": total_schedule_variance_hours,
        "headcount": len({log.get("author").casefold() for row in rows for story in row.get("stories", []) for sub in story.get("subtasks", []) for log in sub.get("worklogs", []) if _to_text(log.get("author"))}),
    }
    totals["schedule_variance_pct"] = _round_pct((totals["schedule_variance_hours"] / totals["planned_to_date_hours"]) * 100) if totals["planned_to_date_hours"] else None
    totals["estimation_accuracy_pct"] = _round_pct((totals["planned_total_hours"] / totals["total_actual_hours"]) * 100) if totals["total_actual_hours"] else None
    return {
        "canonical_run_id": run_id,
        "from_date": range_start.isoformat() if range_start else "",
        "to_date": range_end.isoformat() if range_end else "",
        "date_filter_active": date_filter_active,
        "selected_projects": sorted(selected_project_keys),
        "project_options": [{"project_key": key, "project_name": value} for key, value in sorted(project_options.items())],
        "rows": rows,
        "totals": totals,
        "meta": {
            "hours_per_day": HOURS_PER_DAY,
            "scope_basis": "Default scope includes every canonical Jira epic. Date and project filters only include/exclude epics; nested stories, subtasks, and worklogs remain full epic-lifetime data.",
            "actual_complete_basis": "Resolved-status epics use the later of descendant subtask last worklog date and epic resolved-stable-since date. Unresolved and reopened epics have no Actual Complete Date.",
            "schedule_variance_basis": "Resolved-status epics use Actual Complete Date as the SV reporting date; unresolved and reopened epics use the current date. Date SV is Jira epic planned due date minus that reporting date; hour SV is actual-to-date minus planned-to-date using the Jira epic original estimate shown in Planned vs Actual Hours. TK budget and child estimates are fallback-only when the epic original estimate is missing.",
            "estimation_accuracy_basis": "Jira epic original estimate divided by actual hours multiplied by 100. The ideal range is 85% to 115%; below 70% is marked broken.",
        },
    }
