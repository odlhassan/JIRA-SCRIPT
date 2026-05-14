from __future__ import annotations

import calendar
import sqlite3
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path
from typing import Any, DefaultDict

from canonical_report_data import build_rlt_leave_snapshot
from generate_assignee_hours_report import (
    _list_capacity_profiles,
    _normalize_capacity_payload,
    calculate_capacity_metrics,
)
from generate_employee_performance_report import (
    _list_performance_teams,
    _load_performance_resource_resignation_map,
)


HOURS_PER_DAY = 8.0


def _to_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _parse_iso_date(value: Any) -> date | None:
    text = _to_text(value)
    if not text:
        return None
    try:
        return date.fromisoformat(text[:10])
    except ValueError:
        return None


def _month_bounds(month_value: Any) -> tuple[date, date]:
    text = _to_text(month_value)
    if len(text) != 7 or text[4] != "-":
        raise ValueError("Invalid month. Expected YYYY-MM.")
    try:
        year = int(text[:4])
        month = int(text[5:7])
        last_day = calendar.monthrange(year, month)[1]
    except Exception as exc:
        raise ValueError("Invalid month. Expected YYYY-MM.") from exc
    return date(year, month, 1), date(year, month, last_day)


def _round_hours(value: float) -> float:
    return round(float(value or 0.0) + 1e-9, 2)


def _hours_from_man_days(value: Any) -> float | None:
    text = _to_text(value)
    if not text:
        return None
    try:
        parsed = float(text)
    except Exception:
        return None
    if parsed < 0:
        return None
    return _round_hours(parsed * HOURS_PER_DAY)


def _is_subtask_type(value: Any) -> bool:
    low = _to_text(value).lower()
    return "sub-task" in low or "subtask" in low


def _is_epic_type(value: Any) -> bool:
    return "epic" in _to_text(value).lower()


def _is_story_type(value: Any) -> bool:
    return "story" in _to_text(value).lower()


def _is_resolved_status_text(value: Any) -> bool:
    text = _to_text(value).lower()
    if not text:
        return False
    return text in {"resolved", "resolved!", "done", "closed", "complete", "completed"}


def _is_on_hold_status_text(value: Any) -> bool:
    text = _to_text(value).lower().replace("-", " ").replace("_", " ").strip()
    return text in {"on hold", "hold", "paused", "deferred"}


def _normalize_planner_delivery_status(value: Any) -> str:
    raw = _to_text(value).strip()
    low = raw.lower().replace("-", " ").replace("_", " ").strip()
    if low in {"", "yet to start"}:
        return "Yet to start"
    if low in {"on track", "ontrack"}:
        return "On-track"
    if low == "late":
        return "Late"
    return raw if raw else "Yet to start"


def _delivery_status_view(planner_delivery: Any, jira_status: Any) -> str:
    """Planner delivery_status can lag Jira; align display when planner still says Yet to start."""
    planner = _normalize_planner_delivery_status(planner_delivery)
    if planner != "Yet to start":
        return planner
    jira_raw = _to_text(jira_status)
    if not jira_raw:
        return planner
    if _is_resolved_status_text(jira_raw):
        return "On-track"
    jira_low = jira_raw.lower().replace("-", " ").replace("_", " ")
    if "progress" in jira_low:
        return "On-track"
    return planner


def _chunked(items: list[str], chunk_size: int) -> list[list[str]]:
    return [items[index : index + chunk_size] for index in range(0, len(items), chunk_size)]


def _load_canonical_issue_maps(db_path: Path, run_id: str) -> tuple[dict[str, dict[str, Any]], dict[str, set[str]]]:
    epic_meta: dict[str, dict[str, Any]] = {}
    story_to_epic: dict[str, str] = {}
    subtask_rows: list[dict[str, Any]] = []
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT issue_key, project_key, issue_type, summary, status, parent_issue_key, story_key, epic_key
            FROM canonical_issues
            WHERE run_id = ?
            """,
            (run_id,),
        ).fetchall()

    for row in rows:
        issue_key = _to_text(row["issue_key"]).upper()
        issue_type = _to_text(row["issue_type"])
        epic_key = _to_text(row["epic_key"]).upper()
        parent_key = _to_text(row["parent_issue_key"]).upper()
        story_key = _to_text(row["story_key"]).upper()
        if _is_epic_type(issue_type):
            epic_meta[issue_key] = {
                "jira_status": _to_text(row["status"]),
                "jira_summary": _to_text(row["summary"]),
                "project_key": _to_text(row["project_key"]).upper(),
            }
        elif _is_story_type(issue_type):
            resolved_epic = epic_key or parent_key
            if issue_key and resolved_epic:
                story_to_epic[issue_key] = resolved_epic
        elif _is_subtask_type(issue_type):
            subtask_rows.append(dict(row))

    subtask_keys_by_epic: dict[str, set[str]] = defaultdict(set)
    for row in subtask_rows:
        subtask_key = _to_text(row.get("issue_key")).upper()
        if not subtask_key:
            continue
        row_epic = _to_text(row.get("epic_key")).upper()
        story_key = _to_text(row.get("story_key")).upper()
        parent_key = _to_text(row.get("parent_issue_key")).upper()
        epic_key = row_epic or story_to_epic.get(story_key) or story_to_epic.get(parent_key) or ""
        if epic_key:
            subtask_keys_by_epic[epic_key].add(subtask_key)

    return epic_meta, subtask_keys_by_epic


def _load_worklog_metrics(
    db_path: Path,
    run_id: str,
    subtask_keys_by_epic: dict[str, set[str]],
    month_start: date,
    month_end: date,
) -> tuple[dict[str, float], dict[str, bool]]:
    issue_to_epic: dict[str, str] = {}
    for epic_key, issue_keys in subtask_keys_by_epic.items():
        for issue_key in issue_keys:
            issue_to_epic[issue_key] = epic_key

    actual_hours_by_epic: dict[str, float] = defaultdict(float)
    has_worklog_through_month_end: dict[str, bool] = defaultdict(bool)
    issue_keys = sorted(issue_to_epic)
    if not issue_keys:
        return actual_hours_by_epic, has_worklog_through_month_end

    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        for chunk in _chunked(issue_keys, 400):
            placeholders = ",".join("?" for _ in chunk)
            rows = conn.execute(
                f"""
                SELECT issue_key, started_date, hours_logged
                FROM canonical_worklogs
                WHERE run_id = ?
                  AND issue_key IN ({placeholders})
                  AND started_date <= ?
                """,
                [run_id, *chunk, month_end.isoformat()],
            ).fetchall()
            for row in rows:
                issue_key = _to_text(row["issue_key"]).upper()
                epic_key = issue_to_epic.get(issue_key)
                if not epic_key:
                    continue
                started = _parse_iso_date(row["started_date"])
                hours = float(row["hours_logged"] or 0.0)
                if started is None or hours <= 0:
                    continue
                has_worklog_through_month_end[epic_key] = True
                if month_start <= started <= month_end:
                    actual_hours_by_epic[epic_key] += hours

    return actual_hours_by_epic, has_worklog_through_month_end


def _approved_dates(epic_plan: dict[str, Any]) -> tuple[date | None, date | None, str, str]:
    start_text = _to_text(epic_plan.get("tk_budgeted_start_date")) or _to_text(epic_plan.get("start_date"))
    due_text = _to_text(epic_plan.get("tk_budgeted_due_date")) or _to_text(epic_plan.get("due_date"))
    return _parse_iso_date(start_text), _parse_iso_date(due_text), start_text, due_text


def _capacity_profile_key(profile: dict[str, Any]) -> str:
    return f"{_to_text(profile.get('from_date'))}|{_to_text(profile.get('to_date'))}"


def _find_capacity_profile_by_key(db_path: Path, profile_key: str) -> dict[str, Any] | None:
    key = _to_text(profile_key).strip()
    if not key:
        return None
    try:
        profiles = _list_capacity_profiles(db_path)
    except Exception:
        return None
    for p in profiles:
        if isinstance(p, dict) and _capacity_profile_key(p) == key:
            return p
    return None


def _pick_overlapping_capacity_profile(db_path: Path, month_start: date, month_end: date) -> dict[str, Any] | None:
    try:
        profiles = _list_capacity_profiles(db_path)
    except Exception:
        profiles = []
    best: dict[str, Any] | None = None
    best_days = -1
    for p in profiles:
        pf = _parse_iso_date(p.get("from_date"))
        pt = _parse_iso_date(p.get("to_date"))
        if pf is None or pt is None:
            continue
        lo = max(pf, month_start)
        hi = min(pt, month_end)
        if lo <= hi:
            span = (hi - lo).days + 1
            if span > best_days:
                best_days = span
                best = p
    return best


def _month_capacity_settings_from_saved_profile(profile: dict[str, Any], month_start: date, month_end: date) -> dict[str, Any]:
    holidays = profile.get("holiday_dates")
    if not isinstance(holidays, list):
        holidays = []
    merged = {
        "from_date": month_start.isoformat(),
        "to_date": month_end.isoformat(),
        "employee_count": int(profile.get("employee_count") or 0),
        "standard_hours_per_day": float(profile.get("standard_hours_per_day") or 8.0),
        "ramadan_start_date": _to_text(profile.get("ramadan_start_date")),
        "ramadan_end_date": _to_text(profile.get("ramadan_end_date")),
        "ramadan_hours_per_day": float(profile.get("ramadan_hours_per_day") or 6.5),
        "holiday_dates": holidays,
    }
    rs_text = _to_text(merged.get("ramadan_start_date"))
    re_text = _to_text(merged.get("ramadan_end_date"))
    if rs_text and re_text:
        rs_d = _parse_iso_date(rs_text)
        re_d = _parse_iso_date(re_text)
        if rs_d and re_d:
            lo = max(rs_d, month_start)
            hi = min(re_d, month_end)
            if lo > hi:
                merged["ramadan_start_date"] = ""
                merged["ramadan_end_date"] = ""
    return _normalize_capacity_payload(merged, require_all_fields=True)


def _capacity_settings_for_calendar_month(
    db_path: Path,
    month_start: date,
    month_end: date,
    *,
    capacity_profile_key: str | None = None,
) -> tuple[dict[str, Any], str, str]:
    requested = _to_text(capacity_profile_key).strip()
    if requested:
        sel = _find_capacity_profile_by_key(db_path, requested)
        if sel:
            return _month_capacity_settings_from_saved_profile(sel, month_start, month_end), "selected_profile", requested
    profile = _pick_overlapping_capacity_profile(db_path, month_start, month_end)
    applied_key = _capacity_profile_key(profile) if profile else ""
    if not profile:
        normalized = _normalize_capacity_payload(
            {
                "from_date": month_start.isoformat(),
                "to_date": month_end.isoformat(),
                "employee_count": 0,
                "standard_hours_per_day": 8.0,
                "ramadan_start_date": "",
                "ramadan_end_date": "",
                "ramadan_hours_per_day": 6.5,
                "holiday_dates": [],
            },
            require_all_fields=True,
        )
        return normalized, "default", ""
    return _month_capacity_settings_from_saved_profile(profile, month_start, month_end), "overlapping_profile", applied_key


def _nested_aligned_leave_by_assignee(
    snapshot: dict[str, Any],
    month_start: date,
    month_end: date,
) -> tuple[dict[str, str], dict[str, float], str]:
    """Match Nested View Total Leaves Planned: distributed buckets first, else daily planned buckets, else raw overlap."""
    display_by_lower: dict[str, str] = {}
    leave_by_lower: DefaultDict[str, float] = defaultdict(float)

    def track(name_raw: Any, hours: float) -> None:
        raw_name = _to_text(name_raw)
        if not raw_name or hours <= 0:
            return
        key = raw_name.lower()
        display_by_lower.setdefault(key, raw_name)
        leave_by_lower[key] += hours

    distributed = list(snapshot.get("distributed_subtasks") or [])
    dist_used = False
    for row in distributed:
        if not isinstance(row, dict):
            continue
        bucket = _parse_iso_date(_to_text(row.get("planned_date_for_bucket") or row.get("start_date")))
        if bucket is None or not (month_start <= bucket <= month_end):
            continue
        planned_h = max(0.0, float(row.get("original_estimate_hours") or 0.0))
        actual_h = max(0.0, float(row.get("total_worklog_hours") or 0.0))
        if planned_h <= 0 and actual_h <= 0:
            continue
        dist_used = True
        track(row.get("assignee"), planned_h)

    if dist_used:
        for lk in leave_by_lower:
            leave_by_lower[lk] = _round_hours(float(leave_by_lower[lk]))
        return display_by_lower, dict(leave_by_lower), "distributed_subtasks"

    daily = list(snapshot.get("daily") or [])
    embedded_used = False
    for row in daily:
        if not isinstance(row, dict):
            continue
        day = _parse_iso_date(_to_text(row.get("period_day")))
        if day is None or not (month_start <= day <= month_end):
            continue
        pt = float(row.get("planned_taken_hours") or 0.0)
        pn = float(row.get("planned_not_taken_hours") or 0.0)
        if pt + pn <= 0:
            continue
        embedded_used = True
        track(row.get("assignee"), pt + pn)

    if embedded_used:
        for lk in leave_by_lower:
            leave_by_lower[lk] = _round_hours(float(leave_by_lower[lk]))
        return display_by_lower, dict(leave_by_lower), "daily_planned_buckets"

    raw_tasks = list(snapshot.get("raw_subtasks") or [])
    raw_used = False
    for row in raw_tasks:
        if not isinstance(row, dict):
            continue
        start_d = _parse_iso_date(_to_text(row.get("start_date")))
        due_d = _parse_iso_date(_to_text(row.get("due_date")))
        overlaps = False
        if start_d and due_d:
            overlaps = bool(start_d <= month_end and due_d >= month_start)
        elif start_d:
            overlaps = bool(month_start <= start_d <= month_end)
        elif due_d:
            overlaps = bool(month_start <= due_d <= month_end)
        if not overlaps:
            continue
        planned_h = max(0.0, float(row.get("original_estimate_hours") or 0.0))
        actual_h = max(0.0, float(row.get("total_worklog_hours") or 0.0))
        if planned_h <= 0 and actual_h <= 0:
            continue
        raw_used = True
        track(row.get("assignee"), planned_h)

    for lk in leave_by_lower:
        leave_by_lower[lk] = _round_hours(float(leave_by_lower[lk]))
    return display_by_lower, dict(leave_by_lower), ("raw_subtasks_overlap" if raw_used else "none")


def build_workforce_month_payload(
    db_path: Path,
    month_start: date,
    month_end: date,
    canonical_run_id: str,
    *,
    selected_assignees: set[str] | None = None,
    capacity_profile_key: str | None = None,
) -> dict[str, Any]:
    requested_prof_key = _to_text(capacity_profile_key).strip()
    settings, profile_hint, applied_profile_key = _capacity_settings_for_calendar_month(
        db_path,
        month_start,
        month_end,
        capacity_profile_key=capacity_profile_key,
    )
    cap = calculate_capacity_metrics(settings)
    team_capacity_hours = float(cap["metrics"].get("available_capacity_hours") or 0.0)
    n_team = int(settings.get("employee_count") or 0)
    per_person_cap = _round_hours(team_capacity_hours / n_team) if n_team > 0 else 0.0

    try:
        capacity_profiles_out = _list_capacity_profiles(db_path)
    except Exception:
        capacity_profiles_out = []

    run_id = _to_text(canonical_run_id)
    snapshot = (
        build_rlt_leave_snapshot(db_path, run_id, month_start.isoformat(), month_end.isoformat())
        if run_id
        else {"daily": []}
    )
    display_by_lower, leave_by_lower, leave_aggregate_source = _nested_aligned_leave_by_assignee(
        snapshot,
        month_start,
        month_end,
    )
    known_keys = sorted(display_by_lower.keys(), key=lambda k: display_by_lower[k].lower())

    filter_lowers: set[str] | None = None
    if selected_assignees:
        filter_lowers = {_to_text(a).lower() for a in selected_assignees if _to_text(a)}

    if filter_lowers is None:
        k_sel = n_team
        capacity_hours = _round_hours(team_capacity_hours)
        leave_hours = _round_hours(sum(float(leave_by_lower.get(k, 0.0)) for k in known_keys))
        active_for_rows = list(known_keys)
    else:
        active_for_rows = sorted(filter_lowers, key=lambda k: display_by_lower.get(k, k).lower())
        k_sel = len(active_for_rows)
        capacity_hours = _round_hours(team_capacity_hours * (k_sel / n_team)) if n_team > 0 else 0.0
        leave_hours = _round_hours(
            sum(float(leave_by_lower.get(k, 0.0)) for k in filter_lowers)
        )

    assignee_rows: list[dict[str, Any]] = []
    for lk in active_for_rows:
        display = display_by_lower.get(lk) or lk
        lev = float(leave_by_lower.get(lk, 0.0))
        assignee_rows.append(
            {
                "name": display,
                "leave_hours": _round_hours(lev),
                "leave_days": _round_hours(lev / HOURS_PER_DAY),
                "per_person_capacity_hours": per_person_cap,
                "per_person_availability_hours": _round_hours(per_person_cap - lev),
            }
        )

    availability_hours = _round_hours(capacity_hours - leave_hours)

    try:
        raw_perf_teams = _list_performance_teams(db_path)
    except Exception:
        raw_perf_teams = []

    team_display_by_lower: dict[str, str] = {}
    for t in raw_perf_teams or []:
        t_rec = t if isinstance(t, dict) else {}
        for raw_m in t_rec.get("assignees") or []:
            m = _to_text(raw_m)
            if m:
                team_display_by_lower.setdefault(m.lower(), m)

    option_by_lower: dict[str, str] = {}
    option_by_lower.update(team_display_by_lower)
    option_by_lower.update(display_by_lower)
    option_names = [option_by_lower[k] for k in sorted(option_by_lower.keys(), key=lambda x: option_by_lower[x].lower())]
    resignation_by_name: dict[str, dict[str, Any]] = {}
    try:
        resignation_by_name = _load_performance_resource_resignation_map(db_path, option_names)
    except Exception:
        resignation_by_name = {}
    employee_options: list[dict[str, Any]] = []
    for name in option_names:
        rec = resignation_by_name.get(name) or {}
        employee_options.append(
            {
                "name": name,
                "resigned": bool(rec.get("resigned")),
                "resignation_date": rec.get("resignation_date"),
            }
        )
    for row in assignee_rows:
        nm = _to_text(row.get("name"))
        rec = resignation_by_name.get(nm) or {}
        row["resigned"] = bool(rec.get("resigned"))
        row["resignation_date"] = rec.get("resignation_date")

    teams_sections: list[dict[str, Any]] = []
    grouped_name_lower: set[str] = set()
    for t in raw_perf_teams or []:
        t_rec = t if isinstance(t, dict) else {}
        team_name = _to_text(t_rec.get("team_name"))
        leader = _to_text(t_rec.get("team_leader"))
        if not team_name:
            continue
        members_out: list[dict[str, Any]] = []
        for raw_m in t_rec.get("assignees") or []:
            m = _to_text(raw_m)
            if not m:
                continue
            lk = m.lower()
            disp_name = display_by_lower.get(lk) or team_display_by_lower.get(lk) or m
            rec_m = resignation_by_name.get(disp_name) or {}
            members_out.append(
                {
                    "name": disp_name,
                    "resigned": bool(rec_m.get("resigned")),
                    "resignation_date": rec_m.get("resignation_date"),
                }
            )
            grouped_name_lower.add(disp_name.casefold())
        members_out.sort(key=lambda item: _to_text(item.get("name")).lower())
        if members_out:
            teams_sections.append(
                {
                    "team_name": team_name,
                    "team_leader": leader,
                    "members": members_out,
                }
            )

    tree_ungrouped = [dict(e) for e in employee_options if _to_text(e.get("name")).casefold() not in grouped_name_lower]
    employee_tree = {"teams": teams_sections, "ungrouped": tree_ungrouped}

    if profile_hint == "selected_profile":
        cap_basis = (
            "Saved capacity profile explicitly selected for this report (date range clipped to the calendar month)."
        )
    elif profile_hint == "overlapping_profile":
        cap_basis = (
            "assignee_capacity_settings profile with the longest overlap with the month (calendar clipped to month)."
        )
    else:
        cap_basis = "defaults (no saved profile applies to this month)"

    leave_basis_by_src = {
        "distributed_subtasks": (
            "Same as Nested View Total Leaves Planned: canonical RLT snapshot Subtasks_Distributed buckets in the month "
            "(sum of original estimate hours per bucket day)."
        ),
        "daily_planned_buckets": (
            "Canonical RLT Daily_Assignee rows in the month — planned taken + planned not yet taken only "
            "(Nested View embedded fallback; excludes unplanned/unknown taken hours)."
        ),
        "raw_subtasks_overlap": (
            "Canonical RLT raw leave subtasks overlapping the month (original estimate hours; Nested View fallback when "
            "daily/distributed rows are absent)."
        ),
        "none": "No RLT leave rows in this month for the current canonical run.",
    }

    return {
        "capacity_source": profile_hint,
        "requested_capacity_profile_key": requested_prof_key,
        "applied_capacity_profile_key": applied_profile_key,
        "capacity_profiles": capacity_profiles_out,
        "capacity_settings": cap["settings"],
        "team_metrics": cap["metrics"],
        "employee_count_profile": n_team,
        "selected_employee_count": k_sel,
        "assignee_filter_active": filter_lowers is not None,
        "selected_assignees": [display_by_lower.get(k) or team_display_by_lower.get(k) or k for k in active_for_rows],
        "team_capacity_hours": _round_hours(team_capacity_hours),
        "team_capacity_days": _round_hours(team_capacity_hours / HOURS_PER_DAY),
        "capacity_hours": capacity_hours,
        "capacity_days": _round_hours(capacity_hours / HOURS_PER_DAY),
        "leave_hours": leave_hours,
        "leave_days": _round_hours(leave_hours / HOURS_PER_DAY),
        "leave_aggregate_source": leave_aggregate_source,
        "availability_hours": availability_hours,
        "availability_days": _round_hours(availability_hours / HOURS_PER_DAY),
        "assignees": assignee_rows,
        "assignee_options": option_names,
        "employee_options": employee_options,
        "employee_tree": employee_tree,
        "meta": {
            "capacity_basis": cap_basis,
            "leave_basis": leave_basis_by_src.get(leave_aggregate_source, leave_basis_by_src["none"]),
            "leave_aggregate_source": leave_aggregate_source,
            "availability_formula": "capacity_hours - leave_hours; filtered mode uses (K/N)*team capacity",
        },
    }


def build_monthly_epic_plan_payload(
    db_path: Path,
    month: str,
    planner_rows: list[dict[str, Any]],
    canonical_run_id: str,
    *,
    selected_projects: set[str] | None = None,
    selected_assignees: set[str] | None = None,
    capacity_profile_key: str | None = None,
    jira_base_url: str = "",
    overdue_threshold_days: int = 30,
    include_on_hold: bool = False,
) -> dict[str, Any]:
    run_id = _to_text(canonical_run_id)
    if not run_id:
        raise ValueError("No successful canonical refresh found. Run the canonical refresh first.")

    month_start, month_end = _month_bounds(month)
    selected_project_keys = {
        _to_text(project_key).upper()
        for project_key in (selected_projects or set())
        if _to_text(project_key)
    }
    epic_meta, subtask_keys_by_epic = _load_canonical_issue_maps(db_path, run_id)
    actual_hours_by_epic, has_worklog_through_month_end = _load_worklog_metrics(
        db_path,
        run_id,
        subtask_keys_by_epic,
        month_start,
        month_end,
    )

    rows: list[dict[str, Any]] = []
    by_project: DefaultDict[str, dict[str, Any]] = defaultdict(
        lambda: {
            "project_key": "",
            "project_name": "",
            "epic_count": 0,
            "planned_hours": 0.0,
            "actual_hours": 0.0,
            "brought_forward_count": 0,
            "brought_forward_planned_hours": 0.0,
            "carried_forward_count": 0,
            "carried_forward_planned_hours": 0.0,
        }
    )
    totals = {
        "epic_count": 0,
        "planned_hours": 0.0,
        "actual_hours": 0.0,
        "start_slip_count": 0,
        "end_slip_count": 0,
        "brought_forward_count": 0,
        "brought_forward_planned_hours": 0.0,
        "carried_forward_count": 0,
        "carried_forward_planned_hours": 0.0,
    }
    jira_base = _to_text(jira_base_url).rstrip("/")
    overdue_cutoff = month_start - timedelta(days=max(0, int(overdue_threshold_days)))

    for planner_row in planner_rows:
        project_key = _to_text(planner_row.get("project_key")).upper()
        if selected_project_keys and project_key not in selected_project_keys:
            continue
        epic_key = _to_text(planner_row.get("epic_key")).upper()
        if not epic_key:
            continue
        epic_plan = ((planner_row.get("plans") or {}).get("epic_plan") or {})
        if not isinstance(epic_plan, dict):
            epic_plan = {}
        approved_start, approved_due, approved_start_text, approved_due_text = _approved_dates(epic_plan)
        if approved_start is None and approved_due is None:
            continue
        canonical_meta = epic_meta.get(epic_key) or {}
        jira_status = _to_text(canonical_meta.get("jira_status"))
        is_on_hold = _is_on_hold_status_text(jira_status)
        effectively_resolved = _is_resolved_status_text(jira_status) or (is_on_hold and not include_on_hold)
        brought_forward = bool(
            approved_due is not None
            and approved_due < month_start
            and approved_due >= overdue_cutoff
            and not effectively_resolved
        )
        start_in_month = bool(
            approved_start is not None and month_start <= approved_start <= month_end
        )
        due_in_month = bool(approved_due is not None and month_start <= approved_due <= month_end)
        if not (start_in_month or due_in_month) and not brought_forward:
            continue

        planned_hours = _hours_from_man_days(epic_plan.get("man_days")) or 0.0
        actual_hours = _round_hours(float(actual_hours_by_epic.get(epic_key) or 0.0))
        start_slip = bool(
            approved_start is not None
            and month_start <= approved_start <= month_end
            and not has_worklog_through_month_end.get(epic_key, False)
        )
        end_slip = bool(
            approved_due is not None
            and month_start <= approved_due <= month_end
            and not effectively_resolved
        )
        carried_forward = brought_forward or start_slip or end_slip
        jira_url = _to_text(planner_row.get("jira_url"))
        if not jira_url and jira_base:
            jira_url = f"{jira_base}/browse/{epic_key}"

        rows.append(
            {
                "epic_key": epic_key,
                "epic_name": _to_text(planner_row.get("epic_name")) or _to_text(canonical_meta.get("jira_summary")) or epic_key,
                "project_key": project_key,
                "project_name": _to_text(planner_row.get("project_name")) or project_key,
                "product_category": _to_text(planner_row.get("product_category")),
                "component": _to_text(planner_row.get("component")),
                "approved_start": approved_start_text,
                "approved_due": approved_due_text,
                "delivery_status": _delivery_status_view(planner_row.get("delivery_status"), jira_status),
                "jira_status": jira_status,
                "planned_hours": _round_hours(planned_hours),
                "planned_days": _round_hours(planned_hours / HOURS_PER_DAY),
                "actual_hours": actual_hours,
                "actual_days": _round_hours(actual_hours / HOURS_PER_DAY),
                "start_slip": start_slip,
                "end_slip": end_slip,
                "brought_forward": brought_forward,
                "carried_forward": carried_forward,
                "is_on_hold": is_on_hold,
                "jira_url": jira_url,
                "subtask_count": len(subtask_keys_by_epic.get(epic_key, set())),
            }
        )

        totals["epic_count"] += 1
        totals["planned_hours"] += planned_hours
        totals["actual_hours"] += actual_hours
        totals["start_slip_count"] += 1 if start_slip else 0
        totals["end_slip_count"] += 1 if end_slip else 0
        totals["brought_forward_count"] += 1 if brought_forward else 0
        totals["carried_forward_count"] += 1 if carried_forward else 0
        if brought_forward:
            totals["brought_forward_planned_hours"] = float(totals["brought_forward_planned_hours"]) + planned_hours
        if carried_forward:
            totals["carried_forward_planned_hours"] = float(totals["carried_forward_planned_hours"]) + planned_hours
        pj = by_project[project_key]
        pj["project_key"] = project_key
        pj["project_name"] = _to_text(planner_row.get("project_name")) or project_key
        pj["epic_count"] = int(pj["epic_count"]) + 1
        pj["planned_hours"] = float(pj["planned_hours"]) + planned_hours
        pj["actual_hours"] = float(pj["actual_hours"]) + actual_hours
        if brought_forward:
            pj["brought_forward_count"] = int(pj["brought_forward_count"]) + 1
            pj["brought_forward_planned_hours"] = float(pj["brought_forward_planned_hours"]) + planned_hours
        if carried_forward:
            pj["carried_forward_count"] = int(pj["carried_forward_count"]) + 1
            pj["carried_forward_planned_hours"] = float(pj["carried_forward_planned_hours"]) + planned_hours

    rows.sort(key=lambda item: (_to_text(item.get("project_name")).lower(), _to_text(item.get("approved_start")), _to_text(item.get("epic_key"))))
    by_project_rows: list[dict[str, Any]] = []
    for pk in sorted(by_project.keys()):
        agg = by_project[pk]
        planned = _round_hours(float(agg.get("planned_hours") or 0.0))
        actual = _round_hours(float(agg.get("actual_hours") or 0.0))
        by_project_rows.append(
            {
                "project_key": pk,
                "project_name": _to_text(agg.get("project_name")) or pk,
                "epic_count": int(agg.get("epic_count") or 0),
                "planned_hours": planned,
                "planned_days": _round_hours(planned / HOURS_PER_DAY),
                "actual_hours": actual,
                "actual_days": _round_hours(actual / HOURS_PER_DAY),
                "brought_forward_count": int(agg.get("brought_forward_count") or 0),
                "brought_forward_planned_hours": _round_hours(float(agg.get("brought_forward_planned_hours") or 0.0)),
                "brought_forward_planned_days": _round_hours(
                    float(agg.get("brought_forward_planned_hours") or 0.0) / HOURS_PER_DAY
                ),
                "carried_forward_count": int(agg.get("carried_forward_count") or 0),
                "carried_forward_planned_hours": _round_hours(float(agg.get("carried_forward_planned_hours") or 0.0)),
                "carried_forward_planned_days": _round_hours(
                    float(agg.get("carried_forward_planned_hours") or 0.0) / HOURS_PER_DAY
                ),
            }
        )
    rounded_totals = {
        **totals,
        "planned_hours": _round_hours(float(totals["planned_hours"])),
        "planned_days": _round_hours(float(totals["planned_hours"]) / HOURS_PER_DAY),
        "actual_hours": _round_hours(float(totals["actual_hours"])),
        "actual_days": _round_hours(float(totals["actual_hours"]) / HOURS_PER_DAY),
        "brought_forward_planned_hours": _round_hours(float(totals["brought_forward_planned_hours"])),
        "brought_forward_planned_days": _round_hours(float(totals["brought_forward_planned_hours"]) / HOURS_PER_DAY),
        "carried_forward_planned_hours": _round_hours(float(totals["carried_forward_planned_hours"])),
        "carried_forward_planned_days": _round_hours(float(totals["carried_forward_planned_hours"]) / HOURS_PER_DAY),
    }
    workforce = build_workforce_month_payload(
        db_path,
        month_start,
        month_end,
        run_id,
        selected_assignees=selected_assignees,
        capacity_profile_key=capacity_profile_key,
    )
    return {
        "month": month,
        "from_date": month_start.isoformat(),
        "to_date": month_end.isoformat(),
        "canonical_run_id": run_id,
        "selected_projects": sorted(selected_project_keys),
        "overdue_threshold_days": int(overdue_threshold_days),
        "include_on_hold": bool(include_on_hold),
        "rows": rows,
        "by_project": by_project_rows,
        "totals": rounded_totals,
        "workforce": workforce,
        "meta": {
            "hours_per_day": HOURS_PER_DAY,
            "scope_basis": (
                "Epics whose approved start date or approved due date falls inside the selected calendar month "
                "(TK planner dates), plus brought-forward unresolved epics whose approved due is before that month."
            ),
            "actual_hours_basis": "canonical_worklogs.started_date within the selected month",
            "status_basis": "Current Jira epic status from the latest canonical refresh",
        },
    }
