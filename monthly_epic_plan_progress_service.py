from __future__ import annotations

import calendar
import json
import re
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


def _is_bug_subtask_type(value: Any) -> bool:
    return _to_text(value).casefold() == "bug subtask"


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


def _extract_planner_issue_key(value: Any) -> str:
    text = _to_text(value)
    if not text:
        return ""
    if "/" in text:
        text = text.rsplit("/", 1)[-1]
    text = text.split("?", 1)[0].split("#", 1)[0].strip().upper()
    return text if re.match(r"^[A-Z0-9][A-Z0-9_-]*$", text) else ""


def _build_tk_planned_hours_by_issue(planner_rows: list[dict[str, Any]]) -> dict[str, float]:
    result: dict[str, float] = {}
    for row in planner_rows:
        epic_key = _to_text((row or {}).get("epic_key")).upper()
        plans = (row or {}).get("plans") or {}
        if not isinstance(plans, dict):
            continue
        for plan_key, plan_value in plans.items():
            if not isinstance(plan_value, dict):
                continue
            planned_hours = _hours_from_man_days(
                plan_value.get("tk_budgeted_man_days")
                if plan_value.get("tk_budgeted_man_days") not in (None, "")
                else plan_value.get("man_days")
            )
            if planned_hours is None:
                continue
            if plan_key == "epic_plan":
                if epic_key:
                    result[epic_key] = planned_hours
                continue
            linked_issue_key = _extract_planner_issue_key(plan_value.get("jira_url"))
            if linked_issue_key:
                result[linked_issue_key] = planned_hours
    return result


def _date_in_month(date_text: str, month_start: date, month_end: date) -> bool:
    if not date_text:
        return False
    try:
        d = date.fromisoformat(date_text[:10])
    except ValueError:
        return False
    return month_start <= d <= month_end


def _load_story_planned_hours(
    db_path: Path,
    run_id: str,
    epic_keys: list[str],
    month_start: date,
    month_end: date,
    *,
    include_bug_subtasks: bool = True,
) -> dict[str, float]:
    """
    Return planned hours per epic for the given month, computed from stories/subtasks:

    For each story under an in-scope epic:
    - If the story has subtasks with total original_estimate_hours > 0:
        add the sum of original_estimate_hours of those subtasks whose
        start_date or due_date falls in [month_start, month_end].
        (Do NOT add the parent story's own hours — avoids double-counting.)
    - Otherwise (no subtasks, or all subtasks have 0h):
        if the story's own start_date or due_date falls in [month_start, month_end],
        add the story's original_estimate_hours.
    """
    if not epic_keys or not run_id:
        return {}

    ms = month_start.isoformat()
    me = month_end.isoformat()
    placeholders = ",".join("?" * len(epic_keys))

    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        stories = conn.execute(
            f"""
            SELECT issue_key, epic_key, start_date, due_date, original_estimate_hours
            FROM canonical_issues
            WHERE run_id = ? AND epic_key IN ({placeholders}) AND issue_type = 'Story'
            """,
            (run_id, *epic_keys),
        ).fetchall()
        subtasks = conn.execute(
            f"""
            SELECT issue_key, story_key, epic_key, issue_type, start_date, due_date, original_estimate_hours
            FROM canonical_issues
            WHERE run_id = ? AND epic_key IN ({placeholders})
              AND issue_type IN ('Sub-task', 'Bug Subtask')
            """,
            (run_id, *epic_keys),
        ).fetchall()

    subs_by_story: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for sub in subtasks:
        if not include_bug_subtasks and _is_bug_subtask_type(sub["issue_type"]):
            continue
        sk = _to_text(sub["story_key"]).upper()
        if sk:
            subs_by_story[sk].append(dict(sub))

    planned_by_epic: dict[str, float] = defaultdict(float)
    for story in stories:
        story_key = _to_text(story["issue_key"]).upper()
        epic_key  = _to_text(story["epic_key"]).upper()
        story_subs = subs_by_story.get(story_key, [])
        sub_total  = sum(float(s["original_estimate_hours"] or 0) for s in story_subs)

        if story_subs and sub_total > 0:
            # Story has real subtask estimates — use only the subtask hours that fall in month
            for sub in story_subs:
                if _date_in_month(_to_text(sub["start_date"]), month_start, month_end) or \
                   _date_in_month(_to_text(sub["due_date"]),   month_start, month_end):
                    planned_by_epic[epic_key] += float(sub["original_estimate_hours"] or 0)
        else:
            # No subtasks (or all 0h) — use story's own hours if story falls in month
            if _date_in_month(_to_text(story["start_date"]), month_start, month_end) or \
               _date_in_month(_to_text(story["due_date"]),   month_start, month_end):
                planned_by_epic[epic_key] += float(story["original_estimate_hours"] or 0)

    return dict(planned_by_epic)


def _load_story_gantt_data(
    db_path: Path,
    run_id: str,
    epic_keys: list[str],
    *,
    include_bug_subtasks: bool = True,
) -> dict[str, list[dict[str, Any]]]:
    """
    Return per-epic list of story bars for the Gantt chart.
    Each entry: {issue_key, summary, start_date, due_date, has_hours}.
    has_hours=True when any worklog exists for that story or its subtasks.
    No month filter — covers the full epic scope.
    """
    if not epic_keys or not run_id:
        return {}

    placeholders = ",".join("?" * len(epic_keys))
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        stories = conn.execute(
            f"""
            SELECT issue_key, epic_key, summary, start_date, due_date
            FROM canonical_issues
            WHERE run_id = ? AND epic_key IN ({placeholders}) AND issue_type = 'Story'
            ORDER BY epic_key, start_date, issue_key
            """,
            (run_id, *epic_keys),
        ).fetchall()
        subtasks = conn.execute(
            f"""
            SELECT issue_key, story_key, issue_type
            FROM canonical_issues
            WHERE run_id = ? AND epic_key IN ({placeholders})
              AND issue_type IN ('Sub-task', 'Bug Subtask')
            """,
            (run_id, *epic_keys),
        ).fetchall()

    subtask_to_story: dict[str, str] = {}
    story_keys_set: set[str] = set()
    for sub in subtasks:
        if not include_bug_subtasks and _is_bug_subtask_type(sub["issue_type"]):
            continue
        sk = _to_text(sub["story_key"]).upper()
        ik = _to_text(sub["issue_key"]).upper()
        if sk and ik:
            subtask_to_story[ik] = sk
    for s in stories:
        story_keys_set.add(_to_text(s["issue_key"]).upper())

    all_issue_keys = list(story_keys_set | set(subtask_to_story.keys()))
    stories_with_hours: set[str] = set()
    worklog_dates_raw: DefaultDict[str, set] = defaultdict(set)
    if all_issue_keys:
        with sqlite3.connect(db_path) as conn:
            conn.row_factory = sqlite3.Row
            for chunk in _chunked(all_issue_keys, 400):
                chunk_ph = ",".join("?" for _ in chunk)
                wl_rows = conn.execute(
                    f"""
                    SELECT issue_key, started_date, hours_logged
                    FROM canonical_worklogs
                    WHERE run_id = ? AND issue_key IN ({chunk_ph})
                      AND hours_logged > 0
                    """,
                    [run_id, *chunk],
                ).fetchall()
                for wl in wl_rows:
                    ik = _to_text(wl["issue_key"]).upper()
                    story_key = subtask_to_story.get(ik, ik if ik in story_keys_set else None)
                    if story_key:
                        stories_with_hours.add(story_key)
                        sd = _to_text(wl["started_date"])
                        if sd:
                            worklog_dates_raw[story_key].add(sd[:10])

    worklog_dates_by_story: dict[str, list[str]] = {
        k: sorted(v) for k, v in worklog_dates_raw.items()
    }

    result: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for story in stories:
        story_key = _to_text(story["issue_key"]).upper()
        epic_key = _to_text(story["epic_key"]).upper()
        sd = _to_text(story["start_date"])
        dd = _to_text(story["due_date"])
        if not sd and not dd:
            continue
        result[epic_key].append({
            "issue_key": story_key,
            "summary": _to_text(story["summary"]),
            "start_date": sd,
            "due_date": dd,
            "has_hours": story_key in stories_with_hours,
            "worklog_dates": worklog_dates_by_story.get(story_key, []),
        })

    return dict(result)


def _load_total_planned_hours(
    db_path: Path,
    run_id: str,
    epic_keys: list[str],
    *,
    include_bug_subtasks: bool = True,
) -> dict[str, float]:
    """
    Return total planned hours per epic across ALL stories/subtasks (no month filter).
    Uses the same subtask-priority logic as _load_story_planned_hours but without
    any date scoping — this gives the full epic scope estimate.
    """
    if not epic_keys or not run_id:
        return {}

    placeholders = ",".join("?" * len(epic_keys))
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        stories = conn.execute(
            f"""
            SELECT issue_key, epic_key, original_estimate_hours
            FROM canonical_issues
            WHERE run_id = ? AND epic_key IN ({placeholders}) AND issue_type = 'Story'
            """,
            (run_id, *epic_keys),
        ).fetchall()
        subtasks = conn.execute(
            f"""
            SELECT story_key, epic_key, issue_type, original_estimate_hours
            FROM canonical_issues
            WHERE run_id = ? AND epic_key IN ({placeholders})
              AND issue_type IN ('Sub-task', 'Bug Subtask')
            """,
            (run_id, *epic_keys),
        ).fetchall()

    subs_by_story: dict[str, list[float]] = defaultdict(list)
    for sub in subtasks:
        if not include_bug_subtasks and _is_bug_subtask_type(sub["issue_type"]):
            continue
        sk = _to_text(sub["story_key"]).upper()
        if sk:
            subs_by_story[sk].append(float(sub["original_estimate_hours"] or 0))

    total_by_epic: dict[str, float] = defaultdict(float)
    for story in stories:
        story_key = _to_text(story["issue_key"]).upper()
        epic_key  = _to_text(story["epic_key"]).upper()
        sub_hours = subs_by_story.get(story_key, [])
        sub_total = sum(sub_hours)
        if sub_hours and sub_total > 0:
            total_by_epic[epic_key] += sub_total
        else:
            total_by_epic[epic_key] += float(story["original_estimate_hours"] or 0)

    return dict(total_by_epic)


def _load_estimate_rollup_for_epics(
    db_path: Path,
    run_id: str,
    epic_keys: list[str],
    month_start: date,
    month_end: date,
    jira_base: str = "",
    tk_planned_hours_by_issue: dict[str, float] | None = None,
    include_bug_subtasks: bool = True,
) -> dict[str, Any]:
    """
    Build month-scoped estimate hierarchy totals for the report's planned-this-month epic set.

    The rollup intentionally keeps direct Epic, Story, and Sub-task original estimates
    separate so the report can show how each Jira hierarchy level compares.
    """
    selected_epics = sorted({_to_text(key).upper() for key in epic_keys if _to_text(key)})
    if not selected_epics or not run_id:
        return _empty_estimate_rollup(0)
    if tk_planned_hours_by_issue is None:
        tk_planned_hours_by_issue = {}

    epic_original_hours = 0.0
    story_original_hours = 0.0
    subtask_original_hours = 0.0
    subtask_logged_hours = 0.0

    all_story_original_by_key: dict[str, float] = {}
    story_original_by_key: dict[str, float] = {}
    story_to_epic: dict[str, str] = {}
    story_rows_by_key: dict[str, dict[str, Any]] = {}
    subtask_rows_by_key: dict[str, dict[str, Any]] = {}
    month_subtask_keys: set[str] = set()
    jira_base_clean = _to_text(jira_base).rstrip("/")

    def issue_url(issue_key: str) -> str:
        return f"{jira_base_clean}/browse/{issue_key}" if jira_base_clean and issue_key else ""

    def parent_link(issue_key: str, item_type: str, summary: Any = "") -> dict[str, Any]:
        key = _to_text(issue_key).upper()
        return {
            "issue_key": key,
            "item_type": item_type,
            "summary": _to_text(summary),
            "jira_url": issue_url(key),
        }

    def detail_row(
        row: dict[str, Any],
        item_type: str,
        *,
        original_estimate_hours: float | None = None,
        logged_hours: float | None = None,
        overrun_hours: float | None = None,
        non_bug_logged_hours: float | None = None,
        non_bug_overrun_hours: float | None = None,
        bug_logged_hours: float | None = None,
        parents: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        issue_key = _to_text(row.get("issue_key")).upper()
        original = None if original_estimate_hours is None else _round_hours(original_estimate_hours)
        logged = None if logged_hours is None else _round_hours(logged_hours)
        overrun = None if overrun_hours is None else _round_hours(overrun_hours)
        non_bug_logged = None if non_bug_logged_hours is None else _round_hours(non_bug_logged_hours)
        non_bug_overrun = None if non_bug_overrun_hours is None else _round_hours(non_bug_overrun_hours)
        bug_logged = None if bug_logged_hours is None else _round_hours(bug_logged_hours)
        tk_planned = None
        if issue_key and issue_key in tk_planned_hours_by_issue:
            tk_planned = _round_hours(float(tk_planned_hours_by_issue[issue_key] or 0.0))
        result = {
            "issue_key": issue_key,
            "item_type": item_type,
            "summary": _to_text(row.get("summary")) or issue_key,
            "start_date": _to_text(row.get("start_date")),
            "due_date": _to_text(row.get("due_date")),
            "original_estimate_hours": original,
            "original_estimate_days": None if original is None else _round_hours(original / HOURS_PER_DAY),
            "tk_planned_hours": tk_planned,
            "tk_planned_days": None if tk_planned is None else _round_hours(tk_planned / HOURS_PER_DAY),
            "logged_hours": logged,
            "logged_days": None if logged is None else _round_hours(logged / HOURS_PER_DAY),
            "overrun_hours": overrun,
            "overrun_days": None if overrun is None else _round_hours(overrun / HOURS_PER_DAY),
            "non_bug_logged_hours": non_bug_logged,
            "non_bug_logged_days": None if non_bug_logged is None else _round_hours(non_bug_logged / HOURS_PER_DAY),
            "non_bug_overrun_hours": non_bug_overrun,
            "non_bug_overrun_days": None if non_bug_overrun is None else _round_hours(non_bug_overrun / HOURS_PER_DAY),
            "bug_logged_hours": bug_logged,
            "bug_logged_days": None if bug_logged is None else _round_hours(bug_logged / HOURS_PER_DAY),
            "jira_url": issue_url(issue_key),
            "parents": parents or [],
        }
        return result

    def add_detail(epic_key: str, metric: str, row: dict[str, Any]) -> None:
        metric_map = by_epic.get(epic_key, {}).get("details_by_metric")
        if isinstance(metric_map, dict):
            metric_map.setdefault(metric, []).append(row)

    by_epic: dict[str, dict[str, Any]] = {
        key: {
            "epic_count": 1,
            "story_count": 0,
            "subtask_count": 0,
            "overrun_story_count": 0,
            "epic_original_estimate_hours": 0.0,
            "story_original_estimate_hours": 0.0,
            "subtask_original_estimate_hours": 0.0,
            "subtask_logged_hours": 0.0,
            "story_subtask_logged_over_parent_estimate_hours": 0.0,
            "overrun_story_original_estimate_hours": 0.0,
            "story_subtask_logged_over_parent_estimate_pct": 0.0,
            "details_by_metric": {
                "epic_original": [],
                "story_original": [],
                "subtask_original": [],
                "subtask_logged": [],
                "overrun": [],
            },
        }
        for key in selected_epics
    }

    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        for chunk in _chunked(selected_epics, 400):
            placeholders = ",".join("?" for _ in chunk)
            epic_rows = conn.execute(
                f"""
                SELECT issue_key, summary, start_date, due_date, original_estimate_hours
                FROM canonical_issues
                WHERE run_id = ? AND issue_type = 'Epic'
                  AND UPPER(issue_key) IN ({placeholders})
                """,
                [run_id, *chunk],
            ).fetchall()
            for row in epic_rows:
                epic_key = _to_text(row["issue_key"]).upper()
                estimate = float(row["original_estimate_hours"] or 0.0)
                epic_original_hours += estimate
                if epic_key in by_epic:
                    by_epic[epic_key]["epic_original_estimate_hours"] = (
                        float(by_epic[epic_key]["epic_original_estimate_hours"]) + estimate
                    )
                    add_detail(
                        epic_key,
                        "epic_original",
                        detail_row(dict(row), "Epic", original_estimate_hours=estimate),
                    )

            story_rows = conn.execute(
                f"""
                SELECT issue_key, epic_key, parent_issue_key, summary, start_date, due_date, original_estimate_hours
                FROM canonical_issues
                WHERE run_id = ? AND issue_type = 'Story'
                  AND (UPPER(epic_key) IN ({placeholders}) OR UPPER(parent_issue_key) IN ({placeholders}))
                """,
                [run_id, *chunk, *chunk],
            ).fetchall()
            for row in story_rows:
                story_key = _to_text(row["issue_key"]).upper()
                if not story_key:
                    continue
                epic_key = _to_text(row["epic_key"]).upper() or _to_text(row["parent_issue_key"]).upper()
                if epic_key not in selected_epics:
                    continue
                estimate = float(row["original_estimate_hours"] or 0.0)
                all_story_original_by_key[story_key] = estimate
                story_to_epic[story_key] = epic_key
                story_rows_by_key[story_key] = dict(row)
                if _date_in_month(_to_text(row["start_date"]), month_start, month_end) or _date_in_month(_to_text(row["due_date"]), month_start, month_end):
                    story_original_by_key[story_key] = estimate

            subtask_rows = conn.execute(
                f"""
                                SELECT issue_key, epic_key, parent_issue_key, story_key, issue_type, summary, start_date, due_date, original_estimate_hours
                FROM canonical_issues
                WHERE run_id = ? AND issue_type IN ('Sub-task', 'Bug Subtask')
                  AND UPPER(epic_key) IN ({placeholders})
                """,
                [run_id, *chunk],
            ).fetchall()
            for row in subtask_rows:
                if not include_bug_subtasks and _is_bug_subtask_type(row["issue_type"]):
                    continue
                subtask_key = _to_text(row["issue_key"]).upper()
                if subtask_key:
                    subtask_rows_by_key[subtask_key] = dict(row)
                    if _date_in_month(_to_text(row["start_date"]), month_start, month_end) or _date_in_month(_to_text(row["due_date"]), month_start, month_end):
                        month_subtask_keys.add(subtask_key)
                        story_key = _to_text(row["story_key"]).upper() or _to_text(row["parent_issue_key"]).upper()
                        if story_key and story_key in all_story_original_by_key:
                            story_original_by_key[story_key] = all_story_original_by_key[story_key]

        story_keys = sorted(all_story_original_by_key.keys())
        for chunk in _chunked(story_keys, 400):
            placeholders = ",".join("?" for _ in chunk)
            subtask_rows = conn.execute(
                f"""
                                SELECT issue_key, epic_key, parent_issue_key, story_key, issue_type, summary, start_date, due_date, original_estimate_hours
                FROM canonical_issues
                WHERE run_id = ? AND issue_type IN ('Sub-task', 'Bug Subtask')
                  AND (UPPER(story_key) IN ({placeholders}) OR UPPER(parent_issue_key) IN ({placeholders}))
                """,
                [run_id, *chunk, *chunk],
            ).fetchall()
            for row in subtask_rows:
                if not include_bug_subtasks and _is_bug_subtask_type(row["issue_type"]):
                    continue
                subtask_key = _to_text(row["issue_key"]).upper()
                if subtask_key:
                    subtask_rows_by_key[subtask_key] = dict(row)
                    if _date_in_month(_to_text(row["start_date"]), month_start, month_end) or _date_in_month(_to_text(row["due_date"]), month_start, month_end):
                        month_subtask_keys.add(subtask_key)
                        story_key = _to_text(row["story_key"]).upper() or _to_text(row["parent_issue_key"]).upper()
                        if story_key and story_key in all_story_original_by_key:
                            story_original_by_key[story_key] = all_story_original_by_key[story_key]

        for story_key, estimate in story_original_by_key.items():
            epic_key = story_to_epic.get(story_key)
            if epic_key not in selected_epics:
                continue
            story_original_hours += estimate
            by_epic[epic_key]["story_count"] = int(by_epic[epic_key]["story_count"]) + 1
            by_epic[epic_key]["story_original_estimate_hours"] = (
                float(by_epic[epic_key]["story_original_estimate_hours"]) + estimate
            )
            story_row = story_rows_by_key.get(story_key, {"issue_key": story_key})
            add_detail(
                epic_key,
                "story_original",
                detail_row(
                    story_row,
                    "Story",
                    original_estimate_hours=estimate,
                    parents=[parent_link(epic_key, "Epic")],
                ),
            )

        subtask_to_story: dict[str, str] = {}
        subtask_to_epic: dict[str, str] = {}
        subtask_logged_by_key: dict[str, float] = defaultdict(float)
        story_subtask_logged: dict[str, float] = defaultdict(float)
        story_non_bug_subtask_logged: dict[str, float] = defaultdict(float)
        story_bug_subtask_logged: dict[str, float] = defaultdict(float)
        for row in subtask_rows_by_key.values():
            subtask_key = _to_text(row.get("issue_key")).upper()
            if subtask_key not in month_subtask_keys:
                continue
            epic_key = _to_text(row.get("epic_key")).upper()
            story_key = _to_text(row.get("story_key")).upper() or _to_text(row.get("parent_issue_key")).upper()
            if story_key and story_key in story_to_epic:
                epic_key = story_to_epic[story_key]
            if epic_key not in selected_epics:
                continue
            estimate = float(row.get("original_estimate_hours") or 0.0)
            subtask_original_hours += estimate
            by_epic[epic_key]["subtask_count"] = int(by_epic[epic_key]["subtask_count"]) + 1
            by_epic[epic_key]["subtask_original_estimate_hours"] = (
                float(by_epic[epic_key]["subtask_original_estimate_hours"]) + estimate
            )
            parent_story = story_rows_by_key.get(story_key, {"issue_key": story_key})
            add_detail(
                epic_key,
                "subtask_original",
                detail_row(
                    row,
                    _to_text(row.get("issue_type")) or "Sub-task",
                    original_estimate_hours=estimate,
                    parents=[
                        parent_link(story_key, "Story", parent_story.get("summary")),
                        parent_link(epic_key, "Epic"),
                    ],
                ),
            )
            if subtask_key:
                subtask_to_epic[subtask_key] = epic_key
            if subtask_key and story_key:
                subtask_to_story[subtask_key] = story_key

        # Extend worklog scope to subtasks whose dates are outside the month.
        # The estimate pass above (month_subtask_keys) is intentionally date-scoped to
        # only reflect work planned for this month in the OE numbers. But for logged hours
        # we want every subtask that belongs to the in-scope epics — the same population
        # used by _load_worklog_metrics — so "Subtask Logged" matches "Epic Work Summary
        # actual hours" instead of silently dropping worklogs on out-of-month subtasks.
        for row in subtask_rows_by_key.values():
            subtask_key = _to_text(row.get("issue_key")).upper()
            if not subtask_key or subtask_key in subtask_to_story:
                continue
            epic_key = _to_text(row.get("epic_key")).upper()
            story_key = _to_text(row.get("story_key")).upper() or _to_text(row.get("parent_issue_key")).upper()
            if story_key and story_key in story_to_epic:
                epic_key = story_to_epic[story_key]
            if epic_key not in selected_epics:
                continue
            if subtask_key and epic_key:
                subtask_to_epic[subtask_key] = epic_key
            if subtask_key and story_key:
                subtask_to_story[subtask_key] = story_key

        subtask_keys = sorted(subtask_to_story.keys())
        if subtask_keys:
            wl_latest = conn.execute(
                "SELECT run_id FROM canonical_worklogs ORDER BY rowid DESC LIMIT 1"
            ).fetchone()
            wl_run_id = wl_latest[0] if wl_latest else run_id
            today = date.today().isoformat()
            for chunk in _chunked(subtask_keys, 400):
                placeholders = ",".join("?" for _ in chunk)
                worklog_rows = conn.execute(
                    f"""
                    SELECT issue_key, hours_logged, started_date
                    FROM canonical_worklogs
                    WHERE run_id = ?
                      AND issue_key IN ({placeholders})
                      AND started_date >= ?
                      AND started_date <= ?
                    """,
                    [wl_run_id, *chunk, month_start.isoformat(), min(month_end.isoformat(), today)],
                ).fetchall()
                for row in worklog_rows:
                    subtask_key = _to_text(row["issue_key"]).upper()
                    story_key = subtask_to_story.get(subtask_key)
                    hours = float(row["hours_logged"] or 0.0)
                    if not story_key or hours <= 0:
                        continue
                    epic_key = subtask_to_epic.get(subtask_key)
                    subtask_issue = subtask_rows_by_key.get(subtask_key) or {}
                    issue_type = _to_text(subtask_issue.get("issue_type"))
                    is_bug_subtask = issue_type.casefold() == "bug subtask"
                    subtask_logged_hours += hours
                    subtask_logged_by_key[subtask_key] += hours
                    story_subtask_logged[story_key] += hours
                    if is_bug_subtask:
                        story_bug_subtask_logged[story_key] += hours
                    else:
                        story_non_bug_subtask_logged[story_key] += hours
                    if epic_key in by_epic:
                        by_epic[epic_key]["subtask_logged_hours"] = (
                            float(by_epic[epic_key]["subtask_logged_hours"]) + hours
                        )

        for subtask_key, logged in subtask_logged_by_key.items():
            row = subtask_rows_by_key.get(subtask_key)
            story_key = subtask_to_story.get(subtask_key, "")
            epic_key = subtask_to_epic.get(subtask_key, "")
            if not row or epic_key not in selected_epics:
                continue
            parent_story = story_rows_by_key.get(story_key, {"issue_key": story_key})
            add_detail(
                epic_key,
                "subtask_logged",
                detail_row(
                    row,
                    _to_text(row.get("issue_type")) or "Sub-task",
                    original_estimate_hours=float(row.get("original_estimate_hours") or 0.0),
                    logged_hours=logged,
                    parents=[
                        parent_link(story_key, "Story", parent_story.get("summary")),
                        parent_link(epic_key, "Epic"),
                    ],
                ),
            )

    overrun_hours = 0.0
    overrun_story_estimate_hours = 0.0
    overrun_story_count = 0
    for story_key, logged in story_subtask_logged.items():
        parent_estimate = float(story_original_by_key.get(story_key) or 0.0)
        if logged > parent_estimate:
            epic_key = story_to_epic.get(story_key)
            non_bug_logged = float(story_non_bug_subtask_logged.get(story_key) or 0.0)
            bug_logged = float(story_bug_subtask_logged.get(story_key) or 0.0)
            overrun_value = logged - parent_estimate
            non_bug_overrun_value = max(non_bug_logged - parent_estimate, 0.0)
            overrun_hours += logged - parent_estimate
            overrun_story_estimate_hours += parent_estimate
            overrun_story_count += 1
            if epic_key in by_epic:
                by_epic[epic_key]["story_subtask_logged_over_parent_estimate_hours"] = (
                    float(by_epic[epic_key]["story_subtask_logged_over_parent_estimate_hours"]) + logged - parent_estimate
                )
                by_epic[epic_key]["overrun_story_original_estimate_hours"] = (
                    float(by_epic[epic_key]["overrun_story_original_estimate_hours"]) + parent_estimate
                )
                by_epic[epic_key]["overrun_story_count"] = int(by_epic[epic_key]["overrun_story_count"]) + 1
                story_row = story_rows_by_key.get(story_key, {"issue_key": story_key})
                add_detail(
                    epic_key,
                    "overrun",
                    detail_row(
                        story_row,
                        "Story",
                        original_estimate_hours=parent_estimate,
                        logged_hours=logged,
                        overrun_hours=overrun_value,
                        non_bug_logged_hours=non_bug_logged,
                        non_bug_overrun_hours=non_bug_overrun_value,
                        bug_logged_hours=bug_logged,
                        parents=[parent_link(epic_key, "Epic")],
                    ),
                )

    if overrun_story_estimate_hours > 0:
        overrun_pct = overrun_hours / overrun_story_estimate_hours * 100.0
    elif overrun_hours > 0:
        overrun_pct = 100.0
    else:
        overrun_pct = 0.0

    for epic_rollup in by_epic.values():
        baseline = float(epic_rollup.get("overrun_story_original_estimate_hours") or 0.0)
        overrun = float(epic_rollup.get("story_subtask_logged_over_parent_estimate_hours") or 0.0)
        if baseline > 0:
            epic_rollup["story_subtask_logged_over_parent_estimate_pct"] = overrun / baseline * 100.0
        elif overrun > 0:
            epic_rollup["story_subtask_logged_over_parent_estimate_pct"] = 100.0

    rollup = _round_estimate_rollup({
        "epic_count": len(selected_epics),
        "story_count": len(story_original_by_key),
        "subtask_count": sum(int(item.get("subtask_count") or 0) for item in by_epic.values()),
        "overrun_story_count": overrun_story_count,
        "epic_original_estimate_hours": epic_original_hours,
        "story_original_estimate_hours": story_original_hours,
        "subtask_original_estimate_hours": subtask_original_hours,
        "subtask_logged_hours": subtask_logged_hours,
        "story_subtask_logged_over_parent_estimate_hours": overrun_hours,
        "overrun_story_original_estimate_hours": overrun_story_estimate_hours,
        "story_subtask_logged_over_parent_estimate_pct": overrun_pct,
    })
    rollup["by_epic"] = {
        epic_key: _round_estimate_rollup(epic_rollup)
        for epic_key, epic_rollup in by_epic.items()
    }
    return rollup


def _empty_estimate_rollup(epic_count: int = 0) -> dict[str, Any]:
    rollup = _round_estimate_rollup({
        "epic_count": epic_count,
        "story_count": 0,
        "subtask_count": 0,
        "overrun_story_count": 0,
        "epic_original_estimate_hours": 0.0,
        "story_original_estimate_hours": 0.0,
        "subtask_original_estimate_hours": 0.0,
        "subtask_logged_hours": 0.0,
        "story_subtask_logged_over_parent_estimate_hours": 0.0,
        "overrun_story_original_estimate_hours": 0.0,
        "story_subtask_logged_over_parent_estimate_pct": 0.0,
    })
    rollup["by_epic"] = {}
    rollup["details_by_metric"] = {
        "epic_original": [],
        "story_original": [],
        "subtask_original": [],
        "subtask_logged": [],
        "overrun": [],
    }
    return rollup


def _round_estimate_rollup(rollup: dict[str, Any]) -> dict[str, Any]:
    hour_fields = [
        "epic_original_estimate_hours",
        "story_original_estimate_hours",
        "subtask_original_estimate_hours",
        "subtask_logged_hours",
        "story_subtask_logged_over_parent_estimate_hours",
        "overrun_story_original_estimate_hours",
    ]
    rounded = dict(rollup)
    for field in hour_fields:
        value = _round_hours(float(rounded.get(field) or 0.0))
        rounded[field] = value
        rounded[field.replace("_hours", "_days")] = _round_hours(value / HOURS_PER_DAY)
    rounded["story_subtask_logged_over_parent_estimate_pct"] = _round_hours(
        float(rounded.get("story_subtask_logged_over_parent_estimate_pct") or 0.0)
    )
    rounded["epic_count"] = int(rounded.get("epic_count") or 0)
    rounded["story_count"] = int(rounded.get("story_count") or 0)
    rounded["subtask_count"] = int(rounded.get("subtask_count") or 0)
    rounded["overrun_story_count"] = int(rounded.get("overrun_story_count") or 0)
    return rounded


def _load_child_items_for_month(
    db_path: Path,
    run_id: str,
    epic_keys: list[str],
    month_start: date,
    month_end: date,
    jira_base: str = "",
    *,
    include_bug_subtasks: bool = True,
) -> dict[str, list[dict[str, Any]]]:
    """
    Return per-epic list of story/subtask items that drove the epic's appearance in the
    selected month. Stories with real subtask estimates surface their subtasks;
    stories without subtasks appear directly. All items include in_month flags and worklogs.
    """
    if not epic_keys or not run_id:
        return {}

    placeholders = ",".join("?" * len(epic_keys))
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        stories = conn.execute(
            f"""
            SELECT issue_key, epic_key, summary, status, start_date, due_date, original_estimate_hours,
                   resolved_stable_since_date
            FROM canonical_issues
            WHERE run_id = ? AND epic_key IN ({placeholders}) AND issue_type = 'Story'
            ORDER BY issue_key
            """,
            (run_id, *epic_keys),
        ).fetchall()
        subtasks = conn.execute(
            f"""
            SELECT issue_key, story_key, epic_key, issue_type, summary, status, start_date, due_date, original_estimate_hours,
                   resolved_stable_since_date
            FROM canonical_issues
            WHERE run_id = ? AND epic_key IN ({placeholders})
              AND issue_type IN ('Sub-task', 'Bug Subtask')
            ORDER BY story_key, issue_key
            """,
            (run_id, *epic_keys),
        ).fetchall()

        # Load worklogs for all issues in these epics (both stories and subtasks)
        # Pick the most-recently-inserted run from canonical_worklogs. Sorting by run_id
        # lexicographically is wrong: prefixes like "canonical-epic-*" / "canonical-assignee-*"
        # sort above "canonical-<ts>-*" even when the latter is the newer full refresh, which
        # caused the report to read worklogs from a partial/older run that doesn't contain the
        # subtask's logged hours. rowid grows monotonically with insertion order.
        try:
            wl_latest_run = conn.execute(
                "SELECT run_id FROM canonical_worklogs ORDER BY rowid DESC LIMIT 1"
            ).fetchone()
            wl_run_id = wl_latest_run[0] if wl_latest_run else run_id
            worklogs_raw = conn.execute(
                f"""
                SELECT issue_key, worklog_author, started_date, hours_logged
                FROM canonical_worklogs
                WHERE run_id = ?
                  AND started_date >= ? AND started_date <= ?
                  AND issue_key IN (
                      SELECT issue_key FROM canonical_issues
                      WHERE run_id = ? AND epic_key IN ({placeholders})
                        AND issue_type IN ('Story', 'Sub-task', 'Bug Subtask')
                  )
                ORDER BY issue_key, started_date
                """,
                (wl_run_id, month_start.isoformat(), month_end.isoformat(), run_id, *epic_keys),
            ).fetchall()
        except Exception:
            worklogs_raw = []

    # Group worklogs by issue_key
    worklogs_by_issue: dict[str, list[dict]] = defaultdict(list)
    for wl in worklogs_raw:
        ik = _to_text(wl["issue_key"]).upper()
        worklogs_by_issue[ik].append({
            "date": _to_text(wl["started_date"]),
            "author": _to_text(wl["worklog_author"]),
            "hours": _round_hours(float(wl["hours_logged"] or 0)),
        })

    subs_by_story: dict[str, list[dict]] = defaultdict(list)
    for sub in subtasks:
        if not include_bug_subtasks and _is_bug_subtask_type(sub["issue_type"]):
            continue
        sk = _to_text(sub["story_key"]).upper()
        if sk:
            subs_by_story[sk].append(dict(sub))

    result: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for story in stories:
        story_key = _to_text(story["issue_key"]).upper()
        epic_key  = _to_text(story["epic_key"]).upper()
        story_subs = subs_by_story.get(story_key, [])
        sub_total  = sum(float(s["original_estimate_hours"] or 0) for s in story_subs)
        story_url  = f"{jira_base}/browse/{story_key}" if jira_base and story_key else ""

        if story_subs and sub_total > 0:
            for sub in story_subs:
                sub_key = _to_text(sub["issue_key"]).upper()
                sd = _to_text(sub["start_date"])
                dd = _to_text(sub["due_date"])
                sim = _date_in_month(sd, month_start, month_end)
                dim = _date_in_month(dd, month_start, month_end)
                result[epic_key].append({
                    "issue_key": sub_key,
                    "parent_story_key": story_key,
                    "parent_story_summary": _to_text(story["summary"]),
                    "summary": _to_text(sub["summary"]),
                    "issue_type": _to_text(sub["issue_type"]) or "Sub-task",
                    "status": _to_text(sub["status"]),
                    "start_date": sd,
                    "due_date": dd,
                    "actual_completed_date": _to_text(sub["resolved_stable_since_date"]),
                    "planned_hours": _round_hours(float(sub["original_estimate_hours"] or 0)),
                    "start_in_month": sim,
                    "due_in_month": dim,
                    "in_month": sim or dim,
                    "jira_url": f"{jira_base}/browse/{sub_key}" if jira_base and sub_key else "",
                    "story_jira_url": story_url,
                    "worklogs": worklogs_by_issue.get(sub_key, []),
                })
        else:
            sd = _to_text(story["start_date"])
            dd = _to_text(story["due_date"])
            sim = _date_in_month(sd, month_start, month_end)
            dim = _date_in_month(dd, month_start, month_end)
            result[epic_key].append({
                "issue_key": story_key,
                "parent_story_key": "",
                "parent_story_summary": "",
                "summary": _to_text(story["summary"]),
                "issue_type": "Story",
                "status": _to_text(story["status"]),
                "start_date": sd,
                "due_date": dd,
                "actual_completed_date": _to_text(story["resolved_stable_since_date"]),
                "planned_hours": _round_hours(float(story["original_estimate_hours"] or 0)),
                "start_in_month": sim,
                "due_in_month": dim,
                "in_month": sim or dim,
                "jira_url": story_url,
                "story_jira_url": "",
                "worklogs": worklogs_by_issue.get(story_key, []),
            })

    return dict(result)


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




def _chunked(items: list[str], chunk_size: int) -> list[list[str]]:
    return [items[index : index + chunk_size] for index in range(0, len(items), chunk_size)]


def _load_canonical_issue_maps(
    db_path: Path,
    run_id: str,
    *,
    include_bug_subtasks: bool = True,
) -> tuple[dict[str, dict[str, Any]], dict[str, set[str]]]:
    epic_meta: dict[str, dict[str, Any]] = {}
    story_to_epic: dict[str, str] = {}
    subtask_rows: list[dict[str, Any]] = []
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT issue_key, project_key, issue_type, summary, status, parent_issue_key, story_key, epic_key,
                   resolved_stable_since_date
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
                "resolved_stable_since_date": _to_text(row["resolved_stable_since_date"]),
            }
        elif _is_story_type(issue_type):
            resolved_epic = epic_key or parent_key
            if issue_key and resolved_epic:
                story_to_epic[issue_key] = resolved_epic
        elif _is_subtask_type(issue_type):
            if not include_bug_subtasks and _is_bug_subtask_type(issue_type):
                continue
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
) -> tuple[dict[str, float], dict[str, bool], dict[str, date], dict[str, list[str]], dict[str, float]]:
    actual_hours_by_epic: dict[str, float] = defaultdict(float)
    total_actual_hours_by_epic: dict[str, float] = defaultdict(float)
    has_worklog_through_month_end: dict[str, bool] = defaultdict(bool)
    last_worklog_date_by_epic: dict[str, date] = {}
    worklog_dates_by_epic: dict[str, set[str]] = defaultdict(set)

    epic_keys = sorted(subtask_keys_by_epic.keys())
    if not epic_keys:
        return actual_hours_by_epic, has_worklog_through_month_end, last_worklog_date_by_epic, {}, total_actual_hours_by_epic

    # Use the pre-computed subtask→epic mapping (subtask_keys_by_epic). It already resolves
    # an epic for every subtask via the chain epic_key → story_key → parent_issue_key, so it
    # captures subtasks whose canonical_issues row has a NULL/blank epic_key but whose parent
    # story belongs to a known epic. Re-querying canonical_issues with `epic_key IN (...)` alone
    # would silently drop those subtasks and their worklogs from the epic's actual hours.
    issue_to_epic: dict[str, str] = {}
    for ek, sk_set in subtask_keys_by_epic.items():
        ek_up = _to_text(ek).upper()
        if not ek_up:
            continue
        for sk in sk_set:
            sk_up = _to_text(sk).upper()
            if sk_up:
                issue_to_epic[sk_up] = ek_up
    today = date.today()

    issue_keys = sorted(issue_to_epic)
    if not issue_keys:
        return actual_hours_by_epic, has_worklog_through_month_end, last_worklog_date_by_epic, {}, total_actual_hours_by_epic

    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        # Use latest worklog run by insertion order (rowid). The report's canonical issues
        # run_id may belong to an epic-only or older partial refresh whose canonical_worklogs
        # rows don't cover all subtasks; reading from the newest worklog run avoids missing
        # worklogs that were captured by a later full refresh.
        wl_latest = conn.execute(
            "SELECT run_id FROM canonical_worklogs ORDER BY rowid DESC LIMIT 1"
        ).fetchone()
        wl_run_id = wl_latest[0] if wl_latest else run_id
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
                [wl_run_id, *chunk, today.isoformat()],
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
                if started <= month_end:
                    has_worklog_through_month_end[epic_key] = True
                if month_start <= started <= month_end:
                    actual_hours_by_epic[epic_key] += hours
                # All-time total (any date up to today)
                total_actual_hours_by_epic[epic_key] += hours
                # Track the most recent worklog date across all time (up to today).
                prev = last_worklog_date_by_epic.get(epic_key)
                if prev is None or started > prev:
                    last_worklog_date_by_epic[epic_key] = started
                # Collect every unique work date for the Gantt worklog bar.
                worklog_dates_by_epic[epic_key].add(row["started_date"][:10])

    worklog_dates_sorted: dict[str, list[str]] = {
        k: sorted(v) for k, v in worklog_dates_by_epic.items()
    }
    return actual_hours_by_epic, has_worklog_through_month_end, last_worklog_date_by_epic, worklog_dates_sorted, total_actual_hours_by_epic


def _approved_dates(
    epic_plan: dict[str, Any],
    planner_row: dict[str, Any] | None = None,
) -> tuple[date | None, date | None, str, str]:
    # Priority: planner row top-level dates (from import) → epic_plan TK-budgeted → epic_plan dates
    # planner_row.start_date/due_date are the import-sourced dates shown in Epics Planner UI.
    # epic_plan.start_date/due_date are TK budgeting sub-object dates and should only be fallbacks.
    start_text = (
        _to_text((planner_row or {}).get("start_date"))
        or _to_text(epic_plan.get("tk_budgeted_start_date"))
        or _to_text(epic_plan.get("start_date"))
    )
    due_text = (
        _to_text((planner_row or {}).get("due_date"))
        or _to_text(epic_plan.get("tk_budgeted_due_date"))
        or _to_text(epic_plan.get("due_date"))
    )
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
) -> tuple[dict[str, str], dict[str, float], dict[str, float], str]:
    """
    Returns (display_by_lower, planned_by_lower, unplanned_by_lower, source).
    planned_by_lower is used for availability (reduces capacity).
    unplanned_by_lower is shown in the UI but does NOT reduce capacity.
    "Unknown" classification is treated as planned (conservative).
    """
    display_by_lower: dict[str, str] = {}
    planned_by_lower: DefaultDict[str, float] = defaultdict(float)
    unplanned_by_lower: DefaultDict[str, float] = defaultdict(float)

    def track(name_raw: Any, hours: float, classification: str = "Planned") -> None:
        raw_name = _to_text(name_raw)
        if not raw_name or hours <= 0:
            return
        key = raw_name.lower()
        display_by_lower.setdefault(key, raw_name)
        if classification == "Unplanned":
            unplanned_by_lower[key] += hours
        else:
            planned_by_lower[key] += hours

    def _round_both() -> None:
        for d in (planned_by_lower, unplanned_by_lower):
            for lk in list(d.keys()):
                d[lk] = _round_hours(float(d[lk]))

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
        clf = _to_text(row.get("leave_classification"))
        track(row.get("assignee"), planned_h, clf)

    if dist_used:
        _round_both()
        return display_by_lower, dict(planned_by_lower), dict(unplanned_by_lower), "distributed_subtasks"

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
        track(row.get("assignee"), pt + pn, "Planned")

    if embedded_used:
        _round_both()
        return display_by_lower, dict(planned_by_lower), dict(unplanned_by_lower), "daily_planned_buckets"

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
        clf = _to_text(row.get("leave_classification"))
        track(row.get("assignee"), planned_h, clf)

    _round_both()
    return display_by_lower, dict(planned_by_lower), dict(unplanned_by_lower), ("raw_subtasks_overlap" if raw_used else "none")


def build_workforce_month_payload(
    db_path: Path,
    month_start: date,
    month_end: date,
    canonical_run_id: str,
    *,
    selected_assignees: set[str] | None = None,
    capacity_profile_key: str | None = None,
    jira_base: str = "",
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
    # Actual headcount: all team members minus process team minus resigned
    n_actual = _compute_actual_headcount(db_path)
    actual_capacity_hours = _round_hours(n_actual * per_person_cap) if per_person_cap > 0 else _round_hours(team_capacity_hours)
    actual_capacity_days = _round_hours(actual_capacity_hours / HOURS_PER_DAY)

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
    display_by_lower, planned_by_lower, unplanned_by_lower, leave_aggregate_source = _nested_aligned_leave_by_assignee(
        snapshot,
        month_start,
        month_end,
    )
    # Combined total for display purposes only; availability uses planned leaves only.
    leave_by_lower: dict[str, float] = {
        lk: _round_hours(planned_by_lower.get(lk, 0.0) + unplanned_by_lower.get(lk, 0.0))
        for lk in set(planned_by_lower) | set(unplanned_by_lower)
    }
    known_keys = sorted(display_by_lower.keys(), key=lambda k: display_by_lower[k].lower())

    filter_lowers: set[str] | None = None
    if selected_assignees:
        filter_lowers = {_to_text(a).lower() for a in selected_assignees if _to_text(a)}

    if filter_lowers is None:
        k_sel = n_team
        capacity_hours = _round_hours(team_capacity_hours)
        leave_hours = _round_hours(sum(float(planned_by_lower.get(k, 0.0)) for k in known_keys))
        unplanned_leave_hours = _round_hours(sum(float(unplanned_by_lower.get(k, 0.0)) for k in known_keys))
        active_for_rows = list(known_keys)
    else:
        active_for_rows = sorted(filter_lowers, key=lambda k: display_by_lower.get(k, k).lower())
        k_sel = len(active_for_rows)
        capacity_hours = _round_hours(team_capacity_hours * (k_sel / n_team)) if n_team > 0 else 0.0
        leave_hours = _round_hours(sum(float(planned_by_lower.get(k, 0.0)) for k in filter_lowers))
        unplanned_leave_hours = _round_hours(sum(float(unplanned_by_lower.get(k, 0.0)) for k in filter_lowers))

    assignee_rows: list[dict[str, Any]] = []
    for lk in active_for_rows:
        display = display_by_lower.get(lk) or lk
        planned_lv = float(planned_by_lower.get(lk, 0.0))
        unplanned_lv = float(unplanned_by_lower.get(lk, 0.0))
        lev = planned_lv  # only planned leaves reduce availability
        assignee_rows.append(
            {
                "name": display,
                "leave_hours": _round_hours(lev),
                "leave_days": _round_hours(lev / HOURS_PER_DAY),
                "planned_leave_hours": _round_hours(planned_lv),
                "unplanned_leave_hours": _round_hours(unplanned_lv),
                "total_leave_hours": _round_hours(planned_lv + unplanned_lv),
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
        _nlk = name.lower()
        employee_options.append(
            {
                "name": name,
                "resigned": bool(rec.get("resigned")),
                "resignation_date": rec.get("resignation_date"),
                "leave_hours": _round_hours(float(planned_by_lower.get(_nlk, 0.0))),
                "planned_leave_hours": _round_hours(float(planned_by_lower.get(_nlk, 0.0))),
                "unplanned_leave_hours": _round_hours(float(unplanned_by_lower.get(_nlk, 0.0))),
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
                    "leave_hours": _round_hours(float(planned_by_lower.get(lk, 0.0))),
                    "planned_leave_hours": _round_hours(float(planned_by_lower.get(lk, 0.0))),
                    "unplanned_leave_hours": _round_hours(float(unplanned_by_lower.get(lk, 0.0))),
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

    _jira_base = _to_text(jira_base).rstrip("/")
    daily_leaves: list[dict[str, Any]] = []
    dist_rows = list(snapshot.get("distributed_subtasks") or [])
    if dist_rows:
        for _row in dist_rows:
            if not isinstance(_row, dict):
                continue
            _bucket = _parse_iso_date(_to_text(_row.get("planned_date_for_bucket") or _row.get("start_date")))
            if _bucket is None or not (month_start <= _bucket <= month_end):
                continue
            _ph = max(0.0, float(_row.get("original_estimate_hours") or 0.0))
            if _ph <= 0:
                continue
            _ik = _to_text(_row.get("issue_key")).upper()
            daily_leaves.append({
                "assignee": _to_text(_row.get("assignee")),
                "date": _bucket.isoformat(),
                "hours": _round_hours(_ph),
                "leave_type": _to_text(_row.get("leave_type_raw")),
                "leave_classification": _to_text(_row.get("leave_classification")),
                "summary": _to_text(_row.get("summary")),
                "issue_key": _ik,
                "jira_url": f"{_jira_base}/browse/{_ik}" if _jira_base and _ik else "",
            })
    else:
        for _row in list(snapshot.get("daily") or []):
            if not isinstance(_row, dict):
                continue
            _day = _parse_iso_date(_to_text(_row.get("period_day")))
            if _day is None or not (month_start <= _day <= month_end):
                continue
            _pt = float(_row.get("planned_taken_hours") or 0.0)
            _pn = float(_row.get("planned_not_taken_hours") or 0.0)
            _total = _pt + _pn
            if _total <= 0:
                continue
            daily_leaves.append({
                "assignee": _to_text(_row.get("assignee") or ""),
                "date": _day.isoformat(),
                "hours": _round_hours(_total),
                "leave_type": "",
                "leave_classification": "Planned",
                "summary": "",
                "issue_key": "",
                "jira_url": "",
            })

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
        "employee_count_actual": n_actual,
        "actual_capacity_hours": actual_capacity_hours,
        "actual_capacity_days": actual_capacity_days,
        "selected_employee_count": k_sel,
        "assignee_filter_active": filter_lowers is not None,
        "selected_assignees": [display_by_lower.get(k) or team_display_by_lower.get(k) or k for k in active_for_rows],
        "team_capacity_hours": _round_hours(team_capacity_hours),
        "team_capacity_days": _round_hours(team_capacity_hours / HOURS_PER_DAY),
        "capacity_hours": capacity_hours,
        "capacity_days": _round_hours(capacity_hours / HOURS_PER_DAY),
        "leave_hours": leave_hours,
        "leave_days": _round_hours(leave_hours / HOURS_PER_DAY),
        "planned_leave_hours": leave_hours,
        "planned_leave_days": _round_hours(leave_hours / HOURS_PER_DAY),
        "unplanned_leave_hours": unplanned_leave_hours,
        "unplanned_leave_days": _round_hours(unplanned_leave_hours / HOURS_PER_DAY),
        "total_leave_hours": _round_hours(leave_hours + unplanned_leave_hours),
        "total_leave_days": _round_hours((leave_hours + unplanned_leave_hours) / HOURS_PER_DAY),
        "leave_aggregate_source": leave_aggregate_source,
        "availability_hours": availability_hours,
        "availability_days": _round_hours(availability_hours / HOURS_PER_DAY),
        "assignees": assignee_rows,
        "assignee_options": option_names,
        "employee_options": employee_options,
        "employee_tree": employee_tree,
        "daily_leaves": daily_leaves,
        "support_team": _build_support_team_payload(
            db_path, planned_by_lower, unplanned_by_lower, display_by_lower, per_person_cap, availability_hours,
        ),
        "meta": {
            "capacity_basis": cap_basis,
            "leave_basis": leave_basis_by_src.get(leave_aggregate_source, leave_basis_by_src["none"]),
            "leave_aggregate_source": leave_aggregate_source,
            "availability_formula": "capacity_hours - leave_hours; filtered mode uses (K/N)*team capacity",
        },
    }


def _build_by_project_rows(by_project: DefaultDict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for pk in sorted(by_project.keys()):
        agg = by_project[pk]
        planned = _round_hours(float(agg.get("planned_hours") or 0.0))
        actual = _round_hours(float(agg.get("actual_hours") or 0.0))
        result.append({
            "project_key": pk,
            "project_name": _to_text(agg.get("project_name")) or pk,
            "epic_count": int(agg.get("epic_count") or 0),
            "planned_hours": planned,
            "planned_days": _round_hours(planned / HOURS_PER_DAY),
            "actual_hours": actual,
            "actual_days": _round_hours(actual / HOURS_PER_DAY),
            "brought_forward_count": int(agg.get("brought_forward_count") or 0),
            "brought_forward_planned_hours": _round_hours(float(agg.get("brought_forward_planned_hours") or 0.0)),
            "brought_forward_planned_days": _round_hours(float(agg.get("brought_forward_planned_hours") or 0.0) / HOURS_PER_DAY),
            "carried_forward_count": int(agg.get("carried_forward_count") or 0),
            "carried_forward_planned_hours": _round_hours(float(agg.get("carried_forward_planned_hours") or 0.0)),
            "carried_forward_planned_days": _round_hours(float(agg.get("carried_forward_planned_hours") or 0.0) / HOURS_PER_DAY),
        })
    return result


def _round_totals(totals: dict[str, Any]) -> dict[str, Any]:
    return {
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


def _load_all_jira_epic_rows(
    db_path: Path,
    run_id: str,
    selected_project_keys: set[str],
    month_start: date,
    month_end: date,
    overdue_cutoff: date,
    include_on_hold: bool,
    jira_base: str,
    actual_hours_by_epic: dict[str, float],
    has_worklog_through_month_end: dict[str, bool],
    story_planned_hours_by_epic: dict[str, float],
    last_worklog_date_by_epic: dict[str, date] | None = None,
    worklog_dates_by_epic: dict[str, list[str]] | None = None,
    total_planned_hours_by_epic: dict[str, float] | None = None,
    story_gantt_by_epic: dict[str, list[dict[str, Any]]] | None = None,
    total_actual_hours_by_epic: dict[str, float] | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Build rows and excluded_epics for the ALL JIRA EPICS mode.

    Epics are sourced directly from canonical_issues (not the Epics Planner).
    Month scoping uses Jira start_date / due_date from canonical_issues.
    Planned hours come from story/subtask original estimates (no TK budget).
    """
    if last_worklog_date_by_epic is None:
        last_worklog_date_by_epic = {}
    if worklog_dates_by_epic is None:
        worklog_dates_by_epic = {}
    if total_planned_hours_by_epic is None:
        total_planned_hours_by_epic = {}
    if story_gantt_by_epic is None:
        story_gantt_by_epic = {}
    if total_actual_hours_by_epic is None:
        total_actual_hours_by_epic = {}
    today = date.today()
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        raw_epics = conn.execute(
            """
            SELECT issue_key, project_key, summary, status, start_date, due_date,
                   original_estimate_hours, raw_payload_json, resolved_stable_since_date
            FROM canonical_issues
            WHERE run_id = ? AND issue_type = 'Epic'
            """,
            (run_id,),
        ).fetchall()

    rows: list[dict[str, Any]] = []
    excluded_epics: list[dict[str, Any]] = []

    for epic in raw_epics:
        epic_key = _to_text(epic["issue_key"]).upper()
        project_key = _to_text(epic["project_key"]).upper()
        if selected_project_keys and project_key not in selected_project_keys:
            continue

        jira_status = _to_text(epic["status"])
        is_on_hold = _is_on_hold_status_text(jira_status)
        is_resolved = _is_resolved_status_text(jira_status)
        scope_excluded = is_on_hold and not include_on_hold

        start_text = _to_text(epic["start_date"])
        due_text = _to_text(epic["due_date"])
        actual_completed_text = _to_text(epic["resolved_stable_since_date"])
        approved_start = _parse_iso_date(start_text)
        approved_due = _parse_iso_date(due_text)
        actual_completed_date = _parse_iso_date(actual_completed_text)

        if approved_start is None and approved_due is None:
            continue

        # An epic completed before the selected month is not overdue from this month's perspective.
        completed_before_month = actual_completed_date is not None and actual_completed_date < month_start

        start_in_month = bool(approved_start is not None and month_start <= approved_start <= month_end)
        due_in_month = bool(approved_due is not None and month_start <= approved_due <= month_end)
        brought_forward = bool(
            approved_due is not None
            and approved_due < month_start
            and approved_due >= overdue_cutoff
            and not scope_excluded
            and not completed_before_month
        )

        if not (start_in_month or due_in_month) and not brought_forward:
            overlaps_month = (
                approved_start is not None and approved_due is not None
                and approved_start <= month_end and approved_due >= month_start
            )
            would_be_brought_forward = (
                approved_due is not None
                and approved_due < month_start
                and approved_due >= overdue_cutoff
            )
            if (overlaps_month or would_be_brought_forward) and scope_excluded:
                excl_url = f"{jira_base}/browse/{epic_key}" if jira_base else ""
                excluded_epics.append({
                    "epic_key": epic_key,
                    "epic_name": _to_text(epic["summary"]) or epic_key,
                    "project_key": project_key,
                    "project_name": project_key,
                    "approved_start": start_text,
                    "approved_due": due_text,
                    "jira_status": jira_status,
                    "jira_url": excl_url,
                    "exclusion_reasons": ["Epic is on hold (enable 'Include On Hold' to see it)"],
                })
            continue

        planned_hours = _round_hours(story_planned_hours_by_epic.get(epic_key, 0.0))
        tk_epic_budget_hours = None
        actual_hours = _round_hours(float(actual_hours_by_epic.get(epic_key) or 0.0))

        start_slip = bool(
            approved_start is not None
            and month_start <= approved_start <= month_end
            and not has_worklog_through_month_end.get(epic_key, False)
        )
        end_slip = bool(
            approved_due is not None
            and month_start <= approved_due <= month_end
            and not is_resolved
        )

        # Carry-forward logic:
        # - On-hold epics are never carried forward; they stay in their own month.
        # - Brought-forward overdue epics: carried forward only if NOT resolved in the selected month.
        # - In-month epics: carried forward when one or more risk conditions are met.
        carry_forward_reasons: list[str] = []
        if not is_on_hold:
            if brought_forward:
                if not is_resolved:
                    carry_forward_reasons.append("Brought forward — not resolved in selected month")
            elif start_in_month or due_in_month:
                if start_slip:
                    carry_forward_reasons.append("Start slipped — no worklog logged by month end")
                if end_slip:
                    carry_forward_reasons.append("End slipped — not resolved by due date in selected month")
                due_past_today = bool(
                    approved_due is not None and approved_due < today and not is_resolved
                )
                if due_past_today:
                    carry_forward_reasons.append(
                        f"Due date passed ({due_text}) — epic still open"
                    )
                due_within_or_before_month = approved_due is None or approved_due <= month_end
                last_wl = last_worklog_date_by_epic.get(epic_key)
                no_recent_activity = (
                    due_within_or_before_month
                    and not is_resolved
                    and (last_wl is None or last_wl < today - timedelta(days=7))
                )
                if no_recent_activity:
                    last_wl_str = last_wl.isoformat() if last_wl else "never"
                    carry_forward_reasons.append(
                        f"No worklog activity in last 7 days (last logged: {last_wl_str})"
                    )
                if due_within_or_before_month and not is_resolved and planned_hours > 0 and actual_hours > planned_hours * 1.5 and (actual_hours - planned_hours) >= 4:
                    pct = int(round(actual_hours / planned_hours * 100))
                    carry_forward_reasons.append(
                        f"Over budget — {actual_hours:.1f}h logged vs {planned_hours:.1f}h planned ({pct}%)"
                    )

        carried_forward = bool(carry_forward_reasons)

        no_stories_in_month = bool(not brought_forward and planned_hours == 0.0)
        if no_stories_in_month:
            excl_url = f"{jira_base}/browse/{epic_key}" if jira_base else ""
            excluded_epics.append({
                "epic_key": epic_key,
                "epic_name": _to_text(epic["summary"]) or epic_key,
                "project_key": project_key,
                "project_name": project_key,
                "approved_start": start_text,
                "approved_due": due_text,
                "jira_status": jira_status,
                "jira_url": excl_url,
                "exclusion_reasons": ["No stories planned in this month"],
            })

        jira_url = f"{jira_base}/browse/{epic_key}" if jira_base else ""
        rows.append({
            "epic_key": epic_key,
            "epic_name": _to_text(epic["summary"]) or epic_key,
            "project_key": project_key,
            "project_name": project_key,
            "product_category": "",
            "component": "",
            "approved_start": start_text,
            "approved_due": due_text,
            "actual_completed_date": actual_completed_text,
            "jira_status": jira_status,
            "planned_hours": _round_hours(planned_hours),
            "planned_days": _round_hours(planned_hours / HOURS_PER_DAY),
            "tk_epic_budget_hours": tk_epic_budget_hours,
            "tk_epic_budget_days": None,
            "actual_hours": actual_hours,
            "actual_days": _round_hours(actual_hours / HOURS_PER_DAY),
            "total_actual_hours": _round_hours(float(total_actual_hours_by_epic.get(epic_key, 0.0))),
            "start_slip": start_slip,
            "end_slip": end_slip,
            "brought_forward": brought_forward,
            "carried_forward": carried_forward,
            "carry_forward_reasons": carry_forward_reasons,
            "is_on_hold": is_on_hold,
            "no_stories_in_month": no_stories_in_month,
            "jira_url": jira_url,
            "subtask_count": 0,
            "worklog_dates": worklog_dates_by_epic.get(epic_key, []),
            "total_planned_hours": _round_hours(total_planned_hours_by_epic.get(epic_key, 0.0)),
            "total_planned_days": _round_hours(total_planned_hours_by_epic.get(epic_key, 0.0) / HOURS_PER_DAY),
            "story_bars": story_gantt_by_epic.get(epic_key, []),
        })

    return rows, excluded_epics


def _compute_actual_headcount(db_path: Path) -> int:
    """Count active resources: all team members minus process team minus resigned."""
    try:
        teams = _list_performance_teams(db_path)
    except Exception:
        return 0
    all_members: list[str] = []
    process_members_lower: set[str] = set()
    for team in teams:
        team_name = _to_text(team.get("team_name"))
        for m in (team.get("assignees") or []):
            name = _to_text(m)
            if not name:
                continue
            all_members.append(name)
            if team_name.lower() == "process team":
                process_members_lower.add(name.lower())
    unique_members = sorted(set(m for m in all_members if m), key=str.lower)
    try:
        res_map = _load_performance_resource_resignation_map(db_path, unique_members)
    except Exception:
        res_map = {}
    resigned_lower = {name.lower() for name, rec in res_map.items() if rec.get("resigned")}
    excluded = process_members_lower | resigned_lower
    return sum(1 for m in unique_members if m.lower() not in excluded)


def _build_support_team_payload(
    db_path: Path,
    planned_by_lower: dict[str, float],
    unplanned_by_lower: dict[str, float],
    display_by_lower: dict[str, str],
    per_person_cap: float,
    total_availability_hours: float,
) -> dict[str, Any]:
    members = load_support_team(db_path)
    member_rows: list[dict[str, Any]] = []
    for name in members:
        lk = name.lower()
        planned_lv = _round_hours(float(planned_by_lower.get(lk, 0.0)))
        unplanned_lv = _round_hours(float(unplanned_by_lower.get(lk, 0.0)))
        lv = planned_lv  # only planned leaves reduce availability
        cap = per_person_cap
        avail = _round_hours(cap - lv)
        display = display_by_lower.get(lk) or name
        member_rows.append({
            "name": display,
            "capacity_hours": cap,
            "capacity_days": _round_hours(cap / HOURS_PER_DAY),
            "leave_hours": lv,
            "leave_days": _round_hours(lv / HOURS_PER_DAY),
            "planned_leave_hours": planned_lv,
            "unplanned_leave_hours": unplanned_lv,
            "total_leave_hours": _round_hours(planned_lv + unplanned_lv),
            "availability_hours": avail,
            "availability_days": _round_hours(avail / HOURS_PER_DAY),
        })
    total_support_cap = _round_hours(sum(r["capacity_hours"] for r in member_rows))
    total_support_leave = _round_hours(sum(r["leave_hours"] for r in member_rows))
    total_support_unplanned = _round_hours(sum(r["unplanned_leave_hours"] for r in member_rows))
    total_support_avail = _round_hours(sum(r["availability_hours"] for r in member_rows))
    avail_for_features = _round_hours(total_availability_hours - total_support_avail)
    return {
        "saved_members": members,
        "member_rows": member_rows,
        "total_capacity_hours": total_support_cap,
        "total_capacity_days": _round_hours(total_support_cap / HOURS_PER_DAY),
        "total_leave_hours": total_support_leave,
        "total_leave_days": _round_hours(total_support_leave / HOURS_PER_DAY),
        "total_planned_leave_hours": total_support_leave,
        "total_unplanned_leave_hours": total_support_unplanned,
        "total_availability_hours": total_support_avail,
        "total_availability_days": _round_hours(total_support_avail / HOURS_PER_DAY),
        "availability_for_features_hours": avail_for_features,
        "availability_for_features_days": _round_hours(avail_for_features / HOURS_PER_DAY),
    }


_SUPPORT_TEAM_TABLE = """
CREATE TABLE IF NOT EXISTS support_team_config (
    key TEXT PRIMARY KEY,
    members_json TEXT NOT NULL DEFAULT '[]',
    updated_at TEXT DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ', 'now'))
)
"""


def load_support_team(db_path: Path) -> list[str]:
    try:
        with sqlite3.connect(db_path) as conn:
            conn.execute(_SUPPORT_TEAM_TABLE)
            row = conn.execute(
                "SELECT members_json FROM support_team_config WHERE key = 'members'"
            ).fetchone()
        if not row:
            return []
        parsed = json.loads(row[0])
        return [_to_text(n) for n in parsed if _to_text(n)]
    except Exception:
        return []


def save_support_team(db_path: Path, members: list[str]) -> list[str]:
    cleaned = [_to_text(n) for n in (members or []) if _to_text(n)]
    with sqlite3.connect(db_path) as conn:
        conn.execute(_SUPPORT_TEAM_TABLE)
        conn.execute(
            """
            INSERT INTO support_team_config (key, members_json, updated_at)
            VALUES ('members', ?, strftime('%Y-%m-%dT%H:%M:%SZ', 'now'))
            ON CONFLICT(key) DO UPDATE SET
                members_json = excluded.members_json,
                updated_at   = excluded.updated_at
            """,
            (json.dumps(cleaned),),
        )
    return cleaned


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
    include_bug_subtasks: bool = True,
    epic_mode: str = "tk_epics",
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
    tk_planned_hours_by_issue: dict[str, float] = _build_tk_planned_hours_by_issue(planner_rows)
    epic_meta, subtask_keys_by_epic = _load_canonical_issue_maps(
        db_path,
        run_id,
        include_bug_subtasks=include_bug_subtasks,
    )
    actual_hours_by_epic, has_worklog_through_month_end, last_worklog_date_by_epic, worklog_dates_by_epic, total_actual_hours_by_epic = _load_worklog_metrics(
        db_path,
        run_id,
        subtask_keys_by_epic,
        month_start,
        month_end,
    )

    overdue_cutoff = month_start - timedelta(days=max(0, int(overdue_threshold_days)))
    jira_base = _to_text(jira_base_url).rstrip("/")

    # ALL JIRA EPICS mode: source epics directly from canonical_issues.
    if epic_mode == "all_jira_epics":
        all_jira_keys: list[str] = []
        with sqlite3.connect(db_path) as _conn:
            _conn.row_factory = sqlite3.Row
            _jira_epics = _conn.execute(
                "SELECT issue_key, project_key, start_date, due_date, status, resolved_stable_since_date FROM canonical_issues WHERE run_id = ? AND issue_type = 'Epic'",
                (run_id,),
            ).fetchall()
        for _je in _jira_epics:
            _ek = _to_text(_je["issue_key"]).upper()
            if not _ek:
                continue
            if selected_project_keys and _to_text(_je["project_key"]).upper() not in selected_project_keys:
                continue
            _as = _parse_iso_date(_je["start_date"])
            _ad = _parse_iso_date(_je["due_date"])
            if _as is None and _ad is None:
                continue
            _js = _to_text(_je["status"])
            _scope_excl = _is_on_hold_status_text(_js) and not include_on_hold
            _acd = _parse_iso_date(_je["resolved_stable_since_date"])
            _completed_before = _acd is not None and _acd < month_start
            _bf = _ad is not None and _ad < month_start and _ad >= overdue_cutoff and not _scope_excl and not _completed_before
            _sim = _as is not None and month_start <= _as <= month_end
            _dim = _ad is not None and month_start <= _ad <= month_end
            if _sim or _dim or _bf:
                all_jira_keys.append(_ek)
        story_planned_hours_by_epic = _load_story_planned_hours(
            db_path,
            run_id,
            all_jira_keys,
            month_start,
            month_end,
            include_bug_subtasks=include_bug_subtasks,
        )
        total_planned_hours_by_epic = _load_total_planned_hours(
            db_path,
            run_id,
            all_jira_keys,
            include_bug_subtasks=include_bug_subtasks,
        )
        story_gantt_by_epic = _load_story_gantt_data(
            db_path,
            run_id,
            all_jira_keys,
            include_bug_subtasks=include_bug_subtasks,
        )
        rows, excluded_epics = _load_all_jira_epic_rows(
            db_path, run_id, selected_project_keys, month_start, month_end,
            overdue_cutoff, include_on_hold, jira_base,
            actual_hours_by_epic, has_worklog_through_month_end, story_planned_hours_by_epic,
            last_worklog_date_by_epic, worklog_dates_by_epic, total_planned_hours_by_epic,
            story_gantt_by_epic, total_actual_hours_by_epic,
        )
        # Build by_project and totals from rows
        by_project: DefaultDict[str, dict[str, Any]] = defaultdict(
            lambda: {"project_key": "", "project_name": "", "epic_count": 0, "planned_hours": 0.0,
                     "actual_hours": 0.0, "brought_forward_count": 0, "brought_forward_planned_hours": 0.0,
                     "carried_forward_count": 0, "carried_forward_planned_hours": 0.0}
        )
        totals = {"epic_count": 0, "planned_hours": 0.0, "actual_hours": 0.0,
                  "start_slip_count": 0, "end_slip_count": 0, "brought_forward_count": 0,
                  "brought_forward_planned_hours": 0.0, "carried_forward_count": 0,
                  "carried_forward_planned_hours": 0.0}
        for _r in rows:
            _pk = _to_text(_r.get("project_key")).upper()
            _ph = float(_r.get("planned_hours") or 0.0)
            _ah = float(_r.get("actual_hours") or 0.0)
            _bf2 = bool(_r.get("brought_forward"))
            _cf = bool(_r.get("carried_forward"))
            totals["epic_count"] += 1
            totals["planned_hours"] += _ph
            totals["actual_hours"] += _ah
            totals["start_slip_count"] += 1 if _r.get("start_slip") else 0
            totals["end_slip_count"] += 1 if _r.get("end_slip") else 0
            totals["brought_forward_count"] += 1 if _bf2 else 0
            totals["carried_forward_count"] += 1 if _cf else 0
            if _bf2:
                totals["brought_forward_planned_hours"] = float(totals["brought_forward_planned_hours"]) + _ph
            if _cf:
                totals["carried_forward_planned_hours"] = float(totals["carried_forward_planned_hours"]) + _ph
            pj = by_project[_pk]
            pj["project_key"] = _pk
            pj["project_name"] = _to_text(_r.get("project_name")) or _pk
            pj["epic_count"] = int(pj["epic_count"]) + 1
            pj["planned_hours"] = float(pj["planned_hours"]) + _ph
            pj["actual_hours"] = float(pj["actual_hours"]) + _ah
            if _bf2:
                pj["brought_forward_count"] = int(pj["brought_forward_count"]) + 1
                pj["brought_forward_planned_hours"] = float(pj["brought_forward_planned_hours"]) + _ph
            if _cf:
                pj["carried_forward_count"] = int(pj["carried_forward_count"]) + 1
                pj["carried_forward_planned_hours"] = float(pj["carried_forward_planned_hours"]) + _ph
        rows.sort(key=lambda item: (_to_text(item.get("project_name")).lower(), _to_text(item.get("approved_start")), _to_text(item.get("epic_key"))))
        excluded_epics.sort(key=lambda e: (_to_text(e.get("project_name")).lower(), _to_text(e.get("approved_start")), _to_text(e.get("epic_key"))))
        by_project_rows = _build_by_project_rows(by_project)
        rounded_totals = _round_totals(totals)
        estimate_rollup = _load_estimate_rollup_for_epics(
            db_path,
            run_id,
            [_to_text(row.get("epic_key")).upper() for row in rows if not row.get("brought_forward")],
            month_start,
            month_end,
            jira_base,
            tk_planned_hours_by_issue,
            include_bug_subtasks,
        )
        workforce = build_workforce_month_payload(db_path, month_start, month_end, run_id, selected_assignees=selected_assignees, capacity_profile_key=capacity_profile_key, jira_base=jira_base)
        return {
            "month": month, "from_date": month_start.isoformat(), "to_date": month_end.isoformat(),
            "canonical_run_id": run_id, "selected_projects": sorted(selected_project_keys),
            "overdue_threshold_days": int(overdue_threshold_days), "include_on_hold": bool(include_on_hold),
            "include_bug_subtasks": bool(include_bug_subtasks),
            "epic_mode": epic_mode, "rows": rows, "by_project": by_project_rows,
            "totals": rounded_totals, "estimate_rollup": estimate_rollup,
            "excluded_epics": excluded_epics, "workforce": workforce,
            "meta": {
                "hours_per_day": HOURS_PER_DAY,
                "scope_basis": (
                    "All Jira epics from the latest canonical refresh whose Jira start or due date "
                    "falls inside the selected month, plus unresolved epics overdue by less than the "
                    "overdue threshold. Planned hours use story/subtask original estimates from Jira."
                ),
                "actual_hours_basis": (
                    "canonical_worklogs.started_date within the selected month"
                    + (" including Bug Subtask issues" if include_bug_subtasks else " excluding Bug Subtask issues")
                ),
                "status_basis": "Current Jira epic status from the latest canonical refresh",
            },
        }

    # Collect all in-scope epic keys first so we can batch-load story hours.
    # We do a lightweight pre-pass to identify which epics will be in scope,
    # then load story/subtask planned hours for all of them in one DB query.
    _all_candidate_epic_keys: list[str] = []
    for _pr in planner_rows:
        _ek = _to_text(_pr.get("epic_key")).upper()
        if not _ek:
            continue
        if selected_project_keys and _to_text(_pr.get("project_key")).upper() not in selected_project_keys:
            continue
        _ep = ((_pr.get("plans") or {}).get("epic_plan") or {})
        _as, _ad, _, _ = _approved_dates(_ep if isinstance(_ep, dict) else {}, _pr)
        if _as is None and _ad is None:
            continue
        _js = _to_text((epic_meta.get(_ek) or {}).get("jira_status"))
        _eff_res = _is_resolved_status_text(_js) or (_is_on_hold_status_text(_js) and not include_on_hold)
        _acd_text = _to_text((epic_meta.get(_ek) or {}).get("resolved_stable_since_date"))
        _acd = _parse_iso_date(_acd_text)
        _completed_before = _acd is not None and _acd < month_start
        _bf = _ad is not None and _ad < month_start and _ad >= overdue_cutoff and not _eff_res and not _completed_before
        _sim = _as is not None and month_start <= _as <= month_end
        _dim = _ad is not None and month_start <= _ad <= month_end
        if _sim or _dim or _bf:
            _all_candidate_epic_keys.append(_ek)

    story_planned_hours_by_epic: dict[str, float] = _load_story_planned_hours(
        db_path,
        run_id,
        _all_candidate_epic_keys,
        month_start,
        month_end,
        include_bug_subtasks=include_bug_subtasks,
    )
    tk_planned_hours_by_issue: dict[str, float] = _build_tk_planned_hours_by_issue(planner_rows)
    total_planned_hours_by_epic: dict[str, float] = _load_total_planned_hours(
        db_path,
        run_id,
        _all_candidate_epic_keys,
        include_bug_subtasks=include_bug_subtasks,
    )
    story_gantt_by_epic: dict[str, list[dict[str, Any]]] = _load_story_gantt_data(
        db_path,
        run_id,
        _all_candidate_epic_keys,
        include_bug_subtasks=include_bug_subtasks,
    )
    child_items_by_epic: dict[str, list[dict[str, Any]]] = _load_child_items_for_month(
        db_path,
        run_id,
        _all_candidate_epic_keys,
        month_start,
        month_end,
        jira_base,
        include_bug_subtasks=include_bug_subtasks,
    )

    rows: list[dict[str, Any]] = []
    excluded_epics: list[dict[str, Any]] = []
    today = date.today()
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
        approved_start, approved_due, approved_start_text, approved_due_text = _approved_dates(epic_plan, planner_row)
        if approved_start is None and approved_due is None:
            continue
        canonical_meta = epic_meta.get(epic_key) or {}
        jira_status = _to_text(canonical_meta.get("jira_status"))
        is_on_hold = _is_on_hold_status_text(jira_status)
        is_resolved = _is_resolved_status_text(jira_status)
        actual_completed_text = _to_text(canonical_meta.get("resolved_stable_since_date"))
        actual_completed_date = _parse_iso_date(actual_completed_text)
        # Only on-hold epics (when include_on_hold is off) are excluded from scope.
        # Resolved epics that match the month filter are shown in the table as-is.
        scope_excluded = is_on_hold and not include_on_hold
        # An epic completed before the selected month is not overdue from this month's perspective.
        completed_before_month = actual_completed_date is not None and actual_completed_date < month_start
        brought_forward = bool(
            approved_due is not None
            and approved_due < month_start
            and approved_due >= overdue_cutoff
            and not scope_excluded
            and not completed_before_month
        )
        start_in_month = bool(
            approved_start is not None and month_start <= approved_start <= month_end
        )
        due_in_month = bool(approved_due is not None and month_start <= approved_due <= month_end)

        # Capture on-hold epics excluded from scope that have month-relevant dates.
        if not (start_in_month or due_in_month) and not brought_forward:
            overlaps_month = (
                approved_start is not None and approved_due is not None
                and approved_start <= month_end and approved_due >= month_start
            )
            would_be_brought_forward = (
                approved_due is not None
                and approved_due < month_start
                and approved_due >= overdue_cutoff
            )
            if (overlaps_month or would_be_brought_forward) and scope_excluded:
                excl_url = _to_text(planner_row.get("jira_url"))
                if not excl_url and jira_base:
                    excl_url = f"{jira_base}/browse/{epic_key}"
                excluded_epics.append({
                    "epic_key": epic_key,
                    "epic_name": _to_text(planner_row.get("epic_name")) or _to_text(canonical_meta.get("jira_summary")) or epic_key,
                    "project_key": project_key,
                    "project_name": _to_text(planner_row.get("project_name")) or project_key,
                    "approved_start": approved_start_text,
                    "approved_due": approved_due_text,
                    "jira_status": jira_status,
                    "jira_url": excl_url,
                    "exclusion_reasons": ["Epic is on hold (enable 'Include On Hold' to see it)"],
                })
            continue

        # --- Hybrid planned hours ---
        # Brought-forward (overdue) epics: their stories have pre-month dates, so use
        # TK-budgeted man-days as the planned commitment carried into this month.
        # In-month epics (start or due in month): use story/subtask original estimates
        # that fall within the month (bottom-up, reflects actual Jira planning).
        tk_epic_budget_hours = _hours_from_man_days(
            epic_plan.get("tk_budgeted_man_days")
            if epic_plan.get("tk_budgeted_man_days") not in (None, "")
            else epic_plan.get("man_days")
        )

        if brought_forward:
            planned_hours = _round_hours(tk_epic_budget_hours or 0.0)
        else:
            planned_hours = _round_hours(story_planned_hours_by_epic.get(epic_key, 0.0))

        actual_hours = _round_hours(float(actual_hours_by_epic.get(epic_key) or 0.0))
        start_slip = bool(
            approved_start is not None
            and month_start <= approved_start <= month_end
            and not has_worklog_through_month_end.get(epic_key, False)
        )
        end_slip = bool(
            approved_due is not None
            and month_start <= approved_due <= month_end
            and not is_resolved
        )

        # Carry-forward logic:
        # - On-hold epics are never carried forward; they stay in their own month.
        # - Brought-forward overdue epics: carried forward only if NOT resolved in the selected month.
        # - In-month epics: carried forward when one or more risk conditions are met.
        carry_forward_reasons: list[str] = []
        if not is_on_hold:
            if brought_forward:
                if not is_resolved:
                    carry_forward_reasons.append("Brought forward — not resolved in selected month")
            elif start_in_month or due_in_month:
                if start_slip:
                    carry_forward_reasons.append("Start slipped — no worklog logged by month end")
                if end_slip:
                    carry_forward_reasons.append("End slipped — not resolved by due date in selected month")
                due_past_today = bool(
                    approved_due is not None and approved_due < today and not is_resolved
                )
                if due_past_today:
                    carry_forward_reasons.append(
                        f"Due date passed ({approved_due_text}) — epic still open"
                    )
                due_within_or_before_month = approved_due is None or approved_due <= month_end
                last_wl = last_worklog_date_by_epic.get(epic_key)
                no_recent_activity = (
                    due_within_or_before_month
                    and not is_resolved
                    and (last_wl is None or last_wl < today - timedelta(days=7))
                )
                if no_recent_activity:
                    last_wl_str = last_wl.isoformat() if last_wl else "never"
                    carry_forward_reasons.append(
                        f"No worklog activity in last 7 days (last logged: {last_wl_str})"
                    )
                if due_within_or_before_month and not is_resolved and planned_hours > 0 and actual_hours > planned_hours * 1.5 and (actual_hours - planned_hours) >= 4:
                    pct = int(round(actual_hours / planned_hours * 100))
                    carry_forward_reasons.append(
                        f"Over budget — {actual_hours:.1f}h logged vs {planned_hours:.1f}h planned ({pct}%)"
                    )

        carried_forward = bool(carry_forward_reasons)

        # Flag in-month epics that contributed 0 planned hours (no stories in this month).
        no_stories_in_month = bool(not brought_forward and planned_hours == 0.0)
        if no_stories_in_month:
            excl_url = _to_text(planner_row.get("jira_url"))
            if not excl_url and jira_base:
                excl_url = f"{jira_base}/browse/{epic_key}"
            excluded_epics.append({
                "epic_key": epic_key,
                "epic_name": _to_text(planner_row.get("epic_name")) or _to_text(canonical_meta.get("jira_summary")) or epic_key,
                "project_key": project_key,
                "project_name": _to_text(planner_row.get("project_name")) or project_key,
                "approved_start": approved_start_text,
                "approved_due": approved_due_text,
                "jira_status": jira_status,
                "jira_url": excl_url,
                "exclusion_reasons": ["No stories planned in this month"],
            })

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
                "actual_completed_date": actual_completed_text,
                "jira_status": jira_status,
                "planned_hours": _round_hours(planned_hours),
                "planned_days": _round_hours(planned_hours / HOURS_PER_DAY),
                "tk_epic_budget_hours": tk_epic_budget_hours,
                "tk_epic_budget_days": None if tk_epic_budget_hours is None else _round_hours(tk_epic_budget_hours / HOURS_PER_DAY),
                "actual_hours": actual_hours,
                "actual_days": _round_hours(actual_hours / HOURS_PER_DAY),
                "total_actual_hours": _round_hours(float(total_actual_hours_by_epic.get(epic_key, 0.0))),
                "start_slip": start_slip,
                "end_slip": end_slip,
                "brought_forward": brought_forward,
                "carried_forward": carried_forward,
                "carry_forward_reasons": carry_forward_reasons,
                "is_on_hold": is_on_hold,
                "no_stories_in_month": no_stories_in_month,
                "jira_url": jira_url,
                "subtask_count": len(subtask_keys_by_epic.get(epic_key, set())),
                "child_items": child_items_by_epic.get(epic_key, []),
                "worklog_dates": worklog_dates_by_epic.get(epic_key, []),
                "total_planned_hours": _round_hours(total_planned_hours_by_epic.get(epic_key, 0.0)),
                "total_planned_days": _round_hours(total_planned_hours_by_epic.get(epic_key, 0.0) / HOURS_PER_DAY),
                "story_bars": story_gantt_by_epic.get(epic_key, []),
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
    excluded_epics.sort(key=lambda e: (_to_text(e.get("project_name")).lower(), _to_text(e.get("approved_start")), _to_text(e.get("epic_key"))))
    by_project_rows = _build_by_project_rows(by_project)
    rounded_totals = _round_totals(totals)
    estimate_rollup = _load_estimate_rollup_for_epics(
        db_path,
        run_id,
        [_to_text(row.get("epic_key")).upper() for row in rows if not row.get("brought_forward")],
        month_start,
        month_end,
        jira_base,
        tk_planned_hours_by_issue,
        include_bug_subtasks,
    )
    workforce = build_workforce_month_payload(
        db_path, month_start, month_end, run_id,
        selected_assignees=selected_assignees,
        capacity_profile_key=capacity_profile_key,
        jira_base=jira_base,
    )
    return {
        "month": month,
        "from_date": month_start.isoformat(),
        "to_date": month_end.isoformat(),
        "canonical_run_id": run_id,
        "selected_projects": sorted(selected_project_keys),
        "overdue_threshold_days": int(overdue_threshold_days),
        "include_on_hold": bool(include_on_hold),
        "include_bug_subtasks": bool(include_bug_subtasks),
        "epic_mode": epic_mode,
        "rows": rows,
        "by_project": by_project_rows,
        "totals": rounded_totals,
        "estimate_rollup": estimate_rollup,
        "excluded_epics": excluded_epics,
        "workforce": workforce,
        "meta": {
            "hours_per_day": HOURS_PER_DAY,
            "scope_basis": (
                "Epics whose approved start date or approved due date falls inside the selected calendar month "
                "(TK planner dates), plus brought-forward unresolved epics whose approved due is before that month. "
                "Planned hours: brought-forward epics use TK-budgeted man-days; in-month epics use "
                "story/subtask original estimates from Jira (subtask hours replace parent story hours when subtasks carry estimates)."
            ),
            "actual_hours_basis": (
                "canonical_worklogs.started_date within the selected month"
                + (" including Bug Subtask issues" if include_bug_subtasks else " excluding Bug Subtask issues")
            ),
            "status_basis": "Current Jira epic status from the latest canonical refresh",
        },
    }
