"""Data layer for the Support Center report.

Reads the standalone ``support_center.db`` (Support-tagged issue keys) and joins
it, read-only, against the canonical data already produced for the other reports
(`canonical_issues`, `canonical_worklogs`, `canonical_issue_actuals`) plus the
support-team capacity model used by the Monthly Epic Plan Progress report.

NO canonical database is modified here — every query against the capacity DB is a
SELECT.
"""
from __future__ import annotations

import calendar
import re
import sqlite3
from datetime import date, datetime
from pathlib import Path
from typing import Any

from canonical_report_data import (
    load_canonical_actuals_by_issue,
    load_canonical_issues,
    load_canonical_worklogs,
    resolve_canonical_run_id,
)
from monthly_epic_plan_progress_service import (
    HOURS_PER_DAY,
    _month_bounds,
    build_workforce_month_payload,
)

# Booking stories: "Support by <name> (<Month Year>)" or "Support by <name> <Month Year>"
# Also matches "Technical Support by ..." variants.
BOOKING_STORY_RE = re.compile(r"^\s*(?:technical\s+)?support\s+by\s+.+\(.+\)\s*$", re.IGNORECASE)
# Looser pattern: "Support by <name> <Month Year>" without parentheses
BOOKING_STORY_LOOSE_RE = re.compile(
    r"^\s*(?:technical\s+)?support\s+by\s+\S+(?:\s+\S+)*\s+"
    r"(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
    r"\s+\d{4}\s*$",
    re.IGNORECASE,
)

DONE_STATUSES = {"done", "closed", "resolved", "complete", "completed"}


def _to_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _to_float(value: Any) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _round_hours(value: float) -> float:
    return round(float(value or 0.0) + 1e-9, 2)


def _parse_iso_date(value: Any) -> date | None:
    text = _to_text(value)
    if not text:
        return None
    try:
        return date.fromisoformat(text[:10])
    except ValueError:
        return None


def _is_subtask_type(issue_type: str) -> bool:
    low = _to_text(issue_type).lower()
    return "subtask" in low or "sub-task" in low


def _is_story_type(issue_type: str) -> bool:
    return "story" in _to_text(issue_type).lower()


def _is_done(status: str) -> bool:
    return _to_text(status).lower() in DONE_STATUSES


def is_booking_story(summary: str) -> bool:
    text = _to_text(summary)
    return bool(BOOKING_STORY_RE.match(text) or BOOKING_STORY_LOOSE_RE.match(text))


def _months_in_range(from_date: date, to_date: date) -> list[str]:
    months: list[str] = []
    year, month = from_date.year, from_date.month
    while (year, month) <= (to_date.year, to_date.month):
        months.append(f"{year:04d}-{month:02d}")
        if month == 12:
            year, month = year + 1, 1
        else:
            month += 1
    return months


def load_support_keys(support_db_path: Path) -> dict[str, dict[str, Any]]:
    """Return {issue_key: {project_key, issue_type, summary, work_type_value}}."""
    if not support_db_path.exists():
        return {}
    try:
        with sqlite3.connect(support_db_path) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                """
                SELECT issue_key, project_key, issue_type, summary, work_type_value
                FROM support_issues
                """
            ).fetchall()
    except sqlite3.Error:
        return {}
    out: dict[str, dict[str, Any]] = {}
    for row in rows:
        key = _to_text(row["issue_key"]).upper()
        if key:
            out[key] = dict(row)
    return out


def _story_completion_date(issue: dict[str, Any], actual: dict[str, Any]) -> date | None:
    """Actual completion date for a support story, with documented fallbacks."""
    candidate = (
        _parse_iso_date(actual.get("actual_complete_date"))
        or _parse_iso_date(actual.get("last_worklog_date"))
        or _parse_iso_date(issue.get("due_date"))
        or _parse_iso_date(issue.get("start_date"))
    )
    return candidate


def _build_context(
    canonical_db_path: Path,
    support_db_path: Path,
    run_id: str,
) -> dict[str, Any]:
    effective_run_id = resolve_canonical_run_id(canonical_db_path, run_id)
    issues = load_canonical_issues(canonical_db_path, effective_run_id)
    actuals = load_canonical_actuals_by_issue(canonical_db_path, effective_run_id)
    support_keys = load_support_keys(support_db_path)

    worklogs_by_issue: dict[str, float] = {}
    # Track per-issue worklogs with dates for range-based filtering
    worklogs_dated: dict[str, list[tuple[str, float]]] = {}
    for wl in load_canonical_worklogs(canonical_db_path, effective_run_id):
        wl_key = _to_text(wl.get("issue_key")).upper()
        if wl_key:
            hours = _to_float(wl.get("hours_logged"))
            worklogs_by_issue[wl_key] = worklogs_by_issue.get(wl_key, 0.0) + hours
            wl_date = _to_text(wl.get("started_date"))
            worklogs_dated.setdefault(wl_key, []).append((wl_date, hours))

    issue_by_key: dict[str, dict[str, Any]] = {}
    children_by_story: dict[str, list[dict[str, Any]]] = {}
    for item in issues:
        key = _to_text(item.get("issue_key")).upper()
        if not key:
            continue
        issue_by_key[key] = item
        if _is_subtask_type(item.get("issue_type", "")):
            story_key = _to_text(item.get("story_key")).upper()
            if story_key:
                children_by_story.setdefault(story_key, []).append(item)
    return {
        "run_id": effective_run_id,
        "issue_by_key": issue_by_key,
        "children_by_story": children_by_story,
        "actuals": actuals,
        "support_keys": support_keys,
        "worklogs_by_issue": worklogs_by_issue,
        "worklogs_dated": worklogs_dated,
    }


def _classify_support_stories(ctx: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Split the Support-tagged stories into (actual stories, booking stories)."""
    actual_stories: list[dict[str, Any]] = []
    booking_stories: list[dict[str, Any]] = []
    issue_by_key = ctx["issue_by_key"]
    for key, meta in ctx["support_keys"].items():
        issue = issue_by_key.get(key)
        # Prefer canonical issue_type/summary; fall back to the support_issues row.
        issue_type = _to_text((issue or {}).get("issue_type") or meta.get("issue_type"))
        summary = _to_text((issue or {}).get("summary") or meta.get("summary"))
        if not _is_story_type(issue_type):
            continue
        record = issue or {
            "issue_key": key,
            "project_key": _to_text(meta.get("project_key")),
            "issue_type": issue_type,
            "summary": summary,
            "status": "",
        }
        if is_booking_story(summary):
            booking_stories.append(record)
        else:
            actual_stories.append(record)
    return actual_stories, booking_stories


def _subtask_hours(ctx: dict[str, Any], story_key: str) -> tuple[float, int, int]:
    """Sum logged hours across a story's subtasks/bug-subtasks; count done ones."""
    actuals = ctx["actuals"]
    worklogs_by_issue = ctx.get("worklogs_by_issue") or {}
    total_hours = 0.0
    subtask_count = 0
    done_subtasks = 0
    for child in ctx["children_by_story"].get(story_key.upper(), []):
        subtask_count += 1
        child_key = _to_text(child.get("issue_key")).upper()
        actual = actuals.get(child_key, {})
        hours = _to_float(actual.get("total_worklog_hours"))
        if not hours:
            hours = _to_float(worklogs_by_issue.get(child_key))
        if not hours:
            hours = _to_float(child.get("total_hours_logged"))
        total_hours += hours
        if _is_done(child.get("status")):
            done_subtasks += 1
    return total_hours, subtask_count, done_subtasks


def _story_row(ctx: dict[str, Any], story: dict[str, Any]) -> dict[str, Any]:
    key = _to_text(story.get("issue_key")).upper()
    actual = ctx["actuals"].get(key, {})
    completion = _story_completion_date(story, actual)
    hours, subtask_count, done_subtasks = _subtask_hours(ctx, key)
    return {
        "issue_key": key,
        "project_key": _to_text(story.get("project_key")).upper(),
        "summary": _to_text(story.get("summary")),
        "status": _to_text(story.get("status")),
        "assignee": _to_text(story.get("assignee")),
        "start_date": _to_text(story.get("start_date")),
        "due_date": _to_text(story.get("due_date")),
        "actual_completion_date": completion.isoformat() if completion else "",
        "is_resolved": _is_done(story.get("status")),
        "invested_hours": _round_hours(hours),
        "invested_days": _round_hours(hours / HOURS_PER_DAY),
        "subtask_count": subtask_count,
        "subtask_done_count": done_subtasks,
        "_completion_date": completion,
    }


def _in_range(completion: date | None, from_date: date, to_date: date) -> bool:
    if completion is None:
        return False
    return from_date <= completion <= to_date


def _story_overlaps_range(story: dict[str, Any], ctx: dict[str, Any], from_date: date, to_date: date) -> bool:
    """Return True if the story has any activity or date overlap with [from_date, to_date].

    A story overlaps if:
    - Its completion date is in range, OR
    - Its start_date..due_date span overlaps the range (both must exist), OR
    - Any of its subtasks have worklogs within the range
    """
    key = _to_text(story.get("issue_key")).upper()
    issue = ctx["issue_by_key"].get(key, story)
    actual = ctx["actuals"].get(key, {})

    # Check completion date in range
    completion = _story_completion_date(issue, actual)
    if completion and from_date <= completion <= to_date:
        return True

    # Check story date span overlap (requires both start and due)
    start = _parse_iso_date(issue.get("start_date"))
    due = _parse_iso_date(issue.get("due_date"))
    if start and due and start <= to_date and due >= from_date:
        return True

    # Check if any subtask has worklogs in the range
    worklogs_dated = ctx.get("worklogs_dated") or {}
    for child in ctx["children_by_story"].get(key, []):
        child_key = _to_text(child.get("issue_key")).upper()
        for wl_date_str, _ in worklogs_dated.get(child_key, []):
            wl_date = _parse_iso_date(wl_date_str)
            if wl_date and from_date <= wl_date <= to_date:
                return True

    return False


def _subtask_hours_in_range(ctx: dict[str, Any], story_key: str, from_date: date, to_date: date) -> tuple[float, int, int]:
    """Sum logged hours on subtasks WHERE worklog date is in [from_date, to_date]."""
    worklogs_dated = ctx.get("worklogs_dated") or {}
    actuals = ctx["actuals"]
    total_hours = 0.0
    subtask_count = 0
    done_subtasks = 0
    for child in ctx["children_by_story"].get(story_key.upper(), []):
        subtask_count += 1
        child_key = _to_text(child.get("issue_key")).upper()
        # Sum only worklogs within date range
        child_hours = 0.0
        for wl_date_str, wl_hours in worklogs_dated.get(child_key, []):
            wl_date = _parse_iso_date(wl_date_str)
            if wl_date and from_date <= wl_date <= to_date:
                child_hours += wl_hours
        # If no dated worklogs available, fall back to total (for backwards compat)
        if not child_hours and not worklogs_dated.get(child_key):
            actual = actuals.get(child_key, {})
            child_hours = _to_float(actual.get("total_worklog_hours"))
        total_hours += child_hours
        if _is_done(child.get("status")):
            done_subtasks += 1
    return total_hours, subtask_count, done_subtasks


def _roster_from_booking(ctx: dict[str, Any], booking_stories: list[dict[str, Any]]) -> list[dict[str, Any]]:
    roster: list[dict[str, Any]] = []
    for story in booking_stories:
        summary = _to_text(story.get("summary"))
        # Try parenthesized month first: "Support by Name (June 2026)"
        month_match = re.search(r"\(([^)]+)\)\s*$", summary)
        booked_for = month_match.group(1).strip() if month_match else ""
        # Extract assignee name from "Support by <name>"
        if month_match:
            name_match = re.search(r"support\s+by\s+(.+?)\s*\(", summary, re.IGNORECASE)
        else:
            name_match = re.search(r"support\s+by\s+(.+?)\s+(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)", summary, re.IGNORECASE)
        # If no parenthesized month, try trailing "Month Year" pattern
        if not booked_for:
            trailing_month = re.search(
                r"((?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\s+\d{4})\s*$",
                summary, re.IGNORECASE
            )
            if trailing_month:
                booked_for = trailing_month.group(1).strip()
        roster.append({
            "issue_key": _to_text(story.get("issue_key")).upper(),
            "project_key": _to_text(story.get("project_key")).upper(),
            "assignee": _to_text(story.get("assignee")) or (name_match.group(1).strip() if name_match else ""),
            "booked_for": booked_for,
            "summary": summary,
        })
    roster.sort(key=lambda r: (r["project_key"], r["booked_for"], r["assignee"].lower()))
    return roster


def _roster_month_in_range(booked_for: str, from_date: date, to_date: date) -> bool:
    """Return True if the 'Month YYYY' string overlaps [from_date, to_date]."""
    text = _to_text(booked_for)
    if not text:
        return True
    for fmt in ("%B %Y", "%b %Y"):
        try:
            parsed = datetime.strptime(text, fmt)
            month_start = date(parsed.year, parsed.month, 1)
            month_end = date(parsed.year, parsed.month, calendar.monthrange(parsed.year, parsed.month)[1])
            return month_start <= to_date and month_end >= from_date
        except ValueError:
            continue
    return True


def _load_project_name_map(canonical_db_path: Path) -> dict[str, str]:
    """Return {project_key.upper(): project_name} from managed_projects."""
    try:
        with sqlite3.connect(canonical_db_path) as conn:
            rows = conn.execute("SELECT project_key, project_name FROM managed_projects").fetchall()
        return {_to_text(r[0]).upper(): _to_text(r[1]) for r in rows if _to_text(r[0])}
    except sqlite3.Error:
        return {}


def _available_hours_for_range(
    canonical_db_path: Path,
    run_id: str,
    from_date: date,
    to_date: date,
) -> dict[str, Any]:
    """Sum support-team availability across the calendar months overlapping the range."""
    total_available = 0.0
    total_capacity = 0.0
    per_month: list[dict[str, Any]] = []
    members: list[str] = []
    for month in _months_in_range(from_date, to_date):
        m_start, m_end = _month_bounds(month)
        try:
            workforce = build_workforce_month_payload(canonical_db_path, m_start, m_end, run_id)
        except Exception:
            continue
        support = workforce.get("support_team") or {}
        avail = _to_float(support.get("total_availability_hours"))
        cap = _to_float(support.get("total_capacity_hours"))
        total_available += avail
        total_capacity += cap
        if not members:
            members = support.get("saved_members") or []
        per_month.append({
            "month": month,
            "available_hours": _round_hours(avail),
            "capacity_hours": _round_hours(cap),
        })
    return {
        "support_members": members,
        "support_member_count": len(members),
        "available_hours": _round_hours(total_available),
        "available_days": _round_hours(total_available / HOURS_PER_DAY),
        "capacity_hours": _round_hours(total_capacity),
        "capacity_days": _round_hours(total_capacity / HOURS_PER_DAY),
        "per_month": per_month,
    }


def build_support_center_overview(
    canonical_db_path: Path,
    support_db_path: Path,
    run_id: str,
    from_date: date,
    to_date: date,
    selected_projects: set[str] | None = None,
) -> dict[str, Any]:
    ctx = _build_context(canonical_db_path, support_db_path, run_id)
    actual_stories, booking_stories = _classify_support_stories(ctx)
    project_name_map = _load_project_name_map(canonical_db_path)

    project_filter = {p.upper() for p in selected_projects} if selected_projects else None

    in_range_rows: list[dict[str, Any]] = []
    by_project: dict[str, dict[str, Any]] = {}
    for story in actual_stories:
        row = _story_row(ctx, story)
        if project_filter and row["project_key"] not in project_filter:
            continue
        if not _story_overlaps_range(story, ctx, from_date, to_date):
            continue
        # Recalculate hours using only worklogs within the date range
        hours_in_range, _, _ = _subtask_hours_in_range(ctx, row["issue_key"], from_date, to_date)
        row["invested_hours"] = _round_hours(hours_in_range)
        row["invested_days"] = _round_hours(hours_in_range / HOURS_PER_DAY)
        in_range_rows.append(row)
        agg = by_project.setdefault(row["project_key"], {
            "project_key": row["project_key"],
            "story_count": 0,
            "resolved_count": 0,
            "invested_hours": 0.0,
            "subtask_count": 0,
        })
        agg["story_count"] += 1
        agg["resolved_count"] += 1 if row["is_resolved"] else 0
        agg["invested_hours"] += row["invested_hours"]
        agg["subtask_count"] += row["subtask_count"]

    available = _available_hours_for_range(canonical_db_path, ctx["run_id"], from_date, to_date)

    total_invested = _round_hours(sum(r["invested_hours"] for r in in_range_rows))
    resolved_count = sum(1 for r in in_range_rows if r["is_resolved"])

    roster = _roster_from_booking(ctx, booking_stories)
    roster = [r for r in roster if _roster_month_in_range(r["booked_for"], from_date, to_date)]
    if project_filter:
        roster = [r for r in roster if r["project_key"] in project_filter]
    for r in roster:
        r["project_name"] = project_name_map.get(r["project_key"], r["project_key"])

    # Also count invested hours from booking stories' subtasks within the range
    booking_invested = 0.0
    for story in booking_stories:
        pk = _to_text(story.get("project_key")).upper()
        if project_filter and pk not in project_filter:
            continue
        story_key = _to_text(story.get("issue_key")).upper()
        hours_in_range, subtask_count, _ = _subtask_hours_in_range(ctx, story_key, from_date, to_date)
        if hours_in_range > 0:
            booking_invested += hours_in_range
            agg = by_project.setdefault(pk, {
                "project_key": pk,
                "story_count": 0,
                "resolved_count": 0,
                "invested_hours": 0.0,
                "subtask_count": 0,
            })
            agg["invested_hours"] += hours_in_range
            agg["subtask_count"] += subtask_count

    total_invested = _round_hours(total_invested + booking_invested)

    project_rows = []
    for pk in sorted(by_project.keys()):
        agg = by_project[pk]
        project_rows.append({
            "project_key": pk,
            "project_name": project_name_map.get(pk, pk),
            "story_count": agg["story_count"],
            "resolved_count": agg["resolved_count"],
            "invested_hours": _round_hours(agg["invested_hours"]),
            "invested_days": _round_hours(agg["invested_hours"] / HOURS_PER_DAY),
            "subtask_count": agg["subtask_count"],
        })

    return {
        "range": {"from": from_date.isoformat(), "to": to_date.isoformat()},
        "birds_eye": {
            "available_hours": available["available_hours"],
            "available_days": available["available_days"],
            "capacity_hours": available["capacity_hours"],
            "invested_hours": total_invested,
            "invested_days": _round_hours(total_invested / HOURS_PER_DAY),
            "resolved_support_stories": resolved_count,
            "support_story_count": len(in_range_rows),
            "support_member_count": available["support_member_count"],
            "utilization_pct": (
                _round_hours(100.0 * total_invested / available["available_hours"])
                if available["available_hours"] > 0 else 0.0
            ),
        },
        "support_members": available["support_members"],
        "availability_per_month": available["per_month"],
        "by_project": project_rows,
        "roster": roster,
        "run_id": ctx["run_id"],
        "project_names": project_name_map,
    }


def build_support_center_project_detail(
    canonical_db_path: Path,
    support_db_path: Path,
    run_id: str,
    project_key: str,
    from_date: date,
    to_date: date,
) -> dict[str, Any]:
    ctx = _build_context(canonical_db_path, support_db_path, run_id)
    actual_stories, booking_stories = _classify_support_stories(ctx)
    pk = _to_text(project_key).upper()

    stories_out: list[dict[str, Any]] = []
    for story in actual_stories:
        row = _story_row(ctx, story)
        if row["project_key"] != pk:
            continue
        if not _story_overlaps_range(story, ctx, from_date, to_date):
            continue
        # Recalculate hours using only worklogs within the date range
        hours_in_range, _, _ = _subtask_hours_in_range(ctx, row["issue_key"], from_date, to_date)
        row["invested_hours"] = _round_hours(hours_in_range)
        row["invested_days"] = _round_hours(hours_in_range / HOURS_PER_DAY)
        story_key = row["issue_key"]
        subtasks = []
        for child in ctx["children_by_story"].get(story_key, []):
            child_key = _to_text(child.get("issue_key")).upper()
            actual = ctx["actuals"].get(child_key, {})
            subtasks.append({
                "issue_key": child_key,
                "issue_type": _to_text(child.get("issue_type")),
                "summary": _to_text(child.get("summary")),
                "status": _to_text(child.get("status")),
                "assignee": _to_text(child.get("assignee")),
                "logged_hours": _round_hours(_to_float(actual.get("total_worklog_hours")) or _to_float(child.get("total_hours_logged"))),
                "actual_complete_date": _to_text(actual.get("actual_complete_date")),
            })
        subtasks.sort(key=lambda s: s["issue_key"])
        row_out = {k: v for k, v in row.items() if not k.startswith("_")}
        row_out["subtasks"] = subtasks
        stories_out.append(row_out)

    stories_out.sort(key=lambda s: s["actual_completion_date"] or "")
    roster = [r for r in _roster_from_booking(ctx, booking_stories) if r["project_key"] == pk]
    roster = [r for r in roster if _roster_month_in_range(r["booked_for"], from_date, to_date)]
    project_name_map = _load_project_name_map(canonical_db_path)
    for r in roster:
        r["project_name"] = project_name_map.get(r["project_key"], r["project_key"])

    total_invested = _round_hours(sum(s["invested_hours"] for s in stories_out))
    resolved_count = sum(1 for s in stories_out if s["is_resolved"])

    return {
        "project_key": pk,
        "range": {"from": from_date.isoformat(), "to": to_date.isoformat()},
        "summary": {
            "story_count": len(stories_out),
            "resolved_count": resolved_count,
            "invested_hours": total_invested,
            "invested_days": _round_hours(total_invested / HOURS_PER_DAY),
        },
        "stories": stories_out,
        "roster": roster,
        "run_id": ctx["run_id"],
    }
