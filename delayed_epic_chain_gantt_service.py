from __future__ import annotations

import json
import sqlite3
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any


DEFAULT_ASSIGNEE_MODE = "subtask_assignee"
VALID_ASSIGNEE_MODES = {
    "subtask_assignee": "Subtask Assignee",
    "story_assignee": "Story Assignee",
    "epic_assignee": "Epic Assignee",
}
DEFAULT_UI_SETTINGS = {
    "show_full_year": True,
    "week_bucket_width_px": 100,
}
MIN_WEEK_BUCKET_WIDTH_PX = 60
MAX_WEEK_BUCKET_WIDTH_PX = 220


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


def _parse_plan_json(value: Any) -> dict[str, Any]:
    text = _to_text(value)
    if not text:
        return {}
    try:
        parsed = json.loads(text)
    except Exception:
        return {}
    return parsed if isinstance(parsed, dict) else {}


def _normalize_assignee(value: Any) -> str:
    return _to_text(value).casefold()


def _is_resolved_status_text(value: Any) -> bool:
    text = _to_text(value).strip().lower()
    if not text:
        return False
    return text in {"resolved", "resolved!", "done", "closed", "complete", "completed"}


def _is_in_progress_status_text(value: Any) -> bool:
    text = _to_text(value).strip().lower()
    return any(token in text for token in ("progress", "hold", "review", "testing", "qa"))


def _status_icon(value: Any) -> str:
    if _is_resolved_status_text(value):
        return "check_circle"
    if _is_in_progress_status_text(value):
        return "hourglass_top"
    return "help"


def _is_story_type(value: Any) -> bool:
    return _to_text(value).strip().lower() == "story"


def _is_subtask_type(value: Any) -> bool:
    low = _to_text(value).strip().lower()
    return "sub-task" in low or "subtask" in low


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def normalize_assignee_mode(value: Any) -> str:
    mode = _to_text(value).lower() or DEFAULT_ASSIGNEE_MODE
    if mode not in VALID_ASSIGNEE_MODES:
        raise ValueError(
            f"Invalid assignee_mode. Expected one of: {', '.join(sorted(VALID_ASSIGNEE_MODES))}."
        )
    return mode


def init_ui_settings_db(db_path: Path) -> None:
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS delayed_epic_chain_gantt_ui_settings (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                show_full_year INTEGER NOT NULL DEFAULT 1,
                week_bucket_width_px INTEGER NOT NULL DEFAULT 100,
                updated_at_utc TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def _normalize_week_bucket_width(value: Any) -> int:
    try:
        width = int(float(_to_text(value)))
    except Exception:
        raise ValueError("week_bucket_width_px must be a number.")
    if width < MIN_WEEK_BUCKET_WIDTH_PX or width > MAX_WEEK_BUCKET_WIDTH_PX:
        raise ValueError(
            f"week_bucket_width_px must be between {MIN_WEEK_BUCKET_WIDTH_PX} and {MAX_WEEK_BUCKET_WIDTH_PX}."
        )
    return width


def load_ui_settings(db_path: Path) -> dict[str, Any]:
    init_ui_settings_db(db_path)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        row = conn.execute(
            """
            SELECT show_full_year, week_bucket_width_px, updated_at_utc
            FROM delayed_epic_chain_gantt_ui_settings
            WHERE id = 1
            """
        ).fetchone()
    finally:
        conn.close()
    if not row:
        return {**DEFAULT_UI_SETTINGS, "updated_at_utc": ""}
    week_width = int(row["week_bucket_width_px"] or DEFAULT_UI_SETTINGS["week_bucket_width_px"])
    if week_width < MIN_WEEK_BUCKET_WIDTH_PX or week_width > MAX_WEEK_BUCKET_WIDTH_PX:
        week_width = DEFAULT_UI_SETTINGS["week_bucket_width_px"]
    return {
        "show_full_year": bool(int(row["show_full_year"] or 0)),
        "week_bucket_width_px": week_width,
        "updated_at_utc": _to_text(row["updated_at_utc"]),
    }


def save_ui_settings(db_path: Path, payload: object) -> dict[str, Any]:
    init_ui_settings_db(db_path)
    raw = payload if isinstance(payload, dict) else {}
    show_full_year = bool(raw.get("show_full_year", DEFAULT_UI_SETTINGS["show_full_year"]))
    week_bucket_width_px = _normalize_week_bucket_width(
        raw.get("week_bucket_width_px", DEFAULT_UI_SETTINGS["week_bucket_width_px"])
    )
    updated_at = _utc_now_iso()
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            INSERT INTO delayed_epic_chain_gantt_ui_settings(id, show_full_year, week_bucket_width_px, updated_at_utc)
            VALUES (1, ?, ?, ?)
            ON CONFLICT(id) DO UPDATE SET
                show_full_year = excluded.show_full_year,
                week_bucket_width_px = excluded.week_bucket_width_px,
                updated_at_utc = excluded.updated_at_utc
            """,
            (1 if show_full_year else 0, week_bucket_width_px, updated_at),
        )
        conn.commit()
    finally:
        conn.close()
    return {
        "show_full_year": show_full_year,
        "week_bucket_width_px": week_bucket_width_px,
        "updated_at_utc": updated_at,
    }


def _load_latest_run_id(conn: sqlite3.Connection) -> str:
    row = conn.execute(
        "SELECT last_success_run_id FROM canonical_refresh_state WHERE id = 1"
    ).fetchone()
    run_id = _to_text(row[0] if row else "")
    if not run_id:
        raise ValueError("No successful canonical refresh found.")
    return run_id


def _load_latest_epf_run_id(conn: sqlite3.Connection) -> str:
    row = conn.execute(
        "SELECT last_success_run_id FROM epf_refresh_state WHERE id = 1"
    ).fetchone()
    return _to_text(row[0] if row else "")


def list_filter_options(db_path: Path) -> dict[str, Any]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        run_id = _load_latest_run_id(conn)
        rows = conn.execute(
            """
            SELECT DISTINCT assignee
            FROM canonical_issues
            WHERE run_id = ? AND trim(assignee) <> ''
            ORDER BY lower(trim(assignee))
            """,
            (run_id,),
        ).fetchall()
    finally:
        conn.close()
    return {
        "assignees": [_to_text(row["assignee"]) for row in rows],
        "assignee_modes": [
            {"value": key, "label": label}
            for key, label in VALID_ASSIGNEE_MODES.items()
        ],
        "default_assignee_mode": DEFAULT_ASSIGNEE_MODE,
    }


def _load_epic_plan_rows(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    rows = conn.execute(
        """
        SELECT epic_key, project_key, project_name, epic_name, jira_url, delivery_status, epic_plan_json
        FROM epics_management
        WHERE trim(epic_key) <> ''
        ORDER BY upper(epic_key)
        """
    ).fetchall()
    out: list[dict[str, Any]] = []
    for row in rows:
        plan = _parse_plan_json(row["epic_plan_json"])
        planned_start = _parse_iso_date(plan.get("start_date"))
        planned_due = _parse_iso_date(plan.get("due_date"))
        if planned_start is None or planned_due is None:
            continue
        out.append(
            {
                "epic_key": _to_text(row["epic_key"]).upper(),
                "project_key": _to_text(row["project_key"]).upper(),
                "project_name": _to_text(row["project_name"]),
                "epic_name": _to_text(row["epic_name"]),
                "jira_url": _to_text(row["jira_url"]),
                "delivery_status": _to_text(row["delivery_status"]),
                "planned_start": planned_start,
                "planned_due": planned_due,
                "planned_man_days": float(plan.get("man_days") or 0.0),
            }
        )
    return out


def _build_issue_maps(
    conn: sqlite3.Connection,
    run_id: str,
) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]]]:
    issue_rows = conn.execute(
        """
        SELECT issue_key, project_key, issue_type, summary, status, assignee, start_date, due_date,
               resolved_stable_since_date, parent_issue_key, story_key, epic_key
        FROM canonical_issues
        WHERE run_id = ?
        """,
        (run_id,),
    ).fetchall()
    link_rows = conn.execute(
        """
        SELECT issue_key, parent_issue_key, story_key, epic_key, hierarchy_level
        FROM canonical_issue_links
        WHERE run_id = ?
        """,
        (run_id,),
    ).fetchall()
    actual_rows = conn.execute(
        """
        SELECT issue_key, assignee, first_worklog_date, last_worklog_date, actual_complete_date,
               actual_complete_source, due_completion_bucket, total_worklog_hours, worklog_count
        FROM canonical_issue_actuals
        WHERE run_id = ?
        """,
        (run_id,),
    ).fetchall()

    links_by_issue = {
        _to_text(row["issue_key"]).upper(): dict(row)
        for row in link_rows
    }
    actuals_by_issue = {
        _to_text(row["issue_key"]).upper(): dict(row)
        for row in actual_rows
    }

    issues_by_key: dict[str, dict[str, Any]] = {}
    for row in issue_rows:
        issue_key = _to_text(row["issue_key"]).upper()
        link = links_by_issue.get(issue_key, {})
        actual = actuals_by_issue.get(issue_key, {})
        issue = dict(row)
        issue["issue_key"] = issue_key
        issue["project_key"] = _to_text(row["project_key"]).upper()
        issue["parent_issue_key"] = _to_text(row["parent_issue_key"]).upper() or _to_text(link.get("parent_issue_key")).upper()
        issue["story_key"] = _to_text(row["story_key"]).upper() or _to_text(link.get("story_key")).upper()
        issue["epic_key"] = _to_text(row["epic_key"]).upper() or _to_text(link.get("epic_key")).upper()
        issue["first_worklog_date"] = _parse_iso_date(actual.get("first_worklog_date"))
        issue["last_worklog_date"] = _parse_iso_date(actual.get("last_worklog_date"))
        issue["actual_complete_date"] = _parse_iso_date(actual.get("actual_complete_date"))
        issue["actual_complete_source"] = _to_text(actual.get("actual_complete_source"))
        issue["total_worklog_hours"] = float(actual.get("total_worklog_hours") or 0.0)
        issue["worklog_count"] = int(actual.get("worklog_count") or 0)
        issues_by_key[issue_key] = issue
    return issues_by_key, actuals_by_issue


def _load_leave_map(conn: sqlite3.Connection) -> dict[str, set[date]]:
    run_id = _load_latest_epf_run_id(conn)
    if not run_id:
        return {}
    rows = conn.execute(
        """
        SELECT assignee, period_day, unplanned_taken_hours, planned_taken_hours
        FROM epf_leave_rows
        WHERE run_id = ?
        """,
        (run_id,),
    ).fetchall()
    out: dict[str, set[date]] = {}
    for row in rows:
        day = _parse_iso_date(row["period_day"])
        if day is None:
            continue
        if float(row["unplanned_taken_hours"] or 0.0) <= 0 and float(row["planned_taken_hours"] or 0.0) <= 0:
            continue
        assignee_key = _normalize_assignee(row["assignee"])
        if not assignee_key:
            continue
        out.setdefault(assignee_key, set()).add(day)
    return out


def _matched_assignees(issue_rows: list[dict[str, Any]], mode: str, epic_row: dict[str, Any]) -> list[str]:
    assignees: set[str] = set()
    if mode == "epic_assignee":
        epic_assignee = _to_text(epic_row.get("assignee"))
        if epic_assignee:
            assignees.add(epic_assignee)
    elif mode == "story_assignee":
        for issue in issue_rows:
            if _is_story_type(issue.get("issue_type")) and _to_text(issue.get("assignee")):
                assignees.add(_to_text(issue.get("assignee")))
    else:
        for issue in issue_rows:
            if _is_subtask_type(issue.get("issue_type")) and _to_text(issue.get("assignee")):
                assignees.add(_to_text(issue.get("assignee")))
    return sorted(assignees, key=lambda item: item.casefold())


def _range_overlap(start_a: date, end_a: date, start_b: date, end_b: date) -> bool:
    return start_a <= end_b and start_b <= end_a


def _find_leave_overlap(
    assignees: list[str],
    leave_days_by_assignee: dict[str, set[date]],
    planned_start: date,
    actual_start: date | None,
) -> bool:
    gap_end = actual_start if actual_start and actual_start >= planned_start else planned_start
    for assignee in assignees:
        days = leave_days_by_assignee.get(_normalize_assignee(assignee), set())
        if not days:
            continue
        day_cursor = planned_start
        while day_cursor <= gap_end:
            if day_cursor in days:
                return True
            day_cursor = day_cursor.fromordinal(day_cursor.toordinal() + 1)
    return False


def _build_epic_rows(
    conn: sqlite3.Connection,
    run_id: str,
    assignee_mode: str,
    selected_assignee: str,
) -> list[dict[str, Any]]:
    planner_rows = _load_epic_plan_rows(conn)
    issues_by_key, _actuals_by_issue = _build_issue_maps(conn, run_id)
    leave_days_by_assignee = _load_leave_map(conn)
    selected_assignee_norm = _normalize_assignee(selected_assignee)
    out: list[dict[str, Any]] = []

    for planner in planner_rows:
        epic_key = planner["epic_key"]
        epic_issue = issues_by_key.get(epic_key, {})
        epic_family = [
            issue
            for issue in issues_by_key.values()
            if _to_text(issue.get("issue_key")).upper() == epic_key
            or _to_text(issue.get("epic_key")).upper() == epic_key
        ]
        actual_starts = [
            issue["first_worklog_date"]
            for issue in epic_family
            if isinstance(issue.get("first_worklog_date"), date)
        ]
        actual_completions = [
            issue["actual_complete_date"]
            for issue in epic_family
            if isinstance(issue.get("actual_complete_date"), date)
        ]
        matched_assignees = _matched_assignees(epic_family, assignee_mode, epic_issue)
        if selected_assignee_norm and selected_assignee_norm not in {
            _normalize_assignee(name) for name in matched_assignees
        }:
            continue
        status = _to_text(epic_issue.get("status")) or planner["delivery_status"] or "To Do"
        actual_start = min(actual_starts) if actual_starts else None
        actual_complete = max(actual_completions) if actual_completions else None
        unresolved = not _is_resolved_status_text(status)
        late = bool(actual_complete and actual_complete > planner["planned_due"]) or unresolved
        leave_overlap = _find_leave_overlap(
            matched_assignees,
            leave_days_by_assignee,
            planner["planned_start"],
            actual_start,
        )
        out.append(
            {
                "epic_key": epic_key,
                "epic_name": planner["epic_name"] or epic_key,
                "project_key": planner["project_key"] or _to_text(epic_issue.get("project_key")).upper(),
                "project_name": planner["project_name"],
                "jira_url": planner["jira_url"],
                "planned_start": planner["planned_start"],
                "planned_due": planner["planned_due"],
                "planned_man_days": round(float(planner["planned_man_days"] or 0.0), 2),
                "status": status,
                "status_icon": _status_icon(status),
                "epic_assignee": _to_text(epic_issue.get("assignee")),
                "matched_assignee_names": matched_assignees,
                "actual_start": actual_start,
                "actual_complete_date": actual_complete,
                "unresolved": unresolved,
                "is_late": late,
                "leave_overlap": leave_overlap,
                "assignee_mode": assignee_mode,
            }
        )
    return sorted(
        out,
        key=lambda row: (
            row["planned_start"].isoformat(),
            row["planned_due"].isoformat(),
            row["project_key"],
            row["epic_key"],
        ),
    )


def _annotate_delay_chain(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    sorted_rows = sorted(
        rows,
        key=lambda row: (
            row["planned_start"].isoformat(),
            row["planned_due"].isoformat(),
            row["epic_key"],
        ),
    )
    for index, row in enumerate(sorted_rows):
        shared_assignee_set = {_normalize_assignee(name) for name in row["matched_assignee_names"]}
        blocking: dict[str, Any] | None = None
        for candidate in sorted_rows[:index]:
            if candidate["epic_key"] == row["epic_key"] or not candidate["is_late"]:
                continue
            candidate_assignee_set = {_normalize_assignee(name) for name in candidate["matched_assignee_names"]}
            if not shared_assignee_set or not candidate_assignee_set or not (shared_assignee_set & candidate_assignee_set):
                continue
            candidate_end = candidate["actual_complete_date"]
            if candidate_end is None:
                continue
            if candidate["planned_start"] > row["planned_start"]:
                continue
            if candidate_end < row["planned_start"]:
                continue
            if blocking is None or candidate["planned_start"] > blocking["planned_start"]:
                blocking = candidate
        row["blocking_epic_key"] = _to_text(blocking["epic_key"]) if blocking else ""
        row["blocking_epic_url"] = _to_text(blocking["jira_url"]) if blocking else ""
        row["delay_cause"] = ""
        if blocking is not None:
            row["delay_cause"] = "Previous Epic Delayed"
        elif row["leave_overlap"]:
            row["delay_cause"] = "Leave Overlap"
        elif row["actual_start"] and row["actual_start"] > row["planned_start"]:
            row["delay_cause"] = "Late Start"
        row["is_impacted"] = bool(row["delay_cause"])
    return sorted_rows


def build_report_payload(
    db_path: Path,
    from_date: str,
    to_date: str,
    *,
    assignee: str = "",
    assignee_mode: str = DEFAULT_ASSIGNEE_MODE,
) -> dict[str, Any]:
    range_start = _parse_iso_date(from_date)
    range_end = _parse_iso_date(to_date)
    if range_start is None or range_end is None:
        raise ValueError("Invalid 'from'/'to' date range. Expected YYYY-MM-DD.")
    if range_end < range_start:
        raise ValueError("'to' must be on or after 'from'.")

    mode = normalize_assignee_mode(assignee_mode)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        run_id = _load_latest_run_id(conn)
        filter_options = list_filter_options(db_path)
        epic_rows = _annotate_delay_chain(_build_epic_rows(conn, run_id, mode, assignee))
    finally:
        conn.close()

    selected_rows: list[dict[str, Any]] = []
    for row in epic_rows:
        planned_start: date = row["planned_start"]
        planned_due: date = row["planned_due"]
        predecessor = (
            (planned_start < range_start or planned_due < range_start)
            and row["is_late"]
        )
        planned_in_selected_range = _range_overlap(planned_start, planned_due, range_start, range_end)
        impacted = planned_due >= range_start and row["is_impacted"]
        if not predecessor and not impacted and not planned_in_selected_range:
            continue
        selected_rows.append(
            {
                **row,
                "planned_start": planned_start.isoformat(),
                "planned_due": planned_due.isoformat(),
                "actual_start": row["actual_start"].isoformat() if row["actual_start"] else "",
                "actual_complete_date": row["actual_complete_date"].isoformat() if row["actual_complete_date"] else "",
            }
        )

    return {
        "rows": selected_rows,
        "filter_options": filter_options,
        "selected": {
            "from_date": range_start.isoformat(),
            "to_date": range_end.isoformat(),
            "assignee": _to_text(assignee),
            "assignee_mode": mode,
        },
        "meta": {
            "generated_at_utc": _utc_now_iso(),
            "timeline_year": range_start.year,
            "row_count": len(selected_rows),
        },
    }
