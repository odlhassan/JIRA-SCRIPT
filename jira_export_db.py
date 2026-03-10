"""
SQLite storage for Jira export pipeline data.

Database: jira_exports.db (path from JIRA_EXPORTS_DB_PATH).
Primary storage for worklogs, work items, rollup, and nested view nodes.
"""
from __future__ import annotations

import os
import sqlite3
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from ipp_meeting_utils import normalize_issue_key
DEFAULT_EXPORTS_DB = "jira_exports.db"

# Column names for subtask_worklogs (16 cols) - snake_case for SQL
SUBTASK_WORKLOGS_COLS = [
    "issue_link",
    "issue_id",
    "issue_title",
    "issue_type",
    "parent_story_link",
    "parent_story_id",
    "parent_epic_id",
    "issue_assignee",
    "latest_ipp_meeting",
    "jira_ipp_rmi_dates_altered",
    "ipp_actual_date",
    "ipp_remarks",
    "ipp_actual_date_matches_jira_end_date",
    "worklog_started",
    "hours_logged",
    "worklog_author",
]

# Column names for work_items (29 cols) - match xlsx headers (already snake_case or with spaces; we use snake_case in DB)
WORK_ITEMS_COLS = [
    "project_key",
    "issue_key",
    "work_item_id",
    "work_item_type",
    "jira_issue_type",
    "fix_type",
    "summary",
    "status",
    "resolved_stable_since_date",
    "start_date",
    "end_date",
    "actual_start_date",
    "actual_end_date",
    "original_estimate",
    "original_estimate_hours",
    "assignee",
    "total_hours_logged",
    "priority",
    "parent_issue_key",
    "parent_work_item_id",
    "parent_jira_url",
    "jira_url",
    "latest_ipp_meeting",
    "jira_ipp_rmi_dates_altered",
    "ipp_actual_date",
    "ipp_remarks",
    "ipp_actual_date_matches_jira_end_date",
    "created",
    "updated",
]

# Column names for subtask_worklog_rollup (18 cols) - snake_case
SUBTASK_ROLLUP_COLS = [
    "issue_link",
    "issue_id",
    "issue_title",
    "issue_type",
    "parent_story_link",
    "parent_story_id",
    "parent_epic_id",
    "issue_assignee",
    "latest_ipp_meeting",
    "jira_ipp_rmi_dates_altered",
    "ipp_actual_date",
    "ipp_remarks",
    "ipp_actual_date_matches_jira_end_date",
    "planned_start_date",
    "planned_end_date",
    "actual_start_date",
    "actual_end_date",
    "total_hours_logged",
]

# Column names for nested_view_nodes (flattened IssueNode)
NESTED_VIEW_NODE_COLS = [
    "key",
    "kind",
    "project_key",
    "summary",
    "parent_key",
    "assignee",
    "product_category",
    "epic_key",
    "story_key",
    "man_hours",
    "man_days",
    "actual_hours",
    "actual_days",
    "planned_start",
    "planned_end",
]

# Input headers expected by rollup script (from worklogs) - keys for read_subtask_worklogs dicts
ROLLUP_INPUT_HEADERS = [
    "issue_link",
    "issue_id",
    "issue_title",
    "issue_type",
    "parent_story_link",
    "parent_story_id",
    "parent_epic_id",
    "issue_assignee",
    "worklog_started",
    "hours_logged",
]

# Map from DB column name to rollup input key (same for these)
def _worklog_row_to_rollup_input(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "issue_link": row["issue_link"] or "",
        "issue_id": row["issue_id"] or "",
        "issue_title": row["issue_title"] or "",
        "issue_type": row["issue_type"] or "",
        "parent_story_link": row["parent_story_link"] or "",
        "parent_story_id": row["parent_story_id"] or "",
        "parent_epic_id": row["parent_epic_id"] or "",
        "issue_assignee": row["issue_assignee"] or "",
        "worklog_started": row["worklog_started"] or "",
        "hours_logged": row["hours_logged"],
    }


def _parse_iso_utc(value: str | None) -> datetime | None:
    if not value or not str(value).strip():
        return None
    text = str(value).strip()
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(text)
        if dt.tzinfo is None:
            return dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except ValueError:
        return None


def get_exports_db_path() -> Path:
    raw = os.getenv("JIRA_EXPORTS_DB_PATH", DEFAULT_EXPORTS_DB).strip() or DEFAULT_EXPORTS_DB
    path = Path(raw)
    if path.is_absolute():
        return path
    return Path(__file__).resolve().parent / path


def connect() -> sqlite3.Connection:
    path = get_exports_db_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path))
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.row_factory = sqlite3.Row
    return conn


def ensure_schema(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS export_runs (
            run_id TEXT PRIMARY KEY,
            pipeline TEXT NOT NULL,
            started_at_utc TEXT NOT NULL,
            finished_at_utc TEXT NOT NULL,
            row_count INTEGER,
            status TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS subtask_worklogs (
            issue_link TEXT,
            issue_id TEXT,
            issue_title TEXT,
            issue_type TEXT,
            parent_story_link TEXT,
            parent_story_id TEXT,
            parent_epic_id TEXT,
            issue_assignee TEXT,
            latest_ipp_meeting TEXT,
            jira_ipp_rmi_dates_altered TEXT,
            ipp_actual_date TEXT,
            ipp_remarks TEXT,
            ipp_actual_date_matches_jira_end_date TEXT,
            worklog_started TEXT,
            hours_logged REAL,
            worklog_author TEXT
        );

        CREATE TABLE IF NOT EXISTS work_items (
            project_key TEXT,
            issue_key TEXT,
            work_item_id TEXT,
            work_item_type TEXT,
            jira_issue_type TEXT,
            fix_type TEXT,
            summary TEXT,
            status TEXT,
            resolved_stable_since_date TEXT,
            start_date TEXT,
            end_date TEXT,
            actual_start_date TEXT,
            actual_end_date TEXT,
            original_estimate TEXT,
            original_estimate_hours REAL,
            assignee TEXT,
            total_hours_logged REAL,
            priority TEXT,
            parent_issue_key TEXT,
            parent_work_item_id TEXT,
            parent_jira_url TEXT,
            jira_url TEXT,
            latest_ipp_meeting TEXT,
            jira_ipp_rmi_dates_altered TEXT,
            ipp_actual_date TEXT,
            ipp_remarks TEXT,
            ipp_actual_date_matches_jira_end_date TEXT,
            created TEXT,
            updated TEXT
        );

        CREATE TABLE IF NOT EXISTS subtask_worklog_rollup (
            issue_link TEXT,
            issue_id TEXT,
            issue_title TEXT,
            issue_type TEXT,
            parent_story_link TEXT,
            parent_story_id TEXT,
            parent_epic_id TEXT,
            issue_assignee TEXT,
            latest_ipp_meeting TEXT,
            jira_ipp_rmi_dates_altered TEXT,
            ipp_actual_date TEXT,
            ipp_remarks TEXT,
            ipp_actual_date_matches_jira_end_date TEXT,
            planned_start_date TEXT,
            planned_end_date TEXT,
            actual_start_date TEXT,
            actual_end_date TEXT,
            total_hours_logged REAL
        );

        CREATE TABLE IF NOT EXISTS nested_view_nodes (
            key TEXT,
            kind TEXT,
            project_key TEXT,
            summary TEXT,
            parent_key TEXT,
            assignee TEXT,
            product_category TEXT,
            epic_key TEXT,
            story_key TEXT,
            man_hours REAL,
            man_days REAL,
            actual_hours REAL,
            actual_days REAL,
            planned_start TEXT,
            planned_end TEXT
        );

        CREATE TABLE IF NOT EXISTS dashboard_refresh_runs (
            run_id TEXT PRIMARY KEY,
            mode TEXT NOT NULL,
            started_at_utc TEXT NOT NULL,
            ended_at_utc TEXT,
            status TEXT NOT NULL,
            progress_step TEXT,
            progress_pct INTEGER DEFAULT 0,
            progress_total INTEGER DEFAULT 0,
            progress_done INTEGER DEFAULT 0,
            progress_current_label TEXT DEFAULT '',
            cancel_requested INTEGER DEFAULT 0,
            error_message TEXT,
            stats_json TEXT,
            resume_from_run_id TEXT,
            completed_steps_json TEXT,
            updated_at_utc TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS dashboard_refresh_state (
            id INTEGER PRIMARY KEY CHECK(id = 1),
            active_run_id TEXT NOT NULL DEFAULT '',
            last_success_run_id TEXT NOT NULL DEFAULT '',
            updated_at_utc TEXT NOT NULL
        );
        """
    )
    conn.commit()
    _migrate_dashboard_refresh_runs_item_progress(conn)


def _migrate_dashboard_refresh_runs_item_progress(conn: sqlite3.Connection) -> None:
    """Add progress_total / progress_done / progress_current_label if missing (pre-existing DB)."""
    cols = {row[1].lower() for row in conn.execute("PRAGMA table_info(dashboard_refresh_runs)").fetchall()}
    for col, col_type, default in [
        ("progress_total", "INTEGER", "0"),
        ("progress_done", "INTEGER", "0"),
        ("progress_current_label", "TEXT", "''"),
    ]:
        if col not in cols:
            conn.execute(f"ALTER TABLE dashboard_refresh_runs ADD COLUMN {col} {col_type} DEFAULT {default}")
    conn.commit()


def _cell(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    return str(value).strip() if str(value).strip() else None


def write_subtask_worklogs(conn: sqlite3.Connection, rows: list[list[object]]) -> None:
    """Full-replace: delete all, then bulk insert. rows are 16-element lists in SUBTASK_WORKLOGS_COLS order."""
    conn.execute("DELETE FROM subtask_worklogs")
    if not rows:
        conn.commit()
        return
    placeholders = ",".join(["?"] * len(SUBTASK_WORKLOGS_COLS))
    sql = f"INSERT INTO subtask_worklogs ({','.join(SUBTASK_WORKLOGS_COLS)}) VALUES ({placeholders})"
    for row in rows:
        values = [_cell(row[i]) if i < len(row) else None for i in range(len(SUBTASK_WORKLOGS_COLS))]
        conn.execute(sql, values)
    conn.commit()


def read_subtask_worklogs(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    """Return list of dicts with keys matching rollup INPUT_HEADERS (for export_jira_subtask_worklog_rollup)."""
    cur = conn.execute(
        "SELECT issue_link, issue_id, issue_title, issue_type, parent_story_link, parent_story_id, "
        "parent_epic_id, issue_assignee, worklog_started, hours_logged FROM subtask_worklogs"
    )
    return [_worklog_row_to_rollup_input(row) for row in cur.fetchall()]


def read_worklogs_actual_dates(conn: sqlite3.Connection) -> tuple[dict[str, dict], dict[str, dict], dict[str, dict]]:
    """Return (subtask_dates, story_dates, epic_dates) each keyed by issue key with {min_dt, max_dt} (datetime)."""
    subtask_dates: dict[str, dict] = {}
    story_dates: dict[str, dict] = {}
    epic_dates: dict[str, dict] = {}

    cur = conn.execute(
        "SELECT issue_id, parent_story_id, parent_epic_id, worklog_started FROM subtask_worklogs WHERE worklog_started IS NOT NULL AND worklog_started != ''"
    )
    for row in cur.fetchall():
        issue_id = normalize_issue_key(str(row[0] or ""))
        parent_story_id = normalize_issue_key(str(row[1] or ""))
        parent_epic_id = normalize_issue_key(str(row[2] or ""))
        started_text = str(row[3] or "").strip()
        if not started_text:
            continue
        dt = _parse_iso_utc(started_text)
        if dt is None:
            continue

        def update(d: dict[str, dict], key: str) -> None:
            if not key:
                return
            if key not in d:
                d[key] = {"min_dt": dt, "max_dt": dt}
            else:
                if dt < d[key]["min_dt"]:
                    d[key]["min_dt"] = dt
                if dt > d[key]["max_dt"]:
                    d[key]["max_dt"] = dt

        update(subtask_dates, issue_id)
        update(story_dates, parent_story_id)
        update(epic_dates, parent_epic_id)

    return subtask_dates, story_dates, epic_dates


def _work_items_row_from_list(row: list) -> list:
    """Convert work items row (list of 29 values) to list for INSERT; original_estimate_hours and total_hours_logged may be float."""
    out = []
    for i in range(len(WORK_ITEMS_COLS)):
        v = row[i] if i < len(row) else None
        if v is None:
            out.append(None)
        elif isinstance(v, (int, float)):
            out.append(v)
        else:
            out.append(str(v).strip() or None)
    return out


def write_work_items(conn: sqlite3.Connection, rows: list[list]) -> None:
    """Full-replace work_items table. rows are 29-element lists in WORK_ITEMS_COLS order."""
    conn.execute("DELETE FROM work_items")
    if not rows:
        conn.commit()
        return
    placeholders = ",".join(["?"] * len(WORK_ITEMS_COLS))
    sql = f"INSERT INTO work_items ({','.join(WORK_ITEMS_COLS)}) VALUES ({placeholders})"
    for row in rows:
        values = _work_items_row_from_list(row)
        conn.execute(sql, values)
    conn.commit()


def write_subtask_worklog_rollup(conn: sqlite3.Connection, rows: list[list[object]]) -> None:
    """Full-replace. rows are 18-element lists; we map OUTPUT_HEADERS to SUBTASK_ROLLUP_COLS.
    OUTPUT_HEADERS has 'Latest IPP Meeting', 'planned start date', etc. Our DB uses snake_case.
    The rollup script produces: [issue_link, issue_id, issue_title, issue_type, parent_story_link, parent_story_id,
    parent_epic_id, issue_assignee, yes_no_in_ipp(...), yes_no_dates_altered(...), ipp_actual_date, ipp_remarks,
    yes_no_ipp_actual_matches_jira_end(...), planned_start, planned_end, actual_min_raw, actual_max_raw, total_hours]
    So indices 0-7 same, 8=latest_ipp_meeting, 9=jira_ipp_rmi_dates_altered, 10=ipp_actual_date, 11=ipp_remarks,
    12=ipp_actual_date_matches_jira_end_date, 13=planned_start_date, 14=planned_end_date, 15=actual_start_date,
    16=actual_end_date, 17=total_hours_logged.
    """
    conn.execute("DELETE FROM subtask_worklog_rollup")
    if not rows:
        conn.commit()
        return
    placeholders = ",".join(["?"] * len(SUBTASK_ROLLUP_COLS))
    sql = f"INSERT INTO subtask_worklog_rollup ({','.join(SUBTASK_ROLLUP_COLS)}) VALUES ({placeholders})"
    for row in rows:
        values = [_cell(row[i]) if i < len(row) else None for i in range(len(SUBTASK_ROLLUP_COLS))]
        conn.execute(sql, values)
    conn.commit()


def read_subtask_worklog_rollup(conn: sqlite3.Connection) -> dict[str, dict[str, object]]:
    """Return dict keyed by issue_id (normalized) with {actual_hours, planned_start, planned_end} for nested view."""
    result: dict[str, dict[str, object]] = {}
    cur = conn.execute(
        "SELECT issue_id, total_hours_logged, planned_start_date, planned_end_date FROM subtask_worklog_rollup"
    )
    for row in cur.fetchall():
        issue_id = normalize_issue_key(str(row[0] or ""))
        if not issue_id:
            continue
        try:
            actual_hours = round(float(row[1] or 0), 2)
        except (TypeError, ValueError):
            actual_hours = 0.0
        result[issue_id] = {
            "actual_hours": actual_hours,
            "planned_start": str(row[2] or "").strip(),
            "planned_end": str(row[3] or "").strip(),
        }
    return result


def write_nested_view_nodes(
    conn: sqlite3.Connection,
    epics: dict[str, Any],
    stories: dict[str, Any],
    subtasks: dict[str, Any],
) -> None:
    """Full-replace nested_view_nodes from IssueNode-like dicts (each value has key, kind, project_key, summary, parent_key, assignee, product_category, epic_key, story_key, man_hours, man_days, actual_hours, planned_start, planned_end)."""
    conn.execute("DELETE FROM nested_view_nodes")
    rows = []
    for node in list(epics.values()) + list(stories.values()) + list(subtasks.values()):
        key = getattr(node, "key", None) or (node.get("key") if isinstance(node, dict) else None)
        kind = getattr(node, "kind", None) or (node.get("kind") if isinstance(node, dict) else None)
        project_key = getattr(node, "project_key", None) or (node.get("project_key") if isinstance(node, dict) else None)
        summary = getattr(node, "summary", None) or (node.get("summary") if isinstance(node, dict) else None)
        parent_key = getattr(node, "parent_key", None) or (node.get("parent_key") if isinstance(node, dict) else None)
        assignee = getattr(node, "assignee", None) or (node.get("assignee") if isinstance(node, dict) else None)
        product_category = getattr(node, "product_category", None) or (node.get("product_category") if isinstance(node, dict) else None)
        epic_key = getattr(node, "epic_key", None) or (node.get("epic_key") if isinstance(node, dict) else None)
        story_key = getattr(node, "story_key", None) or (node.get("story_key") if isinstance(node, dict) else None)
        man_hours = getattr(node, "man_hours", None) or (node.get("man_hours") if isinstance(node, dict) else None)
        man_days = getattr(node, "man_days", None) or (node.get("man_days") if isinstance(node, dict) else None)
        actual_hours = getattr(node, "actual_hours", None) or (node.get("actual_hours") if isinstance(node, dict) else None)
        planned_start = getattr(node, "planned_start", None) or (node.get("planned_start") if isinstance(node, dict) else None)
        planned_end = getattr(node, "planned_end", None) or (node.get("planned_end") if isinstance(node, dict) else None)

        if man_days is None and man_hours is not None:
            try:
                man_days = round(float(man_hours) / 8.0, 2)
            except (TypeError, ValueError):
                man_days = None
        if actual_hours is not None and getattr(node, "actual_days", None) is None and (not isinstance(node, dict) or node.get("actual_days") is None):
            try:
                actual_days = round(float(actual_hours) / 8.0, 2)
            except (TypeError, ValueError):
                actual_days = None
        else:
            actual_days = getattr(node, "actual_days", None) or (node.get("actual_days") if isinstance(node, dict) else None)

        rows.append((
            _cell(key),
            _cell(kind),
            _cell(project_key),
            _cell(summary),
            _cell(parent_key),
            _cell(assignee),
            _cell(product_category),
            _cell(epic_key),
            _cell(story_key),
            float(man_hours) if man_hours is not None else None,
            float(man_days) if man_days is not None else None,
            float(actual_hours) if actual_hours is not None else None,
            float(actual_days) if actual_days is not None else None,
            _cell(planned_start),
            _cell(planned_end),
        ))
    if not rows:
        conn.commit()
        return
    placeholders = ",".join(["?"] * len(NESTED_VIEW_NODE_COLS))
    sql = f"INSERT INTO nested_view_nodes ({','.join(NESTED_VIEW_NODE_COLS)}) VALUES ({placeholders})"
    for row in rows:
        conn.execute(sql, row)
    conn.commit()


def record_export_run(
    conn: sqlite3.Connection,
    pipeline: str,
    row_count: int,
    status: str = "success",
    run_id: str | None = None,
    started_at_utc: str | None = None,
    finished_at_utc: str | None = None,
) -> None:
    run_id = run_id or uuid.uuid4().hex
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    started = started_at_utc or now
    finished = finished_at_utc or now
    conn.execute(
        """
        INSERT INTO export_runs (run_id, pipeline, started_at_utc, finished_at_utc, row_count, status)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (run_id, pipeline, started, finished, row_count, status),
    )
    conn.commit()


def has_subtask_worklogs(conn: sqlite3.Connection) -> bool:
    row = conn.execute("SELECT 1 FROM subtask_worklogs LIMIT 1").fetchone()
    return row is not None


def has_subtask_worklog_rollup(conn: sqlite3.Connection) -> bool:
    row = conn.execute("SELECT 1 FROM subtask_worklog_rollup LIMIT 1").fetchone()
    return row is not None


def _placeholders(n: int) -> str:
    return ",".join(["?"] * n)


def upsert_work_items_by_keys(conn: sqlite3.Connection, rows: list[list], keys: set[str]) -> None:
    """Delete work_items rows whose issue_key is in keys, then insert the given rows. keys normalized uppercase."""
    if not keys:
        if rows:
            placeholders = _placeholders(len(WORK_ITEMS_COLS))
            sql = f"INSERT INTO work_items ({','.join(WORK_ITEMS_COLS)}) VALUES ({placeholders})"
            for row in rows:
                conn.execute(sql, _work_items_row_from_list(row))
        conn.commit()
        return
    keys_list = list(keys)
    for i in range(0, len(keys_list), 500):
        chunk = keys_list[i : i + 500]
        conn.execute(
            "DELETE FROM work_items WHERE issue_key IN (" + _placeholders(len(chunk)) + ")",
            chunk,
        )
    if rows:
        placeholders = _placeholders(len(WORK_ITEMS_COLS))
        sql = f"INSERT INTO work_items ({','.join(WORK_ITEMS_COLS)}) VALUES ({placeholders})"
        for row in rows:
            conn.execute(sql, _work_items_row_from_list(row))
    conn.commit()


def upsert_subtask_worklogs_by_keys(conn: sqlite3.Connection, rows: list[list[object]], keys: set[str]) -> None:
    """Delete subtask_worklogs rows whose issue_id is in keys, then insert the given rows."""
    if not keys:
        if rows:
            placeholders = _placeholders(len(SUBTASK_WORKLOGS_COLS))
            sql = f"INSERT INTO subtask_worklogs ({','.join(SUBTASK_WORKLOGS_COLS)}) VALUES ({placeholders})"
            for row in rows:
                values = [_cell(row[j]) if j < len(row) else None for j in range(len(SUBTASK_WORKLOGS_COLS))]
                conn.execute(sql, values)
        conn.commit()
        return
    keys_list = list(keys)
    for i in range(0, len(keys_list), 500):
        chunk = keys_list[i : i + 500]
        conn.execute(
            "DELETE FROM subtask_worklogs WHERE issue_id IN (" + _placeholders(len(chunk)) + ")",
            chunk,
        )
    if rows:
        placeholders = _placeholders(len(SUBTASK_WORKLOGS_COLS))
        sql = f"INSERT INTO subtask_worklogs ({','.join(SUBTASK_WORKLOGS_COLS)}) VALUES ({placeholders})"
        for row in rows:
            values = [_cell(row[j]) if j < len(row) else None for j in range(len(SUBTASK_WORKLOGS_COLS))]
            conn.execute(sql, values)
    conn.commit()


def upsert_subtask_worklog_rollup_by_keys(conn: sqlite3.Connection, rows: list[list[object]], keys: set[str]) -> None:
    """Delete subtask_worklog_rollup rows whose issue_id is in keys, then insert the given rows."""
    if not keys:
        if rows:
            placeholders = _placeholders(len(SUBTASK_ROLLUP_COLS))
            sql = f"INSERT INTO subtask_worklog_rollup ({','.join(SUBTASK_ROLLUP_COLS)}) VALUES ({placeholders})"
            for row in rows:
                values = [_cell(row[j]) if j < len(row) else None for j in range(len(SUBTASK_ROLLUP_COLS))]
                conn.execute(sql, values)
        conn.commit()
        return
    keys_list = list(keys)
    for i in range(0, len(keys_list), 500):
        chunk = keys_list[i : i + 500]
        conn.execute(
            "DELETE FROM subtask_worklog_rollup WHERE issue_id IN (" + _placeholders(len(chunk)) + ")",
            chunk,
        )
    if rows:
        placeholders = _placeholders(len(SUBTASK_ROLLUP_COLS))
        sql = f"INSERT INTO subtask_worklog_rollup ({','.join(SUBTASK_ROLLUP_COLS)}) VALUES ({placeholders})"
        for row in rows:
            values = [_cell(row[j]) if j < len(row) else None for j in range(len(SUBTASK_ROLLUP_COLS))]
            conn.execute(sql, values)
    conn.commit()


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def dashboard_refresh_ensure_state_row(conn: sqlite3.Connection) -> None:
    row = conn.execute("SELECT 1 FROM dashboard_refresh_state WHERE id = 1").fetchone()
    if not row:
        now = _utc_now_iso()
        conn.execute(
            "INSERT INTO dashboard_refresh_state (id, active_run_id, last_success_run_id, updated_at_utc) VALUES (1, '', '', ?)",
            (now,),
        )
        conn.commit()


def dashboard_refresh_insert_run(
    conn: sqlite3.Connection,
    run_id: str,
    mode: str,
    resume_from_run_id: str | None = None,
    completed_steps_json: str | None = None,
) -> None:
    dashboard_refresh_ensure_state_row(conn)
    now = _utc_now_iso()
    conn.execute(
        """INSERT INTO dashboard_refresh_runs
        (run_id, mode, started_at_utc, status, progress_step, progress_pct, cancel_requested, updated_at_utc, resume_from_run_id, completed_steps_json)
        VALUES (?, ?, ?, 'running', '', 0, 0, ?, ?, ?)""",
        (run_id, mode, now, now, resume_from_run_id or None, completed_steps_json),
    )
    conn.execute(
        "UPDATE dashboard_refresh_state SET active_run_id = ?, updated_at_utc = ? WHERE id = 1",
        (run_id, now),
    )
    conn.commit()


def dashboard_refresh_update_progress(
    conn: sqlite3.Connection,
    run_id: str,
    progress_step: str,
    progress_pct: int,
    progress_total: int = 0,
    progress_done: int = 0,
    progress_current_label: str = "",
) -> None:
    now = _utc_now_iso()
    conn.execute(
        """UPDATE dashboard_refresh_runs
        SET progress_step = ?, progress_pct = ?,
            progress_total = ?, progress_done = ?, progress_current_label = ?,
            updated_at_utc = ?
        WHERE run_id = ?""",
        (progress_step, progress_pct, progress_total, progress_done, progress_current_label, now, run_id),
    )
    conn.commit()


def dashboard_refresh_set_cancel_requested(conn: sqlite3.Connection, run_id: str) -> None:
    now = _utc_now_iso()
    conn.execute(
        "UPDATE dashboard_refresh_runs SET cancel_requested = 1, updated_at_utc = ? WHERE run_id = ?",
        (now, run_id),
    )
    conn.commit()


def dashboard_refresh_finish_run(
    conn: sqlite3.Connection,
    run_id: str,
    status: str,
    error_message: str | None = None,
    stats_json: str | None = None,
    completed_steps_json: str | None = None,
) -> None:
    dashboard_refresh_ensure_state_row(conn)
    now = _utc_now_iso()
    conn.execute(
        """UPDATE dashboard_refresh_runs SET
        status = ?, ended_at_utc = ?, error_message = ?, stats_json = ?, updated_at_utc = ?,
        progress_pct = 100,
        completed_steps_json = COALESCE(?, completed_steps_json)
        WHERE run_id = ?""",
        (status, now, error_message, stats_json, now, completed_steps_json, run_id),
    )
    cur = conn.execute("SELECT active_run_id, last_success_run_id FROM dashboard_refresh_state WHERE id = 1").fetchone()
    if cur and cur[0] == run_id:
        new_active = ""
        new_last = cur[1]
        if status == "success":
            new_last = run_id
        conn.execute(
            "UPDATE dashboard_refresh_state SET active_run_id = ?, last_success_run_id = ?, updated_at_utc = ? WHERE id = 1",
            (new_active, new_last, now),
        )
    conn.commit()


def dashboard_refresh_get_run(conn: sqlite3.Connection, run_id: str) -> dict[str, Any] | None:
    row = conn.execute(
        """SELECT run_id, mode, started_at_utc, ended_at_utc, status,
                  progress_step, progress_pct, progress_total, progress_done, progress_current_label,
                  cancel_requested, error_message, stats_json, resume_from_run_id, completed_steps_json, updated_at_utc
           FROM dashboard_refresh_runs WHERE run_id = ?""",
        (run_id,),
    ).fetchone()
    if not row:
        return None
    return {
        "run_id": row[0],
        "mode": row[1],
        "started_at_utc": row[2],
        "ended_at_utc": row[3],
        "status": row[4],
        "progress_step": row[5],
        "progress_pct": row[6] or 0,
        "progress_total": row[7] or 0,
        "progress_done": row[8] or 0,
        "progress_current_label": row[9] or "",
        "cancel_requested": bool(row[10]),
        "error_message": row[11],
        "stats_json": row[12],
        "resume_from_run_id": row[13],
        "completed_steps_json": row[14],
        "updated_at_utc": row[15],
    }


def dashboard_refresh_get_active_run_id(conn: sqlite3.Connection) -> str:
    dashboard_refresh_ensure_state_row(conn)
    row = conn.execute("SELECT active_run_id FROM dashboard_refresh_state WHERE id = 1").fetchone()
    return (row[0] or "").strip() if row else ""


def dashboard_refresh_get_last_run(conn: sqlite3.Connection) -> dict[str, Any] | None:
    row = conn.execute(
        "SELECT run_id, mode, started_at_utc, ended_at_utc, status, progress_step, progress_pct, cancel_requested, error_message, completed_steps_json FROM dashboard_refresh_runs ORDER BY started_at_utc DESC LIMIT 1"
    ).fetchone()
    if not row:
        return None
    return {
        "run_id": row[0],
        "mode": row[1],
        "started_at_utc": row[2],
        "ended_at_utc": row[3],
        "status": row[4],
        "progress_step": row[5],
        "progress_pct": row[6] or 0,
        "cancel_requested": bool(row[7]),
        "error_message": row[8],
        "completed_steps_json": row[9],
    }


def dashboard_refresh_is_cancel_requested(conn: sqlite3.Connection, run_id: str) -> bool:
    row = conn.execute("SELECT cancel_requested FROM dashboard_refresh_runs WHERE run_id = ?", (run_id,)).fetchone()
    return bool(row and row[0])
