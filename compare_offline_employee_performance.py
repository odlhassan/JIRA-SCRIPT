"""
Compare offline Employee Performance export with canonical DB.

Reads employee_performance_report.json from an offline bundle folder (e.g. "12 Mar 2026 06-01-29"),
prints the embedded scoped-subtasks summary, then runs the same logic as settings/sql-console
sample queries against the canonical SQLite DB to find discrepancies.

Usage:
  python compare_offline_employee_performance.py "12 Mar 2026 06-01-29"
  python compare_offline_employee_performance.py "12 Mar 2026 06-01-29" --db path/to/assignee_hours_capacity.db

If --db is omitted, JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH or default assignee_hours_capacity.db
in the script directory is used. If the DB does not exist, only the offline JSON summary is printed.
"""

from __future__ import annotations

import argparse
import json
import os
import sqlite3
import sys
from pathlib import Path


def _find_offline_json(offline_folder: Path) -> Path | None:
    for name in ("employee_performance_report.json",):
        p = offline_folder / name
        if p.is_file():
            return p
    return None


def _extract_scoped_summary(bundle: dict) -> dict | None:
    for key, data in bundle.items():
        if "scoped-subtasks" not in key or not isinstance(data, dict):
            continue
        if data.get("ok") is True and "from_date" in data and "total_planned_hours" in data:
            return data
    return None


def _run_canonical_queries(
    db_path: Path,
    from_date: str,
    to_date: str,
) -> dict:
    """Run SQL console-style queries (subtasks in date range; extended vs log_date actuals)."""
    if not db_path.is_file():
        return {"error": f"DB not found: {db_path}"}
    from_d = from_date[:10] if from_date else ""
    to_d = to_date[:10] if to_date else ""
    if not from_d or not to_d:
        return {"error": "from_date and to_date required (YYYY-MM-DD)."}

    # Match API: exclude RLT; include only subtask issue types (sub-task/subtask) for planned/actual hours
    projects_sql = "AND ci.project_key <> 'RLT'"

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        # 1) EXTENDED: worklogs NOT filtered by started_date (all worklogs per issue)
        sql_extended = f"""
WITH latest_run AS (
  SELECT run_id FROM canonical_refresh_runs
  WHERE status = 'success' ORDER BY updated_at_utc DESC LIMIT 1
), scoped_subtasks AS (
  SELECT ci.issue_key, ci.assignee, COALESCE(ci.original_estimate_hours, 0) AS original_estimate_hours
  FROM canonical_issues ci
  WHERE ci.run_id = (SELECT run_id FROM latest_run)
    {projects_sql}
    AND ( LOWER(ci.issue_type) LIKE '%sub-task%' OR LOWER(ci.issue_type) LIKE '%subtask%' )
    AND ( (ci.start_date >= ? AND ci.start_date <= ?) OR (ci.due_date >= ? AND ci.due_date <= ?) )
), worklog_totals AS (
  SELECT cw.issue_key, SUM(COALESCE(cw.hours_logged, 0)) AS logged_hours
  FROM canonical_worklogs cw
  WHERE cw.run_id = (SELECT run_id FROM latest_run)
  GROUP BY cw.issue_key
)
SELECT
  ROUND(SUM(ss.original_estimate_hours), 2) AS total_planned_hours,
  ROUND(SUM(COALESCE(wt.logged_hours, 0)), 2) AS total_actual_hours,
  COUNT(*) AS total_subtasks,
  COUNT(DISTINCT CASE WHEN TRIM(COALESCE(ss.assignee, '')) <> '' THEN ss.assignee END) AS total_assignees
FROM scoped_subtasks ss
LEFT JOIN worklog_totals wt ON wt.issue_key = ss.issue_key
"""
        row_ext = conn.execute(sql_extended, (from_d, to_d, from_d, to_d)).fetchone()

        # 2) LOG_DATE: worklogs filtered by started_date in range
        sql_log_date = f"""
WITH latest_run AS (
  SELECT run_id FROM canonical_refresh_runs
  WHERE status = 'success' ORDER BY updated_at_utc DESC LIMIT 1
), scoped_subtasks AS (
  SELECT ci.issue_key, ci.assignee, COALESCE(ci.original_estimate_hours, 0) AS original_estimate_hours
  FROM canonical_issues ci
  WHERE ci.run_id = (SELECT run_id FROM latest_run)
    {projects_sql}
    AND ( LOWER(ci.issue_type) LIKE '%sub-task%' OR LOWER(ci.issue_type) LIKE '%subtask%' )
    AND ( (ci.start_date >= ? AND ci.start_date <= ?) OR (ci.due_date >= ? AND ci.due_date <= ?) )
), worklog_totals AS (
  SELECT cw.issue_key, ROUND(SUM(COALESCE(cw.hours_logged, 0)), 2) AS logged_hours
  FROM canonical_worklogs cw
  WHERE cw.run_id = (SELECT run_id FROM latest_run)
    AND cw.started_date >= ? AND cw.started_date <= ?
  GROUP BY cw.issue_key
)
SELECT
  ROUND(SUM(ss.original_estimate_hours), 2) AS total_planned_hours,
  ROUND(SUM(COALESCE(wt.logged_hours, 0)), 2) AS total_actual_hours,
  COUNT(*) AS total_subtasks,
  COUNT(DISTINCT CASE WHEN TRIM(COALESCE(ss.assignee, '')) <> '' THEN ss.assignee END) AS total_assignees
FROM scoped_subtasks ss
LEFT JOIN worklog_totals wt ON wt.issue_key = ss.issue_key
"""
        row_log = conn.execute(
            sql_log_date, (from_d, to_d, from_d, to_d, from_d, to_d)
        ).fetchone()

        return {
            "extended": dict(row_ext) if row_ext else {},
            "log_date": dict(row_log) if row_log else {},
        }
    except Exception as e:
        return {"error": str(e)}
    finally:
        conn.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="Compare offline Employee Performance JSON with canonical DB.")
    parser.add_argument("offline_folder", help="Offline bundle folder name (e.g. '12 Mar 2026 06-01-29')")
    parser.add_argument("--db", default=None, help="Path to canonical DB (assignee_hours_capacity.db)")
    parser.add_argument("--base-dir", default=None, help="Base dir; default is script directory")
    args = parser.parse_args()

    base = Path(args.base_dir or __file__).resolve().parent
    folder = base / args.offline_folder
    if not folder.is_dir():
        print(f"Offline folder not found: {folder}", file=sys.stderr)
        return 1

    json_path = _find_offline_json(folder)
    if not json_path:
        print(f"No employee_performance_report.json in {folder}", file=sys.stderr)
        return 1

    with open(json_path, encoding="utf-8") as f:
        bundle = json.load(f)

    summary = _extract_scoped_summary(bundle)
    if not summary:
        print("No scoped-subtasks summary found in JSON (no key contained 'scoped-subtasks' with ok/from_date/total_planned_hours).")
        return 0

    from_date = summary.get("from_date", "")
    to_date = summary.get("to_date", "")
    print("=== Offline bundle (employee_performance_report.json) ===")
    print(f"  from_date: {from_date}")
    print(f"  to_date:   {to_date}")
    print(f"  total_planned_hours: {summary.get('total_planned_hours')}")
    print(f"  total_actual_hours:  {summary.get('total_actual_hours')}")
    print(f"  total_subtasks:      {summary.get('total_subtasks')}")
    print(f"  total_assignees:     {summary.get('total_assignees')}")
    print()

    db_path = Path(args.db) if args.db else base / "assignee_hours_capacity.db"
    if not db_path.is_absolute():
        db_path = base / db_path
    env_db = os.getenv("JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH", "").strip()
    if not args.db and env_db:
        db_path = Path(env_db)
        if not db_path.is_absolute():
            db_path = base / db_path

    result = _run_canonical_queries(db_path, from_date, to_date)
    if "error" in result:
        print("=== Canonical DB ===")
        print(f"  {result['error']}")
        print()
        print("To get canonical stats, run this script with --db path/to/assignee_hours_capacity.db")
        return 0

    print("=== Canonical DB (settings/sql-console style) ===")
    print("  Extended actuals (all worklogs per issue, no date filter on worklogs):")
    for k, v in (result.get("extended") or {}).items():
        print(f"    {k}: {v}")
    print("  Log-date actuals (only worklogs with started_date in range):")
    for k, v in (result.get("log_date") or {}).items():
        print(f"    {k}: {v}")
    print()
    print("Note: Prepare Offline HTML now uses mode=log_date for Employee Performance, so")
    print("      total_actual_hours in new exports matches 'log_date' (only worklogs in range).")
    print("      If your JSON shows extended-style actuals, re-export with the current build.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
