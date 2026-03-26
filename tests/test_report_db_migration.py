from __future__ import annotations

import json
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fetch_jira_dashboard
from generate_assignee_hours_report import _generate_outputs
from generate_leaves_planned_calendar_html import _load_calendar_data_from_canonical
from generate_missed_entries_html import _load_rows_from_canonical_db
from generate_planned_rmis_html import _load_payload_from_canonical
from sync_team_rmi_gantt_sqlite import build_team_rmi_gantt_snapshot


def _seed_canonical_db(db_path: Path, run_id: str = "run-db-migration") -> str:
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE canonical_refresh_state (
                id INTEGER PRIMARY KEY CHECK(id = 1),
                active_run_id TEXT NOT NULL DEFAULT '',
                last_success_run_id TEXT NOT NULL DEFAULT '',
                updated_at_utc TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.execute(
            """
            INSERT INTO canonical_refresh_state (id, active_run_id, last_success_run_id, updated_at_utc)
            VALUES (1, '', ?, '2026-03-26T00:00:00+00:00')
            """,
            (run_id,),
        )
        conn.execute(
            """
            CREATE TABLE canonical_issues (
                run_id TEXT NOT NULL,
                issue_id TEXT NOT NULL DEFAULT '',
                issue_key TEXT NOT NULL,
                project_key TEXT NOT NULL DEFAULT '',
                issue_type TEXT NOT NULL DEFAULT '',
                summary TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT '',
                assignee TEXT NOT NULL DEFAULT '',
                start_date TEXT NOT NULL DEFAULT '',
                due_date TEXT NOT NULL DEFAULT '',
                created_utc TEXT NOT NULL DEFAULT '',
                updated_utc TEXT NOT NULL DEFAULT '',
                resolved_stable_since_date TEXT NOT NULL DEFAULT '',
                original_estimate_hours REAL NOT NULL DEFAULT 0,
                total_hours_logged REAL NOT NULL DEFAULT 0,
                fix_type TEXT NOT NULL DEFAULT '',
                parent_issue_key TEXT NOT NULL DEFAULT '',
                story_key TEXT NOT NULL DEFAULT '',
                epic_key TEXT NOT NULL DEFAULT '',
                raw_payload_json TEXT NOT NULL DEFAULT '{}'
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE canonical_worklogs (
                run_id TEXT NOT NULL,
                worklog_id TEXT NOT NULL,
                issue_key TEXT NOT NULL,
                project_key TEXT NOT NULL DEFAULT '',
                worklog_author TEXT NOT NULL DEFAULT '',
                issue_assignee TEXT NOT NULL DEFAULT '',
                started_utc TEXT NOT NULL DEFAULT '',
                started_date TEXT NOT NULL DEFAULT '',
                updated_utc TEXT NOT NULL DEFAULT '',
                hours_logged REAL NOT NULL DEFAULT 0
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE canonical_issue_actuals (
                run_id TEXT NOT NULL,
                issue_key TEXT NOT NULL,
                project_key TEXT NOT NULL DEFAULT '',
                assignee TEXT NOT NULL DEFAULT '',
                first_worklog_date TEXT NOT NULL DEFAULT '',
                last_worklog_date TEXT NOT NULL DEFAULT '',
                actual_complete_date TEXT NOT NULL DEFAULT '',
                actual_complete_source TEXT NOT NULL DEFAULT '',
                due_completion_bucket TEXT NOT NULL DEFAULT '',
                total_worklog_hours REAL NOT NULL DEFAULT 0,
                worklog_count INTEGER NOT NULL DEFAULT 0
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE managed_projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_key TEXT NOT NULL UNIQUE,
                project_name TEXT NOT NULL,
                display_name TEXT NOT NULL,
                color_hex TEXT NOT NULL,
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at_utc TEXT NOT NULL,
                updated_at_utc TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE epics_management (
                epic_key TEXT PRIMARY KEY,
                project_key TEXT NOT NULL,
                project_name TEXT NOT NULL,
                product_category TEXT NOT NULL DEFAULT '',
                component TEXT NOT NULL DEFAULT '',
                epic_name TEXT NOT NULL,
                epic_plan_json TEXT NOT NULL DEFAULT '{}'
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE performance_teams (
                team_name TEXT PRIMARY KEY,
                team_leader TEXT NOT NULL DEFAULT '',
                assignees_json TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE assignee_capacity_settings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                from_date TEXT NOT NULL,
                to_date TEXT NOT NULL,
                employee_count INTEGER NOT NULL,
                standard_hours_per_day REAL NOT NULL,
                ramadan_start_date TEXT,
                ramadan_end_date TEXT,
                ramadan_hours_per_day REAL NOT NULL,
                holiday_dates_json TEXT NOT NULL,
                created_at_utc TEXT NOT NULL,
                updated_at_utc TEXT NOT NULL,
                UNIQUE(from_date, to_date)
            )
            """
        )
        conn.execute(
            """
            INSERT INTO managed_projects(project_key, project_name, display_name, color_hex, is_active, created_at_utc, updated_at_utc)
            VALUES ('O2', 'OmniConnect', 'OmniConnect', '#336699', 1, '2026-03-26', '2026-03-26')
            """
        )
        conn.execute(
            """
            INSERT INTO epics_management(epic_key, project_key, project_name, product_category, component, epic_name, epic_plan_json)
            VALUES ('O2-EP1', 'O2', 'OmniConnect', 'Platform', '', 'Epic Alpha', '{"start_date":"2026-02-01","due_date":"2026-02-20","man_days":5}')
            """
        )
        conn.execute(
            """
            INSERT INTO performance_teams(team_name, team_leader, assignees_json, updated_at)
            VALUES ('Technical Writing', 'Alice', '["Alice"]', '2026-03-26')
            """
        )
        issue_rows = [
            (run_id, "", "O2-EP1", "O2", "Epic", "Epic Alpha", "In Progress", "Alice", "2026-02-01", "2026-02-20", "", "", "", 40.0, 8.0, "", "", "", "O2-EP1", "{}"),
            (run_id, "", "O2-ST1", "O2", "Story", "Story Alpha", "In Progress", "Alice", "2026-02-03", "2026-02-10", "", "", "", 16.0, 8.0, "", "O2-EP1", "O2-ST1", "O2-EP1", "{}"),
            (run_id, "", "O2-SUB1", "O2", "Sub-task", "Subtask Alpha", "In Progress", "Alice", "2026-02-03", "2026-02-05", "", "", "", 8.0, 8.0, "", "O2-ST1", "O2-ST1", "O2-EP1", "{}"),
            (run_id, "", "RLT-1", "RLT", "Task", "Leave Parent", "In Progress", "Alice", "2026-02-01", "2026-02-28", "", "", "", 0.0, 0.0, "", "", "", "", "{}"),
            (
                run_id,
                "",
                "RLT-2",
                "RLT",
                "Sub-task",
                "Planned Leave",
                "In Progress",
                "Alice",
                "2026-02-11",
                "2026-02-11",
                "",
                "",
                "",
                8.0,
                8.0,
                "",
                "RLT-1",
                "",
                "",
                json.dumps({"fields": {"customfield_10584": "Planned"}}),
            ),
        ]
        conn.executemany(
            """
            INSERT INTO canonical_issues(
                run_id, issue_id, issue_key, project_key, issue_type, summary, status, assignee,
                start_date, due_date, created_utc, updated_utc, resolved_stable_since_date,
                original_estimate_hours, total_hours_logged, fix_type, parent_issue_key,
                story_key, epic_key, raw_payload_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            issue_rows,
        )
        conn.executemany(
            """
            INSERT INTO canonical_worklogs(
                run_id, worklog_id, issue_key, project_key, worklog_author, issue_assignee,
                started_utc, started_date, updated_utc, hours_logged
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (run_id, "wl-1", "O2-SUB1", "O2", "Alice", "Alice", "2026-02-04T09:00:00+00:00", "2026-02-04", "", 8.0),
                (run_id, "wl-2", "RLT-2", "RLT", "Alice", "Alice", "2026-02-11T09:00:00+00:00", "2026-02-11", "", 8.0),
            ],
        )
        conn.executemany(
            """
            INSERT INTO canonical_issue_actuals(
                run_id, issue_key, project_key, assignee, first_worklog_date, last_worklog_date,
                actual_complete_date, actual_complete_source, due_completion_bucket, total_worklog_hours, worklog_count
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (run_id, "O2-SUB1", "O2", "Alice", "2026-02-04", "2026-02-04", "2026-02-04", "test", "", 8.0, 1),
                (run_id, "O2-EP1", "O2", "Alice", "2026-02-04", "2026-02-04", "", "test", "", 8.0, 1),
            ],
        )
        conn.commit()
    return run_id


class ReportDbMigrationTests(unittest.TestCase):
    def test_dashboard_and_missed_entries_load_from_canonical_db(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            run_id = _seed_canonical_db(db_path)
            with patch.dict(
                "os.environ",
                {
                    "JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH": str(db_path),
                    "JIRA_CANONICAL_RUN_ID": run_id,
                },
                clear=False,
            ):
                payload = fetch_jira_dashboard.fetch_dashboard_data()
            self.assertEqual(payload.get("source_file"), "canonical_db")
            self.assertEqual(payload["epics"][0]["issue_key"], "O2-EP1")

            rows, default_from, default_to = _load_rows_from_canonical_db(db_path, run_id)
            self.assertTrue(rows)
            self.assertEqual(rows[0]["issue_key"], "O2-EP1")
            self.assertTrue(default_from)
            self.assertTrue(default_to)

    def test_planned_rmis_and_team_gantt_build_from_canonical_db(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            run_id = _seed_canonical_db(db_path)

            payload = _load_payload_from_canonical(db_path, run_id)
            self.assertEqual(payload.get("source_file"), "canonical_db")
            epic_rows = [row for row in payload.get("rows", []) if row.get("row_kind") == "epic"]
            self.assertTrue(epic_rows)
            self.assertEqual(epic_rows[0]["jira_key"], "O2-EP1")

            snapshot = build_team_rmi_gantt_snapshot(root / "missing.xlsx", db_path)
            self.assertEqual(snapshot["included_story_rows"], 1)
            self.assertEqual(snapshot["items"][0]["team_name"], "Technical Writing")

    def test_leave_calendar_and_assignee_hours_build_from_canonical_db(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            run_id = _seed_canonical_db(db_path)

            counts, planned_hours, skipped, details_by_date, unmatched_by_date, warnings = _load_calendar_data_from_canonical(db_path, run_id)
            self.assertEqual(skipped, 0)
            self.assertIn("2026-02-11", counts)
            self.assertIn("2026-02-11", details_by_date)
            self.assertEqual(unmatched_by_date, {})
            self.assertEqual(warnings, [])

            paths = {
                "input_path": root / "missing_worklogs.xlsx",
                "work_items_path": root / "missing_items.xlsx",
                "summary_path": root / "assignee_hours_report.xlsx",
                "html_path": root / "assignee_hours_report.html",
                "db_path": db_path,
                "leave_report_path": root / "missing_leave.xlsx",
            }
            with patch.dict("os.environ", {"JIRA_CANONICAL_RUN_ID": run_id}, clear=False):
                outputs = _generate_outputs(paths)
            self.assertTrue(outputs["html_path"].exists())
            self.assertEqual(outputs["payload"]["leave_daily_rows"][0]["period_day"], "2026-02-11")
            self.assertEqual(outputs["payload"]["planned_work_items"][0]["issue_key"], "O2-EP1")


if __name__ == "__main__":
    unittest.main()
