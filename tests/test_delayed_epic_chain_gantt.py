from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from delayed_epic_chain_gantt_service import (
    build_report_payload,
    load_ui_settings,
    save_ui_settings,
)
from report_server import create_report_server_app


def _write_support_files(root: Path) -> None:
    (root / "report_html").mkdir(parents=True, exist_ok=True)
    (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
    (root / "report_html" / "shared-nav.js").write_text("console.log('nav');", encoding="utf-8")
    (root / "report_html" / "shared-nav.css").write_text("body{}", encoding="utf-8")
    (root / "report_html" / "shared-date-filter.js").write_text("console.log('date-filter');", encoding="utf-8")
    html_path = Path(__file__).resolve().parents[1] / "delayed_epic_chain_gantt_report.html"
    (root / "delayed_epic_chain_gantt_report.html").write_text(html_path.read_text(encoding="utf-8"), encoding="utf-8")

    wb = Workbook()
    ws = wb.active
    ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
    ws.append(["O2", "2026-03-01", "2026-03-01", "2026-W09", "2026-03", "Alice", 1.0])
    wb.save(root / "assignee_hours_report.xlsx")


def _seed_db(db_path: Path) -> None:
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE canonical_refresh_state (
                id INTEGER PRIMARY KEY,
                active_run_id TEXT NOT NULL DEFAULT '',
                last_success_run_id TEXT NOT NULL DEFAULT '',
                updated_at_utc TEXT NOT NULL DEFAULT ''
            )
            """
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
            CREATE TABLE canonical_issue_links (
                run_id TEXT NOT NULL,
                issue_key TEXT NOT NULL,
                parent_issue_key TEXT NOT NULL DEFAULT '',
                story_key TEXT NOT NULL DEFAULT '',
                epic_key TEXT NOT NULL DEFAULT '',
                hierarchy_level TEXT NOT NULL DEFAULT ''
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
            CREATE TABLE epics_management (
                epic_key TEXT PRIMARY KEY,
                project_key TEXT NOT NULL,
                project_name TEXT NOT NULL,
                product_category TEXT NOT NULL DEFAULT '',
                component TEXT NOT NULL DEFAULT '',
                epic_name TEXT NOT NULL,
                description TEXT NOT NULL DEFAULT '',
                originator TEXT NOT NULL DEFAULT '',
                priority TEXT NOT NULL DEFAULT 'Low',
                plan_status TEXT NOT NULL DEFAULT 'Planned',
                ipp_meeting_planned TEXT NOT NULL DEFAULT 'No',
                actual_production_date TEXT NOT NULL DEFAULT '',
                remarks TEXT NOT NULL DEFAULT '',
                jira_url TEXT NOT NULL DEFAULT '',
                epic_plan_json TEXT NOT NULL DEFAULT '{}',
                research_urs_plan_json TEXT NOT NULL DEFAULT '{}',
                dds_plan_json TEXT NOT NULL DEFAULT '{}',
                development_plan_json TEXT NOT NULL DEFAULT '{}',
                sqa_plan_json TEXT NOT NULL DEFAULT '{}',
                user_manual_plan_json TEXT NOT NULL DEFAULT '{}',
                production_plan_json TEXT NOT NULL DEFAULT '{}',
                delivery_status TEXT NOT NULL DEFAULT 'Yet to start',
                is_sealed INTEGER NOT NULL DEFAULT 0
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE epf_refresh_state (
                id INTEGER PRIMARY KEY,
                active_run_id TEXT NOT NULL DEFAULT '',
                last_success_run_id TEXT NOT NULL DEFAULT '',
                updated_at_utc TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE epf_leave_rows (
                run_id TEXT NOT NULL,
                assignee TEXT NOT NULL DEFAULT '',
                period_day TEXT NOT NULL DEFAULT '',
                unplanned_taken_hours REAL NOT NULL DEFAULT 0,
                planned_taken_hours REAL NOT NULL DEFAULT 0
            )
            """
        )

        run_id = "run-1"
        conn.execute(
            "INSERT INTO canonical_refresh_state(id, active_run_id, last_success_run_id, updated_at_utc) VALUES (1, ?, ?, ?)",
            (run_id, run_id, "2026-03-31T00:00:00Z"),
        )
        conn.execute(
            "INSERT INTO epf_refresh_state(id, active_run_id, last_success_run_id, updated_at_utc) VALUES (1, ?, ?, ?)",
            ("epf-1", "epf-1", "2026-03-31T00:00:00Z"),
        )

        def add_epic(epic_key: str, epic_name: str, plan_start: str, plan_due: str, status: str, delivery_status: str = "In Progress"):
            conn.execute(
                """
                INSERT INTO epics_management(
                    epic_key, project_key, project_name, epic_name, jira_url, epic_plan_json, delivery_status
                ) VALUES (?, 'O2', 'OmniConnect', ?, ?, ?, ?)
                """,
                (
                    epic_key,
                    epic_name,
                    f"https://jira.example/browse/{epic_key}",
                    f'{{"man_days": 10, "start_date": "{plan_start}", "due_date": "{plan_due}", "jira_url": ""}}',
                    delivery_status,
                ),
            )
            conn.execute(
                """
                INSERT INTO canonical_issues(
                    run_id, issue_key, project_key, issue_type, summary, status, assignee, start_date, due_date,
                    resolved_stable_since_date, original_estimate_hours, total_hours_logged, parent_issue_key, story_key, epic_key
                ) VALUES (?, ?, 'O2', 'Epic', ?, ?, 'Evan', ?, ?, '', 80, 0, '', '', ?)
                """,
                (run_id, epic_key, epic_name, status, plan_start, plan_due, epic_key),
            )
            conn.execute(
                """
                INSERT INTO canonical_issue_links(run_id, issue_key, parent_issue_key, story_key, epic_key, hierarchy_level)
                VALUES (?, ?, '', '', ?, 'epic')
                """,
                (run_id, epic_key, epic_key),
            )

        add_epic("O2-A", "Epic A", "2026-02-15", "2026-02-28", "Resolved!", "Late")
        add_epic("O2-B", "Epic B", "2026-03-01", "2026-03-20", "In-Progress", "Late")
        add_epic("O2-C", "Epic C", "2026-04-02", "2026-04-15", "In-Progress", "Yet to start")
        add_epic("O2-D", "Epic D", "2026-04-15", "2026-04-22", "On Hold", "Late")
        add_epic("O2-E", "Epic E", "2026-02-10", "2026-02-20", "To Do", "Yet to start")
        add_epic("O2-F", "Epic F", "2026-03-10", "2026-03-18", "To Do", "Yet to start")

        def add_story_and_subtask(epic_key: str, first_date: str, end_date: str, story_assignee: str = "Sally", subtask_assignee: str = "Alice"):
            story_key = f"{epic_key}-S1"
            subtask_key = f"{epic_key}-T1"
            conn.execute(
                """
                INSERT INTO canonical_issues(
                    run_id, issue_key, project_key, issue_type, summary, status, assignee, start_date, due_date,
                    resolved_stable_since_date, original_estimate_hours, total_hours_logged, parent_issue_key, story_key, epic_key
                ) VALUES (?, ?, 'O2', 'Story', ?, 'In Progress', ?, '', '', '', 16, 0, ?, ?, ?)
                """,
                (run_id, story_key, f"{epic_key} Story", story_assignee, epic_key, story_key, epic_key),
            )
            conn.execute(
                """
                INSERT INTO canonical_issue_links(run_id, issue_key, parent_issue_key, story_key, epic_key, hierarchy_level)
                VALUES (?, ?, ?, ?, ?, 'story')
                """,
                (run_id, story_key, epic_key, story_key, epic_key),
            )
            conn.execute(
                """
                INSERT INTO canonical_issues(
                    run_id, issue_key, project_key, issue_type, summary, status, assignee, start_date, due_date,
                    resolved_stable_since_date, original_estimate_hours, total_hours_logged, parent_issue_key, story_key, epic_key
                ) VALUES (?, ?, 'O2', 'Sub-task', ?, 'In Progress', ?, '', '', '', 16, 0, ?, ?, ?)
                """,
                (run_id, subtask_key, f"{epic_key} Task", subtask_assignee, story_key, story_key, epic_key),
            )
            conn.execute(
                """
                INSERT INTO canonical_issue_links(run_id, issue_key, parent_issue_key, story_key, epic_key, hierarchy_level)
                VALUES (?, ?, ?, ?, ?, 'subtask')
                """,
                (run_id, subtask_key, story_key, story_key, epic_key),
            )
            conn.execute(
                """
                INSERT INTO canonical_issue_actuals(
                    run_id, issue_key, project_key, assignee, first_worklog_date, last_worklog_date,
                    actual_complete_date, actual_complete_source, due_completion_bucket, total_worklog_hours, worklog_count
                ) VALUES (?, ?, 'O2', ?, ?, ?, ?, 'last_worklog_date', 'after_due', 24, 3)
                """,
                (run_id, subtask_key, subtask_assignee, first_date, end_date, end_date),
            )

        add_story_and_subtask("O2-A", "2026-02-15", "2026-03-05")
        add_story_and_subtask("O2-B", "2026-03-06", "2026-03-30")
        add_story_and_subtask("O2-C", "2026-04-06", "2026-04-16")
        add_story_and_subtask("O2-D", "2026-04-18", "2026-04-24")
        add_story_and_subtask("O2-E", "", "")
        add_story_and_subtask("O2-F", "", "")

        conn.execute(
            """
            INSERT INTO canonical_issue_actuals(
                run_id, issue_key, project_key, assignee, first_worklog_date, last_worklog_date,
                actual_complete_date, actual_complete_source, due_completion_bucket, total_worklog_hours, worklog_count
            ) VALUES (?, 'O2-E', 'O2', 'Evan', '', '', '', 'none', 'not_completed', 0, 0)
            """,
            (run_id,),
        )
        conn.execute(
            """
            UPDATE canonical_issue_actuals
            SET first_worklog_date = '', last_worklog_date = '', actual_complete_date = '',
                actual_complete_source = 'none', due_completion_bucket = 'not_completed',
                total_worklog_hours = 0, worklog_count = 0
            WHERE run_id = ? AND issue_key = 'O2-E-T1'
            """,
            (run_id,),
        )

        conn.executemany(
            """
            INSERT INTO epf_leave_rows(run_id, assignee, period_day, unplanned_taken_hours, planned_taken_hours)
            VALUES ('epf-1', 'Alice', ?, 0, 8)
            """,
            [("2026-04-02",), ("2026-04-15",)],
        )
        conn.commit()
    finally:
        conn.close()


class DelayedEpicChainGanttTests(unittest.TestCase):
    def test_service_build_payload_derives_delay_chain_and_actual_dates(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            _write_support_files(root)
            db_path = root / "assignee_hours_capacity.db"
            _seed_db(db_path)

            payload = build_report_payload(
                db_path,
                "2026-03-01",
                "2026-03-31",
                assignee="",
                assignee_mode="subtask_assignee",
            )

            rows = {row["epic_key"]: row for row in payload["rows"]}
            self.assertEqual(rows["O2-A"]["actual_start"], "2026-02-15")
            self.assertEqual(rows["O2-A"]["actual_complete_date"], "2026-03-05")
            self.assertEqual(rows["O2-B"]["delay_cause"], "Previous Epic Delayed")
            self.assertEqual(rows["O2-B"]["blocking_epic_key"], "O2-A")
            self.assertEqual(rows["O2-C"]["delay_cause"], "Leave Overlap")
            self.assertEqual(rows["O2-D"]["delay_cause"], "Previous Epic Delayed")
            self.assertEqual(rows["O2-D"]["blocking_epic_key"], "O2-C")
            self.assertIn("O2-E", rows)
            self.assertIn("O2-F", rows)
            self.assertEqual(rows["O2-F"]["planned_start"], "2026-03-10")

    def test_service_assignee_modes_match_subtask_story_and_epic_assignee(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            _write_support_files(root)
            db_path = root / "assignee_hours_capacity.db"
            _seed_db(db_path)

            subtask_rows = build_report_payload(
                db_path, "2026-03-01", "2026-03-31", assignee="Alice", assignee_mode="subtask_assignee"
            )["rows"]
            story_rows = build_report_payload(
                db_path, "2026-03-01", "2026-03-31", assignee="Sally", assignee_mode="story_assignee"
            )["rows"]
            epic_rows = build_report_payload(
                db_path, "2026-03-01", "2026-03-31", assignee="Evan", assignee_mode="epic_assignee"
            )["rows"]

            self.assertTrue(subtask_rows)
            self.assertTrue(story_rows)
            self.assertTrue(epic_rows)
            self.assertEqual(story_rows[0]["assignee_mode"], "story_assignee")
            self.assertEqual(epic_rows[0]["assignee_mode"], "epic_assignee")

    def test_ui_settings_round_trip(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "assignee_hours_capacity.db"
            self.assertEqual(load_ui_settings(db_path)["week_bucket_width_px"], 100)
            saved = save_ui_settings(
                db_path,
                {"show_full_year": False, "week_bucket_width_px": 120},
            )
            self.assertFalse(saved["show_full_year"])
            self.assertEqual(saved["week_bucket_width_px"], 120)
            latest = load_ui_settings(db_path)
            self.assertFalse(latest["show_full_year"])
            self.assertEqual(latest["week_bucket_width_px"], 120)

    def test_api_endpoints_return_filter_options_data_and_ui_settings(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            _write_support_files(root)
            _seed_db(root / "assignee_hours_capacity.db")
            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            save_filter = client.post(
                "/api/report-date-filter",
                json={"from_date": "2026-03-01", "to_date": "2026-03-31", "source_page": "delayed_epic_chain_gantt"},
            )
            self.assertEqual(save_filter.status_code, 200)

            filter_resp = client.get("/api/delayed-epic-chain-gantt/filter-options")
            self.assertEqual(filter_resp.status_code, 200)
            filter_json = filter_resp.get_json()
            self.assertTrue(filter_json["ok"])
            self.assertIn("Alice", filter_json["filter_options"]["assignees"])
            self.assertIn("Sally", filter_json["filter_options"]["assignees"])
            self.assertEqual(filter_json["global_date_filter"]["from_date"], "2026-03-01")

            ui_resp = client.post(
                "/api/delayed-epic-chain-gantt/ui-settings",
                json={"show_full_year": False, "week_bucket_width_px": 110},
            )
            self.assertEqual(ui_resp.status_code, 200)
            self.assertEqual(ui_resp.get_json()["settings"]["week_bucket_width_px"], 110)

            data_resp = client.get(
                "/api/delayed-epic-chain-gantt/data?assignee_mode=subtask_assignee"
            )
            self.assertEqual(data_resp.status_code, 200)
            data_json = data_resp.get_json()
            self.assertTrue(data_json["ok"])
            rows = {row["epic_key"]: row for row in data_json["rows"]}
            self.assertEqual(rows["O2-B"]["delay_cause"], "Previous Epic Delayed")
            self.assertEqual(rows["O2-C"]["delay_cause"], "Leave Overlap")
            self.assertIn("O2-F", rows)
            self.assertEqual(rows["O2-F"]["planned_start"], "2026-03-10")

    def test_report_page_route_serves_delayed_epic_chain_gantt_html(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            _write_support_files(root)
            _seed_db(root / "assignee_hours_capacity.db")
            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            resp = client.get("/delayed_epic_chain_gantt_report.html")
            self.assertEqual(resp.status_code, 200)
            html = resp.get_data(as_text=True)
            self.assertIn("Delayed Epic Chain Gantt", html)
            self.assertIn("/api/delayed-epic-chain-gantt/data", html)
            self.assertIn("shared-date-filter.js", html)


if __name__ == "__main__":
    unittest.main()
