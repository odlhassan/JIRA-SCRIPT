from __future__ import annotations

import os
import sqlite3
import tempfile
import unittest
from pathlib import Path

from generate_rmi_jira_gantt_html import generate_html_report, load_report_data, render_html
from report_server import (
    REPORT_FILENAME_TO_ID,
    REPORT_REFRESH_CHAINS,
    STATIC_REPORT_NAV_ITEMS,
    create_report_server_app,
    _init_epics_management_db,
    _resolve_report_html_sources,
    _save_epics_management_row,
)
from run_html_only import build_html_only_steps


def _seed_canonical_tables(db_path: Path) -> str:
    run_id = "canonical-test-run"
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS canonical_refresh_state (
                id INTEGER PRIMARY KEY CHECK(id = 1),
                last_success_run_id TEXT NOT NULL
            )
            """
        )
        conn.execute("INSERT OR REPLACE INTO canonical_refresh_state (id, last_success_run_id) VALUES (1, ?)", (run_id,))
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS canonical_issues (
                run_id TEXT NOT NULL,
                issue_id TEXT NOT NULL,
                issue_key TEXT NOT NULL,
                project_key TEXT NOT NULL,
                issue_type TEXT NOT NULL,
                summary TEXT NOT NULL,
                status TEXT NOT NULL,
                assignee TEXT NOT NULL,
                start_date TEXT NOT NULL,
                due_date TEXT NOT NULL,
                created_utc TEXT NOT NULL,
                updated_utc TEXT NOT NULL,
                resolved_stable_since_date TEXT NOT NULL,
                original_estimate_hours REAL NOT NULL,
                total_hours_logged REAL NOT NULL,
                fix_type TEXT NOT NULL,
                parent_issue_key TEXT NOT NULL,
                story_key TEXT NOT NULL,
                epic_key TEXT NOT NULL,
                raw_payload_json TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS canonical_worklogs (
                run_id TEXT NOT NULL,
                worklog_id TEXT NOT NULL,
                issue_key TEXT NOT NULL,
                project_key TEXT NOT NULL,
                worklog_author TEXT NOT NULL,
                issue_assignee TEXT NOT NULL,
                started_utc TEXT NOT NULL,
                started_date TEXT NOT NULL,
                updated_utc TEXT NOT NULL,
                hours_logged REAL NOT NULL
            )
            """
        )
        rows = [
            (run_id, "1", "O2-321", "O2", "Epic", "Canonical Epic Summary", "In Progress", "", "2026-02-01", "2026-02-28", "", "", "", 120.0, 0.0, "", "", "", "", "{}"),
            (run_id, "2", "O2-401", "O2", "Story", "Story A", "To Do", "Alice", "2026-02-03", "2026-02-12", "", "", "", 16.0, 3.0, "", "O2-321", "O2-401", "O2-321", "{}"),
            (run_id, "3", "O2-402", "O2", "Sub-task", "Subtask A", "Done", "Alice", "2026-02-04", "2026-02-05", "", "", "", 8.0, 2.0, "", "O2-401", "O2-401", "O2-321", "{}"),
            (run_id, "4", "O2-999", "O2", "Epic", "Canonical Only Epic", "To Do", "", "2026-03-01", "2026-03-10", "", "", "", 8.0, 0.0, "", "", "", "", "{}"),
        ]
        conn.executemany(
            """
            INSERT INTO canonical_issues (
                run_id, issue_id, issue_key, project_key, issue_type, summary, status, assignee,
                start_date, due_date, created_utc, updated_utc, resolved_stable_since_date,
                original_estimate_hours, total_hours_logged, fix_type, parent_issue_key,
                story_key, epic_key, raw_payload_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            rows,
        )
        conn.execute(
            """
            INSERT INTO canonical_worklogs (
                run_id, worklog_id, issue_key, project_key, worklog_author, issue_assignee,
                started_utc, started_date, updated_utc, hours_logged
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (run_id, "wl-1", "O2-402", "O2", "Alice", "Alice", "2026-02-05T09:00:00Z", "2026-02-05", "", 2.0),
        )
        conn.commit()
    finally:
        conn.close()
    return run_id


def _seed_planner_row(db_path: Path) -> None:
    _init_epics_management_db(db_path)
    _save_epics_management_row(
        db_path,
        {
            "id": "O2-321",
            "epic_key": "O2-321",
            "project_key": "O2",
            "project_name": "OmniConnect",
            "product_category": "Input",
            "component": "Streaming",
            "epic_name": "Planner Epic Summary",
            "priority": "High",
            "plan_status": "Planned",
            "ipp_meeting_planned": "Yes",
            "jira_url": "https://example.atlassian.net/browse/O2-321",
            "remarks": "Held for next IPP — scope TBD.",
            "plans": {
                "research_urs_plan": {"most_likely_man_days": 2, "start_date": "2026-02-01", "due_date": "2026-02-03", "jira_url": "https://example.atlassian.net/browse/O2-401"},
                "dds_plan": {"most_likely_man_days": 4, "start_date": "2026-02-04", "due_date": "2026-02-06"},
                "development_plan": {"most_likely_man_days": 10, "start_date": "2026-02-07", "due_date": "2026-02-20"},
                "sqa_plan": {"most_likely_man_days": 5, "start_date": "2026-02-21", "due_date": "2026-02-25"},
            },
        },
    )


class RmiJiraGanttReportTests(unittest.TestCase):
    def test_load_report_data_falls_back_to_latest_canonical_run_id_without_refresh_state(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "assignee_hours_capacity.db"
            _seed_planner_row(db_path)
            run_id = "fallback-run-id"
            conn = sqlite3.connect(db_path)
            try:
                conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS canonical_issues (
                        run_id TEXT NOT NULL,
                        issue_id TEXT NOT NULL,
                        issue_key TEXT NOT NULL,
                        project_key TEXT NOT NULL,
                        issue_type TEXT NOT NULL,
                        summary TEXT NOT NULL,
                        status TEXT NOT NULL,
                        assignee TEXT NOT NULL,
                        start_date TEXT NOT NULL,
                        due_date TEXT NOT NULL,
                        created_utc TEXT NOT NULL,
                        updated_utc TEXT NOT NULL,
                        resolved_stable_since_date TEXT NOT NULL,
                        original_estimate_hours REAL NOT NULL,
                        total_hours_logged REAL NOT NULL,
                        fix_type TEXT NOT NULL,
                        parent_issue_key TEXT NOT NULL,
                        story_key TEXT NOT NULL,
                        epic_key TEXT NOT NULL,
                        raw_payload_json TEXT NOT NULL
                    )
                    """
                )
                conn.execute(
                    """
                    INSERT INTO canonical_issues (
                        run_id, issue_id, issue_key, project_key, issue_type, summary, status, assignee,
                        start_date, due_date, created_utc, updated_utc, resolved_stable_since_date,
                        original_estimate_hours, total_hours_logged, fix_type, parent_issue_key,
                        story_key, epic_key, raw_payload_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        run_id,
                        "1",
                        "O2-321",
                        "O2",
                        "Epic",
                        "Canonical Epic",
                        "In Progress",
                        "",
                        "2026-02-01",
                        "2026-02-28",
                        "",
                        "",
                        "",
                        120.0,
                        0.0,
                        "",
                        "",
                        "",
                        "",
                        "{}",
                    ),
                )
                conn.execute(
                    """
                    INSERT INTO canonical_issues (
                        run_id, issue_id, issue_key, project_key, issue_type, summary, status, assignee,
                        start_date, due_date, created_utc, updated_utc, resolved_stable_since_date,
                        original_estimate_hours, total_hours_logged, fix_type, parent_issue_key,
                        story_key, epic_key, raw_payload_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        run_id,
                        "2",
                        "O2-401",
                        "O2",
                        "Story",
                        "Fallback Story",
                        "To Do",
                        "Alice",
                        "2026-02-03",
                        "2026-02-12",
                        "",
                        "",
                        "",
                        16.0,
                        0.0,
                        "",
                        "O2-321",
                        "O2-401",
                        "O2-321",
                        "{}",
                    ),
                )
                conn.commit()
            finally:
                conn.close()

            data = load_report_data(db_path)

        self.assertEqual(data["canonical_run_id"], run_id)
        self.assertEqual(len(data["epics"]), 1)
        self.assertEqual(data["epics"][0]["story_count"], 1)

    def test_load_report_data_uses_planner_estimates_and_canonical_hierarchy(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "assignee_hours_capacity.db"
            _seed_planner_row(db_path)
            run_id = _seed_canonical_tables(db_path)

            data = load_report_data(db_path, run_id)

        self.assertEqual(data["canonical_run_id"], run_id)
        self.assertEqual(data["database_path"], str(db_path.resolve()))
        self.assertEqual(data["canonical_database_path"], str(db_path.resolve()))
        self.assertEqual(len(data["epics"]), 1)
        self.assertEqual(len(data["rmi_schedule_records"]), 1)
        epic = data["epics"][0]
        self.assertEqual(epic["jira_id"], "O2-321")
        self.assertEqual(epic["title"], "Planner Epic Summary")
        self.assertEqual(epic["product"], "OmniConnect")
        self.assertEqual(epic["most_likely_seconds"], 21 * 28800)
        self.assertEqual(epic["tk_approved_seconds"], 9.8 * 28800)
        self.assertEqual(epic["story_estimate_seconds"], 16 * 3600)
        self.assertEqual(epic["subtask_estimate_seconds"], 8 * 3600)
        self.assertEqual(epic["logged_seconds"], 2 * 3600)
        self.assertEqual(len(epic["stories"]), 1)
        self.assertEqual(epic["stories"][0]["story_key"], "O2-401")
        self.assertEqual(epic["stories"][0]["subtasks"][0]["issue_key"], "O2-402")
        self.assertEqual(epic["epics_planner_remarks"], "Held for next IPP — scope TBD.")
        self.assertEqual(data["summary"]["worklog_count"], 1)

    def test_load_report_data_splits_planner_and_canonical_database_via_env(self):
        """Epics Planner rows can live in one SQLite file and canonical_issues in another."""
        old = os.environ.get("JIRA_RMI_GANTT_CANONICAL_DB_PATH")
        try:
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
                planner = Path(td) / "planner.db"
                canon = Path(td) / "canonical.db"
                _seed_planner_row(planner)
                run_id = _seed_canonical_tables(canon)
                os.environ["JIRA_RMI_GANTT_CANONICAL_DB_PATH"] = str(canon)
                data = load_report_data(planner, run_id)
            self.assertEqual(data["canonical_run_id"], run_id)
            self.assertEqual(data["database_path"], str(planner.resolve()))
            self.assertEqual(data["canonical_database_path"], str(canon.resolve()))
            self.assertEqual(len(data["epics"]), 1)
            self.assertEqual(data["epics"][0]["story_count"], 1)
            self.assertEqual(len(data["rmi_schedule_records"]), 1)
        finally:
            if old is None:
                os.environ.pop("JIRA_RMI_GANTT_CANONICAL_DB_PATH", None)
            else:
                os.environ["JIRA_RMI_GANTT_CANONICAL_DB_PATH"] = old

    def test_render_html_contains_report_controls_and_no_excel_source_text(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "assignee_hours_capacity.db"
            _seed_planner_row(db_path)
            _seed_canonical_tables(db_path)
            data = load_report_data(db_path)

        html = render_html(data)

        self.assertIn("RMI Jira Gantt", html)
        self.assertIn("Generated at", html)
        self.assertIn("Epics Planner (SQLite) joined with canonical Jira issues/worklogs.", html)
        self.assertIn("Epics Planner DB:", html)
        self.assertIn("Canonical DB:", html)
        self.assertIn("Capacity Calculator", html)
        self.assertIn("Month Story Analysis", html)
        self.assertIn("data-month-analysis-included-list", html)
        self.assertIn("data-month-analysis-excluded-list", html)
        self.assertIn("Gantt View", html)
        self.assertIn("Table View", html)
        self.assertIn("RMI Estimation &amp; Scheduling", html)
        self.assertIn("rmi-schedule-table", html)
        self.assertIn("rmi_schedule_records", html)
        self.assertIn("OmniConnect", html)
        self.assertIn("O2-401", html)
        self.assertIn("worklog-nested-table", html)
        self.assertIn("Worklog ID", html)
        self.assertIn("wl-1", html)
        self.assertIn("Alice", html)
        self.assertIn("data-worklog-panel", html)
        self.assertNotIn('id="capacity-employees" type="number"', html)
        self.assertNotIn('id="capacity-leaves" type="number"', html)
        self.assertNotIn("source_rmi_rows", html)
        self.assertNotIn("Epic Estimates Approved Plan.xlsx", html)

    def test_generate_html_report_writes_output_file(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            output_path = root / "rmi_jira_gantt_report.html"
            _seed_planner_row(db_path)
            _seed_canonical_tables(db_path)

            generated = generate_html_report(db_path, output_path)

            self.assertEqual(generated, output_path)
            self.assertTrue(output_path.exists())
            self.assertIn("rmi-report-data", output_path.read_text(encoding="utf-8"))

    def test_report_is_registered_for_navigation_refresh_and_rebuild(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            sources = _resolve_report_html_sources(root)

        nav_hrefs = {str(item.get("href")) for item in STATIC_REPORT_NAV_ITEMS}
        steps = build_html_only_steps()

        self.assertEqual(REPORT_FILENAME_TO_ID["rmi_jira_gantt_report.html"], "rmi_jira_gantt")
        self.assertIn("/rmi_jira_gantt_report.html", nav_hrefs)
        self.assertIn("generate_rmi_jira_gantt_html.py", REPORT_REFRESH_CHAINS["rmi_jira_gantt"])
        self.assertIn(("rmi-jira-gantt-html", "generate_rmi_jira_gantt_html.py"), steps)
        self.assertEqual(sources["rmi_jira_gantt_report.html"], root / "rmi_jira_gantt_report.html")

    def test_server_serves_generated_report_page(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            report_dir = root / "report_html"
            report_dir.mkdir()
            (report_dir / "rmi_jira_gantt_report.html").write_text("<html><body>RMI Jira Gantt</body></html>", encoding="utf-8")
            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            resp = client.get("/rmi_jira_gantt_report.html")

        self.assertEqual(resp.status_code, 200)
        self.assertIn("RMI Jira Gantt", resp.get_data(as_text=True))


if __name__ == "__main__":
    unittest.main()
