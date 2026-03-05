from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from generate_employee_performance_report import (
    DEFAULT_PERFORMANCE_SETTINGS,
    _build_html,
    _build_payload,
    _init_performance_settings_db,
    _load_leave_issue_keys,
    _load_unplanned_leave_rows,
    _load_work_items,
    _load_worklogs,
    _load_performance_settings,
    _normalize_performance_settings,
    _save_performance_settings,
)
from report_server import create_report_server_app


class EmployeePerformanceReportTests(unittest.TestCase):
    def test_settings_validation(self):
        valid = _normalize_performance_settings(
            {
                "base_score": 100,
                "min_score": 0,
                "max_score": 100,
                "points_per_bug_hour": 1,
                "points_per_bug_late_hour": 2,
                "points_per_unplanned_leave_hour": 3,
                "points_per_subtask_late_hour": 4,
                "points_per_estimate_overrun_hour": 5,
                "points_per_missed_due_date": 2,
                "overloaded_penalty_enabled": 1,
                "overloaded_penalty_threshold_pct": 10,
            }
        )
        self.assertEqual(valid["base_score"], 100)

        with self.assertRaises(ValueError):
            _normalize_performance_settings(
                {
                    "base_score": 100,
                    "min_score": 0,
                    "max_score": 100,
                    "points_per_bug_hour": -1,
                    "points_per_bug_late_hour": 2,
                    "points_per_unplanned_leave_hour": 3,
                    "points_per_subtask_late_hour": 4,
                    "points_per_estimate_overrun_hour": 5,
                    "points_per_missed_due_date": 2,
                    "overloaded_penalty_enabled": 1,
                    "overloaded_penalty_threshold_pct": 10,
                }
            )

    def test_settings_db_roundtrip(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            _init_performance_settings_db(db)
            initial = _load_performance_settings(db)
            self.assertIn("base_score", initial)
            saved = _save_performance_settings(
                db,
                {
                    "base_score": 90,
                    "min_score": 0,
                    "max_score": 100,
                    "points_per_bug_hour": 0.8,
                    "points_per_bug_late_hour": 1.9,
                    "points_per_unplanned_leave_hour": 0.6,
                    "points_per_subtask_late_hour": 1.2,
                    "points_per_estimate_overrun_hour": 1.4,
                    "points_per_missed_due_date": 2.0,
                    "overloaded_penalty_enabled": 1,
                    "overloaded_penalty_threshold_pct": 12.5,
                },
            )
            loaded = _load_performance_settings(db)
            self.assertEqual(saved["base_score"], loaded["base_score"])
            self.assertEqual(saved["points_per_bug_late_hour"], loaded["points_per_bug_late_hour"])
            self.assertEqual(saved["overloaded_penalty_enabled"], loaded["overloaded_penalty_enabled"])
            self.assertEqual(saved["overloaded_penalty_threshold_pct"], loaded["overloaded_penalty_threshold_pct"])

    def test_html_contains_core_controls(self):
        payload = _build_payload([], [], [], dict(DEFAULT_PERFORMANCE_SETTINGS), [], [], [], [])
        html = _build_html(payload)
        self.assertIn('id="from"', html)
        self.assertIn('id="to"', html)
        self.assertIn('id="projects"', html)
        self.assertIn('id="leaderboard"', html)
        self.assertIn('href="/settings/performance"', html)
        self.assertIn("Planned Leaves", html)
        self.assertIn("Unplanned Leaves", html)
        self.assertIn('id="shortcut-current-month"', html)
        self.assertIn('id="shortcut-previous-month"', html)
        self.assertIn('id="shortcut-last-30-days"', html)
        self.assertIn('id="shortcut-quarter-to-date"', html)
        self.assertIn('id="assignee-overloaded-penalty-toggle"', html)

    def test_html_applies_subtask_only_scope_for_performance_kpis(self):
        payload = _build_payload([], [], [], dict(DEFAULT_PERFORMANCE_SETTINGS), [], [], [], [])
        html = _build_html(payload)
        self.assertIn("function issueTypeLabel(t)", html)
        self.assertIn("function isSubtaskPerformanceType(t)", html)
        self.assertIn('const issueType = String(r.item_issue_type || r.issue_type || "");', html)
        self.assertIn("return isSubtaskPerformanceType(issueType);", html)
        self.assertIn(
            'const issueType = String(r.jira_issue_type || r.issue_type || r.work_item_type || "");',
            html,
        )
        self.assertIn("if (!isSubtaskPerformanceType(issueType)) return false;", html)
        self.assertIn("const leaveIssueKeySet = new Set(", html)
        self.assertIn("function isLeaveIssueKey(issueKey)", html)
        self.assertNotIn("if (!matchesPlannedRange(r, from, to)) return false;", html)
        self.assertIn("const assignedItemsWork = assignedItems.filter((r) => !isLeaveIssueKey(String(r.issue_key || \"\")));", html)
        self.assertIn("value: n(item.planned_hours_assigned),", html)
        self.assertIn("value: n(item.actual_hours_stats_total),", html)
        self.assertIn("toggle-actual-hours-breakdown", html)
        self.assertIn("Object.entries(item.issue_logged_hours_stats_by_issue || {})", html)
        self.assertIn('id="assignee-extended-actuals-toggle"', html)

    def test_html_subtask_type_helper_includes_subtask_and_bug_subtask_patterns(self):
        payload = _build_payload([], [], [], dict(DEFAULT_PERFORMANCE_SETTINGS), [], [], [], [])
        html = _build_html(payload)
        self.assertIn('if (label.includes("sub-task") || label.includes("subtask")) return true;', html)
        self.assertIn('return label.includes("bug") && (label.includes("sub-task") || label.includes("subtask"));', html)

    def test_load_leave_issue_keys_prefers_raw_subtasks_and_normalizes(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            path = Path(td) / "rlt_leave_report.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Raw_Subtasks"
            ws.append(["issue_key", "assignee"])
            ws.append(["rlt-172", "Maria"])
            ws.append(["RLT-172", "Maria"])
            ws.append([" RLT-173 ", "Maria"])
            ws.append(["", "Alice"])
            wb.save(path)

            keys = _load_leave_issue_keys(path)
            self.assertEqual(keys, ["RLT-172", "RLT-173"])

    def test_build_payload_includes_leave_issue_keys(self):
        payload = _build_payload(
            [],
            [],
            [],
            dict(DEFAULT_PERFORMANCE_SETTINGS),
            [],
            [],
            [],
            [],
            leave_issue_keys=["RLT-172"],
        )
        self.assertEqual(payload["leave_issue_keys"], ["RLT-172"])

    def test_load_unplanned_leave_rows_from_daily_assignee(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            path = Path(td) / "rlt_leave_report.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Daily_Assignee"
            ws.append(["assignee", "period_day", "planned_taken_hours", "unplanned_taken_hours"])
            ws.append(["Alice", "2026-02-01", 8, 0])
            ws.append(["Bob", "2026-02-01", 0, 4])
            wb.save(path)

            rows = _load_unplanned_leave_rows(path)
            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0]["assignee"], "Alice")
            self.assertEqual(rows[0]["planned_taken_hours"], 8)
            self.assertEqual(rows[1]["unplanned_taken_hours"], 4)

    def test_load_unplanned_leave_rows_from_dedicated_leave_sheet(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            path = Path(td) / "employee_performance_leaves.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Leaves"
            ws.append(["Employee", "Date", "Leave Type", "Leave Hours"])
            ws.append(["Alice", "2026-02-03", "Planned", 8])
            ws.append(["Alice", "2026-02-04", "Unplanned", 4])
            wb.save(path)

            rows = _load_unplanned_leave_rows(path)
            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0]["planned_taken_hours"], 8)
            self.assertEqual(rows[0]["unplanned_taken_hours"], 0)
            self.assertEqual(rows[1]["planned_taken_hours"], 0)
            self.assertEqual(rows[1]["unplanned_taken_hours"], 4)

    def test_performance_settings_api(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            tdp = Path(td)
            (tdp / "report_html").mkdir(parents=True, exist_ok=True)
            (tdp / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            wb = Workbook()
            ws = wb.active
            ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
            ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 2.0])
            wb.save(tdp / "assignee_hours_report.xlsx")
            app = create_report_server_app(base_dir=tdp, folder_raw="report_html")
            client = app.test_client()

            get_resp = client.get("/api/performance/settings")
            self.assertEqual(get_resp.status_code, 200)
            body = get_resp.get_json()
            self.assertIn("settings", body)

            post_resp = client.post(
                "/api/performance/settings",
                json={
                    "base_score": 88,
                    "min_score": 0,
                    "max_score": 100,
                    "points_per_bug_hour": 1.1,
                    "points_per_bug_late_hour": 2.1,
                    "points_per_unplanned_leave_hour": 0.9,
                    "points_per_subtask_late_hour": 1.2,
                    "points_per_estimate_overrun_hour": 1.3,
                    "points_per_missed_due_date": 2.0,
                    "overloaded_penalty_enabled": 1,
                    "overloaded_penalty_threshold_pct": 10.0,
                },
            )
            self.assertEqual(post_resp.status_code, 200)
            saved = post_resp.get_json()
            self.assertEqual(saved["settings"]["base_score"], 88)

    def test_performance_team_management_api(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            tdp = Path(td)
            (tdp / "report_html").mkdir(parents=True, exist_ok=True)
            (tdp / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")

            wb = Workbook()
            ws = wb.active
            ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
            ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 2.0])
            ws.append(["O2", "2026-02-02", "2026-02-02", "2026-W05", "2026-02", "Bob", 3.0])
            wb.save(tdp / "assignee_hours_report.xlsx")

            app = create_report_server_app(base_dir=tdp, folder_raw="report_html")
            client = app.test_client()

            assignees_resp = client.get("/api/performance/assignees")
            self.assertEqual(assignees_resp.status_code, 200)
            assignees_json = assignees_resp.get_json()
            self.assertIn("Alice", assignees_json["assignees"])

            create_resp = client.post(
                "/api/performance/teams",
                json={"team_name": "Core Team", "team_leader": "Alice", "assignees": ["Alice", "Bob"]},
            )
            self.assertEqual(create_resp.status_code, 200)
            created = create_resp.get_json()
            self.assertEqual(created["team"]["team_name"], "Core Team")
            self.assertEqual(created["team"]["team_leader"], "Alice")

            update_resp = client.put(
                "/api/performance/teams/Core%20Team",
                json={"team_name": "Core Platform", "team_leader": "Bob", "assignees": ["Bob"]},
            )
            self.assertEqual(update_resp.status_code, 200)
            updated = update_resp.get_json()
            self.assertEqual(updated["team"]["team_name"], "Core Platform")
            self.assertEqual(updated["team"]["team_leader"], "Bob")
            self.assertEqual(updated["team"]["assignees"], ["Bob"])

            list_resp = client.get("/api/performance/teams")
            self.assertEqual(list_resp.status_code, 200)
            teams = list_resp.get_json()["teams"]
            self.assertFalse(any(t["team_name"] == "Core Team" for t in teams))
            self.assertTrue(any(t["team_name"] == "Core Platform" for t in teams))

            del_resp = client.delete("/api/performance/teams/Core%20Platform")
            self.assertEqual(del_resp.status_code, 200)
            self.assertTrue(del_resp.get_json()["deleted"])

    def test_employee_refresh_cancel_marks_running_run(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            tdp = Path(td)
            (tdp / "report_html").mkdir(parents=True, exist_ok=True)
            (tdp / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            wb = Workbook()
            ws = wb.active
            ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
            ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 2.0])
            wb.save(tdp / "assignee_hours_report.xlsx")
            app = create_report_server_app(base_dir=tdp, folder_raw="report_html")
            client = app.test_client()
            db_path = tdp / "assignee_hours_capacity.db"

            with sqlite3.connect(db_path) as conn:
                conn.execute(
                    """
                    INSERT INTO epf_refresh_runs(
                        run_id, started_at_utc, status, trigger_source, error_message, stats_json,
                        progress_step, progress_pct, cancel_requested, updated_at_utc
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        "epf-test-run-1",
                        "2026-03-05T00:00:00+00:00",
                        "running",
                        "api_refresh_async",
                        "",
                        "{}",
                        "fetching_worklogs",
                        20,
                        0,
                        "2026-03-05T00:00:00+00:00",
                    ),
                )
                conn.commit()

            cancel_resp = client.post("/api/employee-performance/cancel", json={"run_id": "epf-test-run-1"})
            self.assertEqual(cancel_resp.status_code, 200)
            body = cancel_resp.get_json() or {}
            self.assertTrue(body.get("ok"))
            self.assertEqual(body.get("status"), "cancel_requested")

            with sqlite3.connect(db_path) as conn:
                row = conn.execute("SELECT cancel_requested FROM epf_refresh_runs WHERE run_id = ?", ("epf-test-run-1",)).fetchone()
            self.assertIsNotNone(row)
            self.assertEqual(int(row[0] or 0), 1)

    def test_employee_refresh_current_auto_finalizes_stale_canceled_running_run(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            tdp = Path(td)
            (tdp / "report_html").mkdir(parents=True, exist_ok=True)
            (tdp / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            wb = Workbook()
            ws = wb.active
            ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
            ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 2.0])
            wb.save(tdp / "assignee_hours_report.xlsx")
            app = create_report_server_app(base_dir=tdp, folder_raw="report_html")
            client = app.test_client()
            db_path = tdp / "assignee_hours_capacity.db"

            with sqlite3.connect(db_path) as conn:
                conn.execute(
                    """
                    INSERT INTO epf_refresh_runs(
                        run_id, started_at_utc, status, trigger_source, error_message, stats_json,
                        progress_step, progress_pct, cancel_requested, updated_at_utc
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        "epf-test-stale-run-1",
                        "2026-03-05T00:00:00+00:00",
                        "running",
                        "api_refresh_async",
                        "",
                        "{}",
                        "",
                        0,
                        1,
                        "2026-03-05T00:00:00+00:00",
                    ),
                )
                conn.commit()

            resp = client.get("/api/employee-performance/refresh/current")
            self.assertEqual(resp.status_code, 200)
            body = resp.get_json() or {}
            self.assertTrue(body.get("ok"))
            run = body.get("run") or {}
            self.assertEqual(run.get("run_id"), "epf-test-stale-run-1")
            self.assertEqual(run.get("status"), "canceled")
            self.assertEqual(int(run.get("progress") or 0), 100)

            with sqlite3.connect(db_path) as conn:
                row = conn.execute(
                    "SELECT status, progress_step, progress_pct, ended_at_utc FROM epf_refresh_runs WHERE run_id = ?",
                    ("epf-test-stale-run-1",),
                ).fetchone()
            self.assertIsNotNone(row)
            self.assertEqual(str(row[0]), "canceled")
            self.assertEqual(str(row[1]), "canceled")
            self.assertEqual(int(row[2] or 0), 100)
            self.assertTrue(str(row[3] or "").strip())

    def test_fix_type_rework_flows_from_work_items_to_worklogs(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            tdp = Path(td)
            work_items_xlsx = tdp / "1_jira_work_items_export.xlsx"
            worklogs_xlsx = tdp / "2_jira_subtask_worklogs.xlsx"

            wi_wb = Workbook()
            wi_ws = wi_wb.active
            wi_ws.append(
                [
                    "project_key",
                    "issue_key",
                    "work_item_id",
                    "work_item_type",
                    "jira_issue_type",
                    "fix_type",
                    "summary",
                    "status",
                    "resolved_stable_since_date",
                    "end_date",
                    "original_estimate_hours",
                    "parent_issue_key",
                ]
            )
            wi_ws.append(
                ["O2", "O2-101", "O2-101", "Subtask", "Sub-task", "rework", "A", "Open", "2026-02-12", "2026-02-10", 8, "O2-100"]
            )
            wi_wb.save(work_items_xlsx)

            wl_wb = Workbook()
            wl_ws = wl_wb.active
            wl_ws.append(["issue_id", "issue_assignee", "worklog_started", "hours_logged", "issue_type", "parent_story_id"])
            wl_ws.append(["O2-101", "Alice", "2026-02-11T10:00:00+00:00", 2, "Sub-task", "O2-100"])
            wl_wb.save(worklogs_xlsx)

            work_items = _load_work_items(work_items_xlsx)
            self.assertEqual(work_items["O2-101"]["fix_type"], "rework")
            self.assertEqual(work_items["O2-101"]["resolved_stable_since_date"], "2026-02-12")

            worklogs = _load_worklogs(worklogs_xlsx, work_items)
            self.assertEqual(len(worklogs), 1)
            self.assertEqual(worklogs[0]["fix_type"], "rework")
            self.assertEqual(worklogs[0]["item_resolved_stable_since_date"], "2026-02-12")


if __name__ == "__main__":
    unittest.main()
