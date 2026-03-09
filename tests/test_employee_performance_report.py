from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from generate_employee_performance_report import (
    DEFAULT_PERFORMANCE_SETTINGS,
    _derive_actual_completion,
    _build_html,
    _build_payload,
    _init_performance_settings_db,
    _load_leave_issue_keys,
    _load_unplanned_leave_rows,
    _load_simple_scoring,
    _load_work_items,
    _load_worklogs,
    _load_performance_settings,
    _normalize_performance_settings,
    _precompute_simple_scoring,
    _save_performance_settings,
)
import report_server
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
                "planning_realism_enabled": 0,
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
                    "planning_realism_enabled": 0,
                    "overloaded_penalty_threshold_pct": 10,
                }
            )

    def test_settings_db_roundtrip(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            _init_performance_settings_db(db)
            initial = _load_performance_settings(db)
            self.assertIn("base_score", initial)
            self.assertEqual(initial["overloaded_penalty_enabled"], 0)
            self.assertEqual(initial["planning_realism_enabled"], 0)
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
                    "planning_realism_enabled": 1,
                    "overloaded_penalty_threshold_pct": 12.5,
                },
            )
            loaded = _load_performance_settings(db)
            self.assertEqual(saved["base_score"], loaded["base_score"])
            self.assertEqual(saved["points_per_bug_late_hour"], loaded["points_per_bug_late_hour"])
            self.assertEqual(saved["overloaded_penalty_enabled"], loaded["overloaded_penalty_enabled"])
            self.assertEqual(saved["planning_realism_enabled"], loaded["planning_realism_enabled"])
            self.assertEqual(saved["overloaded_penalty_threshold_pct"], loaded["overloaded_penalty_threshold_pct"])

    def test_html_contains_core_controls(self):
        payload = _build_payload([], [], [], dict(DEFAULT_PERFORMANCE_SETTINGS), [], [], [], [])
        html = _build_html(payload)
        self.assertIn('id="from"', html)
        self.assertIn('id="to"', html)
        self.assertIn('id="projects"', html)
        self.assertIn('id="leaderboard"', html)
        self.assertIn('id="leader-sort"', html)
        self.assertIn('id="leader-sort-direction"', html)
        self.assertIn("Available for more work", html)
        self.assertIn('<option value="desc" selected>Descending</option>', html)
        self.assertIn('href="/settings/performance"', html)
        self.assertIn("Planned Leaves", html)
        self.assertIn("Unplanned Leaves", html)
        self.assertIn('id="shortcut-current-month"', html)
        self.assertIn('id="shortcut-previous-month"', html)
        self.assertIn('id="shortcut-last-30-days"', html)
        self.assertIn('id="shortcut-quarter-to-date"', html)
        self.assertIn('id="assignee-overloaded-penalty-toggle"', html)
        self.assertIn('id="assignee-planning-realism-toggle"', html)
        self.assertIn('id="simple-overrun-mode"', html)
        self.assertIn("Overrun Subtask Hours", html)
        self.assertIn("Total Overrun Hours", html)
        self.assertIn('id="header-average-performance-value"', html)
        self.assertIn('id="header-average-performance-mode"', html)
        self.assertIn('id="header-average-performance-meta"', html)
        self.assertIn('fetch("/api/performance/settings")', html)
        self.assertIn("function applyPerformanceSettings(nextSettings)", html)
        self.assertIn("hydratePerformanceSettings().finally(() => {", html)
        self.assertIn("let performanceSettingsReady = false;", html)
        self.assertIn("Loading performance settings before calculating scores.", html)
        self.assertIn("if (!performanceSettingsReady) {", html)
        self.assertIn("performanceSettingsReady = true;", html)
        self.assertIn('document.getElementById("leader-sort-direction").addEventListener("change", renderAll);', html)
        self.assertIn('document.getElementById("leader-sort-direction").value="desc";', html)

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

    def test_html_marks_zero_planned_hours_as_score_na(self):
        payload = _build_payload([], [], [], dict(DEFAULT_PERFORMANCE_SETTINGS), [], [], [], [])
        html = _build_html(payload)
        self.assertIn("function isScoreEligible(item)", html)
        self.assertIn('return Number.isFinite(value) ? value.toFixed(1) : "N/A";', html)
        self.assertIn("Scoring N/A", html)
        self.assertIn("not eligible for scoring", html)
        self.assertIn("Planned Hours Assigned is", html)

    def test_html_contains_simple_score_drawer_controls(self):
        payload = _build_payload([], [], [], dict(DEFAULT_PERFORMANCE_SETTINGS), [], [], [], [])
        html = _build_html(payload)
        self.assertIn('id="score-detail-drawer"', html)
        self.assertIn('id="score-detail-drawer-overlay"', html)
        self.assertIn('id="summary-simple-score-trigger"', html)
        self.assertIn("openScoreDrawerForAssignee(item)", html)
        self.assertIn("Penalized Subtasks", html)
        self.assertIn("Actual Complete Date", html)
        self.assertIn("Last Logged Date", html)
        self.assertIn("Planned Due Date", html)

    def test_html_overloaded_penalty_uses_capacity_threshold_formula(self):
        payload = _build_payload([], [], [], dict(DEFAULT_PERFORMANCE_SETTINGS), [], [], [], [])
        html = _build_html(payload)
        self.assertIn("if Planned > Capacity x (1 + N/100)", html)
        self.assertIn("const effectiveCapacity = Math.max(0, n(it.employee_capacity_hours));", html)
        self.assertIn("const maxPlannedBeforeCap = effectiveCapacity * (1 + overloadedPenaltyThresholdPct / 100);", html)
        self.assertIn("Capacity/Planned", html)
        self.assertIn("planningRealismEnabled", html)
        self.assertIn("simple_score_overloaded_penalty_pct", html)
        self.assertIn("simple_score_overrun_active", html)

    def test_html_recomputes_simple_score_when_extended_actuals_toggle_changes_actuals(self):
        payload = _build_payload([], [], [], dict(DEFAULT_PERFORMANCE_SETTINGS), [], [], [], [])
        html = _build_html(payload)
        self.assertIn("const actual = n(statsByIssue[issueKey]);", html)
        self.assertIn("const overrun = estimate > 0 ? Math.max(0, actual - estimate) : 0;", html)
        self.assertIn('const isCommitment = estimateStatus === "over_estimate" && dueStatus === "on_time" ? 1 : 0;', html)
        self.assertIn("it.ss_total_actual = liveTotalActual;", html)
        self.assertIn("it.ss_total_overrun = liveTotalOverrun;", html)
        self.assertIn('simpleOverrunMode === "total" ? totalOverTotal : totalOverSubtasks', html)
        self.assertIn('syncSimpleOverrunMode("subtasks")', html)

    def test_html_due_completion_mode_mentions_late_completion_penalties(self):
        payload = _build_payload([], [], [], dict(DEFAULT_PERFORMANCE_SETTINGS), [], [], [], [])
        html = _build_html(payload)
        self.assertIn("Late Completion Rule", html)
        self.assertIn("Late-completed subtasks add their original estimate as penalty input", html)
        self.assertIn("Late Completion Estimate Penalty", html)
        self.assertIn("Late estimate penalty:", html)

    def test_html_due_compliance_table_uses_updated_completion_headers(self):
        payload = _build_payload([], [], [], dict(DEFAULT_PERFORMANCE_SETTINGS), [], [], [], [])
        html = _build_html(payload)
        self.assertIn("<th>Due</th><th>Last Logged Date</th><th>Actual Completed Date</th><th>Stable Resolved</th><th>Bucket</th>", html)
        self.assertIn("Actual complete date came from last logged date after resolved date.", html)
        self.assertIn("Actual complete date came from resolved date.", html)

    def test_derive_actual_completion_prefers_last_log_when_after_resolved(self):
        meta = _derive_actual_completion("2026-03-05", "2026-03-07", "2026-03-03")
        self.assertEqual(meta["actual_complete_date"], "2026-03-07")
        self.assertEqual(meta["actual_complete_source"], "last_logged_date")
        self.assertEqual(meta["completion_bucket"], "after_due")

    def test_derive_actual_completion_prefers_resolved_when_after_last_log(self):
        meta = _derive_actual_completion("2026-03-05", "2026-03-06", "2026-03-08")
        self.assertEqual(meta["actual_complete_date"], "2026-03-08")
        self.assertEqual(meta["actual_complete_source"], "resolved_stable_since_date")
        self.assertEqual(meta["completion_bucket"], "after_due")

    def test_derive_actual_completion_handles_missing_dates(self):
        only_last = _derive_actual_completion("2026-03-05", "2026-03-04", "")
        only_resolved = _derive_actual_completion("2026-03-05", "", "2026-03-04")
        none = _derive_actual_completion("2026-03-05", "", "")
        self.assertEqual(only_last["actual_complete_date"], "2026-03-04")
        self.assertEqual(only_resolved["actual_complete_date"], "2026-03-04")
        self.assertEqual(none["actual_complete_date"], "")
        self.assertEqual(none["completion_bucket"], "not_completed")

    def test_precompute_simple_scoring_persists_actual_completion_fields(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            work_items = {
                "O2-101": {
                    "issue_type": "Sub-task",
                    "assignee": "Alice",
                    "original_estimate_hours": 8,
                    "due_date": "2026-03-05",
                    "resolved_stable_since_date": "2026-03-03",
                    "status": "Done",
                },
                "O2-102": {
                    "issue_type": "Sub-task",
                    "assignee": "Alice",
                    "original_estimate_hours": 8,
                    "due_date": "2026-03-08",
                    "resolved_stable_since_date": "2026-03-08",
                    "status": "Done",
                },
            }
            worklogs = [
                {"issue_id": "O2-101", "worklog_date": "2026-03-07", "hours_logged": 2},
                {"issue_id": "O2-102", "worklog_date": "2026-03-06", "hours_logged": 2},
            ]

            _precompute_simple_scoring(db, work_items, worklogs)
            rows = {row["issue_key"]: row for row in _load_simple_scoring(db)}

            self.assertEqual(rows["O2-101"]["planned_due_date"], "2026-03-05")
            self.assertEqual(rows["O2-101"]["last_logged_date"], "2026-03-07")
            self.assertEqual(rows["O2-101"]["actual_complete_date"], "2026-03-07")
            self.assertEqual(rows["O2-101"]["actual_complete_source"], "last_logged_date")
            self.assertEqual(rows["O2-101"]["due_completion_status"], "late")

            self.assertEqual(rows["O2-102"]["actual_complete_date"], "2026-03-08")
            self.assertEqual(rows["O2-102"]["actual_complete_source"], "resolved_stable_since_date")
            self.assertEqual(rows["O2-102"]["due_completion_status"], "on_time")

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
                    "planning_realism_enabled": 0,
                    "overloaded_penalty_threshold_pct": 10.0,
                },
            )
            self.assertEqual(post_resp.status_code, 200)
            saved = post_resp.get_json()
            self.assertEqual(saved["settings"]["base_score"], 88)

    def test_performance_settings_page_uses_planning_realism_labels(self):
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

            resp = client.get("/settings/performance")
            self.assertEqual(resp.status_code, 200)
            html = resp.get_data(as_text=True)
            self.assertIn("Overloaded Penalty", html)
            self.assertIn("Overload Capping/ Planning Realism", html)
            self.assertIn("Overloaded Threshold (%)", html)
            self.assertIn("If Overload Capping/ Planning Realism is OFF, the overload penalty is deducted", html)
            self.assertIn("Overload Capping/ Planning Realism = ", html)
            self.assertIn("Default is OFF", html)

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

    def test_employee_refresh_passes_target_assignee_to_source_scripts(self):
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
            refresh_fn = app.view_functions["employee_performance_refresh"]
            refresh_globals = refresh_fn.__globals__
            db_path = tdp / "assignee_hours_capacity.db"

            run_envs: list[tuple[str, dict[str, str]]] = []

            def fake_run_script_interruptible(script_name, _cwd, env_overrides=None, extra_args=None, cancel_check=None):
                run_envs.append((script_name, dict(env_overrides or {})))
                if script_name == "generate_employee_performance_report.py":
                    return 0, "ok", ""

                if script_name == "export_jira_subtask_worklogs.py":
                    out_path = Path((env_overrides or {})["JIRA_WORKLOG_XLSX_PATH"])
                    book = Workbook()
                    sheet = book.active
                    sheet.append(["issue_id", "issue_assignee", "worklog_started", "hours_logged", "issue_type", "parent_story_id"])
                    sheet.append(["O2-101", "Alice", "2026-02-11T10:00:00+00:00", 2, "Sub-task", "O2-100"])
                    book.save(out_path)
                    return 0, "ok", ""

                if script_name == "export_jira_work_items.py":
                    out_path = Path((env_overrides or {})["JIRA_EXPORT_XLSX_PATH"])
                    book = Workbook()
                    sheet = book.active
                    sheet.append([
                        "project_key", "issue_key", "work_item_id", "work_item_type", "jira_issue_type", "fix_type",
                        "summary", "status", "resolved_stable_since_date", "end_date", "original_estimate_hours",
                        "parent_issue_key", "assignee"
                    ])
                    sheet.append(["O2", "O2-101", "O2-101", "Subtask", "Sub-task", "", "A", "Open", "", "2026-02-10", 8, "O2-100", "Alice"])
                    book.save(out_path)
                    return 0, "ok", ""

                if script_name == "generate_rlt_leave_report.py":
                    out_path = Path(extra_args[-1]) if extra_args else tdp / "epf_leave.xlsx"
                    book = Workbook()
                    daily = book.active
                    daily.title = "Daily"
                    daily.append(["assignee", "period_day", "unplanned_taken_hours", "planned_taken_hours"])
                    daily.append(["Alice", "2026-02-11", 0, 0])
                    raw_sub = book.create_sheet("Raw_Subtasks")
                    raw_sub.append(["issue_key"])
                    book.save(out_path)
                    return 0, "ok", ""

                return 0, "ok", ""

            with sqlite3.connect(db_path) as conn:
                conn.execute(
                    "UPDATE epf_refresh_state SET active_run_id = ?, last_success_run_id = ?, updated_at_utc = ? WHERE id = 1",
                    ("seed-run", "seed-run", "2026-03-05T00:00:00+00:00"),
                )
                conn.execute(
                    "INSERT INTO epf_work_items(run_id, issue_key, project_key, issue_type, fix_type, summary, status, assignee, start_date, due_date, resolved_stable_since_date, original_estimate_hours, parent_issue_key) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    ("seed-run", "O2-101", "O2", "Subtask", "", "A", "Open", "Alice", "", "2026-02-10", "", 8.0, "O2-100"),
                )
                conn.execute(
                    "INSERT INTO epf_worklogs(run_id, issue_id, issue_assignee, worklog_date, hours_logged, project_key, is_bug, fix_type, item_summary, item_status, item_issue_type, item_assignee, item_parent_issue_key, item_start_date, item_due_date, item_resolved_stable_since_date, story_due_date, original_estimate_hours) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    ("seed-run", "O2-101", "Alice", "2026-02-11", 2.0, "O2", 0, "", "A", "Open", "Sub-task", "Alice", "O2-100", "", "2026-02-10", "", "", 8.0),
                )
                conn.commit()

            with patch.dict(refresh_globals, {"_run_script_interruptible": fake_run_script_interruptible, "sync_report_html": lambda *_args, **_kwargs: None}):
                client = app.test_client()
                response = client.post("/api/employee-performance/refresh", json={"assignee": "Alice", "replace_running": False})
                self.assertEqual(response.status_code, 202)

                import time

                deadline = time.time() + 5
                while time.time() < deadline:
                    with sqlite3.connect(db_path) as conn:
                        row = conn.execute("SELECT status FROM epf_refresh_runs ORDER BY started_at_utc DESC LIMIT 1").fetchone()
                    if row and row[0] != "running":
                        break
                    time.sleep(0.05)

            env_by_script = {name: env for name, env in run_envs}
            self.assertEqual(env_by_script["export_jira_subtask_worklogs.py"].get("JIRA_TARGET_ASSIGNEE"), "Alice")
            self.assertEqual(env_by_script["export_jira_work_items.py"].get("JIRA_TARGET_ASSIGNEE"), "Alice")
            self.assertEqual(env_by_script["generate_rlt_leave_report.py"].get("JIRA_TARGET_ASSIGNEE"), "Alice")

            with sqlite3.connect(db_path) as conn:
                run_row = conn.execute(
                    "SELECT run_id, status FROM epf_refresh_runs ORDER BY started_at_utc DESC LIMIT 1"
                ).fetchone()
            self.assertIsNotNone(run_row)
            run_id = str(run_row[0] or "")
            self.assertTrue(run_id)

            with sqlite3.connect(db_path) as conn:
                wi_count = conn.execute(
                    "SELECT COUNT(*) FROM epf_work_items WHERE run_id = ?",
                    (run_id,),
                ).fetchone()
                wl_count = conn.execute(
                    "SELECT COUNT(*) FROM epf_worklogs WHERE run_id = ?",
                    (run_id,),
                ).fetchone()
                lr_count = conn.execute(
                    "SELECT COUNT(*) FROM epf_leave_rows WHERE run_id = ?",
                    (run_id,),
                ).fetchone()
                li_count = conn.execute(
                    "SELECT COUNT(*) FROM epf_leave_issue_keys WHERE run_id = ?",
                    (run_id,),
                ).fetchone()

            self.assertIsNotNone(wi_count)
            self.assertIsNotNone(wl_count)
            self.assertIsNotNone(lr_count)
            self.assertIsNotNone(li_count)
            self.assertGreater(int(wi_count[0] or 0), 0)
            self.assertGreater(int(wl_count[0] or 0), 0)
            self.assertGreater(int(lr_count[0] or 0), 0)
            # leave_issue_keys can legitimately be 0 if there were no defective/no-entry rows
            self.assertGreaterEqual(int(li_count[0] or 0), 0)

    def test_employee_full_refresh_persists_epf_snapshot_via_generic_refresh(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            tdp = Path(td)
            (tdp / "report_html").mkdir(parents=True, exist_ok=True)
            (tdp / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")

            app = create_report_server_app(base_dir=tdp, folder_raw="report_html")
            client = app.test_client()
            db_path = tdp / "assignee_hours_capacity.db"

            original_interruptible = report_server._run_script_interruptible
            original_sync_report_html = report_server.sync_report_html

            def _fake_interruptible(script_name, _cwd, env_overrides=None, extra_args=None, cancel_check=None):
                env_overrides = env_overrides or {}
                if script_name == "export_jira_subtask_worklogs.py":
                    out_path = Path(env_overrides["JIRA_WORKLOG_XLSX_PATH"])
                    book = Workbook()
                    sheet = book.active
                    sheet.append(
                        ["issue_id", "issue_assignee", "worklog_started", "hours_logged", "issue_type", "parent_story_id"]
                    )
                    sheet.append(
                        ["O2-101", "Alice", "2026-02-11T10:00:00+00:00", 2.0, "Sub-task", "O2-100"]
                    )
                    book.save(out_path)
                    return 0, "ok", ""

                if script_name == "export_jira_work_items.py":
                    out_path = Path(env_overrides["JIRA_EXPORT_XLSX_PATH"])
                    book = Workbook()
                    sheet = book.active
                    sheet.append(
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
                            "assignee",
                        ]
                    )
                    sheet.append(
                        [
                            "O2",
                            "O2-101",
                            "O2-101",
                            "Subtask",
                            "Sub-task",
                            "",
                            "A",
                            "Open",
                            "",
                            "2026-02-10",
                            8.0,
                            "O2-100",
                            "Alice",
                        ]
                    )
                    book.save(out_path)
                    return 0, "ok", ""

                if script_name == "generate_rlt_leave_report.py":
                    out_path = Path(extra_args[-1]) if extra_args else tdp / "epf_leave.xlsx"
                    book = Workbook()
                    daily = book.active
                    daily.title = "Daily"
                    daily.append(["assignee", "period_day", "unplanned_taken_hours", "planned_taken_hours"])
                    daily.append(["Alice", "2026-02-11", 0.0, 0.0])
                    raw_sub = book.create_sheet("Raw_Subtasks")
                    raw_sub.append(["issue_key"])
                    raw_sub.append(["O2-101"])
                    book.save(out_path)
                    return 0, "ok", ""

                if script_name == "generate_employee_performance_report.py":
                    return 0, "ok", ""

                return 0, "ok", ""

            report_server._run_script_interruptible = _fake_interruptible
            report_server.sync_report_html = lambda *_args, **_kwargs: None
            try:
                resp = client.post(
                    "/api/report/refresh",
                    json={"report": "employee_performance", "isolated": True},
                )
                self.assertEqual(resp.status_code, 200)
                body = resp.get_json() or {}
                self.assertTrue(body.get("ok"))
                run_id = str(body.get("run_id") or "")
                self.assertTrue(run_id)

                with sqlite3.connect(db_path) as conn:
                    wi_count = conn.execute(
                        "SELECT COUNT(*) FROM epf_work_items WHERE run_id = ?",
                        (run_id,),
                    ).fetchone()
                    wl_count = conn.execute(
                        "SELECT COUNT(*) FROM epf_worklogs WHERE run_id = ?",
                        (run_id,),
                    ).fetchone()
                    lr_count = conn.execute(
                        "SELECT COUNT(*) FROM epf_leave_rows WHERE run_id = ?",
                        (run_id,),
                    ).fetchone()
                    li_count = conn.execute(
                        "SELECT COUNT(*) FROM epf_leave_issue_keys WHERE run_id = ?",
                        (run_id,),
                    ).fetchone()

                self.assertIsNotNone(wi_count)
                self.assertIsNotNone(wl_count)
                self.assertIsNotNone(lr_count)
                self.assertIsNotNone(li_count)
                self.assertGreater(int(wi_count[0] or 0), 0)
                self.assertGreater(int(wl_count[0] or 0), 0)
                self.assertGreater(int(lr_count[0] or 0), 0)
                self.assertGreater(int(li_count[0] or 0), 0)
            finally:
                report_server._run_script_interruptible = original_interruptible
                report_server.sync_report_html = original_sync_report_html

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
