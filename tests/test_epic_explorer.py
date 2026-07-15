from __future__ import annotations

import json
import sqlite3
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

from epic_explorer_service import build_epic_explorer_payload
from report_server import _init_epics_management_db, create_report_server_app, sync_report_html


def _create_canonical_tables(db_path: Path) -> None:
    with sqlite3.connect(db_path) as conn:
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
            "INSERT INTO canonical_refresh_state(id, active_run_id, last_success_run_id, updated_at_utc) VALUES (1, 'run-1', 'run-1', '2026-05-31T00:00:00Z')"
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
                raw_payload_json TEXT NOT NULL DEFAULT '{}',
                PRIMARY KEY (run_id, issue_key)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE canonical_worklogs (
                run_id TEXT NOT NULL,
                worklog_id TEXT NOT NULL,
                issue_key TEXT NOT NULL DEFAULT '',
                project_key TEXT NOT NULL DEFAULT '',
                worklog_author TEXT NOT NULL DEFAULT '',
                issue_assignee TEXT NOT NULL DEFAULT '',
                started_utc TEXT NOT NULL DEFAULT '',
                started_date TEXT NOT NULL DEFAULT '',
                updated_utc TEXT NOT NULL DEFAULT '',
                hours_logged REAL NOT NULL DEFAULT 0,
                PRIMARY KEY (run_id, worklog_id)
            )
            """
        )
        conn.commit()


def _insert_issue(
    conn: sqlite3.Connection,
    key: str,
    issue_type: str,
    *,
    project: str = "O2",
    summary: str = "",
    status: str = "In Progress",
    assignee: str = "",
    start: str = "",
    due: str = "",
    resolved: str = "",
    estimate: float = 0.0,
    actual: float = 0.0,
    parent: str = "",
    story: str = "",
    epic: str = "",
) -> None:
    conn.execute(
        """
        INSERT INTO canonical_issues(
            run_id, issue_key, project_key, issue_type, summary, status, assignee,
            start_date, due_date, resolved_stable_since_date, original_estimate_hours,
            total_hours_logged, parent_issue_key, story_key, epic_key
        ) VALUES ('run-1', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (key, project, issue_type, summary or key, status, assignee, start, due, resolved, estimate, actual, parent, story, epic),
    )


def _seed_epic_explorer_db(db_path: Path) -> None:
    _create_canonical_tables(db_path)
    with sqlite3.connect(db_path) as conn:
        _insert_issue(
            conn,
            "O2-1",
            "Epic",
            summary="Explorer Epic",
            status="Done",
            assignee="Lead",
            start="2026-05-01",
            due="2026-05-31",
            resolved="2026-05-20",
            estimate=100,
            epic="O2-1",
        )
        _insert_issue(conn, "O2-1-S1", "Story", summary="Explorer Story", start="2026-05-03", due="2026-05-18", estimate=40, parent="O2-1", story="O2-1-S1", epic="O2-1")
        _insert_issue(conn, "O2-1-S2", "Story", summary="Cross Month Story", start="2026-05-29", due="2026-06-02", estimate=30, parent="O2-1", story="O2-1-S2", epic="O2-1")
        _insert_issue(conn, "O2-1-S1-T1", "Sub-task", summary="Build table", assignee="Alice", start="2026-05-03", due="2026-05-10", estimate=16, parent="O2-1-S1", story="O2-1-S1", epic="O2-1")
        _insert_issue(conn, "O2-1-S1-B1", "Bug Subtask", summary="Fix drilldown", assignee="Bob", start="2026-05-11", due="2026-05-18", estimate=8, parent="O2-1-S1", story="O2-1-S1", epic="O2-1")
        _insert_issue(conn, "O2-1-T1", "Task", summary="Direct Epic Task", start="2026-05-12", due="2026-05-24", estimate=12, parent="O2-1", epic="O2-1")
        _insert_issue(conn, "O2-1-T1-SUB1", "Sub-task", summary="Task child", assignee="Dana", start="2026-05-12", due="2026-05-20", estimate=6, actual=12, parent="O2-1-T1", story="O2-1-T1", epic="O2-1")
        _insert_issue(
            conn,
            "CRM-2",
            "Epic",
            project="CRM",
            summary="Outside Epic",
            start="2026-04-01",
            due="2026-04-30",
            estimate=20,
            epic="CRM-2",
        )
        _insert_issue(conn, "CRM-2-S1", "Story", project="CRM", start="2026-04-01", due="2026-04-30", estimate=20, parent="CRM-2", story="CRM-2-S1", epic="CRM-2")
        conn.executemany(
            """
            INSERT INTO canonical_worklogs(run_id, worklog_id, issue_key, project_key, worklog_author, issue_assignee, started_date, hours_logged)
            VALUES ('run-1', ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                ("wl-1", "O2-1-S1-T1", "O2", "Alice", "Alice", "2026-05-05", 10),
                ("wl-2", "O2-1-S1-B1", "O2", "Bob", "Bob", "2026-06-01", 2),
                ("wl-3", "O2-1-T1-SUB1", "O2", "Dana", "Dana", "2026-05-15", 4),
            ],
        )
        conn.commit()

    _init_epics_management_db(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO epics_management(id, epic_key, project_key, project_name, product_category, component, epic_name, delivery_status, jira_url, epic_plan_json)
            VALUES (?, ?, ?, ?, ?, '', ?, 'On-track', ?, ?)
            """,
            (
                "O2-1",
                "O2-1",
                "O2",
                "Omni",
                "Core",
                "Explorer Epic",
                "https://jira.example/browse/O2-1",
                json.dumps({"man_days": 5, "start_date": "2026-05-01", "due_date": "2026-05-31"}),
            ),
        )
        conn.commit()


class EpicExplorerTests(unittest.TestCase):
    def test_unresolved_or_reopened_epic_uses_today_instead_of_last_worklog(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "assignee_hours_capacity.db"
            _create_canonical_tables(db_path)
            with sqlite3.connect(db_path) as conn:
                _insert_issue(
                    conn,
                    "O2-OPEN",
                    "Epic",
                    status="In Progress",
                    start="2026-07-01",
                    due="2026-08-31",
                    resolved="2026-07-04",
                    estimate=62,
                    epic="O2-OPEN",
                )
                _insert_issue(
                    conn,
                    "O2-OPEN-S1",
                    "Story",
                    status="In Progress",
                    start="2026-07-01",
                    due="2026-08-31",
                    estimate=62,
                    parent="O2-OPEN",
                    story="O2-OPEN-S1",
                    epic="O2-OPEN",
                )
                _insert_issue(
                    conn,
                    "O2-OPEN-S1-T1",
                    "Sub-task",
                    status="In Progress",
                    start="2026-07-01",
                    due="2026-08-31",
                    parent="O2-OPEN-S1",
                    story="O2-OPEN-S1",
                    epic="O2-OPEN",
                )
                conn.execute(
                    """
                    INSERT INTO canonical_worklogs(
                        run_id, worklog_id, issue_key, project_key, worklog_author,
                        issue_assignee, started_date, hours_logged
                    ) VALUES ('run-1', 'wl-open', 'O2-OPEN-S1-T1', 'O2', 'Alice', 'Alice', '2026-07-05', 2)
                    """
                )
                conn.commit()

            with patch("epic_explorer_service.date", wraps=date) as mocked_date:
                mocked_date.today.return_value = date(2026, 7, 15)
                payload = build_epic_explorer_payload(db_path, [], "run-1")

            row = payload["rows"][0]
            self.assertEqual(row["actual_complete_date"], "")
            self.assertEqual(row["actual_complete_source"], "none")
            self.assertEqual(row["schedule_variance_date_basis"], "2026-07-15")
            self.assertEqual(row["schedule_variance_date_basis_type"], "current_date")
            self.assertEqual(row["schedule_variance_days"], 47)
            self.assertEqual(row["planned_to_date_hours"], 15.0)
            self.assertEqual(row["actual_to_date_hours"], 2.0)
            self.assertEqual(row["schedule_variance_hours"], -13.0)

    def test_payload_rolls_up_full_epic_data_and_filters_only_epic_scope(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "assignee_hours_capacity.db"
            _seed_epic_explorer_db(db_path)

            planner_rows = [
                {
                    "epic_key": "O2-1",
                    "project_key": "O2",
                    "project_name": "Omni",
                    "product_category": "Core",
                    "epic_name": "Explorer Epic",
                    "jira_url": "https://jira.example/browse/O2-1",
                    "plans": {"epic_plan": {"man_days": 5}},
                }
            ]
            with patch(
                "epic_explorer_service.build_rlt_leave_snapshot",
                return_value={
                    "daily": [
                        {"assignee": "Alice", "period_day": "2026-05-06", "planned_taken_hours": 8, "planned_not_taken_hours": 0, "unplanned_taken_hours": 0},
                        {"assignee": "Charlie", "period_day": "2026-05-06", "planned_taken_hours": 8, "planned_not_taken_hours": 0, "unplanned_taken_hours": 4},
                        {"assignee": "Dana", "period_day": "2026-05-16", "planned_taken_hours": 0, "planned_not_taken_hours": 0, "unplanned_taken_hours": 4},
                    ]
                },
            ):
                payload = build_epic_explorer_payload(
                    db_path,
                    planner_rows,
                    "run-1",
                    from_date="2026-05-01",
                    to_date="2026-05-31",
                )

            self.assertTrue(payload["date_filter_active"])
            self.assertEqual([row["epic_key"] for row in payload["rows"]], ["O2-1"])
            row = payload["rows"][0]
            self.assertEqual(row["product"], "Core")
            self.assertEqual(row["tk_budget_hours"], 40.0)
            self.assertEqual(row["jira_original_estimate_hours"], 100.0)
            self.assertEqual(row["story_estimate_hours"], 82.0)
            self.assertEqual(row["subtask_estimate_hours"], 30.0)
            self.assertEqual(row["planned_total_hours"], 100.0)
            self.assertEqual(row["planned_to_date_hours"], 100.0)
            self.assertEqual(row["total_actual_hours"], 24.0)
            self.assertEqual(row["actual_to_date_hours"], 24.0)
            self.assertEqual(row["schedule_variance_days"], -1)
            self.assertEqual(row["schedule_variance_hours"], -76.0)
            self.assertEqual(row["schedule_variance_pct"], -76.0)
            self.assertEqual(row["estimation_accuracy_pct"], 416.7)
            self.assertEqual(row["estimation_accuracy_status"], "outside_ideal")
            self.assertEqual(row["actual_complete_date"], "2026-06-01")
            self.assertEqual(row["actual_complete_source"], "max_last_logged_resolved_stable")
            self.assertEqual(row["headcount"], 3)
            self.assertEqual(row["schedule_variance_date_position"], "behind")
            self.assertEqual(row["schedule_variance_hours_position"], "behind")
            self.assertEqual({item["issue_key"] for item in row["stories"]}, {"O2-1-S1", "O2-1-S2", "O2-1-T1"})
            self.assertTrue(any(item["issue_key"] == "O2-1-T1" and item["issue_type"] == "Task" for item in row["stories"]))
            story_item = next(item for item in row["stories"] if item["issue_key"] == "O2-1-S1")
            task_item = next(item for item in row["stories"] if item["issue_key"] == "O2-1-T1")
            self.assertEqual({sub["issue_key"] for sub in story_item["subtasks"]}, {"O2-1-S1-T1", "O2-1-S1-B1"})
            self.assertEqual({sub["issue_key"] for sub in task_item["subtasks"]}, {"O2-1-T1-SUB1"})
            self.assertEqual(task_item["actual_hours"], 12.0)
            self.assertEqual(task_item["subtasks"][0]["actual_hours"], 12.0)
            self.assertEqual(len(story_item["subtasks"][0]["worklogs"]), 1)
            self.assertTrue(any(option["project_key"] == "O2" and option["project_name"] == "Omni" for option in payload["project_options"]))
            self.assertEqual([leaf["assignee"] for leaf in row["analytics"]["gantt"]["leaves"]], ["Alice", "Dana"])
            self.assertEqual([leaf["assignee"] for leaf in row["analytics"]["leave_summary"]], ["Alice", "Dana"])
            completion_labels = {item["bucket"]: item["label"] for item in row["analytics"]["completion_stats"]}
            self.assertEqual(completion_labels["after_due"], "Completed after due date")
            after_due = next(item for item in row["analytics"]["completion_stats"] if item["bucket"] == "after_due")
            self.assertEqual(after_due["items"][0]["summary"], "Fix drilldown")
            self.assertIn("within_original_estimate_count", row["analytics"]["estimate_quality"])
            self.assertIn("equal_story_estimate_pct", row["analytics"]["estimate_quality"])
            self.assertIn("Build table", {item["summary"] for item in row["analytics"]["estimate_quality"]["details"]})
            sv = row["analytics"]["schedule_variance"]
            self.assertEqual(sv["planned_to_date_hours"], 100.0)
            self.assertEqual(sv["actual_to_date_hours"], 24.0)
            self.assertEqual(sv["schedule_variance_hours"], -76.0)
            self.assertEqual(sv["trend_3_months"]["direction"], "improving")
            self.assertEqual([item["month"] for item in sv["trend_3_months"]["months"]], ["2026-05", "2026-06"])
            self.assertEqual(
                [(item["month"], item["planned_hours"]) for item in sv["trend_3_months"]["months"]],
                [("2026-05", 62.0), ("2026-06", 20.0)],
            )
            self.assertEqual(row["analytics"]["monthly_plan_basis"]["estimate_source"], "story_original_estimate")
            self.assertEqual(row["analytics"]["monthly_plan_basis"]["capacity_basis"], "assignee_capacity_after_leaves")
            cross_month = next(item for item in row["analytics"]["story_planning_distribution"] if item["issue_key"] == "O2-1-S2")
            self.assertEqual(cross_month["monthly_hours"], [{"month": "2026-05", "hours": 10.0}, {"month": "2026-06", "hours": 20.0}])
            unassigned = row["analytics"]["unassigned_planning_stories"]
            self.assertEqual({item["issue_key"] for item in unassigned}, {"O2-1-S1", "O2-1-S2", "O2-1-T1"})
            self.assertEqual(sum(item["original_estimate_hours"] for item in unassigned), 82.0)
            alice_sv = next(item for item in sv["assignees"] if item["assignee"] == "Alice")
            self.assertEqual(alice_sv["planned_to_date_hours"], 0.0)
            self.assertEqual(alice_sv["actual_to_date_hours"], 10.0)
            self.assertEqual(alice_sv["schedule_variance_hours"], 10.0)

            project_payload = build_epic_explorer_payload(db_path, [], "run-1", selected_projects={"CRM"})
            self.assertEqual([r["epic_key"] for r in project_payload["rows"]], ["CRM-2"])

    def test_api_page_catalog_and_html_sync_register_epic_explorer(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            _seed_epic_explorer_db(db_path)
            (root / "report_html").mkdir(parents=True, exist_ok=True)
            (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            (root / "epic_explorer_report.html").write_text("<html><body>epic explorer source</body></html>", encoding="utf-8")

            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            resp = client.get("/api/epic-explorer/summary?from_date=2026-05-01&to_date=2026-05-31")
            self.assertEqual(resp.status_code, 200)
            body = resp.get_json()
            self.assertTrue(body["ok"])
            self.assertEqual(body["ui_settings"]["capacity_basis"], "assignee_capacity_after_leaves")
            self.assertEqual(body["rows"][0]["epic_key"], "O2-1")
            self.assertEqual(body["rows"][0]["analytics"]["monthly_plan_basis"]["capacity_basis"], "assignee_capacity_after_leaves")

            saved = client.post("/api/epic-explorer/ui-settings", json={"capacity_basis": "standard_workdays"})
            self.assertEqual(saved.status_code, 200)
            self.assertEqual(saved.get_json()["settings"]["capacity_basis"], "standard_workdays")
            settings = client.get("/api/epic-explorer/ui-settings")
            self.assertEqual(settings.status_code, 200)
            self.assertEqual(settings.get_json()["settings"]["capacity_basis"], "standard_workdays")
            refreshed = client.get("/api/epic-explorer/summary?from_date=2026-05-01&to_date=2026-05-31")
            self.assertEqual(refreshed.get_json()["ui_settings"]["capacity_basis"], "standard_workdays")

            categories = client.get("/api/page-categories").get_json()
            self.assertTrue(any(item["page_key"] == "epic_explorer" and item["title"] == "Epic Explorer" for item in categories["page_catalog"]))

            synced = sync_report_html(root, "report_html")
            self.assertGreaterEqual(synced, 1)
            served = root / "report_html" / "epic_explorer_report.html"
            self.assertTrue(served.exists())
            self.assertIn("epic explorer source", served.read_text(encoding="utf-8"))

    def test_html_contains_drilldown_drawer_charts_and_csv_controls(self):
        html = (Path(__file__).resolve().parents[1] / "epic_explorer_report.html").read_text(encoding="utf-8")
        self.assertIn("Epic Explorer", html)
        self.assertIn('id="download-csv"', html)
        self.assertIn('id="project-picker"', html)
        self.assertIn('id="from-date"', html)
        self.assertIn('id="to-date"', html)
        self.assertIn('id="apply-date-filters"', html)
        self.assertIn('id="apply-project-filters"', html)
        self.assertIn("pendingProjects", html)
        self.assertIn("appliedProjects", html)
        self.assertIn('class="row-number-col"', html)
        self.assertIn('class="epic-name-col"', html)
        self.assertIn("max-height: calc(100vh - 235px)", html)
        self.assertIn("tr.epic-row.even td", html)
        self.assertIn("epic-title-inline", html)
        self.assertIn("epic-key-inline", html)
        self.assertIn("epicRow(row, rowNumber)", html)
        self.assertIn('data-toggle-epic', html)
        self.assertIn('data-toggle-story', html)
        self.assertIn('data-toggle-subtask', html)
        self.assertIn('data-open-epic', html)
        self.assertIn('id="drawer-resizer"', html)
        self.assertIn("Month Plan vs Actual", html)
        self.assertIn("Planned vs Actual Hours", html)
        self.assertIn("Planned vs Actual Delivery", html)
        self.assertIn("SV Date", html)
        self.assertIn("SV Hours", html)
        self.assertIn("Est. Accuracy", html)
        self.assertIn("scheduleHoursText(row)", html)
        self.assertIn("estimationAccuracyText(row)", html)
        self.assertIn('data-plan-mode="bar"', html)
        self.assertIn('data-plan-mode="line"', html)
        self.assertIn("planActualLineChart", html)
        self.assertIn('id="capacity-basis"', html)
        self.assertIn("/api/epic-explorer/ui-settings", html)
        self.assertIn("Story Plan Allocation", html)
        self.assertIn("storyPlanningAllocationTable", html)
        self.assertIn("unassignedPlanningStoriesAlert", html)
        self.assertIn("Planning alert:", html)
        self.assertIn("Capacity-based monthly allocation cannot use assignee capacity", html)
        self.assertIn("Daily Work And Leave Gantt", html)
        self.assertIn("ganttHeader", html)
        self.assertIn("gantt-header-cell month", html)
        self.assertIn("Schedule Variance KPIs", html)
        self.assertIn("SV By Date", html)
        self.assertIn("SV By Hours", html)
        self.assertIn("Planned vs Actual Hours", html)
        self.assertIn("Planned vs Actual Delivery", html)
        self.assertIn("Estimation Accuracy", html)
        self.assertIn("SV Per Assignee", html)
        self.assertIn("SV Trend Over Last 3 Months", html)
        self.assertIn("scheduleVarianceSection", html)
        self.assertIn("Resource Utilization", html)
        self.assertIn("resourceUtilizationVisual", html)
        self.assertIn("Support Resources", html)
        self.assertIn("Estimate Quality", html)
        self.assertIn("estimateQualityScorecards", html)
        self.assertIn("data-completion-bucket", html)
        self.assertIn("bucketLabel", html)
        self.assertIn("Leaves By Resource", html)
        self.assertIn("downloadCsv", html)
        self.assertLess(html.index("</style>"), html.index("<body>"))

    def test_html_contains_executive_summary_mini_dashboard_controls(self):
        html = (Path(__file__).resolve().parents[1] / "epic_explorer_report.html").read_text(encoding="utf-8")
        self.assertIn("Executive Summary", html)
        self.assertIn('id="exec-summary"', html)
        self.assertIn('id="epic-picker"', html)
        self.assertIn('id="epic-picker-toggle"', html)
        self.assertIn('id="epic-picker-menu"', html)
        self.assertIn('id="epic-picker-search"', html)
        self.assertIn('id="epic-picker-clear"', html)
        self.assertIn('id="epic-picker-add"', html)
        self.assertIn('id="exec-trend-cards"', html)
        self.assertIn('id="exec-table-wrap"', html)
        self.assertIn('id="exec-leadership-chart"', html)
        self.assertIn("data-exec-epic", html)
        self.assertIn("data-exec-toggle", html)
        self.assertIn("data-exec-remove", html)
        self.assertIn("EXEC_SUMMARY_STORAGE_KEY", html)
        self.assertIn("epicExplorerExecSummaryEpics", html)
        self.assertIn("execWeeklyTrend", html)
        self.assertIn("execWeeklyTrendTable", html)
        self.assertIn("execWeeklyTrendLineChart", html)
        self.assertIn("Story-level estimate proration line chart", html)
        self.assertIn("weekly-trend-table", html)
        self.assertIn('colspan="12"', html)
        self.assertIn("storyPlannedToDate", html)
        self.assertIn("storyActualToDate", html)
        self.assertIn("execMonthlyAverages", html)
        self.assertIn("Month Over Month Average Schedule Variance Trend", html)
        self.assertIn("execTrendChartHtml", html)
        self.assertIn("svTrendLineChart", html)
        self.assertIn("syncExecSummarySelection", html)
        self.assertIn("renderEpicPickerOptions", html)
        self.assertIn('id="epic-picker-options"', html)
        self.assertIn("<th>Status</th><th>Budget</th>", html)
        self.assertIn("statusPill(row.epic_status)", html)
        self.assertIn("function execBudgetHours(row)", html)
        self.assertIn("jiraOriginalEstimate > 0 ? jiraOriginalEstimate : n(row.planned_total_hours)", html)
        self.assertNotIn("tkBudget > 0 ? tkBudget : n(row.planned_total_hours)", html)
        self.assertIn("execBudgetHours(row) > 0 ? hours(execBudgetHours(row)) : \"-\"", html)
        self.assertIn("function execBudgetVarianceText(row)", html)
        self.assertIn("const variance = n(row.total_actual_hours) - budget", html)
        self.assertIn("const variancePct = (variance / budget) * 100", html)
        self.assertIn("<td>${execBudgetVarianceText(row)}</td>", html)
        self.assertIn("Portfolio Budget vs Actual Hours", html)
        self.assertIn("execLeadershipChartHtml", html)
        self.assertIn("Math.max(execBudgetHours(r), n(r.total_actual_hours))", html)
        self.assertIn("const budget = execBudgetHours(r)", html)
        self.assertIn("budgetOverlay", html)
        self.assertIn("actualOverlay", html)
        self.assertIn(".bar.overlay", html)
        self.assertIn("execCompleteTooltip", html)
        self.assertIn("exec-complete-cell", html)
        self.assertIn("renderExecSummary", html)
        self.assertIn('No epics pinned yet', html)


if __name__ == "__main__":
    unittest.main()
