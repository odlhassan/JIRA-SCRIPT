from __future__ import annotations

import sqlite3
import tempfile
import unittest
import re
from pathlib import Path

from datetime import date

from monthly_epic_plan_progress_service import (
    _nested_aligned_leave_by_assignee,
    build_monthly_epic_plan_payload,
)
from report_server import _init_epics_management_db, create_report_server_app
from report_server import sync_report_html


def _planner_row(
    epic_key: str,
    epic_name: str,
    start_date: str,
    due_date: str,
    man_days: float = 10,
    status: str = "Yet to start",
) -> dict:
    return {
        "id": epic_key,
        "project_key": "O2",
        "project_name": "OmniConnect",
        "product_category": "Core",
        "component": "",
        "epic_key": epic_key,
        "epic_name": epic_name,
        "delivery_status": status,
        "jira_url": f"https://jira.example/browse/{epic_key}",
        "plans": {
            "epic_plan": {
                "man_days": man_days,
                "tk_budgeted_start_date": start_date,
                "tk_budgeted_due_date": due_date,
                "start_date": start_date,
                "due_date": due_date,
            }
        },
    }


def _create_canonical_tables(db_path: Path, run_id: str = "run-1") -> None:
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
            "INSERT INTO canonical_refresh_state(id, active_run_id, last_success_run_id, updated_at_utc) VALUES (1, ?, ?, ?)",
            (run_id, run_id, "2026-03-31T00:00:00Z"),
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
    finally:
        conn.close()


def _add_epic_tree(db_path: Path, epic_key: str, status: str, worklogs: list[tuple[str, float]]) -> None:
    story_key = f"{epic_key}-S1"
    subtask_key = f"{epic_key}-T1"
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            INSERT INTO canonical_issues(run_id, issue_key, project_key, issue_type, summary, status, parent_issue_key, story_key, epic_key)
            VALUES ('run-1', ?, 'O2', 'Epic', ?, ?, '', '', ?)
            """,
            (epic_key, f"{epic_key} Epic", status, epic_key),
        )
        conn.execute(
            """
            INSERT INTO canonical_issues(
                run_id, issue_key, project_key, issue_type, summary, status,
                start_date, due_date, original_estimate_hours,
                parent_issue_key, story_key, epic_key
            )
            VALUES ('run-1', ?, 'O2', 'Story', ?, 'In Progress', '2026-03-01', '2026-03-31', 0, ?, ?, ?)
            """,
            (story_key, f"{epic_key} Story", epic_key, story_key, epic_key),
        )
        conn.execute(
            """
            INSERT INTO canonical_issues(
                run_id, issue_key, project_key, issue_type, summary, status,
                start_date, due_date, original_estimate_hours,
                parent_issue_key, story_key, epic_key
            )
            VALUES ('run-1', ?, 'O2', 'Sub-task', ?, 'In Progress', '2026-03-01', '2026-03-31', 80, ?, ?, ?)
            """,
            (subtask_key, f"{epic_key} Task", story_key, story_key, epic_key),
        )
        conn.executemany(
            """
            INSERT INTO canonical_worklogs(run_id, worklog_id, issue_key, project_key, started_date, hours_logged)
            VALUES ('run-1', ?, ?, 'O2', ?, ?)
            """,
            [(f"{subtask_key}-{index}", subtask_key, started_date, hours) for index, (started_date, hours) in enumerate(worklogs, start=1)],
        )
        conn.commit()
    finally:
        conn.close()


def _insert_planner_epic(db_path: Path, row: dict) -> None:
    _init_epics_management_db(db_path)
    epic_plan = row["plans"]["epic_plan"]
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            INSERT INTO epics_management(
                id, epic_key, project_key, project_name, product_category, component,
                epic_name, delivery_status, jira_url, epic_plan_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                row["id"],
                row["epic_key"],
                row["project_key"],
                row["project_name"],
                row["product_category"],
                row["component"],
                row["epic_name"],
                row["delivery_status"],
                row["jira_url"],
                (
                    "{"
                    f'"man_days": {epic_plan["man_days"]}, '
                    f'"tk_budgeted_start_date": "{epic_plan["tk_budgeted_start_date"]}", '
                    f'"tk_budgeted_due_date": "{epic_plan["tk_budgeted_due_date"]}", '
                    f'"start_date": "{epic_plan["start_date"]}", '
                    f'"due_date": "{epic_plan["due_date"]}"'
                    "}"
                ),
            ),
        )
        conn.commit()
    finally:
        conn.close()


def _add_performance_team(db_path: Path, team_name: str, team_leader: str, assignees_json: str) -> None:
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS performance_teams (
                team_name TEXT PRIMARY KEY,
                team_leader TEXT NOT NULL DEFAULT '',
                assignees_json TEXT NOT NULL DEFAULT '[]',
                updated_at TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.execute(
            "INSERT INTO performance_teams(team_name, team_leader, assignees_json, updated_at) VALUES (?, ?, ?, ?)",
            (team_name, team_leader, assignees_json, "2026-03-01T00:00:00Z"),
        )
        conn.commit()
    finally:
        conn.close()


class MonthlyEpicPlanProgressTests(unittest.TestCase):
    def test_nested_aligned_leave_prefers_distributed_buckets(self):
        month_start = date(2026, 4, 1)
        month_end = date(2026, 4, 30)
        snapshot = {
            "distributed_subtasks": [
                {
                    "assignee": "Alpha",
                    "planned_date_for_bucket": "2026-04-10",
                    "original_estimate_hours": 8,
                    "total_worklog_hours": 0,
                }
            ],
            "daily": [
                {
                    "assignee": "Alpha",
                    "period_day": "2026-04-10",
                    "planned_taken_hours": 1,
                    "planned_not_taken_hours": 1,
                    "unplanned_taken_hours": 999,
                    "unknown_taken_hours": 0,
                }
            ],
        }
        _, leave_by, src = _nested_aligned_leave_by_assignee(snapshot, month_start, month_end)
        self.assertEqual(src, "distributed_subtasks")
        self.assertAlmostEqual(leave_by.get("alpha", 0), 8.0, places=2)

    def test_nested_aligned_daily_excludes_unplanned_and_unknown_buckets(self):
        month_start = date(2026, 4, 1)
        month_end = date(2026, 4, 30)
        snapshot = {
            "distributed_subtasks": [],
            "daily": [
                {
                    "assignee": "Beta",
                    "period_day": "2026-04-12",
                    "planned_taken_hours": 2,
                    "planned_not_taken_hours": 3,
                    "unplanned_taken_hours": 44,
                    "unknown_taken_hours": 11,
                }
            ],
        }
        _, leave_by, src = _nested_aligned_leave_by_assignee(snapshot, month_start, month_end)
        self.assertEqual(src, "daily_planned_buckets")
        self.assertAlmostEqual(leave_by.get("beta", 0), 5.0, places=2)

    def test_service_includes_epic_when_start_or_due_in_month_and_month_worklogs(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "assignee_hours_capacity.db"
            _create_canonical_tables(db_path)
            _add_epic_tree(db_path, "O2-SPAN", "In Progress", [("2026-02-20", 4), ("2026-03-05", 6), ("2026-03-20", 2), ("2026-04-01", 7)])

            payload = build_monthly_epic_plan_payload(
                db_path,
                "2026-03",
                [_planner_row("O2-SPAN", "Spanning Epic", "2026-03-01", "2026-04-10")],
                "run-1",
                selected_projects={"O2"},
            )

            self.assertEqual(payload["totals"]["epic_count"], 1)
            row = payload["rows"][0]
            self.assertEqual(row["epic_key"], "O2-SPAN")
            self.assertEqual(row["planned_hours"], 80)
            self.assertEqual(row["actual_hours"], 8)
            self.assertFalse(row["start_slip"])
            self.assertFalse(row["end_slip"])

    def test_service_excludes_epic_when_schedule_spans_month_but_no_endpoint_in_month(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "assignee_hours_capacity.db"
            _create_canonical_tables(db_path)
            _add_epic_tree(db_path, "O2-MID", "In Progress", [("2026-03-10", 2)])
            payload = build_monthly_epic_plan_payload(
                db_path,
                "2026-03",
                [_planner_row("O2-MID", "Spans March", "2026-02-10", "2026-05-10")],
                "run-1",
                selected_projects={"O2"},
            )
            self.assertEqual(payload["totals"]["epic_count"], 0)
            self.assertEqual(len(payload["rows"]), 0)

    def test_service_marks_start_and_end_slips_for_month(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "assignee_hours_capacity.db"
            _create_canonical_tables(db_path)
            _add_epic_tree(db_path, "O2-START", "To Do", [])
            _add_epic_tree(db_path, "O2-END", "In Progress", [("2026-03-04", 3)])
            _add_epic_tree(db_path, "O2-DONE", "Resolved!", [("2026-03-04", 3)])
            _add_epic_tree(db_path, "O2-OLD", "In Progress", [])
            _add_performance_team(db_path, "Delivery Team", "Alice", '["Alice", "Bob"]')

            payload = build_monthly_epic_plan_payload(
                db_path,
                "2026-03",
                [
                    _planner_row("O2-START", "Start Slip", "2026-03-05", "2026-04-15"),
                    _planner_row("O2-END", "End Slip", "2026-02-01", "2026-03-15"),
                    _planner_row("O2-DONE", "Done", "2026-02-01", "2026-03-15"),
                    _planner_row("O2-OLD", "Previous Month Pending", "2026-01-10", "2026-02-25"),
                ],
                "run-1",
                selected_projects={"O2"},
            )

            rows = {row["epic_key"]: row for row in payload["rows"]}
            self.assertTrue(rows["O2-START"]["start_slip"])
            self.assertFalse(rows["O2-START"]["end_slip"])
            self.assertTrue(rows["O2-END"]["end_slip"])
            self.assertFalse(rows["O2-END"]["start_slip"])
            self.assertFalse(rows["O2-DONE"]["end_slip"])
            self.assertTrue(rows["O2-OLD"]["brought_forward"])
            self.assertFalse(rows["O2-OLD"]["start_slip"])
            self.assertFalse(rows["O2-OLD"]["end_slip"])
            # O2-OLD is "In Progress" (not resolved) and brought_forward → carried forward
            self.assertTrue(rows["O2-OLD"]["carried_forward"])
            self.assertEqual(payload["totals"]["brought_forward_count"], 1)
            self.assertEqual(payload["totals"]["brought_forward_planned_hours"], 80.0)
            # carried_forward_planned_hours: O2-START(80) + O2-END(80) + O2-OLD(80) = 240
            self.assertEqual(payload["totals"]["carried_forward_planned_hours"], 240.0)
            self.assertEqual(payload["totals"]["start_slip_count"], 1)
            self.assertEqual(payload["totals"]["end_slip_count"], 1)
            self.assertEqual(payload["totals"]["carried_forward_count"], 3)
            self.assertIn("workforce", payload)
            self.assertIn("assignee_options", payload["workforce"])
            self.assertIn("employee_options", payload["workforce"])
            self.assertIn("employee_tree", payload["workforce"])
            etree = payload["workforce"]["employee_tree"]
            self.assertIn("teams", etree)
            self.assertIn("ungrouped", etree)
            team = next(t for t in etree["teams"] if t["team_name"] == "Delivery Team")
            self.assertEqual([m["name"] for m in team["members"]], ["Alice", "Bob"])
            by_proj = {p["project_key"]: p for p in payload["by_project"]}
            self.assertIn("O2", by_proj)
            self.assertEqual(by_proj["O2"]["epic_count"], 4)
            self.assertEqual(by_proj["O2"]["brought_forward_count"], 1)
            self.assertEqual(by_proj["O2"]["brought_forward_planned_hours"], 80.0)
            self.assertEqual(by_proj["O2"]["carried_forward_count"], 3)
            self.assertEqual(by_proj["O2"]["carried_forward_planned_hours"], 240.0)

    def test_delivery_status_aligns_when_jira_in_progress_but_planner_yet_to_start(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "assignee_hours_capacity.db"
            _create_canonical_tables(db_path)
            _add_epic_tree(db_path, "O2-WIP", "In Progress", [])
            payload = build_monthly_epic_plan_payload(
                db_path,
                "2026-04",
                [_planner_row("O2-WIP", "WIP Epic", "2026-04-01", "2026-04-27")],
                "run-1",
                selected_projects={"O2"},
            )
            self.assertEqual(len(payload["rows"]), 1)
            self.assertEqual(payload["rows"][0]["delivery_status"], "On-track")

    def test_delivery_status_keeps_yet_to_start_when_jira_to_do(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "assignee_hours_capacity.db"
            _create_canonical_tables(db_path)
            _add_epic_tree(db_path, "O2-TODO", "To Do", [])
            payload = build_monthly_epic_plan_payload(
                db_path,
                "2026-04",
                [_planner_row("O2-TODO", "Todo Epic", "2026-04-04", "2026-04-06")],
                "run-1",
                selected_projects={"O2"},
            )
            self.assertEqual(payload["rows"][0]["delivery_status"], "Yet to start")

    def test_delivery_status_keeps_planner_late_when_jira_also_active(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "assignee_hours_capacity.db"
            _create_canonical_tables(db_path)
            _add_epic_tree(db_path, "O2-LATE", "In Progress", [])
            payload = build_monthly_epic_plan_payload(
                db_path,
                "2026-04",
                [_planner_row("O2-LATE", "Late Epic", "2026-04-01", "2026-04-27", status="Late")],
                "run-1",
                selected_projects={"O2"},
            )
            self.assertEqual(payload["rows"][0]["delivery_status"], "Late")

    def test_api_summary_loads_epics_management_rows(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir(parents=True, exist_ok=True)
            (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            (root / "report_html" / "shared-nav.js").write_text("console.log('nav');", encoding="utf-8")
            (root / "report_html" / "shared-nav.css").write_text("body{}", encoding="utf-8")
            (root / "monthly_epic_plan_progress_report.html").write_text("<html><body>monthly</body></html>", encoding="utf-8")
            db_path = root / "assignee_hours_capacity.db"
            _create_canonical_tables(db_path)
            _add_epic_tree(db_path, "O2-API", "In Progress", [("2026-03-08", 5)])
            _insert_planner_epic(db_path, _planner_row("O2-API", "API Epic", "2026-03-01", "2026-03-31", 4, "Late"))

            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            resp = client.get("/api/monthly-epic-plan-progress/summary?month=2026-03&projects=O2")
            self.assertEqual(resp.status_code, 200)
            body = resp.get_json()
            self.assertTrue(body["ok"])
            self.assertEqual(body["rows"][0]["epic_key"], "O2-API")
            self.assertEqual(body["rows"][0]["actual_hours"], 5)

            html_resp = client.get("/monthly_epic_plan_progress_report.html")
            self.assertEqual(html_resp.status_code, 200)
            self.assertIn("monthly", html_resp.get_data(as_text=True))

    def test_report_html_sync_copies_monthly_source_report(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            source = root / "monthly_epic_plan_progress_report.html"
            source.write_text(
                '<html><script>function rowToHtml(row) {'
                'const monthYm = els.month ? String(els.month.value || "").trim() : "";'
                'return monthYm;}</script></html>',
                encoding="utf-8",
            )

            synced = sync_report_html(root, "report_html")

            served = root / "report_html" / "monthly_epic_plan_progress_report.html"
            self.assertGreaterEqual(synced, 1)
            self.assertTrue(served.exists())
            self.assertIn("function rowToHtml(row)", served.read_text(encoding="utf-8"))
            self.assertIn('const monthYm = els.month ? String(els.month.value || "").trim() : "";', served.read_text(encoding="utf-8"))

    def test_report_ui_has_project_filter_and_icon_only_epic_opener(self):
        html_path = Path(__file__).resolve().parents[1] / "monthly_epic_plan_progress_report.html"
        html = html_path.read_text(encoding="utf-8")

        self.assertIn('id="project-dropdown"', html)
        self.assertIn('id="project-dropdown-toggle"', html)
        self.assertIn('id="project-dropdown-list"', html)
        self.assertNotIn('id="project-select" multiple', html)
        self.assertIn("open_in_new", html)
        self.assertIn('aria-label="Open epic in Jira"', html)
        self.assertNotIn("epicKey + '</a>", html)
        self.assertIn("projectMetaByKey", html)
        self.assertIn("project-chip", html)
        self.assertIn("--project-color", html)
        self.assertIn("function projectDisplayName", html)
        self.assertIn("projectName = projectDisplayName", html)
        self.assertNotIn("esc(projectKey || \"-\") + '</span>'", html)
        self.assertIn("function formatDisplayDate", html)
        self.assertIn('"-" + months[month - 1] + "-"', html)
        self.assertIn("function statusPill", html)
        self.assertIn(".pill.info", html)
        self.assertIn('normalized === "yet to start"', html)
        self.assertIn('normalized === "on track"', html)
        self.assertIn("start < todayIso()", html)
        self.assertIn('id="kpi-brought-forward-count"', html)
        self.assertIn('id="kpi-brought-forward-planned"', html)
        self.assertIn('id="kpi-carried-forward-count"', html)
        self.assertIn('id="kpi-carried-forward-planned"', html)
        self.assertIn("Epics in scope this month", html)
        self.assertIn("Planned Hours", html)
        self.assertNotIn('id="kpi-start-slip"', html)
        self.assertIn("proj-card-filter", html)
        self.assertIn("filteredRowsForTable", html)
        self.assertIn("resetTableProjectFilter", html)
        self.assertIn("function renderProjectCards", html)
        self.assertIn("totals.brought_forward_planned_days", html)
        self.assertIn("totals.brought_forward_count", html)
        self.assertIn("totals.carried_forward_planned_hours", html)
        self.assertIn("totals.carried_forward_count", html)
        self.assertIn("workforce", html)
        self.assertIn('id="capacity-profile-select"', html)
        self.assertIn("renderCapacityProfileDropdown", html)
        self.assertIn("capacity_profile", html)
        self.assertIn('id="employee-dropdown-toggle"', html)
        self.assertIn('id="employee-dropdown-search"', html)
        self.assertIn('id="employee-dropdown-empty"', html)
        self.assertIn('id="assignee-select-all"', html)
        self.assertIn("function renderEmployeeDropdown", html)
        self.assertIn("function applyEmployeeDropdownFilter", html)
        self.assertIn('const bid = "emp-cb-leaf-" + (rowSeq++);', html)
        self.assertIn('const tid = "emp-team-cb-" + (teamSeq++);', html)
        self.assertIn("emp-dd-team-cb", html)
        self.assertIn("emp-dd-meta", html)
        self.assertIn("Brought Forward", html)
        self.assertIn("Brought forward", html)
        row_to_html = re.search(r"function rowToHtml\(row\) \{(?P<body>.*?)\n    // Sort order:", html, re.S)
        self.assertIsNotNone(row_to_html)
        self.assertIn('const monthYm = els.month ? String(els.month.value || "").trim() : "";', row_to_html.group("body"))


if __name__ == "__main__":
    unittest.main()
