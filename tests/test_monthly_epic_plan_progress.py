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


def _planner_row_with_story_plan(
    epic_key: str,
    epic_name: str,
    start_date: str,
    due_date: str,
    *,
    epic_man_days: float = 10,
    story_key: str,
    story_man_days: float,
) -> dict:
    row = _planner_row(epic_key, epic_name, start_date, due_date, man_days=epic_man_days)
    row["plans"]["research_urs_plan"] = {
        "jira_url": f"https://jira.example/browse/{story_key}",
        "tk_budgeted_man_days": story_man_days,
        "man_days": story_man_days,
        "start_date": start_date,
        "due_date": due_date,
    }
    return row


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


def _add_estimate_rollup_tree(
    db_path: Path,
    epic_key: str,
    *,
    epic_original: float,
    story_rows: list[tuple[str, float, list[tuple[str, float, float]]]],
    status: str = "In Progress",
    start_date: str = "2026-03-01",
    due_date: str = "2026-03-31",
) -> None:
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            INSERT INTO canonical_issues(
                run_id, issue_key, project_key, issue_type, summary, status,
                start_date, due_date, original_estimate_hours, parent_issue_key, story_key, epic_key
            )
            VALUES ('run-1', ?, 'O2', 'Epic', ?, ?, ?, ?, ?, '', '', ?)
            """,
            (epic_key, f"{epic_key} Epic", status, start_date, due_date, epic_original, epic_key),
        )
        worklog_rows = []
        for story_suffix, story_estimate, subtask_specs in story_rows:
            story_key = f"{epic_key}-{story_suffix}"
            conn.execute(
                """
                INSERT INTO canonical_issues(
                    run_id, issue_key, project_key, issue_type, summary, status,
                    start_date, due_date, original_estimate_hours, parent_issue_key, story_key, epic_key
                )
                VALUES ('run-1', ?, 'O2', 'Story', ?, 'In Progress', ?, ?, ?, ?, ?, ?)
                """,
                (story_key, f"{story_key} Story", start_date, due_date, story_estimate, epic_key, story_key, epic_key),
            )
            for idx, subtask_spec in enumerate(subtask_specs, start=1):
                if len(subtask_spec) == 4:
                    sub_suffix, subtask_estimate, logged_hours, issue_type = subtask_spec
                else:
                    sub_suffix, subtask_estimate, logged_hours = subtask_spec
                    issue_type = "Sub-task"
                subtask_key = f"{story_key}-{sub_suffix}"
                conn.execute(
                    """
                    INSERT INTO canonical_issues(
                        run_id, issue_key, project_key, issue_type, summary, status,
                        start_date, due_date, original_estimate_hours, parent_issue_key, story_key, epic_key
                    )
                    VALUES ('run-1', ?, 'O2', ?, ?, 'In Progress', ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        subtask_key,
                        issue_type,
                        f"{subtask_key} Subtask",
                        start_date,
                        due_date,
                        subtask_estimate,
                        story_key,
                        story_key,
                        epic_key,
                    ),
                )
                if logged_hours:
                    worklog_rows.append((f"{subtask_key}-{idx}", subtask_key, "2026-03-10", logged_hours))
        conn.executemany(
            """
            INSERT INTO canonical_worklogs(run_id, worklog_id, issue_key, project_key, started_date, hours_logged)
            VALUES ('run-1', ?, ?, 'O2', ?, ?)
            """,
            worklog_rows,
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

    def test_estimate_rollup_uses_same_in_scope_epics(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "assignee_hours_capacity.db"
            _create_canonical_tables(db_path)
            _add_estimate_rollup_tree(
                db_path,
                "O2-IN",
                epic_original=120,
                story_rows=[
                    ("S1", 40, [("T1", 10, 30), ("T2", 15, 20)]),
                    ("S2", 16, [("T1", 8, 4)]),
                ],
            )
            _add_estimate_rollup_tree(
                db_path,
                "O2-OUT",
                epic_original=999,
                story_rows=[("S1", 999, [("T1", 999, 999)])],
                start_date="2026-01-01",
                due_date="2026-01-15",
            )
            _add_estimate_rollup_tree(
                db_path,
                "O2-BF",
                epic_original=777,
                story_rows=[("S1", 777, [("T1", 777, 777)])],
                start_date="2026-02-01",
                due_date="2026-02-25",
            )

            payload = build_monthly_epic_plan_payload(
                db_path,
                "2026-03",
                [
                    _planner_row("O2-IN", "Included Epic", "2026-03-01", "2026-03-31"),
                    _planner_row("O2-OUT", "Excluded Epic", "2026-01-01", "2026-01-15"),
                    _planner_row("O2-BF", "Brought Forward Epic", "2026-02-01", "2026-02-25"),
                ],
                "run-1",
                selected_projects={"O2"},
            )

            rows_by_key = {row["epic_key"]: row for row in payload["rows"]}
            self.assertEqual(set(rows_by_key), {"O2-BF", "O2-IN"})
            self.assertTrue(rows_by_key["O2-BF"]["brought_forward"])
            rollup = payload["estimate_rollup"]
            self.assertEqual(rollup["epic_count"], 1)
            self.assertEqual(rollup["story_count"], 2)
            self.assertEqual(rollup["subtask_count"], 3)
            self.assertEqual(rollup["epic_original_estimate_hours"], 120.0)
            self.assertEqual(rollup["story_original_estimate_hours"], 56.0)
            self.assertEqual(rollup["subtask_original_estimate_hours"], 33.0)
            self.assertEqual(rollup["subtask_logged_hours"], 54.0)
            self.assertEqual(rollup["story_subtask_logged_over_parent_estimate_hours"], 10.0)
            self.assertEqual(rollup["story_subtask_logged_over_parent_estimate_pct"], 25.0)
            self.assertEqual(rollup["overrun_story_count"], 1)
            by_epic = rollup["by_epic"]["O2-IN"]
            self.assertIn("details_by_metric", by_epic)
            self.assertEqual(len(by_epic["details_by_metric"]["epic_original"]), 1)
            self.assertEqual(len(by_epic["details_by_metric"]["story_original"]), 2)
            self.assertEqual(len(by_epic["details_by_metric"]["subtask_original"]), 3)
            self.assertEqual(len(by_epic["details_by_metric"]["subtask_logged"]), 3)
            self.assertEqual(len(by_epic["details_by_metric"]["overrun"]), 1)
            subtask_detail = by_epic["details_by_metric"]["subtask_logged"][0]
            self.assertEqual(subtask_detail["item_type"], "Sub-task")
            self.assertIn("parents", subtask_detail)
            self.assertEqual([parent["item_type"] for parent in subtask_detail["parents"]], ["Story", "Epic"])

    def test_estimate_rollup_detail_rows_include_tk_planned_hours_for_matching_epics_and_stories(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "assignee_hours_capacity.db"
            _create_canonical_tables(db_path)
            _add_estimate_rollup_tree(
                db_path,
                "O2-IN",
                epic_original=120,
                story_rows=[
                    ("S1", 40, [("T1", 10, 0)]),
                    ("S2", 16, [("T1", 8, 0)]),
                ],
            )

            payload = build_monthly_epic_plan_payload(
                db_path,
                "2026-03",
                [
                    _planner_row_with_story_plan(
                        "O2-IN",
                        "Included Epic",
                        "2026-03-01",
                        "2026-03-31",
                        epic_man_days=11,
                        story_key="O2-IN-S1",
                        story_man_days=3,
                    )
                ],
                "run-1",
                selected_projects={"O2"},
            )

            by_epic = payload["estimate_rollup"]["by_epic"]["O2-IN"]
            epic_detail = by_epic["details_by_metric"]["epic_original"][0]
            self.assertEqual(epic_detail["tk_planned_hours"], 88.0)
            self.assertEqual(epic_detail["tk_planned_days"], 11.0)
            story_details = {item["issue_key"]: item for item in by_epic["details_by_metric"]["story_original"]}
            self.assertEqual(story_details["O2-IN-S1"]["tk_planned_hours"], 24.0)
            self.assertEqual(story_details["O2-IN-S1"]["tk_planned_days"], 3.0)
            self.assertIsNone(story_details["O2-IN-S2"]["tk_planned_hours"])

    def test_estimate_rollup_overrun_detail_rows_split_bug_and_non_bug_logged_hours(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "assignee_hours_capacity.db"
            _create_canonical_tables(db_path)
            _add_estimate_rollup_tree(
                db_path,
                "O2-IN",
                epic_original=120,
                story_rows=[
                    (
                        "S1",
                        40,
                        [
                            ("T1", 10, 30, "Sub-task"),
                            ("B1", 5, 18, "Bug Subtask"),
                        ],
                    ),
                ],
            )

            payload = build_monthly_epic_plan_payload(
                db_path,
                "2026-03",
                [_planner_row("O2-IN", "Included Epic", "2026-03-01", "2026-03-31")],
                "run-1",
                selected_projects={"O2"},
            )

            overrun_detail = payload["estimate_rollup"]["by_epic"]["O2-IN"]["details_by_metric"]["overrun"][0]
            self.assertEqual(overrun_detail["issue_key"], "O2-IN-S1")
            self.assertEqual(overrun_detail["original_estimate_hours"], 40.0)
            self.assertEqual(overrun_detail["logged_hours"], 48.0)
            self.assertEqual(overrun_detail["overrun_hours"], 8.0)
            self.assertEqual(overrun_detail["non_bug_logged_hours"], 30.0)
            self.assertEqual(overrun_detail["non_bug_overrun_hours"], 0.0)
            self.assertEqual(overrun_detail["bug_logged_hours"], 18.0)

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
        self.assertIn("Estimate hierarchy stats", html)
        self.assertIn('id="estimate-rollup-panel"', html)
        self.assertIn('id="estimate-rollup-chart"', html)
        self.assertIn("estimate-bar-row", html)
        self.assertIn("estimate-bar-fill", html)
        self.assertIn("Month Plan", html)
        self.assertIn("Epic Estimate", html)
        self.assertIn("Story Estimate", html)
        self.assertIn("Subtask Estimate", html)
        self.assertIn("Subtask Logged", html)
        self.assertIn("Story Overrun", html)
        self.assertIn("epics planned this month", html)
        self.assertIn("Brought-forward overdue epics are shown in Executive summary separately", html)
        self.assertIn('id="est-month-plan"', html)
        self.assertIn('id="est-month-plan-bar"', html)
        self.assertIn("matches Executive summary planned hours", html)
        self.assertIn('id="est-epic-original"', html)
        self.assertIn('id="est-epic-original-bar"', html)
        self.assertIn('id="est-story-original"', html)
        self.assertIn('id="est-story-original-bar"', html)
        self.assertIn('id="est-subtask-original"', html)
        self.assertIn('id="est-subtask-original-bar"', html)
        self.assertIn('id="est-subtask-logged"', html)
        self.assertIn('id="est-subtask-logged-bar"', html)
        self.assertIn('id="est-overrun-value"', html)
        self.assertIn('id="est-overrun-bar"', html)
        self.assertIn('id="estimate-detail-overlay"', html)
        self.assertIn('id="estimate-detail-resize"', html)
        self.assertIn('id="estimate-detail-controls"', html)
        self.assertIn('id="estimate-detail-include-bugs"', html)
        self.assertIn("Include bug subtasks", html)
        self.assertIn("function openEstimateDetail", html)
        self.assertIn("function renderEstimateDetailDrawer", html)
        self.assertIn("function estimateDetailRows", html)
        self.assertIn("function estimateDetailDisplayedRows", html)
        self.assertIn("estimateDetailIncludeBugSubtasks", html)
        self.assertIn("non_bug_logged_hours", html)
        self.assertIn("non_bug_overrun_hours", html)
        self.assertIn("Main Story Overrun bars stay unchanged", html)
        self.assertIn("data-estimate-metric=\"subtask_logged\"", html)
        self.assertIn("estimate-detail-table", html)
        self.assertIn("TK planned", html)
        self.assertIn('detailHoursText(row, "tk_planned")', html)
        self.assertIn("TK planned: ", html)
        self.assertIn("estimate-parent-link", html)
        self.assertIn("Drag to resize", html)
        self.assertIn("function renderEstimateRollup", html)
        self.assertIn("aggregateEstimateRollupFromRows", html)
        self.assertIn("const setBar = (el, value)", html)
        self.assertIn("estimate_rollup", html)
        row_to_html = re.search(r"function rowToHtml\(row\) \{(?P<body>.*?)\n    // Sort order:", html, re.S)
        self.assertIsNotNone(row_to_html)
        self.assertIn('const monthYm = els.month ? String(els.month.value || "").trim() : "";', row_to_html.group("body"))

    def test_resource_planning_panel_present_in_html(self):
        html_path = Path(__file__).resolve().parents[1] / "monthly_epic_plan_progress_report.html"
        html = html_path.read_text(encoding="utf-8")

        self.assertIn('id="res-summary-panel"', html)
        self.assertIn('id="res-total-headcount"', html)
        self.assertIn('id="res-total-capacity"', html)
        self.assertIn('id="res-dev-headcount"', html)
        self.assertIn('id="res-dev-capacity"', html)
        self.assertIn('id="res-dev-leaves"', html)
        self.assertIn('id="res-dev-avail"', html)
        self.assertIn('id="res-support-group"', html)
        self.assertIn('id="res-support-headcount"', html)
        self.assertIn('id="res-support-avail"', html)
        self.assertIn("function renderResourceSummary", html)
        self.assertIn("_resSummaryState", html)
        self.assertIn("Resource Planning", html)
        self.assertIn("Total resources", html)
        self.assertIn("Dev Resources", html)
        self.assertIn("Support Resources", html)

    def test_process_team_auto_exclusion_present_in_html(self):
        html_path = Path(__file__).resolve().parents[1] / "monthly_epic_plan_progress_report.html"
        html = html_path.read_text(encoding="utf-8")

        self.assertIn("Auto-exclude Process team when no explicit filter is active", html)
        self.assertIn(".includes(\"process\")", html)
        self.assertIn("_didAutoExclude", html)
        self.assertIn("setTimeout(loadSummary, 0)", html)


if __name__ == "__main__":
    unittest.main()
