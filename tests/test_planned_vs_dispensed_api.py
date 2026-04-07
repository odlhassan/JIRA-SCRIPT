from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

import report_server
from report_server import create_report_server_app


def _build_app(root: Path):
    (root / "report_html").mkdir(parents=True, exist_ok=True)
    (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
    wb = Workbook()
    ws = wb.active
    ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
    ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
    wb.save(root / "assignee_hours_report.xlsx")
    return create_report_server_app(base_dir=root, folder_raw="report_html")


def _seed_epic_planner_data(root: Path, epic_key: str, man_days: float | None, story_key: str):
    db_path = root / "assignee_hours_capacity.db"
    report_server._init_epics_management_db(db_path)
    conn = sqlite3.connect(db_path)
    try:
        plan_json = "{}"
        if man_days is not None:
            plan_json = (
                '{"man_days":'
                + str(man_days)
                + ',"start_date":"2026-02-01","due_date":"2026-02-25","jira_url":""}'
            )
        conn.execute(
            """
            INSERT INTO epics_management (
                epic_key, project_key, project_name, product_category, component, epic_name,
                description, originator, priority, plan_status, ipp_meeting_planned, actual_production_date, remarks, jira_url,
                epic_plan_json, research_urs_plan_json, dds_plan_json,
                development_plan_json, sqa_plan_json, user_manual_plan_json, production_plan_json
            ) VALUES (?, ?, ?, ?, ?, ?, '', '', 'Low', 'Not Planned Yet', 'No', '', '', '', ?, '{}', '{}', '{}', '{}', '{}', '{}')
            """,
            (
                epic_key,
                "O2",
                "O2",
                "General",
                "",
                "Epic from Planner",
                plan_json,
            ),
        )
        conn.execute(
            """
            INSERT INTO epics_management_story_sync (
                story_key, epic_key, project_key, story_name, story_status, jira_url,
                start_date, due_date, estimate_hours, payload_json, synced_at_utc
            ) VALUES (?, ?, 'O2', 'Story from Planner', 'In Progress', '', '2026-02-02', '2026-02-22', 40.0, '{}', '2026-03-03T00:00:00Z')
            """,
            (story_key, epic_key),
        )
        conn.commit()
    finally:
        conn.close()


def _cache_db_path(root: Path) -> Path:
    return root / "assignee_hours_capacity.db"


class PlannedVsDispensedApiTests(unittest.TestCase):
    def test_summary_and_details_from_mocked_hierarchy(self):
        hierarchy = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 100.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-20",
                }
            ],
            "stories": [
                {
                    "issue_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Story One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 0.0,
                    "planned_start": "2026-02-02",
                    "planned_due": "2026-02-12",
                }
            ],
            "subtasks": [
                {
                    "issue_key": "O2-SUB1",
                    "story_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "issue_type_name": "Sub-task",
                    "summary": "Subtask One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 25.0,
                    "planned_start": "2026-02-03",
                    "planned_due": "2026-02-05",
                },
                {
                    "issue_key": "O2-SUB2",
                    "story_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "issue_type_name": "Bug Subtask",
                    "summary": "Subtask Two",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 20.0,
                    "planned_start": "2026-02-06",
                    "planned_due": "2026-02-08",
                },
            ],
        }

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            client = app.test_client()
            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server._get_planned_vs_dispensed_hierarchy_cached", return_value=(hierarchy, "canonical_db")),
                patch(
                    "report_server._load_canonical_worklogs_by_issue",
                    return_value={
                        "O2-SUB1": [
                            {"issue_key": "O2-SUB1", "started_date": "2026-02-03", "hours_logged": 3.0},
                            {"issue_key": "O2-SUB1", "started_date": "2026-02-05", "hours_logged": 2.0},
                        ],
                        "O2-SUB2": [
                            {"issue_key": "O2-SUB2", "started_date": "2026-02-07", "hours_logged": 4.0},
                            {"issue_key": "O2-SUB2", "started_date": "2026-03-03", "hours_logged": 6.0},
                        ],
                        "O2-ST1": [
                            {"issue_key": "O2-ST1", "started_date": "2026-02-04", "hours_logged": 99.0},
                        ],
                    },
                ),
            ):
                summary_resp = client.get(
                    "/api/planned-vs-dispensed/summary?from=2026-02-01&to=2026-02-28&mode=log_date"
                )
                self.assertEqual(summary_resp.status_code, 200)
                summary = summary_resp.get_json()
                self.assertTrue(summary.get("ok"))
                row = next(
                    (item for item in (summary.get("rows", []) or []) if item.get("project_key") == "O2"),
                    None,
                )
                self.assertIsNotNone(row)
                self.assertEqual(row["project_key"], "O2")
                self.assertEqual(row["planned_epic_hours"], 100.0)
                self.assertEqual(row["dispensed_subtask_hours"], 45.0)
                self.assertEqual(row["remaining_hours"], 55.0)
                self.assertEqual(row.get("dispensed_bucket_mode"), "week")
                self.assertEqual(
                    row.get("dispensed_buckets"),
                    [
                        {"bucket_key": "2026-W06", "bucket_label": "2026-W06", "hours": 45.0},
                        {"bucket_key": "remaining_hours", "bucket_label": "Remaining Hours", "hours": 55.0},
                    ],
                )
                self.assertEqual(row.get("dispensed_stack_hours"), 100.0)
                self.assertEqual(row.get("remaining_hours_outside_range"), 55.0)
                self.assertEqual(float(row.get("actual_hours") or 0.0), 15.0)
                self.assertEqual(float(row.get("actual_in_range_hours") or 0.0), 9.0)
                self.assertEqual(row.get("actual_bucket_mode"), "week")
                self.assertEqual(
                    row.get("actual_buckets"),
                    [
                        {"bucket_key": "2026-W06", "bucket_label": "2026-W06", "hours": 9.0},
                        {"bucket_key": "remaining_hours", "bucket_label": "Remaining Hours", "hours": 6.0},
                    ],
                )

                details_resp = client.get(
                    "/api/planned-vs-dispensed/details?from=2026-02-01&to=2026-02-28&mode=log_date&project_key=O2"
                )
                self.assertEqual(details_resp.status_code, 200)
                details = details_resp.get_json()
                self.assertTrue(details.get("ok"))
                self.assertEqual(details["totals"]["planned_epic_hours"], 100.0)
                self.assertEqual(details["totals"]["dispensed_subtask_hours"], 45.0)
                self.assertEqual(details["totals"]["actual_hours"], 15.0)
                self.assertEqual(details["totals"]["remaining_hours"], 55.0)
                self.assertEqual(len(details.get("epics", [])), 1)
                self.assertEqual(details["epics"][0]["stories"][0]["subtasks"][0]["planned_start"], "2026-02-03")
                self.assertEqual(details["epics"][0]["stories"][0]["subtasks"][0]["planned_due"], "2026-02-05")
                self.assertEqual(
                    details["epics"][0]["dispensed_buckets"],
                    [
                        {"bucket_key": "2026-W06", "bucket_label": "2026-W06", "hours": 45.0},
                        {"bucket_key": "remaining_hours", "bucket_label": "Remaining Hours", "hours": 55.0},
                    ],
                )
                self.assertEqual(float(details["epics"][0]["dispensed_stack_hours"]), 100.0)
                self.assertEqual(float(details["epics"][0]["actual_hours"]), 15.0)
                self.assertEqual(float(details["epics"][0]["actual_in_range_hours"]), 9.0)
                self.assertEqual(
                    details["epics"][0]["actual_buckets"],
                    [
                        {"bucket_key": "2026-W06", "bucket_label": "2026-W06", "hours": 9.0},
                        {"bucket_key": "remaining_hours", "bucket_label": "Remaining Hours", "hours": 6.0},
                    ],
                )

    def test_summary_uses_monthly_dispensed_buckets_for_multi_month_range(self):
        hierarchy = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 100.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-03-20",
                }
            ],
            "stories": [
                {
                    "issue_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Story One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 0.0,
                    "planned_start": "2026-02-02",
                    "planned_due": "2026-03-12",
                }
            ],
            "subtasks": [
                {
                    "issue_key": "O2-SUB1",
                    "story_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Subtask One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 25.0,
                    "planned_start": "2026-02-27",
                    "planned_due": "2026-02-28",
                },
                {
                    "issue_key": "O2-SUB2",
                    "story_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Subtask Two",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 20.0,
                    "planned_start": "2026-03-01",
                    "planned_due": "2026-03-02",
                },
            ],
        }
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            client = app.test_client()
            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server._get_planned_vs_dispensed_hierarchy_cached", return_value=(hierarchy, "canonical_db")),
            ):
                summary_resp = client.get(
                    "/api/planned-vs-dispensed/summary?from=2026-02-28&to=2026-03-01&mode=log_date"
                )
                self.assertEqual(summary_resp.status_code, 200)
                summary = summary_resp.get_json()
                self.assertTrue(summary.get("ok"))
                self.assertEqual(summary.get("dispensed_bucket_mode"), "month")
                row = next(
                    (item for item in (summary.get("rows", []) or []) if item.get("project_key") == "O2"),
                    None,
                )
                self.assertIsNotNone(row)
                self.assertEqual(row.get("dispensed_bucket_mode"), "month")
                self.assertEqual(
                    row.get("dispensed_buckets"),
                    [
                        {"bucket_key": "2026-02", "bucket_label": "2026-02", "hours": 25.0},
                        {"bucket_key": "2026-03", "bucket_label": "2026-03", "hours": 20.0},
                        {"bucket_key": "remaining_hours", "bucket_label": "Remaining Hours", "hours": 55.0},
                    ],
                )

    def test_summary_assigns_full_subtask_hours_to_anchor_week(self):
        hierarchy = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 120.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-28",
                }
            ],
            "stories": [
                {
                    "issue_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Story One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 0.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-28",
                }
            ],
            "subtasks": [
                {
                    "issue_key": "O2-SUB1",
                    "story_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Long Subtask",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 40.0,
                    "planned_start": "2026-02-03",
                    "planned_due": "2026-02-22",
                },
            ],
        }
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            client = app.test_client()
            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server._get_planned_vs_dispensed_hierarchy_cached", return_value=(hierarchy, "canonical_db")),
            ):
                summary_resp = client.get(
                    "/api/planned-vs-dispensed/summary?from=2026-02-01&to=2026-02-28&mode=log_date"
                )
                self.assertEqual(summary_resp.status_code, 200)
                summary = summary_resp.get_json()
                self.assertTrue(summary.get("ok"))
                row = next(
                    (item for item in (summary.get("rows", []) or []) if item.get("project_key") == "O2"),
                    None,
                )
                self.assertIsNotNone(row)
                self.assertEqual(row.get("dispensed_bucket_mode"), "week")
                buckets = row.get("dispensed_buckets") or []
                keys = [item.get("bucket_key") for item in buckets]
                self.assertEqual(keys, ["2026-W06", "remaining_hours"])
                hours = [float(item.get("hours") or 0.0) for item in buckets]
                self.assertAlmostEqual(sum(hours), 120.0, places=1)
                self.assertEqual(float(buckets[0].get("hours") or 0.0), 40.0)

    def test_summary_places_remaining_after_in_range_bucket_when_outside_subtasks_excluded(self):
        hierarchy = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 100.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-28",
                }
            ],
            "stories": [
                {
                    "issue_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Story One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 0.0,
                    "planned_start": "2026-01-01",
                    "planned_due": "2026-02-28",
                }
            ],
            "subtasks": [
                {
                    "issue_key": "O2-SUB1",
                    "story_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Subtask Before Range",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 20.0,
                    "planned_start": "2026-01-10",
                    "planned_due": "2026-01-20",
                },
                {
                    "issue_key": "O2-SUB2",
                    "story_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Subtask In Range",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 40.0,
                    "planned_start": "2026-02-03",
                    "planned_due": "2026-02-09",
                },
            ],
        }
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            client = app.test_client()
            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server._get_planned_vs_dispensed_hierarchy_cached", return_value=(hierarchy, "canonical_db")),
            ):
                summary_resp = client.get(
                    "/api/planned-vs-dispensed/summary?from=2026-02-01&to=2026-02-28&mode=log_date"
                )
                self.assertEqual(summary_resp.status_code, 200)
                summary = summary_resp.get_json()
                self.assertTrue(summary.get("ok"))
                row = next(
                    (item for item in (summary.get("rows", []) or []) if item.get("project_key") == "O2"),
                    None,
                )
                self.assertIsNotNone(row)
                buckets = row.get("dispensed_buckets") or []
                self.assertGreaterEqual(len(buckets), 2)
                self.assertNotEqual(buckets[0].get("bucket_key"), "remaining_hours")
                self.assertEqual(buckets[-1].get("bucket_key"), "remaining_hours")

    def test_summary_counts_subtask_when_only_due_date_is_in_range(self):
        hierarchy = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 100.0,
                    "planned_start": "2026-01-01",
                    "planned_due": "2026-02-28",
                }
            ],
            "stories": [
                {
                    "issue_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Story One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 0.0,
                    "planned_start": "2026-01-01",
                    "planned_due": "2026-02-28",
                }
            ],
            "subtasks": [
                {
                    "issue_key": "O2-SUB1",
                    "story_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Cross Range Subtask",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 18.0,
                    "planned_start": "2026-01-20",
                    "planned_due": "2026-02-05",
                }
            ],
        }
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            client = app.test_client()
            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server._get_planned_vs_dispensed_hierarchy_cached", return_value=(hierarchy, "canonical_db")),
            ):
                summary_resp = client.get(
                    "/api/planned-vs-dispensed/summary?from=2026-02-01&to=2026-02-28&mode=log_date"
                )
                self.assertEqual(summary_resp.status_code, 200)
                summary = summary_resp.get_json()
                self.assertTrue(summary.get("ok"))
                row = next(
                    (item for item in (summary.get("rows", []) or []) if item.get("project_key") == "O2"),
                    None,
                )
                self.assertIsNotNone(row)
                self.assertEqual(float(row.get("dispensed_in_range_hours") or 0.0), 18.0)

    def test_summary_does_not_create_week_buckets_for_fully_outside_subtasks(self):
        hierarchy = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 100.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-28",
                }
            ],
            "stories": [
                {
                    "issue_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Story One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 0.0,
                    "planned_start": "2026-01-01",
                    "planned_due": "2026-03-31",
                }
            ],
            "subtasks": [
                {
                    "issue_key": "O2-SUB1",
                    "story_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Before Range",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 20.0,
                    "planned_start": "2026-01-05",
                    "planned_due": "2026-01-20",
                },
                {
                    "issue_key": "O2-SUB2",
                    "story_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "After Range",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 30.0,
                    "planned_start": "2026-03-10",
                    "planned_due": "2026-03-25",
                },
            ],
        }
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            client = app.test_client()
            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server._get_planned_vs_dispensed_hierarchy_cached", return_value=(hierarchy, "canonical_db")),
            ):
                summary_resp = client.get(
                    "/api/planned-vs-dispensed/summary?from=2026-02-01&to=2026-02-28&mode=log_date"
                )
                self.assertEqual(summary_resp.status_code, 200)
                summary = summary_resp.get_json()
                self.assertTrue(summary.get("ok"))
                row = next(
                    (item for item in (summary.get("rows", []) or []) if item.get("project_key") == "O2"),
                    None,
                )
                self.assertIsNotNone(row)
                buckets = row.get("dispensed_buckets") or []
                self.assertEqual(
                    buckets,
                    [{"bucket_key": "remaining_hours", "bucket_label": "Remaining Hours", "hours": 100.0}],
                )

    def test_validation_errors(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            client = app.test_client()

            bad_mode = client.get(
                "/api/planned-vs-dispensed/summary?from=2026-02-01&to=2026-02-28&mode=bad"
            )
            self.assertEqual(bad_mode.status_code, 400)

            missing_project = client.get(
                "/api/planned-vs-dispensed/details?from=2026-02-01&to=2026-02-28&mode=log_date"
            )
            self.assertEqual(missing_project.status_code, 400)

            bad_plan_source = client.get(
                "/api/planned-vs-dispensed/summary?from=2026-02-01&to=2026-02-28&mode=log_date&plan_source=bad"
            )
            self.assertEqual(bad_plan_source.status_code, 400)

    def test_summary_keeps_selected_projects_with_zero_rows(self):
        hierarchy = {"epics": [], "stories": [], "subtasks": []}
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            client = app.test_client()
            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server._get_planned_vs_dispensed_hierarchy_cached", return_value=(hierarchy, "canonical_db")),
            ):
                resp = client.get(
                    "/api/planned-vs-dispensed/summary?from=2026-02-01&to=2026-02-28&mode=log_date&projects=O2,OC"
                )
                self.assertEqual(resp.status_code, 200)
                payload = resp.get_json()
                self.assertTrue(payload.get("ok"))
                rows = payload.get("rows", [])
                keys = sorted(row.get("project_key") for row in rows)
                self.assertEqual(keys, ["O2", "OC"])
                for row in rows:
                    self.assertEqual(float(row.get("planned_epic_hours", 1.0)), 0.0)
                    self.assertEqual(float(row.get("dispensed_subtask_hours", 1.0)), 0.0)
                    self.assertEqual(float(row.get("remaining_hours", 1.0)), 0.0)

    def test_hierarchy_cache_persists_across_app_restart(self):
        hierarchy = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 16.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-10",
                }
            ],
            "stories": [],
            "subtasks": [],
        }
        query = "/api/planned-vs-dispensed/summary?from=2026-02-01&to=2026-02-28&mode=log_date&projects=O2"

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app_first = _build_app(root)
            client_first = app_first.test_client()
            with patch("report_server._get_planned_vs_dispensed_hierarchy_cached", return_value=(hierarchy, "canonical_db")) as mock_loader:
                first_resp = client_first.get(query)
                self.assertEqual(first_resp.status_code, 200)
                first_payload = first_resp.get_json()
                self.assertTrue(first_payload.get("ok"))
                self.assertEqual(first_payload.get("source"), "canonical_db")
                self.assertEqual(mock_loader.call_count, 1)

            app_second = _build_app(root)
            client_second = app_second.test_client()
            with patch("report_server._get_planned_vs_dispensed_hierarchy_cached", side_effect=AssertionError("should use cache")):
                second_resp = client_second.get(query)
                self.assertEqual(second_resp.status_code, 200)
                second_payload = second_resp.get_json()
                self.assertTrue(second_payload.get("ok"))
                self.assertEqual(second_payload.get("source"), "canonical_db")

    def test_clear_planned_vs_dispensed_cache_tables_removes_only_cache_rows(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = _cache_db_path(root)
            report_server.clear_planned_vs_dispensed_cache_tables(db_path)
            with sqlite3.connect(db_path) as conn:
                conn.execute("CREATE TABLE IF NOT EXISTS preserved_table (id INTEGER PRIMARY KEY, value TEXT)")
                conn.execute("INSERT INTO preserved_table (value) VALUES ('keep')")
                conn.execute(
                    """
                    INSERT INTO planned_vs_dispensed_cache (
                        cache_key, from_date, to_date, mode, project_scope, payload_json, fetched_at
                    ) VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
                    """,
                    ("summary-cache", "2026-02-01", "2026-02-28", "log_date", "O2", '{"ok":true}'),
                )
                conn.execute(
                    """
                    INSERT INTO planned_vs_dispensed_response_cache (
                        cache_key, endpoint, from_date, to_date, mode,
                        project_scope, statuses_scope, assignees_scope, project_key,
                        payload_json, fetched_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
                    """,
                    ("response-cache", "summary", "2026-02-01", "2026-02-28", "log_date", "O2", "", "", "", '{"ok":true}'),
                )
                conn.commit()

            report_server.clear_planned_vs_dispensed_cache_tables(db_path)

            with sqlite3.connect(db_path) as conn:
                self.assertEqual(
                    conn.execute("SELECT COUNT(*) FROM planned_vs_dispensed_cache").fetchone()[0],
                    0,
                )
                self.assertEqual(
                    conn.execute("SELECT COUNT(*) FROM planned_vs_dispensed_response_cache").fetchone()[0],
                    0,
                )
                self.assertEqual(
                    conn.execute("SELECT value FROM preserved_table").fetchone()[0],
                    "keep",
                )

    def test_startup_cache_clear_forces_summary_recompute_after_restart(self):
        hierarchy_first = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 10.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-10",
                }
            ],
            "stories": [],
            "subtasks": [],
        }
        hierarchy_second = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 20.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-10",
                }
            ],
            "stories": [],
            "subtasks": [],
        }
        query = "/api/planned-vs-dispensed/summary?from=2026-02-01&to=2026-02-28&mode=log_date&projects=O2"

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app_first = _build_app(root)
            client_first = app_first.test_client()
            with patch("report_server._get_planned_vs_dispensed_hierarchy_cached", return_value=(hierarchy_first, "canonical_db")):
                first_resp = client_first.get(query)
                self.assertEqual(first_resp.status_code, 200)
                first_payload = first_resp.get_json()
                self.assertEqual(float(first_payload["rows"][0]["planned_epic_hours"]), 10.0)

            report_server.clear_planned_vs_dispensed_cache_tables(_cache_db_path(root))

            app_second = _build_app(root)
            client_second = app_second.test_client()
            with patch("report_server._get_planned_vs_dispensed_hierarchy_cached", return_value=(hierarchy_second, "canonical_db")) as mock_loader:
                second_resp = client_second.get(query)
                self.assertEqual(second_resp.status_code, 200)
                second_payload = second_resp.get_json()
                self.assertEqual(float(second_payload["rows"][0]["planned_epic_hours"]), 20.0)
                self.assertEqual(mock_loader.call_count, 1)

    def test_refresh_flag_bypasses_cached_summary_response(self):
        hierarchy_first = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 10.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-10",
                }
            ],
            "stories": [],
            "subtasks": [],
        }
        hierarchy_second = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 20.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-10",
                }
            ],
            "stories": [],
            "subtasks": [],
        }
        base_q = "/api/planned-vs-dispensed/summary?from=2026-02-01&to=2026-02-28&mode=log_date&projects=O2"
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            client = app.test_client()
            with patch("report_server._get_planned_vs_dispensed_hierarchy_cached", return_value=(hierarchy_first, "canonical_db")):
                first_resp = client.get(base_q)
                self.assertEqual(first_resp.status_code, 200)
                first_payload = first_resp.get_json()
                o2_first = next((r for r in first_payload.get("rows", []) if r.get("project_key") == "O2"), None)
                self.assertIsNotNone(o2_first)
                self.assertEqual(float(o2_first.get("planned_epic_hours", 0.0)), 10.0)

            with patch("report_server._get_planned_vs_dispensed_hierarchy_cached", return_value=(hierarchy_second, "canonical_db")):
                second_resp = client.get(base_q + "&refresh=1")
                self.assertEqual(second_resp.status_code, 200)
                second_payload = second_resp.get_json()
                o2_second = next((r for r in second_payload.get("rows", []) if r.get("project_key") == "O2"), None)
                self.assertIsNotNone(o2_second)
                self.assertEqual(float(o2_second.get("planned_epic_hours", 0.0)), 20.0)

    def test_details_uses_bottom_up_dispensed_and_rollup_dates(self):
        hierarchy = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 100.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-20",
                }
            ],
            "stories": [
                {
                    "issue_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Story One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 80.0,
                    "planned_start": "2026-02-02",
                    "planned_due": "2026-02-12",
                }
            ],
            "subtasks": [
                {
                    "issue_key": "O2-SUB1",
                    "story_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Subtask One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 5.0,
                    "planned_start": "2026-02-04",
                    "planned_due": "2026-02-06",
                },
                {
                    "issue_key": "O2-SUB2",
                    "story_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Subtask Two",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 3.0,
                    "planned_start": "2026-02-03",
                    "planned_due": "2026-02-10",
                },
            ],
        }
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            client = app.test_client()
            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server._get_planned_vs_dispensed_hierarchy_cached", return_value=(hierarchy, "canonical_db")),
            ):
                resp = client.get(
                    "/api/planned-vs-dispensed/details?from=2026-02-01&to=2026-02-28&mode=log_date&project_key=O2"
                )
                self.assertEqual(resp.status_code, 200)
                payload = resp.get_json()
                epic = payload["epics"][0]
                story = epic["stories"][0]
                self.assertEqual(float(story["dispensed_estimates"]), 8.0)
                self.assertEqual(float(epic["dispensed_estimates"]), 8.0)
                self.assertEqual(story["dispensed_start"], "2026-02-03")
                self.assertEqual(story["dispensed_due"], "2026-02-10")
                self.assertEqual(epic["dispensed_start"], "2026-02-03")
                self.assertEqual(epic["dispensed_due"], "2026-02-10")
                self.assertIsNone(story["subtasks"][0]["planned_hours"])

    def test_epic_planner_source_uses_db_and_reports_missing_epics(self):
        hierarchy = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 100.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-20",
                },
                {
                    "issue_key": "O2-EP2",
                    "project_key": "O2",
                    "summary": "Epic Two",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 50.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-20",
                },
            ],
            "stories": [
                {
                    "issue_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Story One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 10.0,
                    "planned_start": "2026-02-02",
                    "planned_due": "2026-02-12",
                }
            ],
            "subtasks": [
                {
                    "issue_key": "O2-SUB1",
                    "story_key": "O2-ST1",
                    "epic_key": "O2-EP1",
                    "summary": "Subtask One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 5.0,
                    "planned_start": "2026-02-03",
                    "planned_due": "2026-02-05",
                }
            ],
        }
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            _seed_epic_planner_data(root, epic_key="O2-EP1", man_days=5.0, story_key="O2-ST1")
            _seed_epic_planner_data(root, epic_key="O2-EP2", man_days=None, story_key="O2-ST2")
            app = _build_app(root)
            client = app.test_client()
            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server._get_planned_vs_dispensed_hierarchy_cached", return_value=(hierarchy, "canonical_db")),
            ):
                summary_resp = client.get(
                    "/api/planned-vs-dispensed/summary?from=2026-02-01&to=2026-02-28&mode=log_date&plan_source=epic_planner&projects=O2"
                )
                self.assertEqual(summary_resp.status_code, 200)
                summary = summary_resp.get_json()
                self.assertEqual(summary.get("plan_source"), "epic_planner")
                row = summary["rows"][0]
                self.assertEqual(float(row["planned_epic_hours"]), 40.0)
                self.assertEqual(int(summary["missing_plan"]["count"]), 1)
                self.assertIn("O2-EP2", summary["missing_plan"]["projects"][0]["epic_keys"])

                details_resp = client.get(
                    "/api/planned-vs-dispensed/details?from=2026-02-01&to=2026-02-28&mode=log_date&plan_source=epic_planner&project_key=O2"
                )
                self.assertEqual(details_resp.status_code, 200)
                details = details_resp.get_json()
                self.assertEqual(details.get("plan_source"), "epic_planner")
                epic_one = next((e for e in details["epics"] if e["issue_key"] == "O2-EP1"), None)
                self.assertIsNotNone(epic_one)
                self.assertEqual(float(epic_one["planned_hours"]), 40.0)
                self.assertEqual(epic_one["planned_start"], "2026-02-01")
                story_one = epic_one["stories"][0]
                self.assertEqual(float(story_one["planned_hours"]), 40.0)
                self.assertEqual(story_one["planned_start"], "2026-02-02")

    def test_response_cache_separated_by_plan_source(self):
        hierarchy = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 10.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-10",
                }
            ],
            "stories": [],
            "subtasks": [],
        }
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            _seed_epic_planner_data(root, epic_key="O2-EP1", man_days=2.0, story_key="O2-STX")
            app = _build_app(root)
            client = app.test_client()
            q_jira = "/api/planned-vs-dispensed/summary?from=2026-02-01&to=2026-02-28&mode=log_date&projects=O2&plan_source=jira_estimates"
            q_planner = "/api/planned-vs-dispensed/summary?from=2026-02-01&to=2026-02-28&mode=log_date&projects=O2&plan_source=epic_planner"
            with patch("report_server._get_planned_vs_dispensed_hierarchy_cached", return_value=(hierarchy, "canonical_db")) as mock_loader:
                resp_jira = client.get(q_jira)
                resp_planner = client.get(q_planner)
                self.assertEqual(resp_jira.status_code, 200)
                self.assertEqual(resp_planner.status_code, 200)
                self.assertEqual(mock_loader.call_count, 2)
                payload_jira = resp_jira.get_json()
                payload_planner = resp_planner.get_json()
                self.assertEqual(float(payload_jira["rows"][0]["planned_epic_hours"]), 10.0)
                self.assertEqual(float(payload_planner["rows"][0]["planned_epic_hours"]), 16.0)


if __name__ == "__main__":
    unittest.main()
