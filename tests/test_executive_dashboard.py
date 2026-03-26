from __future__ import annotations

import json
import sqlite3
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

import report_server
from report_server import create_report_server_app


def _build_app(root: Path):
    (root / "report_html").mkdir(parents=True, exist_ok=True)
    (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
    source_html = (Path(__file__).resolve().parents[1] / "executive_dashboard.html").read_text(encoding="utf-8")
    (root / "executive_dashboard.html").write_text(source_html, encoding="utf-8")
    wb = Workbook()
    ws = wb.active
    ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
    ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
    wb.save(root / "assignee_hours_report.xlsx")
    return create_report_server_app(base_dir=root, folder_raw="report_html")


def _seed_canonical_run(db_path: Path, run_id: str = "exec-run") -> str:
    with sqlite3.connect(db_path) as conn:
        now = "2026-03-10T00:00:00+00:00"
        conn.execute(
            """
            INSERT OR REPLACE INTO canonical_refresh_runs(
                run_id, scope_year, managed_project_keys_json, started_at_utc, ended_at_utc,
                status, trigger_source, error_message, stats_json,
                progress_step, progress_pct, cancel_requested, updated_at_utc
            ) VALUES (?, 2026, '["O2"]', ?, ?, 'success', 'test', '', '{}', 'done', 100, 0, ?)
            """,
            (run_id, now, now, now),
        )
        conn.execute(
            "UPDATE canonical_refresh_state SET active_run_id=?, last_success_run_id=?, updated_at_utc=? WHERE id=1",
            (run_id, run_id, now),
        )
        issues = [
            ("E1", "O2-EP1", "O2", "Epic", "Epic One", "Resolved", "Lead", "", "", "", "", 0.0, 0.0, "", "", "O2-EP1"),
            ("S1", "O2-DEV1", "O2", "Story", "Development Story", "Resolved", "Lead", "", "", "", "O2-EP1", 0.0, 0.0, "", "O2-DEV1", "O2-EP1"),
            ("S2", "O2-SQA1", "O2", "Story", "SQA Story", "Resolved", "Lead", "", "", "", "O2-EP1", 0.0, 0.0, "", "O2-SQA1", "O2-EP1"),
            ("S3", "O2-UM1", "O2", "Story", "User Manual Story", "Resolved", "Lead", "", "", "", "O2-EP1", 0.0, 0.0, "", "O2-UM1", "O2-EP1"),
            ("S4", "O2-PROD1", "O2", "Story", "Production Story", "Resolved", "Lead", "", "", "", "O2-EP1", 0.0, 0.0, "", "O2-PROD1", "O2-EP1"),
            ("T1", "O2-SUB1", "O2", "Sub-task", "Committed Subtask A", "Resolved!", "Alice", "2026-02-05", "2026-02-06", "O2-DEV1", "O2-EP1", 8.0, 0.0, "O2-DEV1", "O2-DEV1", "O2-EP1"),
            ("T2", "O2-SUB2", "O2", "Sub-task", "Committed Subtask B", "In Progress", "Alice", "2026-02-07", "", "O2-DEV1", "O2-EP1", 4.0, 0.0, "O2-DEV1", "O2-DEV1", "O2-EP1"),
            ("T3", "O2-PROD-SUB1", "O2", "Sub-task", "Production Subtask", "Resolved", "Alice", "", "", "O2-PROD1", "O2-EP1", 2.0, 0.0, "O2-PROD1", "O2-PROD1", "O2-EP1"),
            ("E2", "O2-EP2", "O2", "Epic", "Epic Two", "In Progress", "Lead", "", "", "", "", 0.0, 0.0, "", "", "O2-EP2"),
            ("S5", "O2-DEV2", "O2", "Story", "Development Story Two", "Resolved", "Lead", "", "", "", "O2-EP2", 0.0, 0.0, "", "O2-DEV2", "O2-EP2"),
            ("S6", "O2-SQA2", "O2", "Story", "SQA Story Two", "Resolved", "Lead", "", "", "", "O2-EP2", 0.0, 0.0, "", "O2-SQA2", "O2-EP2"),
            ("S7", "O2-UM2", "O2", "Story", "User Manual Story Two", "Resolved", "Lead", "", "", "", "O2-EP2", 0.0, 0.0, "", "O2-UM2", "O2-EP2"),
            ("S8", "O2-PROD2", "O2", "Story", "Production Story Two", "In Progress", "Lead", "", "", "", "O2-EP2", 0.0, 0.0, "", "O2-PROD2", "O2-EP2"),
            ("T4", "O2-SUBX", "O2", "Sub-task", "Committed Subtask Zero", "Resolved", "Bob", "2026-02-08", "2026-02-09", "O2-DEV2", "O2-EP2", 0.0, 0.0, "O2-DEV2", "O2-DEV2", "O2-EP2"),
        ]
        for row in issues:
            conn.execute(
                """
                INSERT OR REPLACE INTO canonical_issues(
                    run_id, issue_id, issue_key, project_key, issue_type, summary, status, assignee,
                    start_date, due_date, created_utc, updated_utc, resolved_stable_since_date,
                    original_estimate_hours, total_hours_logged, parent_issue_key, story_key, epic_key, raw_payload_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '', ?, ?, ?, ?, ?, '{}')
                """,
                (run_id, *row),
            )
        worklogs = [
            ("wl-1", "O2-SUB1", "2026-02-05T09:00:00+00:00", "2026-02-05", 5.0),
            ("wl-2", "O2-SUB1", "2026-02-06T12:00:00+00:00", "2026-02-06", 5.0),
            ("wl-3", "O2-SUB2", "2026-02-07T11:00:00+00:00", "2026-02-07", 2.0),
            ("wl-4", "O2-PROD-SUB1", "2026-02-10T18:00:00+00:00", "2026-02-10", 1.0),
            ("wl-5", "O2-SUBX", "2026-02-09T09:00:00+00:00", "2026-02-09", 5.0),
        ]
        for worklog_id, issue_key, started_utc, started_date, hours in worklogs:
            conn.execute(
                """
                INSERT OR REPLACE INTO canonical_worklogs(
                    run_id, worklog_id, issue_key, project_key, worklog_author, issue_assignee,
                    started_utc, started_date, updated_utc, hours_logged
                ) VALUES (?, ?, ?, 'O2', 'Tester', 'Tester', ?, ?, ?, ?)
                """,
                (run_id, worklog_id, issue_key, started_utc, started_date, started_utc, hours),
            )
        conn.commit()
    return run_id


def _seed_epics_management(db_path: Path):
    report_server._init_epics_management_db(db_path)
    conn = sqlite3.connect(db_path)
    try:
        plans_ep1 = {
            "epic_plan_json": json.dumps({"man_days": 2.0, "start_date": "2026-02-01", "due_date": "2026-02-12", "jira_url": "https://jira.example.com/browse/O2-EP1"}),
            "research_urs_plan_json": "{}",
            "dds_plan_json": "{}",
            "development_plan_json": json.dumps({"man_days": 1.0, "start_date": "2026-02-01", "due_date": "2026-02-05", "jira_url": "https://jira.example.com/browse/O2-DEV1"}),
            "sqa_plan_json": json.dumps({"man_days": 0.5, "start_date": "2026-02-06", "due_date": "2026-02-07", "jira_url": "https://jira.example.com/browse/O2-SQA1"}),
            "user_manual_plan_json": json.dumps({"man_days": 0.25, "start_date": "2026-02-08", "due_date": "2026-02-08", "jira_url": "https://jira.example.com/browse/O2-UM1"}),
            "production_plan_json": json.dumps({"man_days": 0.25, "start_date": "2026-02-09", "due_date": "2026-02-10", "jira_url": "https://jira.example.com/browse/O2-PROD1"}),
        }
        plans_ep2 = {
            "epic_plan_json": json.dumps({"man_days": 3.0, "start_date": "2026-02-01", "due_date": "2026-02-15", "jira_url": "https://jira.example.com/browse/O2-EP2"}),
            "research_urs_plan_json": "{}",
            "dds_plan_json": "{}",
            "development_plan_json": json.dumps({"man_days": 1.0, "start_date": "2026-02-01", "due_date": "2026-02-05", "jira_url": "https://jira.example.com/browse/O2-DEV2"}),
            "sqa_plan_json": json.dumps({"man_days": 0.5, "start_date": "2026-02-06", "due_date": "2026-02-07", "jira_url": "https://jira.example.com/browse/O2-SQA2"}),
            "user_manual_plan_json": json.dumps({"man_days": 0.25, "start_date": "2026-02-08", "due_date": "2026-02-08", "jira_url": "https://jira.example.com/browse/O2-UM2"}),
            "production_plan_json": json.dumps({"man_days": 0.25, "start_date": "2026-02-09", "due_date": "2026-02-10", "jira_url": "https://jira.example.com/browse/O2-PROD2"}),
        }
        for epic_key, epic_name, plans in (
            ("O2-EP1", "Epic One", plans_ep1),
            ("O2-EP2", "Epic Two", plans_ep2),
        ):
            conn.execute(
                """
                INSERT OR REPLACE INTO epics_management (
                    epic_key, project_key, project_name, product_category, component, epic_name,
                    description, originator, priority, plan_status, ipp_meeting_planned, actual_production_date,
                    delivery_status, remarks, jira_url, is_sealed,
                    epic_plan_json, research_urs_plan_json, dds_plan_json,
                    development_plan_json, sqa_plan_json, user_manual_plan_json, production_plan_json
                ) VALUES (?, 'O2', 'O2', 'General', '', ?, '', '', 'Low', 'Planned', 'No', '', 'Yet to start', '', '', 0, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    epic_key,
                    epic_name,
                    plans["epic_plan_json"],
                    plans["research_urs_plan_json"],
                    plans["dds_plan_json"],
                    plans["development_plan_json"],
                    plans["sqa_plan_json"],
                    plans["user_manual_plan_json"],
                    plans["production_plan_json"],
                ),
            )
        sealed_snapshot_old = {
            "epic_key": "O2-EP1",
            "epic_name": "Epic One",
            "plans": {"epic_plan": {"man_days": 1.5, "start_date": "2026-02-01", "due_date": "2026-02-12", "jira_url": "https://jira.example.com/browse/O2-EP1"}},
        }
        sealed_snapshot_new = {
            "epic_key": "O2-EP1",
            "epic_name": "Epic One",
            "plans": {"epic_plan": {"man_days": 2.5, "start_date": "2026-02-01", "due_date": "2026-02-20", "jira_url": "https://jira.example.com/browse/O2-EP1"}},
        }
        conn.execute(
            "INSERT OR REPLACE INTO epics_management_approved_dates (epic_key, approved_at_utc, snapshot_json) VALUES ('O2-EP1', '2026-02-09T10:00:00Z', ?)",
            (json.dumps(sealed_snapshot_old),),
        )
        conn.execute(
            "INSERT OR REPLACE INTO epics_management_approved_dates (epic_key, approved_at_utc, snapshot_json) VALUES ('O2-EP1', '2026-02-20T10:00:00Z', ?)",
            (json.dumps(sealed_snapshot_new),),
        )
        conn.commit()
    finally:
        conn.close()


class ExecutiveDashboardTests(unittest.TestCase):
    def test_dashboard_epic_refresh_runs_scoped_canonical_job_and_returns_db_payload(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            db_path = root / "assignee_hours_capacity.db"
            _seed_canonical_run(db_path)
            _seed_epics_management(db_path)
            with sqlite3.connect(db_path) as conn:
                conn.execute(
                    """
                    INSERT OR REPLACE INTO managed_projects(
                        project_key, project_name, display_name, color_hex, is_active, created_at_utc, updated_at_utc
                    ) VALUES ('O2', 'Octopus 2', 'Octopus 2', '#336699', 1, '2026-03-10T00:00:00+00:00', '2026-03-10T00:00:00+00:00')
                    """
                )
                conn.execute(
                    """
                    UPDATE epics_management
                    SET ipp_meeting_planned='Yes',
                        actual_production_date='2026-02-12',
                        remarks='DB remark'
                    WHERE epic_key='O2-EP1'
                    """
                )
                conn.commit()

            issue_catalog = {
                "O2-EP1": {
                    "id": "1001",
                    "key": "O2-EP1",
                    "fields": {
                        "project": {"key": "O2"},
                        "issuetype": {"name": "Epic"},
                        "summary": "Epic One",
                        "status": {"name": "On Hold"},
                        "assignee": {"displayName": "Lead"},
                        "parent": {},
                        "customfield_10014": "",
                        "customfield_1": "2026-02-01",
                        "duedate": "2026-02-12",
                        "created": "2025-12-20T00:00:00+00:00",
                        "updated": "2026-03-11T00:00:00+00:00",
                        "timeoriginalestimate": 57600,
                        "aggregatetimespent": 10800,
                    },
                },
                "O2-DEV1": {
                    "id": "1002",
                    "key": "O2-DEV1",
                    "fields": {
                        "project": {"key": "O2"},
                        "issuetype": {"name": "Story"},
                        "summary": "Development Story",
                        "status": {"name": "In Progress"},
                        "assignee": {"displayName": "Lead"},
                        "parent": {"key": "O2-EP1"},
                        "customfield_10014": "O2-EP1",
                        "customfield_1": "2026-02-01",
                        "duedate": "2026-02-05",
                        "created": "2026-01-02T00:00:00+00:00",
                        "updated": "2026-03-11T00:00:00+00:00",
                        "timeoriginalestimate": 14400,
                        "aggregatetimespent": 7200,
                    },
                },
                "O2-SUB1": {
                    "id": "1003",
                    "key": "O2-SUB1",
                    "fields": {
                        "project": {"key": "O2"},
                        "issuetype": {"name": "Sub-task"},
                        "summary": "Committed Subtask A",
                        "status": {"name": "In Progress"},
                        "assignee": {"displayName": "Alice"},
                        "parent": {"key": "O2-DEV1"},
                        "customfield_10014": "",
                        "customfield_1": "2026-02-05",
                        "duedate": "2026-02-06",
                        "created": "2026-02-01T00:00:00+00:00",
                        "updated": "2026-03-11T00:00:00+00:00",
                        "timeoriginalestimate": 7200,
                        "aggregatetimespent": 10800,
                    },
                },
            }

            def _fake_fetch_issue_keys(_session, issue_keys, fields):
                return [issue_catalog[key] for key in issue_keys if key in issue_catalog]

            def _fake_fetch_stories(_session, epic_keys, fields, project_keys=None):
                return [issue_catalog["O2-DEV1"]] if "O2-EP1" in epic_keys else []

            def _fake_fetch_subtasks(_session, story_keys, fields, project_keys=None):
                return [issue_catalog["O2-SUB1"]] if "O2-DEV1" in story_keys else []

            def _fake_worklogs(_session, issue_key, delay_seconds=0.2, max_retries=5, request_timeout_seconds=30.0):
                if issue_key == "O2-SUB1":
                    return [
                        {
                            "id": "wl-refresh-1",
                            "started": "2026-02-05T09:00:00.000+0000",
                            "updated": "2026-03-11T10:00:00.000+0000",
                            "timeSpentSeconds": 10800,
                            "author": {"displayName": "Alice"},
                        }
                    ]
                return []

            client = app.test_client()
            with (
                patch.object(report_server, "get_session", return_value=object()),
                patch.object(report_server, "resolve_jira_start_date_field_id", return_value="customfield_1"),
                patch.object(report_server, "resolve_jira_end_date_field_ids", return_value=["duedate"]),
                patch.object(report_server, "export_resolve_fix_type_field_id", return_value=""),
                patch.object(report_server, "_fetch_jira_issues_by_keys", side_effect=_fake_fetch_issue_keys),
                patch.object(report_server, "_fetch_story_issues_for_epics", side_effect=_fake_fetch_stories),
                patch.object(report_server, "_fetch_subtask_issues_for_stories", side_effect=_fake_fetch_subtasks),
                patch.object(report_server, "export_fetch_worklogs_for_issue", side_effect=_fake_worklogs),
                patch.object(report_server, "_canonical_rebuild_compatibility_artifacts", return_value={"source": "canonical_db"}),
            ):
                start_resp = client.post("/api/dashboard/refresh-epic/O2-EP1", json={"from": "2026-02-01", "to": "2026-02-28"})
                self.assertEqual(start_resp.status_code, 202)
                run_id = str((start_resp.get_json() or {}).get("run_id") or "")
                self.assertTrue(run_id)

                final_body = {}
                for _ in range(60):
                    poll_resp = client.get(f"/api/dashboard/refresh-epic/{run_id}")
                    self.assertEqual(poll_resp.status_code, 200)
                    final_body = poll_resp.get_json() or {}
                    if str(final_body.get("status") or "").lower() in {"success", "failed", "canceled"}:
                        break
                    time.sleep(0.1)

                self.assertEqual(str(final_body.get("status") or "").lower(), "success")
                self.assertEqual(final_body["epic"]["status"], "On Hold")
                self.assertEqual(final_body["epic"]["latest_ipp_meeting"], "Yes")
                self.assertEqual(final_body["epic"]["ipp_remarks"], "DB remark")
                self.assertEqual(final_body["epic"]["ipp_actual_matches_jira_end_date"], "Yes")
                self.assertEqual(final_body["epic"]["subtask_hours_logged_total"], 3.0)

                with sqlite3.connect(db_path) as conn:
                    state_row = conn.execute(
                        "SELECT last_success_run_id FROM canonical_refresh_state WHERE id = 1"
                    ).fetchone()
                    latest_run_id = str((state_row[0] if state_row else "") or "")
                    self.assertEqual(latest_run_id, run_id)
                    unrelated = conn.execute(
                        "SELECT status FROM canonical_issues WHERE run_id = ? AND issue_key = 'O2-EP2'",
                        (run_id,),
                    ).fetchone()
                    refreshed = conn.execute(
                        "SELECT status FROM canonical_issues WHERE run_id = ? AND issue_key = 'O2-EP1'",
                        (run_id,),
                    ).fetchone()
                    derived = conn.execute(
                        "SELECT COUNT(*) FROM canonical_issue_actuals WHERE run_id = ?",
                        (run_id,),
                    ).fetchone()
                self.assertEqual(str(unrelated[0] or ""), "In Progress")
                self.assertEqual(str(refreshed[0] or ""), "On Hold")
                self.assertGreater(int((derived[0] if derived else 0) or 0), 0)

    def test_release_phase_metadata_returns_active_unresolved_phases_in_order(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            _seed_canonical_run(root / "assignee_hours_capacity.db")
            _seed_epics_management(root / "assignee_hours_capacity.db")

            with sqlite3.connect(root / "assignee_hours_capacity.db") as conn:
                conn.execute(
                    "UPDATE canonical_issues SET status='In Progress' WHERE run_id=? AND issue_key='O2-SQA1'",
                    ("exec-run",),
                )
                conn.execute(
                    "UPDATE canonical_issues SET status='On Hold' WHERE run_id=? AND issue_key='O2-UM1'",
                    ("exec-run",),
                )
                conn.commit()

            client = app.test_client()
            resp = client.get("/api/dashboard/releases/phase-metadata?epic_key=O2-EP1&epic_key=O2-EP2&epic_key=O2-EP3")
            self.assertEqual(resp.status_code, 200)
            body = resp.get_json()
            entries = {item["epic_key"]: item for item in body["epics"]}

            self.assertEqual(entries["O2-EP1"]["current_phase_names"], ["SQA Plan", "User Manual Plan"])
            self.assertIn("SQA Plan: In Progress", entries["O2-EP1"]["phase_statuses"])
            self.assertIn("User Manual Plan: On Hold", entries["O2-EP1"]["phase_statuses"])
            self.assertEqual(entries["O2-EP2"]["current_phase_names"], ["Production Plan"])
            self.assertEqual(entries["O2-EP3"]["current_phase_names"], [])
            self.assertIn("Development Plan: Missing", entries["O2-EP3"]["phase_statuses"])

    def test_settings_get_post_and_validation(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            client = app.test_client()

            resp = client.get("/api/executive-dashboard/settings")
            self.assertEqual(resp.status_code, 200)
            self.assertEqual(resp.get_json()["settings"]["estimation_basis"], "subtask_planned_hours")

            save_resp = client.post("/api/executive-dashboard/settings", json={"estimation_basis": "epic_planned_hours"})
            self.assertEqual(save_resp.status_code, 200)
            self.assertEqual(save_resp.get_json()["settings"]["estimation_basis"], "epic_planned_hours")

            bad_resp = client.post("/api/executive-dashboard/settings", json={"estimation_basis": "bad"})
            self.assertEqual(bad_resp.status_code, 400)
            self.assertIn("estimation_basis", bad_resp.get_json()["error"])

    def test_summary_subtask_basis_and_cycle_time(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            _seed_canonical_run(root / "assignee_hours_capacity.db")
            _seed_epics_management(root / "assignee_hours_capacity.db")
            client = app.test_client()

            resp = client.get("/api/executive-dashboard/summary?from=2026-02-01&to=2026-02-15&projects=O2")
            self.assertEqual(resp.status_code, 200)
            body = resp.get_json()
            insights = body["delivery_insights"]
            self.assertEqual(insights["total_committed_items"], 3)
            self.assertEqual(insights["completed_items"], 2)
            self.assertAlmostEqual(float(insights["estimation_accuracy_pct"]), 66.67, places=2)
            self.assertEqual(insights["estimation_accuracy_excluded_count"], 1)
            self.assertEqual(insights["cycle_time_item_count"], 1)
            self.assertEqual(insights["cycle_time_blocked_count"], 1)
            self.assertAlmostEqual(float(insights["cycle_time_avg_hours"]), 129.0, places=2)
            blocked = body["blocked_epics"][0]
            self.assertEqual(blocked["epic_key"], "O2-EP2")
            self.assertEqual(blocked["blocking_reason"], "unresolved phase story")
            self.assertIn("Production Plan: In Progress", blocked["current_phase_statuses"])

    def test_summary_epic_planned_and_sealed_basis(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            _seed_canonical_run(root / "assignee_hours_capacity.db")
            _seed_epics_management(root / "assignee_hours_capacity.db")
            client = app.test_client()

            planned_save = client.post("/api/executive-dashboard/settings", json={"estimation_basis": "epic_planned_hours"})
            self.assertEqual(planned_save.status_code, 200)
            planned_resp = client.get("/api/executive-dashboard/summary?from=2026-02-01&to=2026-02-15&projects=O2")
            self.assertEqual(planned_resp.status_code, 200)
            planned_pct = float(planned_resp.get_json()["delivery_insights"]["estimation_accuracy_pct"])
            self.assertAlmostEqual(planned_pct, 42.5, places=2)

            sealed_save = client.post("/api/executive-dashboard/settings", json={"estimation_basis": "epic_sealed_hours"})
            self.assertEqual(sealed_save.status_code, 200)
            sealed_resp = client.get("/api/executive-dashboard/summary?from=2026-02-01&to=2026-02-15&projects=O2")
            self.assertEqual(sealed_resp.status_code, 200)
            sealed_body = sealed_resp.get_json()
            sealed_pct = float(sealed_body["delivery_insights"]["estimation_accuracy_pct"])
            self.assertAlmostEqual(sealed_pct, 100.0, places=2)
            self.assertEqual(sealed_body["delivery_insights"]["estimation_accuracy_excluded_count"], 1)


if __name__ == "__main__":
    unittest.main()
