from __future__ import annotations

import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from report_server import create_report_server_app


def _build_app(root: Path):
    (root / "report_html").mkdir(parents=True, exist_ok=True)
    (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
    (root / "report_html" / "planned_actual_table_view.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
    wb = Workbook()
    ws = wb.active
    ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
    ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
    wb.save(root / "assignee_hours_report.xlsx")
    return create_report_server_app(base_dir=root, folder_raw="report_html")


VIEWER_HEADERS = {"X-Role": "viewer"}
OPERATOR_HEADERS = {"X-Role": "operator"}
ADMIN_HEADERS = {"X-Role": "admin", "X-Actor": "qa-user"}


class PlannedActualTableViewApiTests(unittest.TestCase):
    def test_validation_errors(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = _build_app(Path(td))
            client = app.test_client()

            bad = client.get("/api/planned-actual-table-view/summary?from=2026-02-01&to=2026-02-28&mode=bad")
            self.assertEqual(bad.status_code, 400)

            diff_missing = client.get("/api/planned-actual-table-view/diff", headers=VIEWER_HEADERS)
            self.assertEqual(diff_missing.status_code, 400)

    def test_refresh_summary_history_and_diff(self):
        hierarchy = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "project_name": "OmniConnect",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 40.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-20",
                }
            ],
            "stories": [
                {
                    "issue_key": "O2-ST1",
                    "project_key": "O2",
                    "project_name": "OmniConnect",
                    "epic_key": "O2-EP1",
                    "summary": "Story One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 16.0,
                    "planned_start": "2026-02-02",
                    "planned_due": "2026-02-12",
                }
            ],
            "subtasks": [
                {
                    "issue_key": "O2-SUB1",
                    "project_key": "O2",
                    "project_name": "OmniConnect",
                    "epic_key": "O2-EP1",
                    "story_key": "O2-ST1",
                    "summary": "Subtask One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 8.0,
                    "planned_start": "2026-02-03",
                    "planned_due": "2026-02-05",
                }
            ],
        }
        hierarchy2 = {
            **hierarchy,
            "subtasks": [
                {
                    **hierarchy["subtasks"][0],
                    "estimate_hours": 10.0,
                }
            ],
        }
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = _build_app(Path(td))
            client = app.test_client()
            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server._load_planned_vs_dispensed_hierarchy", return_value=hierarchy),
                patch("report_server._fetch_subtask_actual_hours_by_keys", return_value={"O2-SUB1": 6.0}),
            ):
                refresh_resp = client.post(
                    "/api/planned-actual-table-view/refresh",
                    json={"from": "2026-02-01", "to": "2026-02-28", "mode": "log_date", "projects": "O2", "run_sync": True},
                    headers=OPERATOR_HEADERS,
                )
                self.assertEqual(refresh_resp.status_code, 200)
                self.assertTrue(refresh_resp.get_json().get("ok"))
                self.assertEqual(int(refresh_resp.get_json().get("attempt") or 0), 1)
                self.assertEqual(int(refresh_resp.get_json().get("max_attempts") or 0), 1)

            summary = client.get("/api/planned-actual-table-view/summary?from=2026-02-01&to=2026-02-28&mode=log_date&projects=O2")
            self.assertEqual(summary.status_code, 200)
            payload = summary.get_json()
            self.assertTrue(payload.get("ok"))
            self.assertFalse(payload.get("needs_refresh"))
            self.assertGreater(len(payload.get("rows", [])), 0)
            self.assertEqual(float(payload.get("totals", {}).get("planned_hours", 0.0)), 40.0)
            self.assertEqual(float(payload.get("totals", {}).get("actual_hours", 0.0)), 6.0)

            first_snapshot = str(payload.get("snapshot_id"))

            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server._load_planned_vs_dispensed_hierarchy", return_value=hierarchy2),
                patch("report_server._fetch_subtask_actual_hours_by_keys", return_value={"O2-SUB1": 9.0}),
            ):
                refresh_resp_2 = client.post(
                    "/api/planned-actual-table-view/refresh",
                    json={"from": "2026-02-01", "to": "2026-02-28", "mode": "log_date", "projects": "O2", "run_sync": True, "force_full": True},
                    headers=OPERATOR_HEADERS,
                )
                self.assertEqual(refresh_resp_2.status_code, 200)

            summary2 = client.get("/api/planned-actual-table-view/summary?from=2026-02-01&to=2026-02-28&mode=log_date&projects=O2")
            second_snapshot = str((summary2.get_json() or {}).get("snapshot_id"))
            self.assertNotEqual(first_snapshot, second_snapshot)

            history = client.get("/api/planned-actual-table-view/history?limit=10", headers=VIEWER_HEADERS)
            self.assertEqual(history.status_code, 200)
            self.assertTrue((history.get_json() or {}).get("rows"))

            queue = client.get("/api/planned-actual-table-view/queue", headers=VIEWER_HEADERS)
            self.assertEqual(queue.status_code, 200)
            self.assertIn("rows", queue.get_json() or {})

            diff = client.get(
                f"/api/planned-actual-table-view/diff?left_snapshot_id={first_snapshot}&right_snapshot_id={second_snapshot}",
                headers=VIEWER_HEADERS,
            )
            self.assertEqual(diff.status_code, 200)
            diff_payload = diff.get_json() or {}
            self.assertTrue(diff_payload.get("ok"))
            self.assertIn("delta", diff_payload)

            export_csv = client.post(
                "/api/planned-actual-table-view/export",
                json={"snapshot_id": second_snapshot, "format": "csv"},
                headers=VIEWER_HEADERS,
            )
            self.assertEqual(export_csv.status_code, 200)
            self.assertIn("text/csv", str(export_csv.content_type))

            export_xlsx = client.post(
                "/api/planned-actual-table-view/export",
                json={"snapshot_id": second_snapshot, "format": "xlsx"},
                headers=VIEWER_HEADERS,
            )
            self.assertEqual(export_xlsx.status_code, 200)
            self.assertIn("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", str(export_xlsx.content_type))

            pin_resp = client.post(
                f"/api/planned-actual-table-view/snapshots/{second_snapshot}/pin-official",
                headers=ADMIN_HEADERS,
            )
            self.assertEqual(pin_resp.status_code, 200)
            pin_payload = pin_resp.get_json() or {}
            self.assertTrue(pin_payload.get("ok"))
            self.assertTrue((pin_payload.get("snapshot") or {}).get("is_official"))
            self.assertEqual((pin_payload.get("snapshot") or {}).get("official_pinned_by"), "qa-user")

            unpin_resp = client.post(
                f"/api/planned-actual-table-view/snapshots/{second_snapshot}/unpin-official",
                headers=ADMIN_HEADERS,
            )
            self.assertEqual(unpin_resp.status_code, 200)
            unpin_payload = unpin_resp.get_json() or {}
            self.assertTrue(unpin_payload.get("ok"))
            self.assertFalse((unpin_payload.get("snapshot") or {}).get("is_official"))

    def test_permissions_enforced_for_sensitive_endpoints(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = _build_app(Path(td))
            client = app.test_client()
            refresh_forbidden = client.post(
                "/api/planned-actual-table-view/refresh",
                json={"from": "2026-02-01", "to": "2026-02-28", "mode": "log_date", "projects": "O2", "run_sync": True},
            )
            self.assertEqual(refresh_forbidden.status_code, 403)

            history_forbidden = client.get("/api/planned-actual-table-view/history")
            self.assertEqual(history_forbidden.status_code, 403)

            export_forbidden = client.post("/api/planned-actual-table-view/export", json={"format": "csv"})
            self.assertEqual(export_forbidden.status_code, 403)

    def test_filter_options_projects_include_all_managed_projects(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = _build_app(Path(td))
            client = app.test_client()
            create_o2 = client.post(
                "/api/projects",
                json={
                    "project_key": "ZZ1",
                    "project_name": "Project ZZ1",
                    "display_name": "Project ZZ1",
                    "color_hex": "#336699",
                },
            )
            self.assertEqual(create_o2.status_code, 200)
            create_o3 = client.post(
                "/api/projects",
                json={
                    "project_key": "ZZ2",
                    "project_name": "Project ZZ2",
                    "display_name": "Project ZZ2",
                    "color_hex": "#669933",
                },
            )
            self.assertEqual(create_o3.status_code, 200)

            options_resp = client.get("/api/planned-actual-table-view/filter-options")
            self.assertEqual(options_resp.status_code, 200)
            options_body = options_resp.get_json() or {}
            options = (options_body.get("options") or {}) if isinstance(options_body, dict) else {}
            keys = set(str(item).upper() for item in (options.get("projects") or []))
            self.assertIn("ZZ1", keys)
            self.assertIn("ZZ2", keys)

    def test_summary_uses_managed_project_display_name_for_project_rows(self):
        hierarchy = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "project_name": "OmniConnect",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 40.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-20",
                }
            ],
            "stories": [],
            "subtasks": [],
        }
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = _build_app(Path(td))
            client = app.test_client()
            create_project = client.post(
                "/api/projects",
                json={
                    "project_key": "O2",
                    "project_name": "OmniConnect",
                    "display_name": "O2 Product Team",
                    "color_hex": "#336699",
                },
            )
            if create_project.status_code not in {200, 409}:
                self.fail(f"Unexpected status while creating project: {create_project.status_code}")
            update_project = client.put(
                "/api/projects/O2",
                json={
                    "project_name": "OmniConnect",
                    "display_name": "O2 Product Team",
                    "color_hex": "#336699",
                    "is_active": True,
                },
            )
            self.assertEqual(update_project.status_code, 200)
            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server._load_planned_vs_dispensed_hierarchy", return_value=hierarchy),
                patch("report_server._fetch_subtask_actual_hours_by_keys", return_value={}),
            ):
                refresh_resp = client.post(
                    "/api/planned-actual-table-view/refresh",
                    json={"from": "2026-02-01", "to": "2026-02-28", "mode": "log_date", "projects": "O2", "run_sync": True},
                    headers=OPERATOR_HEADERS,
                )
                self.assertEqual(refresh_resp.status_code, 200)

            summary = client.get("/api/planned-actual-table-view/summary?from=2026-02-01&to=2026-02-28&mode=log_date&projects=O2")
            self.assertEqual(summary.status_code, 200)
            payload = summary.get_json() or {}
            rows = payload.get("rows", []) if isinstance(payload, dict) else []
            project_rows = [item for item in rows if str((item or {}).get("row_type")) == "project" and str((item or {}).get("project_key")).upper() == "O2"]
            self.assertTrue(project_rows)
            self.assertEqual(str(project_rows[0].get("summary")), "O2 Product Team")

    def test_cancel_and_rollback_for_queued_run(self):
        hierarchy = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "project_name": "OmniConnect",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 40.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-20",
                }
            ],
            "stories": [],
            "subtasks": [],
        }

        def _slow_hierarchy(*_args, **_kwargs):
            time.sleep(0.8)
            return hierarchy

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = _build_app(Path(td))
            client = app.test_client()
            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server._load_planned_vs_dispensed_hierarchy", side_effect=_slow_hierarchy),
                patch("report_server._fetch_subtask_actual_hours_by_keys", return_value={}),
            ):
                first = client.post(
                    "/api/planned-actual-table-view/refresh",
                    json={"from": "2026-02-01", "to": "2026-02-28", "mode": "log_date", "projects": "O2"},
                    headers=OPERATOR_HEADERS,
                )
                self.assertEqual(first.status_code, 202)
                first_run_id = str((first.get_json() or {}).get("run_id"))
                second = client.post(
                    "/api/planned-actual-table-view/refresh",
                    json={"from": "2026-02-01", "to": "2026-02-28", "mode": "log_date", "projects": "O2"},
                    headers=OPERATOR_HEADERS,
                )
                self.assertEqual(second.status_code, 202)
                second_payload = second.get_json() or {}
                self.assertEqual(second_payload.get("status"), "queued")
                queued_run_id = str(second_payload.get("run_id"))

                cancel_resp = client.post(
                    "/api/planned-actual-table-view/cancel",
                    json={"run_id": queued_run_id},
                    headers=OPERATOR_HEADERS,
                )
                self.assertEqual(cancel_resp.status_code, 200)
                cancel_body = cancel_resp.get_json() or {}
                self.assertTrue(cancel_body.get("ok"))
                self.assertEqual(cancel_body.get("status"), "canceled")

                status_resp = client.get(f"/api/planned-actual-table-view/refresh/{queued_run_id}", headers=VIEWER_HEADERS)
                self.assertEqual(status_resp.status_code, 200)
                status_body = status_resp.get_json() or {}
                self.assertEqual(status_body.get("status"), "canceled")

                # Avoid background thread warnings by waiting for the running run to finish.
                for _ in range(20):
                    first_status = client.get(f"/api/planned-actual-table-view/refresh/{first_run_id}", headers=VIEWER_HEADERS)
                    self.assertEqual(first_status.status_code, 200)
                    first_body = first_status.get_json() or {}
                    if str(first_body.get("status")) in {"success", "failed", "canceled"}:
                        break
                    time.sleep(0.15)


if __name__ == "__main__":
    unittest.main()
