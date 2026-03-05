from __future__ import annotations

import sqlite3
import tempfile
import time
import unittest
from pathlib import Path

from openpyxl import Workbook

import report_server
from report_server import create_report_server_app


def _write_work_items_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(
        [
            "issue_key",
            "jira_issue_type",
            "assignee",
            "summary",
            "start_date",
            "end_date",
            "original_estimate",
            "total_hours_logged",
            "jira_url",
        ]
    )
    ws.append(
        [
            "O2-101",
            "Sub-task",
            "Alice",
            "Implement sync",
            "2026-03-01",
            "2026-03-04",
            "8h",
            2.0,
            "https://example.atlassian.net/browse/O2-101",
        ]
    )
    wb.save(path)


class MissedEntriesRefreshApiTests(unittest.TestCase):
    def test_refresh_saves_snapshot_rows_to_database(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            tdp = Path(td)
            (tdp / "report_html").mkdir(parents=True, exist_ok=True)
            (tdp / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            _write_work_items_xlsx(tdp / "1_jira_work_items_export.xlsx")
            app = create_report_server_app(base_dir=tdp, folder_raw="report_html")
            client = app.test_client()

            original_interruptible = report_server._run_script_interruptible

            def _fake_interruptible(script_name, base_dir, extra_args=None, env_overrides=None, cancel_check=None, poll_interval_sec=0.5):
                return 0, f"ok:{script_name}", ""

            report_server._run_script_interruptible = _fake_interruptible
            try:
                start_resp = client.post("/api/missed-entries/refresh", json={})
                self.assertEqual(start_resp.status_code, 202)
                run_id = str((start_resp.get_json() or {}).get("run_id") or "")
                self.assertTrue(run_id)

                status_value = ""
                for _ in range(60):
                    status_resp = client.get(f"/api/missed-entries/refresh/{run_id}")
                    self.assertEqual(status_resp.status_code, 200)
                    run = (status_resp.get_json() or {}).get("run") or {}
                    status_value = str(run.get("status") or "").lower()
                    if status_value in {"success", "failed", "canceled"}:
                        break
                    time.sleep(0.1)
                self.assertEqual(status_value, "success")

                db_path = tdp / "assignee_hours_capacity.db"
                with sqlite3.connect(db_path) as conn:
                    snapshot_count_row = conn.execute(
                        "SELECT COUNT(*) FROM me_snapshot_rows WHERE run_id = ?",
                        (run_id,),
                    ).fetchone()
                    state_row = conn.execute(
                        "SELECT active_run_id, last_success_run_id FROM me_refresh_state WHERE id = 1"
                    ).fetchone()
                self.assertIsNotNone(snapshot_count_row)
                self.assertGreater(int(snapshot_count_row[0] or 0), 0)
                self.assertIsNotNone(state_row)
                self.assertEqual(str(state_row[0] or ""), run_id)
                self.assertEqual(str(state_row[1] or ""), run_id)
            finally:
                report_server._run_script_interruptible = original_interruptible

    def test_cancel_marks_running_missed_entries_refresh(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            tdp = Path(td)
            (tdp / "report_html").mkdir(parents=True, exist_ok=True)
            (tdp / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            _write_work_items_xlsx(tdp / "1_jira_work_items_export.xlsx")
            app = create_report_server_app(base_dir=tdp, folder_raw="report_html")
            client = app.test_client()
            db_path = tdp / "assignee_hours_capacity.db"

            original_interruptible = report_server._run_script_interruptible

            def _slow_interruptible(script_name, base_dir, extra_args=None, env_overrides=None, cancel_check=None, poll_interval_sec=0.5):
                for _ in range(40):
                    if callable(cancel_check) and bool(cancel_check()):
                        return -1, "", "Canceled by user."
                    time.sleep(0.05)
                return 0, f"ok:{script_name}", ""

            report_server._run_script_interruptible = _slow_interruptible
            try:
                start_resp = client.post("/api/missed-entries/refresh", json={})
                self.assertEqual(start_resp.status_code, 202)
                body = start_resp.get_json() or {}
                run_id = str(body.get("run_id") or "")
                self.assertTrue(run_id)

                cancel_resp = client.post("/api/missed-entries/cancel", json={"run_id": run_id})
                self.assertEqual(cancel_resp.status_code, 200)
                cancel_body = cancel_resp.get_json() or {}
                self.assertTrue(cancel_body.get("ok"))
                self.assertEqual(str(cancel_body.get("status") or ""), "cancel_requested")

                cancel_requested = 0
                for _ in range(30):
                    with sqlite3.connect(db_path) as conn:
                        row = conn.execute(
                            "SELECT cancel_requested FROM me_refresh_runs WHERE run_id = ?",
                            (run_id,),
                        ).fetchone()
                    cancel_requested = int((row[0] if row else 0) or 0)
                    if cancel_requested == 1:
                        break
                    time.sleep(0.05)
                self.assertEqual(cancel_requested, 1)
            finally:
                report_server._run_script_interruptible = original_interruptible


if __name__ == "__main__":
    unittest.main()
