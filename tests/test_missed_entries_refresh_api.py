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


def _seed_canonical_run(db_path: Path, run_id: str = "canonical-test-run") -> str:
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
        conn.execute(
            """
            INSERT OR REPLACE INTO canonical_issues(
                run_id, issue_id, issue_key, project_key, issue_type, summary, status, assignee,
                start_date, due_date, created_utc, updated_utc, resolved_stable_since_date,
                original_estimate_hours, total_hours_logged, fix_type, parent_issue_key, story_key, epic_key, raw_payload_json
            ) VALUES (?, '101', 'O2-101', 'O2', 'Sub-task', 'Implement sync', 'Done', 'Alice',
                      '2026-03-01', '2026-03-04', ?, ?, '', 8.0, 2.0, '', '', 'O2-100', 'O2-1', '{}')
            """,
            (run_id, now, now),
        )
        conn.commit()
    return run_id


class MissedEntriesRefreshApiTests(unittest.TestCase):
    def test_refresh_saves_snapshot_rows_to_database(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            tdp = Path(td)
            (tdp / "report_html").mkdir(parents=True, exist_ok=True)
            (tdp / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            _write_work_items_xlsx(tdp / "1_jira_work_items_export.xlsx")
            app = create_report_server_app(base_dir=tdp, folder_raw="report_html")
            _seed_canonical_run(tdp / "assignee_hours_capacity.db")
            client = app.test_client()
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
                stats_row = conn.execute(
                    "SELECT stats_json FROM me_refresh_runs WHERE run_id = ?",
                    (run_id,),
                ).fetchone()
            self.assertIsNotNone(snapshot_count_row)
            self.assertGreater(int(snapshot_count_row[0] or 0), 0)
            self.assertIsNotNone(state_row)
            self.assertEqual(str(state_row[0] or ""), run_id)
            self.assertEqual(str(state_row[1] or ""), run_id)
            self.assertIn("canonical_db", str(stats_row[0] or ""))

    def test_cancel_marks_running_missed_entries_refresh(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            tdp = Path(td)
            (tdp / "report_html").mkdir(parents=True, exist_ok=True)
            (tdp / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            _write_work_items_xlsx(tdp / "1_jira_work_items_export.xlsx")
            app = create_report_server_app(base_dir=tdp, folder_raw="report_html")
            client = app.test_client()
            db_path = tdp / "assignee_hours_capacity.db"
            _seed_canonical_run(db_path)

            original_builder = report_server._canonical_build_missed_entries_rows

            def _slow_builder(db_path, run_id):
                for _ in range(40):
                    time.sleep(0.05)
                return original_builder(db_path, run_id)

            report_server._canonical_build_missed_entries_rows = _slow_builder
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
                report_server._canonical_build_missed_entries_rows = original_builder


if __name__ == "__main__":
    unittest.main()
