from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

import report_server
from report_server import create_report_server_app


def _seed_canonical_run(db_path: Path, run_id: str = "canonical-test-run") -> str:
    with sqlite3.connect(db_path) as conn:
        now = "2026-03-10T00:00:00+00:00"
        conn.execute(
            """
            INSERT OR REPLACE INTO canonical_refresh_runs(
                run_id, scope_year, managed_project_keys_json, started_at_utc, ended_at_utc,
                status, trigger_source, error_message, stats_json,
                progress_step, progress_pct, cancel_requested, updated_at_utc
            ) VALUES (?, 2026, '["O2","RLT"]', ?, ?, 'success', 'test', '', '{}', 'done', 100, 0, ?)
            """,
            (run_id, now, now, now),
        )
        conn.execute(
            "UPDATE canonical_refresh_state SET active_run_id=?, last_success_run_id=?, updated_at_utc=? WHERE id=1",
            (run_id, run_id, now),
        )
        conn.commit()
    return run_id


class Group2CanonicalRefreshTests(unittest.TestCase):
    def test_group2_reports_use_canonical_refresh_chain(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir(parents=True, exist_ok=True)
            (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            canonical_run_id = _seed_canonical_run(root / "assignee_hours_capacity.db")
            client = app.test_client()

            original_run_script = report_server._run_script
            original_sync_report_html = report_server.sync_report_html
            calls: list[tuple[str, dict[str, str]]] = []

            def _fake_run_script(script_name, _cwd, extra_args=None, env_overrides=None):
                calls.append((script_name, dict(env_overrides or {})))
                return 0, "ok", ""

            report_server._run_script = _fake_run_script
            report_server.sync_report_html = lambda *_args, **_kwargs: None
            try:
                expectations = {
                    "rlt_leave_report": ["generate_rlt_leave_report.py"],
                    "leaves_planned_calendar": ["generate_rlt_leave_report.py", "generate_leaves_planned_calendar_html.py"],
                    "assignee_hours": ["generate_rlt_leave_report.py", "generate_assignee_hours_report.py"],
                }
                for report_id, expected_steps in expectations.items():
                    calls.clear()
                    response = client.post("/api/report/refresh", json={"report": report_id})
                    self.assertEqual(response.status_code, 200)
                    payload = response.get_json() or {}
                    self.assertTrue(payload.get("ok"))
                    self.assertEqual(payload.get("canonical_run_id"), canonical_run_id)
                    self.assertEqual(payload.get("steps"), expected_steps)
                    called_scripts = [name for name, _env in calls]
                    self.assertEqual(called_scripts, expected_steps)
                    self.assertEqual(calls[0][1].get("JIRA_LEAVE_REPORT_SOURCE"), "canonical_db")
                    self.assertEqual(calls[0][1].get("JIRA_CANONICAL_RUN_ID"), canonical_run_id)
            finally:
                report_server._run_script = original_run_script
                report_server.sync_report_html = original_sync_report_html


if __name__ == "__main__":
    unittest.main()
