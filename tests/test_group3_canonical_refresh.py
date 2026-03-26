from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import report_server
from report_server import create_report_server_app


OPERATOR_HEADERS = {"X-Role": "operator"}


def _build_app(root: Path):
    (root / "report_html").mkdir(parents=True, exist_ok=True)
    for name in [
        "dashboard.html",
        "planned_vs_dispensed_report.html",
        "planned_actual_table_view.html",
        "original_estimates_hierarchy_report.html",
    ]:
        (root / "report_html" / name).write_text("<html><body>ok</body></html>", encoding="utf-8")
    return create_report_server_app(base_dir=root, folder_raw="report_html")


def _seed_canonical_run(db_path: Path, run_id: str = "canonical-phase7-run") -> str:
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
            ("", run_id, now),
        )
        conn.commit()
    return run_id


def _seed_canonical_hierarchy(db_path: Path, run_id: str) -> None:
    with sqlite3.connect(db_path) as conn:
        issue_rows = [
            (run_id, "", "O2-EP1", "O2", "Epic", "Epic One", "In Progress", "Alice", "2026-02-01", "2026-02-20", "", "", "", 40.0, 8.0, "", "", "", "O2-EP1", "{}"),
            (run_id, "", "O2-ST1", "O2", "Story", "Story One", "In Progress", "Alice", "2026-02-02", "2026-02-12", "", "", "", 16.0, 8.0, "", "", "O2-ST1", "O2-EP1", "{}"),
            (run_id, "", "O2-SUB1", "O2", "Sub-task", "Subtask One", "In Progress", "Alice", "2026-02-03", "2026-02-05", "", "", "", 8.0, 6.0, "", "O2-ST1", "O2-ST1", "O2-EP1", "{}"),
            (run_id, "", "O2-SUB2", "O2", "Sub-task", "Subtask Two", "In Progress", "Alice", "2026-02-06", "2026-02-07", "", "", "", 4.0, 2.0, "", "O2-ST1", "O2-ST1", "O2-EP1", "{}"),
        ]
        conn.executemany(
            """
            INSERT OR REPLACE INTO canonical_issues(
                run_id, issue_id, issue_key, project_key, issue_type, summary, status, assignee,
                start_date, due_date, created_utc, updated_utc, resolved_stable_since_date,
                original_estimate_hours, total_hours_logged, fix_type, parent_issue_key,
                story_key, epic_key, raw_payload_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            issue_rows,
        )
        link_rows = [
            (run_id, "O2-EP1", "", "", "O2-EP1", "epic"),
            (run_id, "O2-ST1", "", "O2-ST1", "O2-EP1", "story"),
            (run_id, "O2-SUB1", "O2-ST1", "O2-ST1", "O2-EP1", "subtask"),
            (run_id, "O2-SUB2", "O2-ST1", "O2-ST1", "O2-EP1", "subtask"),
        ]
        conn.executemany(
            """
            INSERT OR REPLACE INTO canonical_issue_links(
                run_id, issue_key, parent_issue_key, story_key, epic_key, hierarchy_level
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            link_rows,
        )
        worklog_rows = [
            (run_id, "wl-1", "O2-SUB1", "O2", "Alice", "Alice", "2026-02-04T09:00:00+00:00", "2026-02-04", "", 6.0),
            (run_id, "wl-2", "O2-SUB2", "O2", "Alice", "Alice", "2026-02-06T09:00:00+00:00", "2026-02-06", "", 2.0),
        ]
        conn.executemany(
            """
            INSERT OR REPLACE INTO canonical_worklogs(
                run_id, worklog_id, issue_key, project_key, worklog_author, issue_assignee,
                started_utc, started_date, updated_utc, hours_logged
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            worklog_rows,
        )
        conn.commit()


class Group3CanonicalRefreshTests(unittest.TestCase):
    def test_canonical_compatibility_rebuild_refreshes_jira_exports_status(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            _build_app(root)
            db_path = root / "assignee_hours_capacity.db"
            run_id = _seed_canonical_run(db_path)
            with sqlite3.connect(db_path) as conn:
                conn.execute(
                    """
                    INSERT OR REPLACE INTO canonical_issues(
                        run_id, issue_id, issue_key, project_key, issue_type, summary, status, assignee,
                        start_date, due_date, created_utc, updated_utc, resolved_stable_since_date,
                        original_estimate_hours, total_hours_logged, fix_type, parent_issue_key,
                        story_key, epic_key, raw_payload_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        run_id,
                        "314",
                        "FF-314",
                        "FF",
                        "Epic",
                        "Configurable Alarms and Sites, Notifications",
                        "On Hold",
                        "Muhammad Usman Javed",
                        "2026-02-06",
                        "2026-02-27",
                        "2026-02-01T00:00:00+00:00",
                        "2026-03-18T10:56:54.515+05:00",
                        "",
                        128.0,
                        0.0,
                        "",
                        "",
                        "FF-314",
                        "FF-314",
                        "{}",
                    ),
                )
                conn.execute(
                    """
                    INSERT OR REPLACE INTO canonical_issue_links(
                        run_id, issue_key, parent_issue_key, story_key, epic_key, hierarchy_level
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (run_id, "FF-314", "", "", "FF-314", "epic"),
                )
                conn.commit()

            exports_db_path = root / "jira_exports.db"
            with sqlite3.connect(exports_db_path) as conn:
                conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS work_items (
                        project_key TEXT,
                        issue_key TEXT,
                        work_item_id TEXT,
                        work_item_type TEXT,
                        jira_issue_type TEXT,
                        fix_type TEXT,
                        summary TEXT,
                        status TEXT,
                        start_date TEXT,
                        end_date TEXT,
                        actual_start_date TEXT,
                        actual_end_date TEXT,
                        original_estimate TEXT,
                        original_estimate_hours REAL,
                        assignee TEXT,
                        total_hours_logged REAL,
                        priority TEXT,
                        parent_issue_key TEXT,
                        parent_work_item_id TEXT,
                        parent_jira_url TEXT,
                        jira_url TEXT,
                        latest_ipp_meeting TEXT,
                        jira_ipp_rmi_dates_altered TEXT,
                        ipp_actual_date TEXT,
                        ipp_remarks TEXT,
                        ipp_actual_date_matches_jira_end_date TEXT,
                        created TEXT,
                        updated TEXT
                    )
                    """
                )
                conn.execute(
                    """
                    INSERT INTO work_items(
                        project_key, issue_key, work_item_id, work_item_type, jira_issue_type, fix_type, summary, status,
                        start_date, end_date, actual_start_date, actual_end_date, original_estimate, original_estimate_hours,
                        assignee, total_hours_logged, priority, parent_issue_key, parent_work_item_id, parent_jira_url,
                        jira_url, latest_ipp_meeting, jira_ipp_rmi_dates_altered, ipp_actual_date, ipp_remarks,
                        ipp_actual_date_matches_jira_end_date, created, updated
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        "FF",
                        "FF-314",
                        "FF-314",
                        "Epic",
                        "Epic",
                        "",
                        "Configurable Alarms and Sites, Notifications",
                        "In-Progress",
                        "2026-02-06",
                        "2026-02-27",
                        "",
                        "",
                        "128",
                        128.0,
                        "Muhammad Usman Javed",
                        0.0,
                        "",
                        "",
                        "",
                        "",
                        "https://octopusdtlsupport.atlassian.net/browse/FF-314",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "2026-02-01T00:00:00+00:00",
                        "2026-03-01T00:00:00+00:00",
                    ),
                )
                conn.commit()

            with patch.dict("os.environ", {"JIRA_EXPORTS_DB_PATH": str(exports_db_path)}, clear=False):
                stats = report_server._canonical_rebuild_compatibility_artifacts(db_path, run_id, root)
            self.assertEqual(int(stats.get("work_items_rows") or 0), 1)

            with sqlite3.connect(exports_db_path) as conn:
                row = conn.execute(
                    "SELECT status FROM work_items WHERE issue_key = ?",
                    ("FF-314",),
                ).fetchone()
            self.assertIsNotNone(row)
            self.assertEqual(str(row[0] or ""), "On Hold")

    def test_planned_vs_dispensed_and_planned_actual_use_canonical_data(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            db_path = root / "assignee_hours_capacity.db"
            run_id = _seed_canonical_run(db_path)
            _seed_canonical_hierarchy(db_path, run_id)
            client = app.test_client()

            summary_resp = client.get(
                "/api/planned-vs-dispensed/summary?from=2026-02-01&to=2026-02-28&mode=log_date&projects=O2"
            )
            self.assertEqual(summary_resp.status_code, 200)
            summary = summary_resp.get_json() or {}
            self.assertTrue(summary.get("ok"))
            self.assertEqual(summary.get("source"), "canonical_db")
            row = next((item for item in (summary.get("rows") or []) if item.get("project_key") == "O2"), None)
            self.assertIsNotNone(row)
            self.assertEqual(float(row.get("planned_epic_hours") or 0.0), 40.0)
            self.assertEqual(float(row.get("dispensed_subtask_hours") or 0.0), 12.0)

            refresh_resp = client.post(
                "/api/planned-actual-table-view/refresh",
                json={"from": "2026-02-01", "to": "2026-02-28", "mode": "log_date", "projects": "O2", "run_sync": True},
                headers=OPERATOR_HEADERS,
            )
            self.assertEqual(refresh_resp.status_code, 200)
            run_payload = refresh_resp.get_json() or {}
            self.assertTrue(run_payload.get("ok"))
            pactv_run_id = str(run_payload.get("run_id") or "")
            self.assertTrue(pactv_run_id)
            with sqlite3.connect(db_path) as conn:
                row = conn.execute(
                    "SELECT source FROM planned_actual_refresh_runs WHERE run_id=?",
                    (pactv_run_id,),
                ).fetchone()
            self.assertIsNotNone(row)
            self.assertEqual(str(row[0] or ""), "canonical_db")

    def test_group3_report_refreshes_use_canonical_refresh_chain(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            canonical_run_id = _seed_canonical_run(root / "assignee_hours_capacity.db")
            client = app.test_client()

            original_run_script = report_server._run_script
            original_sync_report_html = report_server.sync_report_html
            calls: list[str] = []

            def _fake_run_script(script_name, _cwd, extra_args=None, env_overrides=None):
                calls.append(script_name)
                return 0, "ok", ""

            report_server._run_script = _fake_run_script
            report_server.sync_report_html = lambda *_args, **_kwargs: None
            try:
                expectations = {
                    "planned_rmis": ["generate_planned_rmis_html.py"],
                    "gantt_chart": ["generate_gantt_chart_html.py"],
                    "phase_rmi_gantt": ["generate_phase_rmi_gantt_html.py"],
                    "planned_vs_dispensed": ["generate_planned_vs_dispensed_report.py"],
                    "planned_actual_table_view": ["generate_planned_actual_table_view.py"],
                    "original_estimates_hierarchy": ["generate_original_estimates_hierarchy_report.py"],
                }
                for report_id, expected_steps in expectations.items():
                    calls.clear()
                    response = client.post("/api/report/refresh", json={"report": report_id})
                    self.assertEqual(response.status_code, 200)
                    payload = response.get_json() or {}
                    self.assertTrue(payload.get("ok"))
                    self.assertEqual(payload.get("canonical_run_id"), canonical_run_id)
                    self.assertEqual(payload.get("steps"), expected_steps)
                    self.assertEqual(calls, expected_steps)
            finally:
                report_server._run_script = original_run_script
                report_server.sync_report_html = original_sync_report_html


if __name__ == "__main__":
    unittest.main()
