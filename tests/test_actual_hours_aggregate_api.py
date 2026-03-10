from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from report_server import create_report_server_app


def _write_work_items(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(
        [
            "project_key",
            "issue_key",
            "jira_issue_type",
            "parent_issue_key",
            "start_date",
            "end_date",
        ]
    )
    ws.append(["O2", "O2-EP1", "Epic", "", "2026-02-10", "2026-02-20"])
    ws.append(["O2", "O2-ST1", "Story", "O2-EP1", "", ""])
    ws.append(["O2", "O2-SUB1", "Sub-task", "O2-ST1", "2026-02-12", "2026-02-18"])
    ws.append(["O2", "O2-EP2", "Epic", "", "2026-01-01", "2026-01-05"])
    ws.append(["O2", "O2-ST2", "Story", "O2-EP2", "", ""])
    ws.append(["O2", "O2-SUB2", "Sub-task", "O2-ST2", "2026-01-02", "2026-01-04"])
    wb.save(path)


def _write_worklogs(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(
        [
            "issue_id",
            "parent_epic_id",
            "issue_assignee",
            "worklog_author",
            "worklog_started",
            "hours_logged",
        ]
    )
    ws.append(["O2-SUB1", "O2-EP1", "Alice", "Alice", "2026-02-15T10:00:00+0500", 3])
    ws.append(["O2-SUB1", "O2-EP1", "Alice", "Alice", "2026-03-01T10:00:00+0500", 2])
    ws.append(["O2-SUB2", "O2-EP2", "Bob", "Bob", "2026-02-16T10:00:00+0500", 4])
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
        issues = [
            ("O2-EP1", "O2", "Epic", "", "2026-02-10", "2026-02-20", "", "O2-EP1"),
            ("O2-ST1", "O2", "Story", "O2-EP1", "", "", "O2-ST1", "O2-EP1"),
            ("O2-SUB1", "O2", "Sub-task", "O2-ST1", "2026-02-12", "2026-02-18", "O2-ST1", "O2-EP1"),
            ("O2-EP2", "O2", "Epic", "", "2026-01-01", "2026-01-05", "", "O2-EP2"),
            ("O2-ST2", "O2", "Story", "O2-EP2", "", "", "O2-ST2", "O2-EP2"),
            ("O2-SUB2", "O2", "Sub-task", "O2-ST2", "2026-01-02", "2026-01-04", "O2-ST2", "O2-EP2"),
        ]
        for idx, (issue_key, project_key, issue_type, parent_key, start_date, due_date, story_key, epic_key) in enumerate(issues, start=1):
            conn.execute(
                """
                INSERT OR REPLACE INTO canonical_issues(
                    run_id, issue_id, issue_key, project_key, issue_type, summary, status, assignee,
                    start_date, due_date, created_utc, updated_utc, resolved_stable_since_date,
                    original_estimate_hours, total_hours_logged, fix_type, parent_issue_key, story_key, epic_key, raw_payload_json
                ) VALUES (?, ?, ?, ?, ?, ?, 'Open', ?, ?, ?, ?, ?, '', 0, 0, '', ?, ?, ?, '{}')
                """,
                (run_id, str(idx), issue_key, project_key, issue_type, issue_key, "Alice" if "1" in issue_key else "Bob", start_date, due_date, now, now, parent_key, story_key, epic_key),
            )
            conn.execute(
                """
                INSERT OR REPLACE INTO canonical_issue_links(
                    run_id, issue_key, parent_issue_key, story_key, epic_key, hierarchy_level
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                (run_id, issue_key, parent_key, story_key, epic_key, "subtask" if "SUB" in issue_key else ("story" if "ST" in issue_key else "epic")),
            )
        for worklog_id, issue_key, started_date, hours, author in [
            ("wl-1", "O2-SUB1", "2026-02-15", 3.0, "Alice"),
            ("wl-2", "O2-SUB1", "2026-03-01", 2.0, "Alice"),
            ("wl-3", "O2-SUB2", "2026-02-16", 4.0, "Bob"),
        ]:
            conn.execute(
                """
                INSERT OR REPLACE INTO canonical_worklogs(
                    run_id, worklog_id, issue_key, project_key, worklog_author, issue_assignee,
                    started_utc, started_date, updated_utc, hours_logged
                ) VALUES (?, ?, ?, 'O2', ?, ?, ?, ?, ?, ?)
                """,
                (run_id, worklog_id, issue_key, author, author, f"{started_date}T10:00:00+0500", started_date, f"{started_date}T10:00:00+0500", hours),
            )
        conn.commit()
    return run_id


class ActualHoursAggregateApiTests(unittest.TestCase):
    def test_aggregate_mode_behaviors_and_legacy_compat(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir(parents=True, exist_ok=True)
            (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            _write_work_items(root / "1_jira_work_items_export.xlsx")
            _write_worklogs(root / "2_jira_subtask_worklogs.xlsx")

            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            _seed_canonical_run(root / "assignee_hours_capacity.db")
            client = app.test_client()

            log_resp = client.get(
                "/api/actual-hours/aggregate?from=2026-02-01&to=2026-02-28&mode=log_date&report=test"
            )
            self.assertEqual(log_resp.status_code, 200)
            log_payload = log_resp.get_json()
            self.assertTrue(log_payload.get("ok"))
            self.assertEqual(log_payload.get("mode"), "log_date")
            self.assertEqual(log_payload.get("source_file"), "canonical_db")
            self.assertEqual(log_payload["epic_hours_by_issue"].get("O2-EP1"), 3.0)
            self.assertEqual(log_payload["epic_hours_by_issue"].get("O2-EP2"), 4.0)
            self.assertEqual(log_payload["project_hours_by_key"].get("O2"), 7.0)

            planned_resp = client.get(
                "/api/actual-hours/aggregate?from=2026-02-01&to=2026-02-28&mode=planned_dates&report=test"
            )
            self.assertEqual(planned_resp.status_code, 200)
            planned_payload = planned_resp.get_json()
            self.assertTrue(planned_payload.get("ok"))
            self.assertEqual(planned_payload.get("mode"), "planned_dates")
            self.assertEqual(planned_payload["epic_hours_by_issue"].get("O2-EP1"), 5.0)
            self.assertIsNone(planned_payload["epic_hours_by_issue"].get("O2-EP2"))
            self.assertEqual(planned_payload["project_hours_by_key"].get("O2"), 5.0)
            day_map = planned_payload["assignee_hours_by_period"]["day"]
            self.assertEqual(day_map["2026-02-15"]["Alice"], 3.0)
            self.assertIsNone(day_map.get("2026-03-01"))

            invalid_resp = client.get(
                "/api/actual-hours/aggregate?from=2026-02-01&to=2026-02-28&mode=bad_mode&report=test"
            )
            self.assertEqual(invalid_resp.status_code, 400)

            legacy_resp = client.get(
                "/api/nested-view/actual-hours?from=2026-02-01&to=2026-02-28&mode=planned_dates"
            )
            self.assertEqual(legacy_resp.status_code, 200)
            legacy_payload = legacy_resp.get_json()
            self.assertTrue(legacy_payload.get("ok"))
            self.assertEqual(legacy_payload.get("source_file"), "canonical_db")
            self.assertEqual(legacy_payload.get("mode"), "planned_dates")
            self.assertEqual(legacy_payload["subtask_hours_by_issue"].get("O2-SUB1"), 5.0)
            self.assertEqual(legacy_payload["subtask_hours_by_issue"].get("O2-SUB2"), 4.0)

            nested_log_resp = client.get(
                "/api/nested-view/actual-hours?from=2026-02-01&to=2026-02-28&mode=log_date"
            )
            self.assertEqual(nested_log_resp.status_code, 200)
            nested_log_payload = nested_log_resp.get_json()
            self.assertTrue(nested_log_payload.get("ok"))
            self.assertEqual(nested_log_payload.get("mode"), "log_date")
            self.assertEqual(nested_log_payload["subtask_hours_by_issue"].get("O2-SUB1"), 3.0)
            self.assertEqual(nested_log_payload["subtask_hours_by_issue"].get("O2-SUB2"), 4.0)


if __name__ == "__main__":
    unittest.main()
