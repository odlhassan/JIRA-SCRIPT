from __future__ import annotations

import sqlite3
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import patch

import report_server
from db_migration import plan_migration


class CanonicalTwoPhaseTests(unittest.TestCase):
    def test_compute_promotes_only_after_successful_compute(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "capacity.db"
            report_server._init_canonical_refresh_db(db_path)
            fetch_run_id = "fetch-1"
            now = report_server._canonical_now_utc()
            with sqlite3.connect(db_path) as conn:
                conn.execute(
                    """INSERT INTO canonical_refresh_runs(
                        run_id, scope_year, managed_project_keys_json, started_at_utc,
                        ended_at_utc, status, trigger_source, error_message, stats_json,
                        progress_step, progress_pct, cancel_requested, updated_at_utc
                    ) VALUES (?, 2026, '[\"O2\"]', ?, ?, 'success', 'test', '', '{}', 'fetch_done', 100, 0, ?)""",
                    (fetch_run_id, now, now, now),
                )
                conn.execute(
                    "INSERT INTO canonical_issues(run_id, issue_key, project_key) VALUES (?, 'O2-1', 'O2')",
                    (fetch_run_id,),
                )
                conn.execute(
                    "UPDATE canonical_refresh_state SET last_success_run_id='old-report', last_success_fetch_run_id=? WHERE id=1",
                    (fetch_run_id,),
                )
                conn.commit()
            compute_run_id = report_server._canonical_create_compute_run(db_path, fetch_run_id, "test")
            with (
                patch.object(report_server, "_canonical_rebuild_derived_data", return_value={"issues": 1}),
                patch.object(report_server, "_sync_epics_management_from_canonical", return_value={"updated": 0}),
                patch.object(report_server, "_canonical_rebuild_compatibility_artifacts", return_value={"ok": True}),
                patch.object(report_server, "_run_script", return_value=(0, "", "")),
                patch.object(report_server, "sync_report_html", return_value=0),
            ):
                result, status = report_server._run_canonical_compute(
                    db_path, Path(td), compute_run_id, fetch_run_id, "test"
                )
            self.assertEqual(status, 200)
            self.assertTrue(result["ok"])
            with sqlite3.connect(db_path) as conn:
                state = conn.execute(
                    "SELECT last_success_run_id, last_success_compute_run_id FROM canonical_refresh_state WHERE id=1"
                ).fetchone()
            self.assertEqual(state, (fetch_run_id, compute_run_id))

    def test_incremental_discovery_uses_checkpoint_overlap(self) -> None:
        captured: list[str] = []

        def fake_fetch(_session, jql: str, fields: list[str]):
            captured.append(jql)
            return [{"key": "O2-9", "fields": {"project": {"key": "O2"}}}]

        checkpoint = "2026-08-01T10:00:00+00:00"
        with patch.object(report_server, "export_fetch_issues", side_effect=fake_fetch):
            issues, reasons = report_server._canonical_collect_incremental_candidates(
                object(), ["O2"], checkpoint, "customfield_1", ["duedate"]
            )
        self.assertEqual(set(issues), {"O2-9"})
        self.assertIn("updated_since_checkpoint", reasons["O2-9"])
        self.assertIn('updated >= "2026-08-01 09:50"', captured[0])

    def test_reconciliation_becomes_due_after_seven_days(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "capacity.db"
            report_server._init_canonical_refresh_db(db_path)
            old = (datetime.now(timezone.utc) - timedelta(days=8)).isoformat()
            with sqlite3.connect(db_path) as conn:
                conn.execute(
                    "UPDATE canonical_refresh_state SET last_full_reconciliation_at_utc=? WHERE id=1",
                    (old,),
                )
                conn.commit()
            self.assertTrue(report_server._canonical_reconciliation_due(db_path))

    def test_production_migration_plan_includes_two_phase_schema(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            local_db = root / "local.db"
            prod_db = root / "production.db"
            report_server._init_canonical_refresh_db(local_db)
            with sqlite3.connect(prod_db) as conn:
                conn.execute(
                    """CREATE TABLE canonical_refresh_state (
                        id INTEGER PRIMARY KEY CHECK(id = 1), active_run_id TEXT NOT NULL DEFAULT '',
                        last_success_run_id TEXT NOT NULL DEFAULT '', updated_at_utc TEXT NOT NULL
                    )"""
                )
                conn.execute("INSERT INTO canonical_refresh_state(id, updated_at_utc) VALUES(1, '')")
                conn.commit()
            plan = plan_migration(prod_db, local_db)
            tables = {step["table"] for step in plan["steps"]}
            self.assertTrue({"canonical_fetch_runs", "canonical_compute_runs", "employee_performance_scoped_runs"} <= tables)
            self.assertIn("canonical_refresh_state", tables)


if __name__ == "__main__":
    unittest.main()
