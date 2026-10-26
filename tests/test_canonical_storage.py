import json
import sqlite3
import tempfile
import threading
import time
import unittest
from contextlib import closing
from pathlib import Path
from unittest.mock import patch
from types import SimpleNamespace

import report_server
import canonical_storage


class CanonicalStorageTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.db = self.root / "capacity.db"
        report_server._init_canonical_refresh_db(self.db)
        with closing(sqlite3.connect(self.db)) as conn, conn:
            for run, status in (("old", "success"), ("new", "fetch_ready")):
                conn.execute("""INSERT INTO canonical_refresh_runs
                    (run_id, started_at_utc, status, stats_json) VALUES (?, '2026-10-20', ?, ?)""",
                    (run, status, json.dumps({"mode": "full", "replace_previous_data": True})))
                conn.execute("INSERT INTO canonical_issues(run_id, issue_key, raw_payload_json) VALUES (?, 'O2-1', ?)",
                             (run, "x" * (2_000_000 if run == "old" else 10)))
            conn.execute("CREATE TABLE manual_plans (value TEXT)")
            conn.execute("INSERT INTO manual_plans VALUES ('keep my plan')")
            conn.execute("UPDATE canonical_refresh_state SET last_success_run_id='old' WHERE id=1")
        for run in ("old", "new"):
            report_server._canonical_fetch_run_upsert(self.db, run_id=run, scope_year=2026,
                managed_project_keys=["O2"], requested_mode="full", effective_mode="full",
                status="success", completed=True)
        self.compute_id = report_server._canonical_create_compute_run(self.db, "new", "test")

    def run_compute(self, *, script_code=0, sync_callback=None):
        with (
            patch.object(report_server, "_canonical_rebuild_derived_data", return_value={}),
            patch.object(report_server, "_sync_epics_management_from_canonical", return_value={}),
            patch.object(report_server, "_canonical_rebuild_compatibility_artifacts", return_value={}),
            patch.object(report_server, "_run_script", return_value=(script_code, "", "test failure")),
            patch.object(report_server, "sync_report_html", side_effect=sync_callback),
        ):
            return report_server._run_canonical_compute(self.db, self.root, self.compute_id, "new")

    def runs(self):
        with closing(sqlite3.connect(self.db)) as conn:
            return {r[0] for r in conn.execute("SELECT DISTINCT run_id FROM canonical_issues")}

    def test_success_replaces_old_data_preserves_plans_and_shrinks_file(self):
        with closing(sqlite3.connect(self.db)) as conn, conn:
            for run in ("old", "new"):
                conn.execute("INSERT INTO canonical_worklogs(run_id, worklog_id) VALUES (?, 'log-1')", (run,))
                conn.execute("INSERT INTO canonical_issue_actuals(run_id, issue_key) VALUES (?, 'O2-1')", (run,))
        before = self.db.stat().st_size
        result, status = self.run_compute()
        self.assertEqual(status, 200)
        self.assertTrue(result["ok"])
        self.assertEqual(self.runs(), {"new"})
        self.assertLess(self.db.stat().st_size, before / 2)
        with closing(sqlite3.connect(self.db)) as conn:
            self.assertEqual(conn.execute("SELECT DISTINCT run_id FROM canonical_worklogs").fetchall(), [("new",)])
            self.assertEqual(conn.execute("SELECT DISTINCT run_id FROM canonical_issue_actuals").fetchall(), [("new",)])
            self.assertEqual(conn.execute("SELECT * FROM manual_plans").fetchall(), [("keep my plan",)])
            self.assertEqual(conn.execute("PRAGMA integrity_check").fetchone()[0], "ok")
            self.assertEqual(conn.execute("SELECT last_success_fetch_run_id, last_success_compute_run_id FROM canonical_refresh_state").fetchone(), ("new", self.compute_id))
        stats = json.loads(report_server._canonical_get_run(self.db, "new")["stats_json"])
        self.assertEqual(stats["storage_cleanup"]["status"], "compacted")

    def test_failed_compute_keeps_previous_snapshot(self):
        _, status = self.run_compute(script_code=1)
        self.assertEqual(status, 500)
        self.assertEqual(self.runs(), {"old", "new"})
        self.assertEqual(report_server._canonical_last_success_run_id(self.db), "old")

    def test_cifs_allocation_failure_is_visible_maintenance_warning(self):
        with patch.object(canonical_storage, "_trim_cifs_allocation", side_effect=OSError("SMB allocation retained")):
            _, status = self.run_compute()
        self.assertEqual(status, 200)
        self.assertEqual(self.runs(), {"new"})
        stats = json.loads(report_server._canonical_get_run(self.db, "new")["stats_json"])
        self.assertEqual(stats["storage_cleanup"]["status"], "pruned")
        self.assertIn("SMB allocation retained", stats["storage_cleanup"]["warning"])

    def test_reclaimed_bytes_use_server_allocation_instead_of_logical_length(self):
        with patch.object(canonical_storage, "_trim_cifs_allocation", return_value={
            "allocated_bytes_before": 10_000_000, "allocated_bytes_after": 1_000_000,
        }):
            self.run_compute()
        stats = json.loads(report_server._canonical_get_run(self.db, "new")["stats_json"])
        self.assertEqual(stats["storage_cleanup"]["bytes_reclaimed"], 9_000_000)

    def test_cifs_helper_uses_separate_process(self):
        with patch.object(canonical_storage.sys, "platform", "linux"), patch.object(
            canonical_storage.subprocess, "run", return_value=SimpleNamespace(stdout='{"allocated_bytes_after": 123}')
        ) as run:
            self.assertEqual(canonical_storage._trim_cifs_allocation(self.db), {"allocated_bytes_after": 123})
        self.assertEqual(run.call_args.args[0][-2:], ["--trim-cifs", str(self.db.resolve())])

    def test_canceled_compute_keeps_previous_snapshot(self):
        with closing(sqlite3.connect(self.db)) as conn, conn:
            conn.execute("UPDATE canonical_compute_runs SET cancel_requested=1")
        result, _ = self.run_compute()
        self.assertEqual(result["status"], "canceled")
        self.assertEqual(self.runs(), {"old", "new"})

    def test_cancel_after_reports_finish_prevents_replacement(self):
        def cancel_after_sync(*args):
            with closing(sqlite3.connect(self.db)) as conn, conn:
                conn.execute("UPDATE canonical_compute_runs SET cancel_requested=1")
        with patch.object(report_server, "replace_previous_snapshots") as cleanup:
            with patch.object(report_server, "sync_report_html", side_effect=cancel_after_sync):
                with (
                    patch.object(report_server, "_canonical_rebuild_derived_data", return_value={}),
                    patch.object(report_server, "_sync_epics_management_from_canonical", return_value={}),
                    patch.object(report_server, "_canonical_rebuild_compatibility_artifacts", return_value={}),
                    patch.object(report_server, "_run_script", return_value=(0, "", "")),
                ):
                    result, _ = report_server._run_canonical_compute(self.db, self.root, self.compute_id, "new")
            cleanup.assert_not_called()
        self.assertEqual(result["status"], "canceled")
        self.assertEqual(report_server._canonical_last_success_run_id(self.db), "old")

    def test_fetch_failure_keeps_choice_for_retry_and_preserves_old_snapshot(self):
        with patch.object(report_server, "get_session", side_effect=RuntimeError("Jira offline")):
            result, status = report_server._run_canonical_phase1_refresh(
                db_path=self.db, artifact_base_dir=self.root, run_id="new", scope_year=2026,
                managed_project_keys=["O2"], mode="full", fetch_only=True, keep_run_running=True)
        self.assertEqual(status, 500)
        self.assertFalse(result["ok"])
        stats = json.loads(report_server._canonical_get_run(self.db, "new")["stats_json"])
        self.assertTrue(stats["replace_previous_data"])
        self.assertEqual(self.runs(), {"old", "new"})
        self.assertEqual(report_server._canonical_last_success_run_id(self.db), "old")

    def test_unchecked_and_smart_refresh_do_not_replace(self):
        for stats in ({"mode": "full", "replace_previous_data": False},
                      {"mode": "smart", "replace_previous_data": True}):
            with closing(sqlite3.connect(self.db)) as conn, conn:
                conn.execute("UPDATE canonical_refresh_runs SET stats_json=? WHERE run_id='new'", (json.dumps(stats),))
            with patch.object(report_server, "replace_previous_snapshots") as cleanup:
                _, status = self.run_compute()
                self.assertEqual(status, 200)
                cleanup.assert_not_called()
            self.assertEqual(self.runs(), {"old", "new"})

    def test_low_disk_reports_pending_compaction_without_failing_refresh(self):
        with patch.object(canonical_storage.shutil, "disk_usage") as usage:
            usage.return_value.free = 0
            _, status = self.run_compute()
        self.assertEqual(status, 200)
        self.assertEqual(self.runs(), {"new"})
        stats = json.loads(report_server._canonical_get_run(self.db, "new")["stats_json"])
        self.assertEqual(stats["storage_cleanup"]["status"], "pruned")
        self.assertIn("compaction pending", stats["current_detail"])

    def test_replacement_does_not_run_general_retention_before_publishing(self):
        with patch.object(report_server, "_canonical_prune_two_phase_runs",
                          side_effect=sqlite3.DatabaseError("database disk image is malformed")) as prune:
            result, status = self.run_compute()
        self.assertEqual(status, 200)
        self.assertTrue(result["ok"])
        prune.assert_not_called()
        self.assertEqual(self.runs(), {"new"})

    def test_startup_switches_existing_wal_before_settings_initialization(self):
        db = self.root / "assignee_hours_capacity.db"
        with closing(sqlite3.connect(db)) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
        class StopAfterJournal(Exception):
            pass
        def check_mode(path):
            with closing(sqlite3.connect(path)) as conn:
                self.assertEqual(conn.execute("PRAGMA journal_mode").fetchone()[0], "delete")
            raise StopAfterJournal()
        with patch.object(report_server, "_init_capacity_db", side_effect=check_mode):
            with self.assertRaises(StopAfterJournal):
                report_server.create_report_server_app(self.root, "report_html")

    def test_retention_error_does_not_fail_completed_reports(self):
        with closing(sqlite3.connect(self.db)) as conn, conn:
            conn.execute("UPDATE canonical_refresh_runs SET stats_json=? WHERE run_id='new'",
                         (json.dumps({"mode": "full", "replace_previous_data": False}),))
        with patch.object(report_server, "_canonical_prune_two_phase_runs",
                          side_effect=sqlite3.OperationalError("database or disk is full")):
            result, status = self.run_compute()
        self.assertEqual(status, 200)
        self.assertTrue(result["ok"])
        self.assertEqual(report_server._canonical_last_success_run_id(self.db), "new")
        with closing(sqlite3.connect(self.db)) as conn:
            stats = json.loads(conn.execute("SELECT stats_json FROM canonical_compute_runs").fetchone()[0])
        self.assertIn("disk is full", stats["retention_pruned"]["warning"])
        legacy_stats = json.loads(report_server._canonical_get_run(self.db, "new")["stats_json"])
        self.assertIn("History cleanup pending", legacy_stats["current_detail"])

    def test_completed_reports_wait_for_reader_longer_than_default_sqlite_timeout(self):
        locked = threading.Event()
        readers = []
        def hold_reader():
            with closing(sqlite3.connect(self.db)) as conn:
                conn.execute("BEGIN")
                conn.execute("SELECT value FROM manual_plans").fetchall()
                locked.set()
                time.sleep(6)
                conn.rollback()
        def after_reports(*args):
            reader = threading.Thread(target=hold_reader)
            readers.append(reader)
            reader.start()
            self.assertTrue(locked.wait(3))
        try:
            result, status = self.run_compute(sync_callback=after_reports)
        finally:
            for reader in readers:
                reader.join(10)
        self.assertEqual(status, 200)
        self.assertTrue(result["ok"])
        self.assertEqual(report_server._canonical_last_success_run_id(self.db), "new")
        compute = report_server._canonical_compute_get_run(self.db, self.compute_id)
        self.assertEqual(compute["status"], "success")

    def test_compaction_uses_live_pages_after_deleting_large_history(self):
        before = self.db.stat().st_size
        # Enough room for the retained data, less than twice the bloated file.
        with patch.object(canonical_storage.shutil, "disk_usage", return_value=SimpleNamespace(free=before)):
            result, status = self.run_compute()
        self.assertEqual(status, 200)
        with closing(sqlite3.connect(self.db)) as conn:
            stats = json.loads(conn.execute("SELECT stats_json FROM canonical_refresh_runs WHERE run_id='new'").fetchone()[0])
        self.assertEqual(stats["storage_cleanup"]["status"], "compacted")
        self.assertLess(self.db.stat().st_size, before / 2)

    def test_cleanup_keeps_other_connections_secure_delete_policy(self):
        with closing(sqlite3.connect(self.db)) as conn:
            original = conn.execute("PRAGMA secure_delete").fetchone()[0]
        _, status = self.run_compute()
        self.assertEqual(status, 200)
        with closing(sqlite3.connect(self.db)) as conn:
            self.assertEqual(conn.execute("PRAGMA secure_delete").fetchone()[0], original)
            self.assertEqual(conn.execute("PRAGMA integrity_check").fetchone()[0], "ok")

    def test_failed_later_batch_preserves_published_snapshot_and_can_resume(self):
        with closing(sqlite3.connect(self.db)) as conn, conn:
            conn.executemany("INSERT INTO canonical_issues(run_id, issue_key) VALUES ('old', ?)",
                             [(f"O2-{i}",) for i in range(2, 6002)])
            conn.execute("""CREATE TRIGGER fail_later_delete BEFORE DELETE ON canonical_issues
                WHEN old.issue_key='O2-5500' BEGIN SELECT RAISE(ABORT, 'simulated disk failure'); END""")
        report_server._canonical_mark_run_status(self.db, "new", "success", activate=True)
        with self.assertRaises(sqlite3.IntegrityError):
            canonical_storage.replace_previous_snapshots(self.db, "new", self.compute_id)
        with closing(sqlite3.connect(self.db)) as conn, conn:
            self.assertEqual(conn.execute("SELECT count(*) FROM canonical_issues WHERE run_id='old'").fetchone()[0], 1001)
            self.assertEqual(conn.execute("SELECT count(*) FROM canonical_issues WHERE run_id='new'").fetchone()[0], 1)
            conn.execute("DROP TRIGGER fail_later_delete")
        canonical_storage.replace_previous_snapshots(self.db, "new", self.compute_id)
        self.assertEqual(self.runs(), {"new"})

    def test_unpublished_snapshot_cannot_delete_anything(self):
        with self.assertRaises(ValueError):
            canonical_storage.replace_previous_snapshots(self.db, "new", self.compute_id)
        self.assertEqual(self.runs(), {"old", "new"})

    def test_cleanup_failure_is_visible_and_does_not_fail_reports(self):
        with patch.object(report_server, "replace_previous_snapshots", side_effect=sqlite3.OperationalError("locked")):
            _, status = self.run_compute()
        self.assertEqual(status, 200)
        self.assertEqual(self.runs(), {"old", "new"})
        stats = json.loads(report_server._canonical_get_run(self.db, "new")["stats_json"])
        self.assertIn("replacement pending", stats["current_detail"])

    def test_transient_cleanup_io_error_reopens_and_retries_without_refetch(self):
        calls = []
        def cleanup(*args):
            calls.append(args)
            if len(calls) == 1:
                raise sqlite3.OperationalError("disk I/O error")
            return canonical_storage.replace_previous_snapshots(*args)
        with patch.object(report_server, "replace_previous_snapshots", side_effect=cleanup):
            result, status = self.run_compute()
        self.assertEqual(status, 200)
        self.assertTrue(result["ok"])
        self.assertEqual(len(calls), 2)
        self.assertEqual(self.runs(), {"new"})
        stats = json.loads(report_server._canonical_get_run(self.db, "new")["stats_json"])
        self.assertEqual(stats["storage_cleanup"]["attempts"], 2)
        self.assertEqual(stats["storage_cleanup"]["status"], "compacted")

    def test_persistent_cleanup_io_error_stops_after_three_attempts(self):
        with patch.object(report_server, "replace_previous_snapshots",
                          side_effect=sqlite3.OperationalError("disk I/O error")) as cleanup:
            result, status = self.run_compute()
        self.assertEqual(status, 200)
        self.assertTrue(result["ok"])
        self.assertEqual(cleanup.call_count, 3)
        self.assertEqual(self.runs(), {"old", "new"})
        stats = json.loads(report_server._canonical_get_run(self.db, "new")["stats_json"])
        self.assertEqual(stats["storage_cleanup"]["attempts"], 3)
        self.assertIn("disk I/O error", stats["current_detail"])

    def test_another_active_fetch_prevents_deletion(self):
        report_server._canonical_mark_run_status(self.db, "new", "success", activate=True)
        with closing(sqlite3.connect(self.db)) as conn, conn:
            conn.execute("UPDATE canonical_fetch_runs SET status='running' WHERE fetch_run_id='old'")
            conn.execute("UPDATE canonical_refresh_state SET active_fetch_run_id='old' WHERE id=1")
        with self.assertRaises(ValueError):
            canonical_storage.replace_previous_snapshots(self.db, "new", self.compute_id)
        self.assertEqual(self.runs(), {"old", "new"})

    def test_abandoned_fetch_without_active_pointer_does_not_block_replacement(self):
        with closing(sqlite3.connect(self.db)) as conn, conn:
            conn.execute("UPDATE canonical_fetch_runs SET status='running' WHERE fetch_run_id='old'")
            conn.execute("UPDATE canonical_refresh_state SET active_fetch_run_id='' WHERE id=1")
        result, status = self.run_compute()
        self.assertEqual(status, 200)
        self.assertTrue(result["ok"])
        self.assertEqual(self.runs(), {"new"})

    def test_another_active_compute_prevents_deletion(self):
        other = report_server._canonical_create_compute_run(self.db, "old", "test")
        report_server._canonical_mark_run_status(self.db, "new", "success", activate=True)
        with self.assertRaises(ValueError):
            canonical_storage.replace_previous_snapshots(self.db, "new", self.compute_id)
        self.assertEqual(self.runs(), {"old", "new"})

    def test_api_option_is_explicit_and_applies_only_to_complete_full_refresh(self):
        from tests.test_canonical_refresh_api import _seed_managed_projects
        (self.root / "report_html").mkdir()
        with patch.dict("os.environ", {"JIRA_PROJECT_KEYS": ""}):
            app = report_server.create_report_server_app(base_dir=self.root, folder_raw="report_html")
        db = self.root / "assignee_hours_capacity.db"
        _seed_managed_projects(db)
        seen = []

        def fake_fetch(**kwargs):
            row = report_server._canonical_get_run(db, kwargs["run_id"])
            stats = json.loads(row["stats_json"])
            seen.append(stats["replace_previous_data"])
            report_server._canonical_mark_run_status(db, kwargs["run_id"], "failed", stats=stats)
            return {"ok": False}, 500

        with (
            patch.object(report_server, "_run_canonical_phase1_refresh", side_effect=fake_fetch),
            patch.object(report_server.threading, "Thread", side_effect=lambda target, daemon: SimpleNamespace(start=target)),
        ):
            client = app.test_client()
            for route, payload in (
                ("canonical-refresh", {"mode": "full", "replace_previous_data": True}),
                ("canonical-refresh", {"mode": "full"}),
                ("canonical-refresh", {"mode": "full", "replace_previous_data": False}),
                ("canonical-refresh", {"mode": "smart", "replace_previous_data": True}),
                ("canonical-fetch", {"mode": "full", "replace_previous_data": True}),
            ):
                self.assertEqual(client.post("/api/" + route, json={"year": 2026, **payload}).status_code, 202)
            html = client.get("/settings/canonical-refresh").get_data(as_text=True)
        self.assertEqual(seen, [True, False, False, False, False])
        self.assertIn('id="replace-previous-data" type="checkbox" checked', html)
        self.assertIn('mode === "full" && document.getElementById("replace-previous-data").checked', html)


if __name__ == "__main__":
    unittest.main()
