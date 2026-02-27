from __future__ import annotations

import sqlite3
import unittest

from jira_incremental_cache import (
    apply_overlap,
    get_changed_or_new_issue_keys,
    get_or_init_checkpoint,
    init_db,
    set_checkpoint,
    upsert_issue_index,
    upsert_worklog_payload,
    get_cached_worklogs_for_subtasks,
)


class IncrementalCacheTests(unittest.TestCase):
    def setUp(self):
        self.conn = sqlite3.connect(":memory:")
        init_db(self.conn)

    def tearDown(self):
        self.conn.close()

    def test_init_db_creates_tables(self):
        tables = {
            row[0]
            for row in self.conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        }
        self.assertIn("sync_state", tables)
        self.assertIn("issue_index", tables)
        self.assertIn("issue_payloads", tables)
        self.assertIn("worklog_payloads", tables)
        self.assertIn("pipeline_artifacts", tables)

    def test_changed_new_detection(self):
        upsert_issue_index(
            self.conn,
            [
                {
                    "issue_id": "10001",
                    "issue_key": "O2-1",
                    "updated_utc": "2026-01-01T00:00:00Z",
                    "issue_type": "Story",
                    "project_key": "O2",
                    "last_seen_utc": "2026-01-01T00:00:00Z",
                    "is_deleted": 0,
                }
            ],
        )
        changed = get_changed_or_new_issue_keys(
            self.conn,
            [
                {"issue_id": "10001", "issue_key": "O2-1", "updated_utc": "2026-01-01T00:00:00Z"},
                {"issue_id": "10001", "issue_key": "O2-1", "updated_utc": "2026-01-02T00:00:00Z"},
                {"issue_id": "10002", "issue_key": "O2-2", "updated_utc": "2026-01-01T00:00:00Z"},
            ],
        )
        self.assertIn("O2-1", changed)
        self.assertIn("O2-2", changed)
        self.assertEqual(len(changed), 2)

    def test_checkpoint_roundtrip(self):
        cp = get_or_init_checkpoint(self.conn, "work_items", "2025-01-01T00:00:00Z")
        self.assertEqual(cp, "2025-01-01T00:00:00Z")
        set_checkpoint(self.conn, "work_items", "2026-01-01T00:00:00Z")
        cp2 = get_or_init_checkpoint(self.conn, "work_items", "2020-01-01T00:00:00Z")
        self.assertEqual(cp2, "2026-01-01T00:00:00Z")

    def test_overlap_math(self):
        value = apply_overlap("2026-01-01T00:05:00Z", 5)
        self.assertEqual(value, "2026-01-01T00:00:00Z")

    def test_worklog_payload_roundtrip(self):
        upsert_worklog_payload(
            self.conn,
            issue_key="O2-10",
            issue_id="555",
            worklogs=[{"id": "a", "timeSpentSeconds": 3600}],
            worklog_updated_utc="2026-02-01T10:00:00Z",
        )
        data = get_cached_worklogs_for_subtasks(self.conn, ["O2-10"])
        self.assertIn("O2-10", data)
        self.assertEqual(len(data["O2-10"]), 1)
        self.assertEqual(data["O2-10"][0]["id"], "a")


if __name__ == "__main__":
    unittest.main()
