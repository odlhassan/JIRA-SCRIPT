from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

from generate_nested_view_html import _load_nested_rows_from_canonical_db


class NestedViewDbLoaderTests(unittest.TestCase):
    def test_loads_nested_rows_directly_from_canonical_db(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            db_path = root / "assignee_hours_capacity.db"
            conn = sqlite3.connect(db_path)
            try:
                conn.execute(
                    """
                    CREATE TABLE canonical_refresh_state (
                        id INTEGER PRIMARY KEY CHECK(id = 1),
                        last_success_run_id TEXT NOT NULL DEFAULT ''
                    )
                    """
                )
                conn.execute(
                    """
                    INSERT INTO canonical_refresh_state (id, last_success_run_id)
                    VALUES (1, 'run-1')
                    """
                )
                conn.execute(
                    """
                    CREATE TABLE canonical_issues (
                        run_id TEXT NOT NULL,
                        issue_key TEXT NOT NULL,
                        project_key TEXT NOT NULL DEFAULT '',
                        issue_type TEXT NOT NULL DEFAULT '',
                        summary TEXT NOT NULL DEFAULT '',
                        assignee TEXT NOT NULL DEFAULT '',
                        start_date TEXT NOT NULL DEFAULT '',
                        due_date TEXT NOT NULL DEFAULT '',
                        original_estimate_hours REAL NOT NULL DEFAULT 0,
                        total_hours_logged REAL NOT NULL DEFAULT 0,
                        parent_issue_key TEXT NOT NULL DEFAULT '',
                        story_key TEXT NOT NULL DEFAULT '',
                        epic_key TEXT NOT NULL DEFAULT ''
                    )
                    """
                )
                conn.execute(
                    """
                    CREATE TABLE canonical_issue_actuals (
                        run_id TEXT NOT NULL,
                        issue_key TEXT NOT NULL,
                        total_worklog_hours REAL NOT NULL DEFAULT 0
                    )
                    """
                )
                conn.execute(
                    """
                    CREATE TABLE managed_projects (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        project_key TEXT NOT NULL UNIQUE,
                        project_name TEXT NOT NULL,
                        display_name TEXT NOT NULL,
                        color_hex TEXT NOT NULL,
                        is_active INTEGER NOT NULL DEFAULT 1,
                        created_at_utc TEXT NOT NULL,
                        updated_at_utc TEXT NOT NULL
                    )
                    """
                )
                conn.execute(
                    """
                    CREATE TABLE epics_management (
                        epic_key TEXT PRIMARY KEY,
                        project_key TEXT NOT NULL,
                        project_name TEXT NOT NULL,
                        product_category TEXT NOT NULL,
                        component TEXT NOT NULL DEFAULT '',
                        epic_name TEXT NOT NULL
                    )
                    """
                )
                conn.execute(
                    """
                    INSERT INTO managed_projects (
                        project_key, project_name, display_name, color_hex, is_active, created_at_utc, updated_at_utc
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    ("O2", "OmniConnect", "OmniConnect", "#336699", 1, "2026-03-11 00:00:00", "2026-03-11 00:00:00"),
                )
                conn.execute(
                    """
                    INSERT INTO epics_management (
                        epic_key, project_key, project_name, product_category, component, epic_name
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    ("O2-100", "O2", "OmniConnect", "Platform", "", "Epic Alpha"),
                )
                conn.executemany(
                    """
                    INSERT INTO canonical_issues (
                        run_id, issue_key, project_key, issue_type, summary, assignee,
                        start_date, due_date, original_estimate_hours, total_hours_logged,
                        parent_issue_key, story_key, epic_key
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    [
                        ("run-1", "O2-100", "O2", "Epic", "Epic Alpha", "Alice", "2026-02-01", "2026-02-20", 16.0, 2.0, "", "", ""),
                        ("run-1", "O2-200", "O2", "Story", "Story Alpha", "Alice", "2026-02-02", "2026-02-10", 8.0, 3.0, "O2-100", "", "O2-100"),
                        ("run-1", "O2-300", "O2", "Sub-task", "Subtask Alpha", "Alice", "2026-02-03", "2026-02-04", 4.0, 1.0, "O2-200", "O2-200", "O2-100"),
                        ("run-1", "O2-201", "O2", "Story", "Orphan Story", "Bob", "2026-02-05", "2026-02-12", 6.0, 0.0, "", "", ""),
                    ],
                )
                conn.execute(
                    """
                    INSERT INTO canonical_issue_actuals (run_id, issue_key, total_worklog_hours)
                    VALUES (?, ?, ?)
                    """,
                    ("run-1", "O2-300", 5.5),
                )
                conn.commit()
            finally:
                conn.close()

            rows = _load_nested_rows_from_canonical_db(db_path, run_id="run-1")

        aspects = [str(row.get("aspect")) for row in rows]
        self.assertIn("O2 - OmniConnect", aspects)
        self.assertIn("Platform", aspects)
        self.assertIn("Epic Alpha", aspects)
        self.assertIn("Story Alpha", aspects)
        self.assertIn("Subtask Alpha", aspects)
        self.assertIn("Alice", aspects)
        self.assertIn("No RMI", aspects)
        self.assertIn("Orphan Story", aspects)

        project_row = next(row for row in rows if row.get("aspect") == "O2 - OmniConnect")
        epic_row = next(row for row in rows if row.get("aspect") == "Epic Alpha")
        story_row = next(row for row in rows if row.get("aspect") == "Story Alpha")
        subtask_row = next(row for row in rows if row.get("aspect") == "Subtask Alpha")
        orphan_story_row = next(row for row in rows if row.get("aspect") == "Orphan Story")
        assignee_row = next(row for row in rows if row.get("aspect") == "Alice")

        self.assertEqual(epic_row.get("row_type"), "rmi")
        self.assertEqual(story_row.get("row_type"), "story")
        self.assertEqual(subtask_row.get("row_type"), "subtask")
        self.assertEqual(epic_row.get("jira_key"), "O2-100")
        self.assertEqual(story_row.get("jira_key"), "O2-200")
        self.assertEqual(subtask_row.get("jira_key"), "O2-300")
        self.assertTrue(str(subtask_row.get("jira_url")).endswith("/browse/O2-300"))
        self.assertEqual(subtask_row.get("actual_hours"), 5.5)
        self.assertEqual(project_row.get("approved_hours"), 16.0)
        self.assertEqual(project_row.get("planned_hours"), 4.0)
        self.assertEqual(epic_row.get("approved_hours"), 16.0)
        self.assertEqual(epic_row.get("planned_hours"), 4.0)
        self.assertEqual(story_row.get("approved_hours"), 8.0)
        self.assertEqual(story_row.get("planned_hours"), 4.0)
        self.assertEqual(subtask_row.get("approved_hours"), 4.0)
        self.assertEqual(subtask_row.get("planned_hours"), 4.0)
        self.assertEqual(orphan_story_row.get("missing_parent_reason"), "missing_rmi_parent")
        self.assertTrue(bool(orphan_story_row.get("is_missing_parent")))
        self.assertEqual(assignee_row.get("row_type"), "assignee")
        self.assertEqual(assignee_row.get("jira_key"), "")


if __name__ == "__main__":
    unittest.main()
