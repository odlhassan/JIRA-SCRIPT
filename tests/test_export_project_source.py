from __future__ import annotations

import os
import sqlite3
import tempfile
import unittest
from pathlib import Path

from export_jira_nested_view import _get_project_keys as nested_get_project_keys
from export_jira_subtask_worklogs import _get_project_keys as worklogs_get_project_keys
from export_jira_work_items import _get_project_keys as work_items_get_project_keys


class ExportProjectSourceTests(unittest.TestCase):
    def _create_managed_projects_db(self, db_path: Path, keys: list[str]) -> None:
        conn = sqlite3.connect(db_path)
        try:
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
            for key in keys:
                conn.execute(
                    """
                    INSERT INTO managed_projects (
                        project_key, project_name, display_name, color_hex, is_active, created_at_utc, updated_at_utc
                    ) VALUES (?, ?, ?, '#1D4ED8', 1, '2026-01-01 00:00:00', '2026-01-01 00:00:00')
                    """,
                    (key, key, key),
                )
            conn.commit()
        finally:
            conn.close()

    def test_db_first_then_env_fallback(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            old_db = os.environ.get("JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH")
            old_keys = os.environ.get("JIRA_PROJECT_KEYS")
            try:
                os.environ["JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH"] = str(db_path)
                os.environ["JIRA_PROJECT_KEYS"] = "ENV1,ENV2"

                self._create_managed_projects_db(db_path, ["O2", "FF"])

                keys, source = work_items_get_project_keys()
                self.assertEqual(keys, ["FF", "O2"])
                self.assertEqual(source, "managed_projects_db")

                keys, source = worklogs_get_project_keys()
                self.assertEqual(keys, ["FF", "O2"])
                self.assertEqual(source, "managed_projects_db")

                keys, source = nested_get_project_keys()
                self.assertEqual(keys, ["FF", "O2"])
                self.assertEqual(source, "managed_projects_db")
            finally:
                if old_db is None:
                    os.environ.pop("JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH", None)
                else:
                    os.environ["JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH"] = old_db
                if old_keys is None:
                    os.environ.pop("JIRA_PROJECT_KEYS", None)
                else:
                    os.environ["JIRA_PROJECT_KEYS"] = old_keys

    def test_env_fallback_when_db_missing(self):
        old_db = os.environ.get("JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH")
        old_keys = os.environ.get("JIRA_PROJECT_KEYS")
        try:
            os.environ["JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH"] = str(Path("Z:/no/such/managed_projects.db"))
            os.environ["JIRA_PROJECT_KEYS"] = "ENV1,ENV2"

            keys, source = work_items_get_project_keys()
            self.assertEqual(keys, ["ENV1", "ENV2"])
            self.assertEqual(source, "env_fallback")
        finally:
            if old_db is None:
                os.environ.pop("JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH", None)
            else:
                os.environ["JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH"] = old_db
            if old_keys is None:
                os.environ.pop("JIRA_PROJECT_KEYS", None)
            else:
                os.environ["JIRA_PROJECT_KEYS"] = old_keys


if __name__ == "__main__":
    unittest.main()
