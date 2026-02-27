from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from managed_projects_registry import (
    create_managed_project,
    deterministic_color_for_project_key,
    list_active_project_keys,
    list_managed_projects,
    parse_project_keys_from_env,
    restore_managed_project,
    seed_managed_projects,
    soft_delete_managed_project,
    update_managed_project,
)


class ManagedProjectsRegistryTests(unittest.TestCase):
    def test_crud_soft_delete_restore(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "settings.db"
            created = create_managed_project(
                db_path,
                {
                    "project_key": "O2",
                    "project_name": "OmniConnect",
                    "display_name": "Omni Connect",
                    "color_hex": "#1D4ED8",
                },
            )
            self.assertEqual(created["project_key"], "O2")
            self.assertTrue(created["is_active"])

            updated = update_managed_project(
                db_path,
                "O2",
                {
                    "display_name": "Omni Connect Updated",
                    "color_hex": "#334455",
                },
            )
            self.assertEqual(updated["display_name"], "Omni Connect Updated")
            self.assertEqual(updated["color_hex"], "#334455")

            deleted = soft_delete_managed_project(db_path, "O2")
            self.assertFalse(deleted["is_active"])
            self.assertEqual(list_active_project_keys(db_path), [])

            restored = restore_managed_project(db_path, "O2")
            self.assertTrue(restored["is_active"])
            self.assertEqual(list_active_project_keys(db_path), ["O2"])

    def test_validation_errors(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "settings.db"
            with self.assertRaises(ValueError):
                create_managed_project(
                    db_path,
                    {
                        "project_key": "BAD KEY",
                        "project_name": "Bad",
                        "display_name": "Bad",
                        "color_hex": "#1D4ED8",
                    },
                )
            with self.assertRaises(ValueError):
                create_managed_project(
                    db_path,
                    {
                        "project_key": "O2",
                        "project_name": "OmniConnect",
                        "display_name": "",
                        "color_hex": "#1D4ED8",
                    },
                )
            with self.assertRaises(ValueError):
                create_managed_project(
                    db_path,
                    {
                        "project_key": "O2",
                        "project_name": "OmniConnect",
                        "display_name": "OmniConnect",
                        "color_hex": "blue",
                    },
                )

    def test_seed_inserts_missing_only_and_no_overwrite(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db_path = Path(td) / "settings.db"
            create_managed_project(
                db_path,
                {
                    "project_key": "O2",
                    "project_name": "Original Name",
                    "display_name": "Custom Display",
                    "color_hex": "#112233",
                },
            )

            stats = seed_managed_projects(
                db_path,
                ["O2", "FF"],
                project_name_resolver=lambda key: {"O2": "Resolved O2", "FF": "Fintech Fuel"}[key],
            )
            self.assertEqual(stats["inserted"], 1)
            self.assertEqual(stats["skipped_existing"], 1)

            rows = list_managed_projects(db_path, include_inactive=True)
            by_key = {row["project_key"]: row for row in rows}
            self.assertEqual(by_key["O2"]["project_name"], "Original Name")
            self.assertEqual(by_key["O2"]["display_name"], "Custom Display")
            self.assertEqual(by_key["O2"]["color_hex"], "#112233")
            self.assertEqual(by_key["FF"]["project_name"], "Fintech Fuel")
            self.assertEqual(by_key["FF"]["display_name"], "Fintech Fuel")
            self.assertEqual(by_key["FF"]["color_hex"], deterministic_color_for_project_key("FF"))

    def test_parse_env_keys(self):
        keys = parse_project_keys_from_env("O2, FF, O2, bad key, DIGITALLOG")
        self.assertEqual(keys, ["O2", "FF", "DIGITALLOG"])


if __name__ == "__main__":
    unittest.main()
