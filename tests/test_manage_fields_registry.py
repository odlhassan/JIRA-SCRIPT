from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from manage_fields_registry import (
    create_manage_field,
    load_manage_fields,
    restore_manage_field,
    soft_delete_manage_field,
    update_manage_field,
)
from report_entity_registry import init_report_entities_db


class ManageFieldsRegistryTests(unittest.TestCase):
    def test_create_load_update_soft_delete_restore(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            init_report_entities_db(db)

            created = create_manage_field(
                db,
                {
                    "field_key": "delivery_gap",
                    "label": "Delivery Gap",
                    "description": "Planned minus actual.",
                    "data_type": "number",
                    "formula_expression": "planned_hours - actual_hours",
                    "formula_version": 1,
                    "formula_meta_json": {"editor": "hybrid"},
                    "is_active": True,
                },
            )
            self.assertEqual(created["field_key"], "delivery_gap")
            self.assertEqual(created["formula_meta_json"]["references"], ["actual_hours", "planned_hours"])

            active_rows = load_manage_fields(db, include_inactive=False)
            self.assertEqual(len(active_rows), 1)
            self.assertEqual(active_rows[0]["field_key"], "delivery_gap")

            updated = update_manage_field(
                db,
                "delivery_gap",
                {
                    "label": "Delivery Gap Updated",
                    "description": "Updated",
                    "data_type": "number",
                    "formula_expression": "sum(planned_hours) - actual_hours",
                    "formula_version": 2,
                    "formula_meta_json": {"via": "edit"},
                    "is_active": True,
                },
            )
            self.assertEqual(updated["label"], "Delivery Gap Updated")
            self.assertEqual(updated["formula_version"], 2)
            self.assertEqual(updated["formula_meta_json"]["via"], "edit")

            deleted = soft_delete_manage_field(db, "delivery_gap")
            self.assertFalse(deleted["is_active"])
            self.assertEqual(len(load_manage_fields(db, include_inactive=False)), 0)
            self.assertEqual(len(load_manage_fields(db, include_inactive=True)), 1)

            restored = restore_manage_field(db, "delivery_gap")
            self.assertTrue(restored["is_active"])
            self.assertEqual(len(load_manage_fields(db, include_inactive=False)), 1)

    def test_auto_key_generation_and_collision_suffix(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            init_report_entities_db(db)
            first = create_manage_field(
                db,
                {
                    "label": "Delivery Gap",
                    "description": "",
                    "data_type": "number",
                    "formula_expression": "planned_hours - actual_hours",
                    "formula_version": 1,
                    "formula_meta_json": {},
                    "is_active": True,
                },
            )
            second = create_manage_field(
                db,
                {
                    "label": "Delivery Gap",
                    "description": "",
                    "data_type": "number",
                    "formula_expression": "planned_hours - actual_hours",
                    "formula_version": 1,
                    "formula_meta_json": {},
                    "is_active": True,
                },
            )
            self.assertEqual(first["field_key"], "delivery_gap")
            self.assertEqual(second["field_key"], "delivery_gap_2")

    def test_client_supplied_key_is_ignored_on_create(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            init_report_entities_db(db)
            created = create_manage_field(
                db,
                {
                    "field_key": "manually_set_key",
                    "label": "Friendly Name",
                    "description": "",
                    "data_type": "number",
                    "formula_expression": "planned_hours",
                    "formula_version": 1,
                    "formula_meta_json": {},
                    "is_active": True,
                },
            )
            self.assertEqual(created["field_key"], "friendly_name")

    def test_formula_validation_unknown_entity_rejected(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            init_report_entities_db(db)
            with self.assertRaises(ValueError) as ctx:
                create_manage_field(
                    db,
                    {
                        "field_key": "bad_formula",
                        "label": "Bad Formula",
                        "description": "",
                        "data_type": "number",
                        "formula_expression": "sum(missing_entity)",
                        "formula_version": 1,
                        "formula_meta_json": {},
                        "is_active": True,
                    },
                )
            self.assertIn("Invalid formula_expression", str(ctx.exception))

    def test_update_missing_key_rejected(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db = Path(td) / "assignee_hours_capacity.db"
            init_report_entities_db(db)
            with self.assertRaises(LookupError) as ctx:
                update_manage_field(
                    db,
                    "unknown",
                    {
                        "label": "x",
                        "description": "",
                        "data_type": "number",
                        "formula_expression": "planned_hours",
                        "formula_version": 1,
                        "formula_meta_json": {},
                        "is_active": True,
                    },
                )
            self.assertIn("not found", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
