from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from report_server import create_report_server_app


class ManageFieldsApiTests(unittest.TestCase):
    def _build_app(self, root: Path):
        (root / "report_html").mkdir(parents=True, exist_ok=True)
        (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
        wb = Workbook()
        ws = wb.active
        ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
        ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
        wb.save(root / "assignee_hours_report.xlsx")
        return create_report_server_app(base_dir=root, folder_raw="report_html")

    def test_crud_and_soft_delete_restore(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = self._build_app(Path(td))
            client = app.test_client()

            create_resp = client.post(
                "/api/manage-fields",
                json={
                    "label": "Delivery Gap",
                    "description": "Planned minus actual.",
                    "data_type": "number",
                    "formula_expression": "planned_hours - actual_hours",
                    "formula_version": 1,
                    "formula_meta_json": {},
                    "is_active": True,
                },
            )
            self.assertEqual(create_resp.status_code, 200)
            created = create_resp.get_json()["field"]
            self.assertEqual(created["field_key"], "delivery_gap")

            get_resp = client.get("/api/manage-fields")
            self.assertEqual(get_resp.status_code, 200)
            body = get_resp.get_json()
            self.assertEqual(len(body["fields"]), 1)
            self.assertGreaterEqual(len(body["entity_catalog"]), 1)

            update_resp = client.put(
                "/api/manage-fields/delivery_gap",
                json={
                    "label": "Delivery Gap Updated",
                    "description": "Updated",
                    "data_type": "number",
                    "formula_expression": "sum(planned_hours) - actual_hours",
                    "formula_version": 2,
                    "formula_meta_json": {"via": "api"},
                    "is_active": True,
                },
            )
            self.assertEqual(update_resp.status_code, 200)
            updated = update_resp.get_json()["field"]
            self.assertEqual(updated["label"], "Delivery Gap Updated")
            self.assertEqual(updated["formula_version"], 2)

            delete_resp = client.delete("/api/manage-fields/delivery_gap")
            self.assertEqual(delete_resp.status_code, 200)
            deleted = delete_resp.get_json()["field"]
            self.assertFalse(deleted["is_active"])

            active_only_resp = client.get("/api/manage-fields")
            self.assertEqual(active_only_resp.status_code, 200)
            self.assertEqual(len(active_only_resp.get_json()["fields"]), 0)

            include_inactive_resp = client.get("/api/manage-fields?include_inactive=1")
            self.assertEqual(include_inactive_resp.status_code, 200)
            self.assertEqual(len(include_inactive_resp.get_json()["fields"]), 1)

            restore_resp = client.post("/api/manage-fields/delivery_gap/restore")
            self.assertEqual(restore_resp.status_code, 200)
            restored = restore_resp.get_json()["field"]
            self.assertTrue(restored["is_active"])

    def test_create_ignores_client_field_key(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = self._build_app(Path(td))
            client = app.test_client()
            resp = client.post(
                "/api/manage-fields",
                json={
                    "field_key": "dont_use_this",
                    "label": "Friendly KPI Field",
                    "description": "desc",
                    "data_type": "number",
                    "formula_expression": "planned_hours",
                    "formula_version": 1,
                    "formula_meta_json": {},
                    "is_active": True,
                },
            )
            self.assertEqual(resp.status_code, 200)
            body = resp.get_json()
            self.assertEqual(body["field"]["field_key"], "friendly_kpi_field")

    def test_validation_and_not_found_errors(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = self._build_app(Path(td))
            client = app.test_client()

            invalid_resp = client.post(
                "/api/manage-fields",
                json={
                    "field_key": "bad_formula",
                    "label": "Bad",
                    "description": "",
                    "data_type": "number",
                    "formula_expression": "sum(missing_entity)",
                    "formula_version": 1,
                    "formula_meta_json": {},
                    "is_active": True,
                },
            )
            self.assertEqual(invalid_resp.status_code, 400)
            self.assertIn("Invalid formula_expression", invalid_resp.get_json()["error"])

            missing_update = client.put(
                "/api/manage-fields/unknown",
                json={
                    "label": "x",
                    "description": "",
                    "data_type": "number",
                    "formula_expression": "planned_hours",
                    "formula_version": 1,
                    "formula_meta_json": {},
                    "is_active": True,
                },
            )
            self.assertEqual(missing_update.status_code, 404)

            missing_delete = client.delete("/api/manage-fields/unknown")
            self.assertEqual(missing_delete.status_code, 404)

            missing_restore = client.post("/api/manage-fields/unknown/restore")
            self.assertEqual(missing_restore.status_code, 404)

    def test_root_settings_links_include_manage_fields(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = self._build_app(Path(td))
            client = app.test_client()
            resp = client.get("/report_html/")
            self.assertEqual(resp.status_code, 200)
            html = resp.get_data(as_text=True)
            self.assertIn("Manage Fields", html)
            self.assertIn("/settings/manage-fields", html)


if __name__ == "__main__":
    unittest.main()
