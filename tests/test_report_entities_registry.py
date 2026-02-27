from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from report_entity_registry import load_report_entities
from report_server import create_report_server_app


class ReportEntityRegistryTests(unittest.TestCase):
    def _build_app(self, root: Path):
        (root / "report_html").mkdir(parents=True, exist_ok=True)
        (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
        wb = Workbook()
        ws = wb.active
        ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
        ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
        wb.save(root / "assignee_hours_report.xlsx")
        return create_report_server_app(base_dir=root, folder_raw="report_html")

    def test_seed_and_get_api(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()
            resp = client.get("/api/report-entities")
            self.assertEqual(resp.status_code, 200)
            body = resp.get_json()
            self.assertEqual(body["source"], "db")
            self.assertEqual(len(body["entities"]), 17)
            keys = {item["entity_key"] for item in body["entities"]}
            self.assertIn("capacity", keys)
            self.assertIn("planned_rmi", keys)
            self.assertNotIn("dispensed_rmi", keys)
            self.assertNotIn("rmi_dispensing_progress", keys)
            self.assertIn("activity", keys)
            self.assertEqual(body["global_settings"]["leave_taken_identification_mode"], "hours")
            self.assertEqual(body["global_settings"]["planned_actual_equality_tolerance_hours"], 0.0)
            capacity = next(x for x in body["entities"] if x["entity_key"] == "capacity")
            self.assertIn("formula_expression", capacity)
            self.assertIn("formula_version", capacity)
            self.assertIn("formula_meta_json", capacity)
            self.assertEqual(capacity["formula_expression"], "")
            self.assertEqual(capacity["formula_version"], 1)
            self.assertIsInstance(capacity["formula_meta_json"], dict)

    def test_put_updates_entities_and_settings(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()
            base = client.get("/api/report-entities").get_json()
            entities = base["entities"]
            for item in entities:
                if item["entity_key"] == "capacity":
                    item["label"] = "Capacity Updated"
                    item["formula_expression"] = "sum(planned_hours) - actual_hours"
                    item["formula_version"] = 2
                    item["formula_meta_json"] = {"editor": "hybrid"}
            payload = {
                "entities": entities,
                "global_settings": {
                    "planned_leave_min_notice_days": 7,
                    "planned_leave_rule_apply_from_date": "2026-01-01",
                    "leave_taken_identification_mode": "status",
                    "leave_taken_rule_apply_from_date": "2026-01-05",
                    "rmi_planned_field_resolution": "hybrid",
                    "planned_actual_equality_tolerance_hours": 0.5,
                },
            }
            put_resp = client.put("/api/report-entities", json=payload)
            self.assertEqual(put_resp.status_code, 200)
            body = put_resp.get_json()
            capacity = next(x for x in body["entities"] if x["entity_key"] == "capacity")
            self.assertEqual(capacity["label"], "Capacity Updated")
            self.assertEqual(capacity["formula_expression"], "sum(planned_hours) - actual_hours")
            self.assertEqual(capacity["formula_version"], 2)
            self.assertEqual(capacity["formula_meta_json"]["editor"], "hybrid")
            self.assertEqual(capacity["formula_meta_json"]["references"], ["actual_hours", "planned_hours"])
            self.assertEqual(body["global_settings"]["planned_leave_min_notice_days"], 7)
            self.assertEqual(body["global_settings"]["leave_taken_identification_mode"], "status")
            self.assertEqual(body["global_settings"]["planned_actual_equality_tolerance_hours"], 0.5)

    def test_put_global_settings_rejects_negative_tolerance(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()
            base = client.get("/api/report-entities").get_json()
            payload = {
                "entities": base["entities"],
                "global_settings": {
                    **base["global_settings"],
                    "planned_actual_equality_tolerance_hours": -0.1,
                },
            }
            resp = client.put("/api/report-entities", json=payload)
            self.assertEqual(resp.status_code, 400)
            body = resp.get_json()
            self.assertIn("planned_actual_equality_tolerance_hours", body["error"])

    def test_put_validation_error(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()
            base = client.get("/api/report-entities").get_json()
            entities = base["entities"]
            entities[0]["selection_rule_json"] = "not-json-object-or-array"
            resp = client.put("/api/report-entities", json={"entities": entities})
            self.assertEqual(resp.status_code, 400)
            body = resp.get_json()
            self.assertIn("selection_rule_json", body["error"])

    def test_put_formula_validation_error_unknown_entity(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()
            base = client.get("/api/report-entities").get_json()
            entities = base["entities"]
            for item in entities:
                if item["entity_key"] == "capacity":
                    item["formula_expression"] = "sum(unknown_entity)"
            resp = client.put("/api/report-entities", json={"entities": entities})
            self.assertEqual(resp.status_code, 400)
            body = resp.get_json()
            self.assertIn("Invalid formula_expression", body["error"])
            self.assertIn("capacity", body["error"])

    def test_put_formula_validation_error_self_reference(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()
            base = client.get("/api/report-entities").get_json()
            entities = base["entities"]
            for item in entities:
                if item["entity_key"] == "capacity":
                    item["formula_expression"] = "capacity + planned_hours"
            resp = client.put("/api/report-entities", json={"entities": entities})
            self.assertEqual(resp.status_code, 400)
            body = resp.get_json()
            self.assertIn("Self reference", body["error"])

    def test_reset_restores_defaults(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()
            base = client.get("/api/report-entities").get_json()
            entities = base["entities"]
            for item in entities:
                if item["entity_key"] == "capacity":
                    item["label"] = "Temporary Label"
            client.put("/api/report-entities", json={"entities": entities})
            reset_resp = client.post("/api/report-entities/reset")
            self.assertEqual(reset_resp.status_code, 200)
            body = reset_resp.get_json()
            capacity = next(x for x in body["entities"] if x["entity_key"] == "capacity")
            self.assertEqual(capacity["label"], "Capacity")
            self.assertEqual(len(body["entities"]), 17)

    def test_settings_route_html(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()
            resp = client.get("/settings/report-entities")
            self.assertEqual(resp.status_code, 200)
            html = resp.get_data(as_text=True)
            self.assertIn("Report Entity Registry", html)
            self.assertIn('id="entity-tbody"', html)
            self.assertIn('id="planned-leave-n"', html)
            self.assertIn('id="planned-actual-tolerance"', html)
            self.assertIn('id="e-formula-expression"', html)
            self.assertIn('id="formula-suggestions"', html)
            self.assertIn('id="formula-validation"', html)
            self.assertIn('id="formula-quick-insert"', html)
            self.assertIn("/api/report-entities", html)

    def test_seed_once_exact_count(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            _ = app
            db = root / "assignee_hours_capacity.db"
            first = load_report_entities(db)
            second = load_report_entities(db)
            self.assertEqual(len(first), 17)
            self.assertEqual(len(second), 17)


if __name__ == "__main__":
    unittest.main()
