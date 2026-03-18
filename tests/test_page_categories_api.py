from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from report_server import create_report_server_app


class PageCategoriesApiTests(unittest.TestCase):
    def _build_app(self, root: Path):
        (root / "report_html").mkdir(parents=True, exist_ok=True)
        (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
        wb = Workbook()
        ws = wb.active
        ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
        ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
        wb.save(root / "assignee_hours_report.xlsx")
        return create_report_server_app(base_dir=root, folder_raw="report_html")

    def test_page_categories_crud_bulk_and_nav_config(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = self._build_app(Path(td))
            client = app.test_client()

            initial = client.get("/api/page-categories")
            self.assertEqual(initial.status_code, 200)
            initial_body = initial.get_json()
            self.assertEqual(initial_body["categories"], [])
            self.assertFalse(initial_body["has_categories"])
            self.assertFalse(initial_body["navigation"]["enabled"])

            create_resp = client.post(
                "/api/page-categories/categories",
                json={"name": "Ops Reports", "icon_name": "analytics", "display_in_navigation": True, "display_order": 5, "is_active": True},
            )
            self.assertEqual(create_resp.status_code, 200)
            created = create_resp.get_json()["category"]
            self.assertEqual(created["name"], "Ops Reports")
            self.assertEqual(created["icon_name"], "analytics")

            cid = int(created["id"])
            update_resp = client.put(
                f"/api/page-categories/categories/{cid}",
                json={"name": "Operations", "icon_name": "monitoring", "display_in_navigation": True, "display_order": 3, "is_active": True},
            )
            self.assertEqual(update_resp.status_code, 200)
            self.assertEqual(update_resp.get_json()["category"]["name"], "Operations")
            self.assertEqual(update_resp.get_json()["category"]["icon_name"], "monitoring")

            bulk_resp = client.put(
                "/api/page-categories",
                json={
                    "categories": [
                        {"id": cid, "name": "Operations", "icon_name": "monitoring", "display_in_navigation": True, "display_order": 3, "is_active": True},
                        {"name": "Admin Core", "icon_name": "settings", "display_in_navigation": True, "display_order": 2, "is_active": True},
                    ],
                    "assignments": [
                        {"page_key": "dashboard", "page_type": "report", "category_id": cid},
                        {"page_key": "capacity_settings", "page_type": "configuration", "category_id": cid},
                    ],
                    "page_overrides": [
                        {"page_key": "dashboard", "display_name": "Executive Dashboard"},
                    ],
                },
            )
            self.assertEqual(bulk_resp.status_code, 200)
            body = bulk_resp.get_json()
            self.assertTrue(body["has_categories"])
            self.assertTrue(body["navigation"]["enabled"])
            self.assertGreaterEqual(len(body["categories"]), 2)
            self.assertTrue(any(item["page_key"] == "dashboard" for item in body["assignments"]))
            self.assertTrue(any(item["page_key"] == "page_categories" for item in body["page_catalog"]))
            self.assertTrue(any(item["page_key"] == "original_estimates_hierarchy_report" for item in body["page_catalog"]))
            self.assertTrue(
                any(
                    item["page_key"] == "original_estimates_hierarchy_report"
                    and item["title"] == "Epic Estimate Report"
                    for item in body["page_catalog"]
                )
            )
            self.assertTrue(
                any(
                    item["page_key"] == "dashboard"
                    and item["title"] == "Executive Dashboard"
                    and item["default_title"] == "Dashboard"
                    for item in body["page_catalog"]
                )
            )

            reports_categories = body["navigation"]["reports"]["categories"]
            self.assertTrue(any(group["name"] == "Operations" for group in reports_categories))
            self.assertTrue(any(group["icon_name"] == "monitoring" for group in reports_categories))
            self.assertTrue(
                any(
                    item["page_key"] == "dashboard" and item["title"] == "Executive Dashboard"
                    for group in reports_categories
                    for item in group["items"]
                )
            )
            admin_categories = body["navigation"]["admin_settings"]["categories"]
            self.assertTrue(any(group["name"] == "Operations" for group in admin_categories))

            dup_resp = client.post(
                "/api/page-categories/categories",
                json={"name": "operations", "display_in_navigation": True, "display_order": 7, "is_active": True},
            )
            self.assertEqual(dup_resp.status_code, 409)

            delete_resp = client.delete(f"/api/page-categories/categories/{cid}")
            self.assertEqual(delete_resp.status_code, 200)
            delete_body = delete_resp.get_json()
            self.assertTrue(delete_body["deleted"])
            self.assertFalse(any(item["category_id"] == cid for item in delete_body["assignments"]))

    def test_hidden_categories_removed_from_navigation_groups(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = self._build_app(Path(td))
            client = app.test_client()

            create_resp = client.post(
                "/api/page-categories/categories",
                json={"name": "Hidden Group", "icon_name": "visibility_off", "display_in_navigation": False, "display_order": 1, "is_active": True},
            )
            self.assertEqual(create_resp.status_code, 200)
            cid = int(create_resp.get_json()["category"]["id"])

            save_resp = client.put(
                "/api/page-categories",
                json={
                    "categories": [{"id": cid, "name": "Hidden Group", "display_in_navigation": False, "display_order": 1, "is_active": True}],
                    "assignments": [{"page_key": "dashboard", "page_type": "report", "category_id": cid}],
                },
            )
            self.assertEqual(save_resp.status_code, 200)
            body = save_resp.get_json()
            self.assertTrue(body["navigation"]["enabled"])
            self.assertEqual(body["navigation"]["reports"]["categories"], [])

    def test_report_display_names_round_trip_and_configuration_rename_rejected(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = self._build_app(Path(td))
            client = app.test_client()

            save_resp = client.put(
                "/api/page-categories",
                json={
                    "categories": [],
                    "assignments": [],
                    "page_overrides": [
                        {"page_key": "dashboard", "display_name": "Executive Dashboard"},
                        {"page_key": "capacity_settings", "display_name": "Capacity Admin"},
                    ],
                },
            )
            self.assertEqual(save_resp.status_code, 400)
            self.assertIn("only supported for report pages", save_resp.get_json()["error"])

            good_save = client.put(
                "/api/page-categories",
                json={
                    "categories": [],
                    "assignments": [],
                    "page_overrides": [
                        {"page_key": "dashboard", "display_name": "Executive Dashboard"},
                        {"page_key": "original_estimates_hierarchy_report", "display_name": "Epic Estimates"},
                    ],
                },
            )
            self.assertEqual(good_save.status_code, 200)
            body = good_save.get_json()
            self.assertFalse(body["navigation"]["enabled"])
            self.assertTrue(any(item["page_key"] == "dashboard" and item["display_name"] == "Executive Dashboard" for item in body["page_overrides"]))
            self.assertTrue(
                any(
                    item["page_key"] == "dashboard"
                    and item["title"] == "Executive Dashboard"
                    and item["default_title"] == "Dashboard"
                    and item["title_editable"] is True
                    for item in body["page_catalog"]
                )
            )
            self.assertTrue(
                any(
                    item["page_key"] == "capacity_settings"
                    and item["title"] == "Capacity Settings"
                    and item["display_name"] == ""
                    and item["title_editable"] is False
                    for item in body["page_catalog"]
                )
            )
            self.assertTrue(
                any(
                    item["page_key"] == "original_estimates_hierarchy_report"
                    and item["title"] == "Epic Estimates"
                    for item in body["page_catalog"]
                )
            )

            reload_resp = client.get("/api/page-categories")
            self.assertEqual(reload_resp.status_code, 200)
            reload_body = reload_resp.get_json()
            self.assertTrue(any(item["page_key"] == "dashboard" and item["display_name"] == "Executive Dashboard" for item in reload_body["page_overrides"]))
            self.assertTrue(
                any(
                    item["page_key"] == "dashboard"
                    and item["title"] == "Executive Dashboard"
                    and item["route_or_file"] == "dashboard.html"
                    for item in reload_body["page_catalog"]
                )
            )


if __name__ == "__main__":
    unittest.main()
