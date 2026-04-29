from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from report_server import (
    CAPACITY_SETTINGS_ROUTE,
    EPIC_PHASES_SETTINGS_ROUTE,
    EPICS_DROPDOWN_OPTIONS_SETTINGS_ROUTE,
    EXECUTIVE_DASHBOARD_SETTINGS_ROUTE,
    EPICS_MANAGEMENT_SETTINGS_ROUTE,
    MANAGE_FIELDS_SETTINGS_ROUTE,
    PAGE_CATEGORIES_SETTINGS_ROUTE,
    PERFORMANCE_SETTINGS_ROUTE,
    PROJECTS_SETTINGS_ROUTE,
    REPORT_ENTITIES_SETTINGS_ROUTE,
    SQL_CONSOLE_SETTINGS_ROUTE,
    create_report_server_app,
)


class AdminNavigationTests(unittest.TestCase):
    def _build_app(self, root: Path):
        (root / "report_html").mkdir(parents=True, exist_ok=True)
        (root / "introduction.html").write_text("<html><body>intro</body></html>", encoding="utf-8")
        (root / "report_html" / "introduction.html").write_text("<html><body>intro</body></html>", encoding="utf-8")
        (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
        (root / "executive_dashboard.html").write_text("<html><body>executive</body></html>", encoding="utf-8")
        wb = Workbook()
        ws = wb.active
        ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
        ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
        wb.save(root / "assignee_hours_report.xlsx")
        return create_report_server_app(base_dir=root, folder_raw="report_html")

    def test_settings_pages_use_shared_side_navigation_without_header_links(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()
            routes = [
                CAPACITY_SETTINGS_ROUTE,
                PERFORMANCE_SETTINGS_ROUTE,
                EXECUTIVE_DASHBOARD_SETTINGS_ROUTE,
                REPORT_ENTITIES_SETTINGS_ROUTE,
                MANAGE_FIELDS_SETTINGS_ROUTE,
                PROJECTS_SETTINGS_ROUTE,
                PAGE_CATEGORIES_SETTINGS_ROUTE,
                EPICS_DROPDOWN_OPTIONS_SETTINGS_ROUTE,
                EPIC_PHASES_SETTINGS_ROUTE,
                EPICS_MANAGEMENT_SETTINGS_ROUTE,
                SQL_CONSOLE_SETTINGS_ROUTE,
            ]
            for route in routes:
                with self.subTest(route=route):
                    resp = client.get(route)
                    self.assertEqual(resp.status_code, 200)
                    html = resp.get_data(as_text=True)
                    self.assertNotIn('href="/dashboard.html"', html)
                    self.assertNotIn('aria-current="page"', html)
                    self.assertIn('href="/shared-nav.css"', html)
                    self.assertIn('src="/shared-nav.js"', html)

    def test_report_html_lists_admin_and_reports_sections(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()
            resp = client.get("/report_html/")
            self.assertEqual(resp.status_code, 200)
            html = resp.get_data(as_text=True)
            self.assertIn("<h2>Admin Settings</h2>", html)
            self.assertIn("<h2>Reports</h2>", html)
            self.assertIn(f'href="{CAPACITY_SETTINGS_ROUTE}"', html)
            self.assertIn(f'href="{PERFORMANCE_SETTINGS_ROUTE}"', html)
            self.assertIn(f'href="{EXECUTIVE_DASHBOARD_SETTINGS_ROUTE}"', html)
            self.assertIn(f'href="{REPORT_ENTITIES_SETTINGS_ROUTE}"', html)
            self.assertIn(f'href="{MANAGE_FIELDS_SETTINGS_ROUTE}"', html)
            self.assertIn(f'href="{PROJECTS_SETTINGS_ROUTE}"', html)
            self.assertIn(f'href="{PAGE_CATEGORIES_SETTINGS_ROUTE}"', html)
            self.assertIn(f'href="{EPICS_DROPDOWN_OPTIONS_SETTINGS_ROUTE}"', html)
            self.assertIn(f'href="{EPIC_PHASES_SETTINGS_ROUTE}"', html)
            self.assertIn(f'href="{EPICS_MANAGEMENT_SETTINGS_ROUTE}"', html)
            self.assertIn(f'href="{SQL_CONSOLE_SETTINGS_ROUTE}"', html)
            self.assertIn('href="/dashboard.html"', html)
            self.assertIn('href="/introduction.html"', html)

    def test_index_redirects_to_introduction_when_available(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()
            resp = client.get("/", follow_redirects=False)
            self.assertEqual(resp.status_code, 302)
            self.assertEqual(resp.headers.get("Location"), "/introduction.html")

    def test_report_html_uses_effective_report_display_names(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = self._build_app(root)
            client = app.test_client()

            save_resp = client.put(
                "/api/page-categories",
                json={
                    "categories": [],
                    "assignments": [],
                    "page_overrides": [{"page_key": "dashboard", "display_name": "Executive Dashboard"}],
                },
            )
            self.assertEqual(save_resp.status_code, 200)

            resp = client.get("/report_html/")
            self.assertEqual(resp.status_code, 200)
            html = resp.get_data(as_text=True)
            self.assertIn('href="/dashboard.html">Executive Dashboard</a>', html)
            self.assertNotIn('href="/dashboard.html">dashboard.html</a>', html)

    def test_shared_nav_contract_contains_admin_section(self):
        nav_js = (Path(__file__).resolve().parents[1] / "shared-nav.js").read_text(encoding="utf-8")
        self.assertIn('title: "Reports"', nav_js)
        self.assertIn('page_key: "introduction"', nav_js)
        self.assertIn("introduction.html", nav_js)
        self.assertIn('unified-nav-title">EPR Tool', nav_js)
        self.assertIn("approved_vs_planned_hours_report.html", nav_js)
        self.assertIn("original_estimates_hierarchy_report.html", nav_js)
        self.assertIn('title: "Admin Settings"', nav_js)
        self.assertIn("applyCatalogTitles", nav_js)
        self.assertIn(CAPACITY_SETTINGS_ROUTE, nav_js)
        self.assertIn(PERFORMANCE_SETTINGS_ROUTE, nav_js)
        self.assertIn(EXECUTIVE_DASHBOARD_SETTINGS_ROUTE, nav_js)
        self.assertIn("executive_dashboard.html", nav_js)
        self.assertIn(REPORT_ENTITIES_SETTINGS_ROUTE, nav_js)
        self.assertIn(MANAGE_FIELDS_SETTINGS_ROUTE, nav_js)
        self.assertIn(PROJECTS_SETTINGS_ROUTE, nav_js)
        self.assertIn(PAGE_CATEGORIES_SETTINGS_ROUTE, nav_js)
        self.assertIn(EPICS_DROPDOWN_OPTIONS_SETTINGS_ROUTE, nav_js)
        self.assertIn(EPIC_PHASES_SETTINGS_ROUTE, nav_js)
        self.assertIn(EPICS_MANAGEMENT_SETTINGS_ROUTE, nav_js)
        self.assertIn(SQL_CONSOLE_SETTINGS_ROUTE, nav_js)


if __name__ == "__main__":
    unittest.main()
