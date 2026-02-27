from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from report_server import create_report_server_app


class ProjectsApiTests(unittest.TestCase):
    def _build_app(self, root: Path):
        (root / "report_html").mkdir(parents=True, exist_ok=True)
        (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
        wb = Workbook()
        ws = wb.active
        ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
        ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
        wb.save(root / "assignee_hours_report.xlsx")
        return create_report_server_app(base_dir=root, folder_raw="report_html")

    def test_projects_crud_and_include_inactive(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = self._build_app(Path(td))
            client = app.test_client()

            create_resp = client.post(
                "/api/projects",
                json={
                    "project_key": "PTEST",
                    "project_name": "Project Test",
                    "display_name": "Project Test",
                    "color_hex": "#1D4ED8",
                },
            )
            self.assertEqual(create_resp.status_code, 200)

            list_resp = client.get("/api/projects")
            self.assertEqual(list_resp.status_code, 200)
            projects = list_resp.get_json()["projects"]
            self.assertTrue(any(item["project_key"] == "PTEST" for item in projects))

            update_resp = client.put(
                "/api/projects/PTEST",
                json={
                    "display_name": "Omni Connect Team",
                    "color_hex": "#334455",
                },
            )
            self.assertEqual(update_resp.status_code, 200)
            self.assertEqual(update_resp.get_json()["project"]["display_name"], "Omni Connect Team")

            delete_resp = client.delete("/api/projects/PTEST")
            self.assertEqual(delete_resp.status_code, 200)

            active_resp = client.get("/api/projects")
            self.assertEqual(active_resp.status_code, 200)
            self.assertFalse(any(item["project_key"] == "PTEST" for item in active_resp.get_json()["projects"]))

            inactive_resp = client.get("/api/projects?include_inactive=1")
            self.assertEqual(inactive_resp.status_code, 200)
            self.assertTrue(any(item["project_key"] == "PTEST" for item in inactive_resp.get_json()["projects"]))

            restore_resp = client.post("/api/projects/PTEST/restore")
            self.assertEqual(restore_resp.status_code, 200)
            self.assertTrue(restore_resp.get_json()["project"]["is_active"])

    @patch("report_server._jira_search_projects")
    def test_jira_projects_search_endpoint(self, mock_search):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            app = self._build_app(Path(td))
            client = app.test_client()

            mock_search.return_value = [
                {"project_key": "O2", "project_name": "OmniConnect"},
                {"project_key": "FF", "project_name": "Fintech Fuel"},
            ]
            resp = client.get("/api/jira/projects/search?q=O&limit=10")
            self.assertEqual(resp.status_code, 200)
            body = resp.get_json()
            self.assertEqual(len(body["projects"]), 2)
            self.assertEqual(body["projects"][0]["project_key"], "O2")

            mock_search.side_effect = RuntimeError("jira unavailable")
            err_resp = client.get("/api/jira/projects/search?q=O")
            self.assertEqual(err_resp.status_code, 502)
            self.assertIn("Failed to fetch Jira projects", err_resp.get_json()["error"])


if __name__ == "__main__":
    unittest.main()
