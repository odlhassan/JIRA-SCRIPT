from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from report_server import create_report_server_app


class ReportHtmlPromotionTests(unittest.TestCase):
    def test_root_redirects_to_introduction_page(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            report_dir = root / "report_html"
            report_dir.mkdir(parents=True, exist_ok=True)
            (report_dir / "dashboard.html").write_text("<html><body>report index</body></html>", encoding="utf-8")
            (root / "introduction.html").write_text("<html><body>intro</body></html>", encoding="utf-8")

            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            with patch("report_server.tempfile.mkstemp", side_effect=OSError("read-only file system")):
                resp = client.get("/")

            self.assertEqual(resp.status_code, 302)
            self.assertEqual(resp.headers["Location"], "/introduction.html")

    def test_monthly_html_serves_existing_copy_when_live_promotion_is_read_only(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            report_dir = root / "report_html"
            report_dir.mkdir(parents=True, exist_ok=True)
            (report_dir / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            (root / "monthly_epic_plan_progress_report.html").write_text(
                "<html><body>source monthly</body></html>",
                encoding="utf-8",
            )
            (report_dir / "monthly_epic_plan_progress_report.html").write_text(
                "<html><body>served monthly</body></html>",
                encoding="utf-8",
            )

            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            with patch("report_server.tempfile.mkstemp", side_effect=OSError("read-only file system")):
                html_resp = client.get("/monthly_epic_plan_progress_report.html")

            self.assertEqual(html_resp.status_code, 200)
            self.assertIn("served monthly", html_resp.get_data(as_text=True))

    def test_monthly_html_serves_root_source_when_report_html_copy_is_missing_and_read_only(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            report_dir = root / "report_html"
            report_dir.mkdir(parents=True, exist_ok=True)
            (report_dir / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            (root / "monthly_epic_plan_progress_report.html").write_text(
                "<html><body>source monthly</body></html>",
                encoding="utf-8",
            )

            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            with patch("report_server.tempfile.mkstemp", side_effect=OSError("read-only file system")):
                html_resp = client.get("/monthly_epic_plan_progress_report.html")

            self.assertEqual(html_resp.status_code, 200)
            self.assertIn("source monthly", html_resp.get_data(as_text=True))

    def test_dashboard_html_serves_generated_fallback_when_package_paths_are_read_only(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            home = root / "home"
            (root / "report_html").mkdir(parents=True, exist_ok=True)

            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            with (
                patch.dict("os.environ", {"HOME": str(home)}, clear=False),
                patch("report_server.dash_fetch_dashboard_data", return_value={"epics": [], "stories": [], "subtasks": [], "bug_subtasks": [], "projects": [], "orphans": {}, "generated_at": "2026-06-11 09:00 UTC"}),
                patch("report_server.dash_generate_dashboard_html", return_value="<html><body>generated dashboard</body></html>"),
                patch("report_server.tempfile.mkstemp", side_effect=OSError("read-only file system")),
            ):
                html_resp = client.get("/dashboard.html")

            self.assertEqual(html_resp.status_code, 200)
            self.assertIn("generated dashboard", html_resp.get_data(as_text=True))
            self.assertTrue((home / "data" / "generated_reports" / "dashboard.html").exists())


if __name__ == "__main__":
    unittest.main()
