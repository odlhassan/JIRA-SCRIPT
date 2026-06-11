from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from report_server import create_report_server_app


class ReportHtmlPromotionTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
