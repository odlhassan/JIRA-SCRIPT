from __future__ import annotations

import re
import unittest
from pathlib import Path

from report_server import (
    REPORT_FILENAME_TO_ID,
    _build_report_info_catalog,
    _inject_refresh_ui,
)


class ReportInfoSyncTests(unittest.TestCase):
    def test_injected_html_contains_drawer_contract(self):
        repo_root = Path(__file__).resolve().parents[1]
        report_dir = repo_root / "report_html"
        for filename, report_id in REPORT_FILENAME_TO_ID.items():
            html_path = report_dir / filename
            if not html_path.exists():
                continue
            base_html = html_path.read_text(encoding="utf-8")
            injected = _inject_refresh_ui(base_html, report_id)
            self.assertIn("window.reportInfoCatalog", injected, msg=filename)
            self.assertIn('id="report-info-drawer"', injected, msg=filename)
            self.assertIn("data-info-id", injected, msg=filename)

    def test_docs_info_ids_match_catalog(self):
        repo_root = Path(__file__).resolve().parents[1]
        docs_dir = repo_root / "docs" / "report-user-guide" / "screens"
        self.assertTrue(docs_dir.exists())
        screen_files = sorted(docs_dir.glob("*.md"))
        self.assertGreaterEqual(len(screen_files), 11)

        for path in screen_files:
            text = path.read_text(encoding="utf-8")
            report_match = re.search(r"Report ID:\s*`([^`]+)`", text)
            ids_match = re.search(r"INFO_IDS:\s*(.+)", text)
            self.assertIsNotNone(report_match, msg=str(path))
            self.assertIsNotNone(ids_match, msg=str(path))
            report_id = report_match.group(1).strip()
            info_ids = {
                token.strip("` ").strip()
                for token in ids_match.group(1).split(",")
                if token.strip()
            }
            expected = {item["id"] for item in _build_report_info_catalog(report_id)}
            self.assertEqual(expected, info_ids, msg=str(path))


if __name__ == "__main__":
    unittest.main()
