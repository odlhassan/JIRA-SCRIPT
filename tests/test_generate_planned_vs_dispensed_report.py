from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import generate_planned_vs_dispensed_report as script


class GeneratePlannedVsDispensedReportTests(unittest.TestCase):
    def test_writes_legacy_and_canonical_report_files(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            report_html_dir = root / "report_html"
            report_html_dir.mkdir(parents=True, exist_ok=True)
            source_path = report_html_dir / script.CANONICAL_OUTPUT
            source_path.write_text("<html><body>approved</body></html>", encoding="utf-8")

            with patch.object(script, "__file__", str(root / "generate_planned_vs_dispensed_report.py")):
                with patch.dict("os.environ", {"JIRA_PLANNED_VS_DISPENSED_HTML_PATH": str(root / script.LEGACY_OUTPUT)}, clear=False):
                    script.main()

            legacy_output = root / script.LEGACY_OUTPUT
            canonical_output = root / script.CANONICAL_OUTPUT
            self.assertTrue(legacy_output.exists())
            self.assertTrue(canonical_output.exists())
            self.assertEqual(legacy_output.read_text(encoding="utf-8"), "<html><body>approved</body></html>")
            self.assertEqual(canonical_output.read_text(encoding="utf-8"), "<html><body>approved</body></html>")


if __name__ == "__main__":
    unittest.main()
