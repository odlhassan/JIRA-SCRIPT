from __future__ import annotations

import unittest
from pathlib import Path


class SharedDateFilterTests(unittest.TestCase):
    def test_phase_rmi_gantt_is_excluded_from_auto_apply(self):
        script = (Path(__file__).resolve().parents[1] / "report_html" / "shared-date-filter.js").read_text(encoding="utf-8")
        self.assertIn('var AUTO_APPLY_EXCLUDED_PAGES = {', script)
        self.assertIn('"phase_rmi_gantt_report": true', script)
        self.assertIn('if (AUTO_APPLY_EXCLUDED_PAGES[currentPageKey()]) return;', script)


if __name__ == "__main__":
    unittest.main()
