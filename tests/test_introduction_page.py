from __future__ import annotations

import unittest
from pathlib import Path


class IntroductionPageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.introduction_html = (Path(__file__).resolve().parents[1] / "introduction.html").read_text(
            encoding="utf-8"
        )
        cls.introduction_docs = (
            Path(__file__).resolve().parents[1]
            / "docs"
            / "report-user-guide"
            / "screens"
            / "00-introduction-epr-tool.md"
        ).read_text(encoding="utf-8")

    def test_introduction_page_contains_architecture_sections(self) -> None:
        html = self.introduction_html
        self.assertIn("Detailed software architecture", html)
        self.assertIn("Detailed data-flow architecture", html)
        self.assertIn("Operational intake module", html)
        self.assertIn("Business context and interpretation module", html)
        self.assertIn("Canonical reporting memory", html)
        self.assertNotIn("fetch_jira_dashboard.py", html)
        self.assertNotIn("generate_assignee_hours_report.py", html)

    def test_introduction_page_bootstraps_mermaid_diagrams(self) -> None:
        html = self.introduction_html
        self.assertIn("cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js", html)
        self.assertIn("mermaid.initialize", html)
        self.assertIn("flowchart LR", html)
        self.assertIn("flowchart TD", html)
        self.assertIn("data-diagram-panel", html)
        self.assertIn("data-zoom-slider", html)
        self.assertIn("initializeDiagramZoom", html)
        self.assertIn("Full-width Mermaid rendering", html)

    def test_introduction_docs_describe_architecture_content(self) -> None:
        docs = self.introduction_docs
        self.assertIn("Detailed Software Architecture", docs)
        self.assertIn("Detailed Data-Flow Architecture", docs)
        self.assertIn("Mermaid-rendered diagrams", docs)
        self.assertIn("functional-module language", docs)


if __name__ == "__main__":
    unittest.main()
