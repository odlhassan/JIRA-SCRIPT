from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class CodexAgentSetupContractTests(unittest.TestCase):
    def _read(self, relative_path: str) -> str:
        return (ROOT / relative_path).read_text(encoding="utf-8")

    def test_required_agent_setup_artifacts_exist(self) -> None:
        required_paths = [
            ".cursor/rules/workspace-context.mdc",
            ".cursor/rules/html-and-associated-scripts.mdc",
            ".cursor/rules/experience-latest-changes.mdc",
            ".cursor/skills/localhost-ready-after-code/SKILL.md",
            ".cursor/skills/update-docs-with-code/SKILL.md",
            "AGENTS.md",
            "docs/codex-agent-gap-analysis.md",
            "docs/codex-task-contract.md",
            "docs/codex-agent-validation-report.md",
        ]
        for relative_path in required_paths:
            self.assertTrue((ROOT / relative_path).exists(), relative_path)

    def test_agents_requires_localhost_and_local_test_sections(self) -> None:
        text = self._read("AGENTS.md")
        self.assertIn("How to experience latest changes on live localhost", text)
        self.assertIn("How to test locally", text)
        self.assertIn("execute the narrowest relevant local CLI commands yourself", text)
        self.assertIn("exact CLI commands you executed", text)

    def test_workspace_context_captures_real_repo_defaults(self) -> None:
        text = self._read(".cursor/rules/workspace-context.mdc")
        for snippet in (
            "E:\\JIRA SCRIPT",
            "python run_server.py",
            "python run_html_only.py --no-server",
            "http://127.0.0.1:3000/introduction.html",
            "report_html/",
            "assignee_hours_capacity.db",
            "jira_sync_cache.db",
        ):
            self.assertIn(snippet, text)

    def test_html_rule_requires_coupled_updates(self) -> None:
        text = self._read(".cursor/rules/html-and-associated-scripts.mdc")
        for snippet in (
            "JavaScript",
            "CSS",
            "generator",
            "tests",
            "documentation",
            "_resolve_report_html_sources()",
            "report_html/",
        ):
            self.assertIn(snippet, text)

    def test_experience_rule_requires_exact_localhost_contract(self) -> None:
        text = self._read(".cursor/rules/experience-latest-changes.mdc")
        for snippet in (
            "How to experience latest changes on live localhost",
            "exact commands you executed",
            "exact additional commands the user can run locally",
            "python run_server.py",
            "expected visible behavior",
            "Not applicable for this change",
        ):
            self.assertIn(snippet, text)

    def test_project_local_skills_have_repo_specific_exit_criteria(self) -> None:
        localhost_skill = self._read(".cursor/skills/localhost-ready-after-code/SKILL.md")
        docs_skill = self._read(".cursor/skills/update-docs-with-code/SKILL.md")
        self.assertIn("## Exit Criteria", localhost_skill)
        self.assertIn("python run_server.py", localhost_skill)
        self.assertIn("python run_html_only.py --no-server", localhost_skill)
        self.assertIn("## Exit Criteria", docs_skill)
        self.assertIn("docs/codex-agent-gap-analysis.md", docs_skill)
        self.assertIn("docs/codex-task-contract.md", docs_skill)

    def test_task_contract_doc_contains_completion_requirements(self) -> None:
        text = self._read("docs/codex-task-contract.md")
        self.assertIn("Use project rules and skills strictly.", text)
        self.assertIn("How to experience latest changes on live localhost", text)
        self.assertIn("How to test locally", text)
        self.assertIn("python run_server.py", text)


if __name__ == "__main__":
    unittest.main()
