from __future__ import annotations

import subprocess
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent


class AzureDeployContractTests(unittest.TestCase):
    def test_project_image_runtime_dependency_is_vendored_for_production(self) -> None:
        requirements = (ROOT / "requirements.txt").read_text(encoding="utf-8")
        requirement_names = {
            line.split(";", 1)[0]
            .strip()
            .split("[", 1)[0]
            .split("=", 1)[0]
            .split(">", 1)[0]
            .split("<", 1)[0]
            .lower()
            for line in requirements.splitlines()
            if line.strip() and not line.lstrip().startswith("#")
        }
        self.assertIn(
            "pillow",
            requirement_names,
            "Pillow must be installed because project_image_registry imports PIL during uploads",
        )

        workflow = (
            ROOT / ".github" / "workflows" / "azure-appservice-deploy.yml"
        ).read_text(encoding="utf-8")
        self.assertIn(
            ".python_packages/lib/site-packages/PIL",
            workflow,
            "The deploy workflow must verify that Pillow was vendored into the production ZIP",
        )

    def test_startup_journal_helper_is_deployable_source(self) -> None:
        helper = ROOT / "db_journal_mode.py"
        self.assertTrue(helper.exists(), "db_journal_mode.py must ship with report_server.py")

        result = subprocess.run(
            ["git", "check-ignore", "-q", "db_journal_mode.py"],
            cwd=ROOT,
            check=False,
        )
        self.assertNotEqual(result.returncode, 0, "db_journal_mode.py must not be gitignored")

    def test_azure_workflow_checks_startup_dependencies_before_deploy(self) -> None:
        workflow = ROOT / ".github" / "workflows" / "azure-appservice-deploy.yml"
        text = workflow.read_text(encoding="utf-8")

        self.assertIn("Verify required startup files", text)
        self.assertIn("db_journal_mode.py", text)
        self.assertLess(
            text.index("Verify required startup files"),
            text.index("Deploy to Azure Web App"),
        )


if __name__ == "__main__":
    unittest.main()
