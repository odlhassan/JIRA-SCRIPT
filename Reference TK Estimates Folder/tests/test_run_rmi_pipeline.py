import unittest
from pathlib import Path
from unittest.mock import ANY, patch

import run_rmi_pipeline as pipeline


class RunRmiPipelineTests(unittest.TestCase):
    def test_run_pipeline_executes_extraction_then_html_generation(self):
        workbook_path = Path("source.xlsx")
        db_path = Path("output.db")
        html_path = Path("report.html")

        with (
            patch.object(pipeline, "load_env_config", return_value={"ENV_PATH": "x.env", "JIRA_SITE": "a", "JIRA_EMAIL": "b", "JIRA_API_TOKEN": "c"}) as load_env,
            patch.object(
                pipeline,
                "run_extraction",
                return_value={
                    "eligible_rows": 5,
                    "epics_fetched": 4,
                    "stories_fetched": 9,
                    "descendants_fetched": 12,
                    "worklogs_fetched": 30,
                    "errors": 1,
                },
            ) as run_extraction,
            patch.object(pipeline, "generate_html_report", return_value=html_path) as generate_html,
        ):
            result = pipeline.run_pipeline(
                workbook_path=workbook_path,
                db_path=db_path,
                html_path=html_path,
                sheet_contains="RMI",
                limit=10,
            )

        load_env.assert_called_once_with()
        run_extraction.assert_called_once_with(
            workbook_path=workbook_path,
            db_path=db_path,
            env_config={"ENV_PATH": "x.env", "JIRA_SITE": "a", "JIRA_EMAIL": "b", "JIRA_API_TOKEN": "c"},
            sheet_contains="RMI",
            limit=10,
            progress_callback=ANY,
        )
        generate_html.assert_called_once_with(db_path, html_path)
        self.assertEqual(result["env_path"], "x.env")
        self.assertEqual(result["db_path"], db_path)
        self.assertEqual(result["html_path"], html_path)
        self.assertEqual(result["summary"]["eligible_rows"], 5)
        self.assertTrue(callable(run_extraction.call_args.kwargs["progress_callback"]))


if __name__ == "__main__":
    unittest.main()
