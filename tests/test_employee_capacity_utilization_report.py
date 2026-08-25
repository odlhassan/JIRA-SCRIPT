import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from generate_employee_capacity_utilization_report import _html
from report_server import REPORT_IDS_WITHOUT_REFRESH_WIDGET, _base_page_catalog, create_report_server_app


class EmployeeCapacityUtilizationReportTests(unittest.TestCase):
    def test_report_has_standalone_filters_and_canonical_payload(self):
        html = _html({"issues": [], "worklogs": [], "leaves": [], "teams": [], "profiles": [], "support": [], "resources": {}})
        self.assertIn("Employee Capacity &amp; Utilization", html)
        self.assertIn('id="month"', html)
        self.assertIn('id="profile"', html)
        self.assertIn('id="teams"', html)
        self.assertIn('id="resigned"', html)
        self.assertIn('id="support"', html)
        self.assertIn("process team", html)
        self.assertIn("Grand Total", html)
        self.assertIn("Canonical Jira worklogs", html)
        self.assertIn("/api/employee-capacity-utilization/data", html)
        self.assertIn("filters apply automatically", html)
        self.assertNotIn(">Refresh<", html)
        self.assertIn("All except Process Team", html)
        self.assertIn("#teams .team-option input[type=checkbox]", html)
        self.assertIn("width:17px!important", html)
        self.assertIn("position:sticky", html)
        self.assertIn("event.key==='Escape'", html)
        self.assertIn('id="util-a"', html)
        self.assertIn('id="util-b"', html)
        self.assertIn('id="util-c"', html)
        self.assertIn('name="color-scope"', html)
        self.assertIn("employeeCapacityUtilization.colorRules.v1", html)
        self.assertIn("utilizationBand", html)
        self.assertIn("Enter increasing values", html)
        self.assertIn("ecu-drawer", html)
        self.assertIn("ecu-resize", html)
        self.assertIn("pointermove", html)
        self.assertIn("wireDetailCells", html)
        self.assertIn("Work item title", html)
        self.assertIn("Hours logged", html)
        self.assertIn("jira_browse_base", html)

    def test_report_is_available_to_page_categorization(self):
        pages = _base_page_catalog()
        self.assertTrue(any(page["page_key"] == "employee_capacity_utilization" for page in pages))

    def test_report_has_no_injected_refresh_widget(self):
        self.assertIn("employee_capacity_utilization", REPORT_IDS_WITHOUT_REFRESH_WIDGET)

    def test_live_api_returns_runtime_canonical_payload(self):
        payload = {
            "issues": [], "worklogs": [{"worklog_author": "Alice", "hours_logged": 7.5}],
            "leaves": [], "teams": [], "profiles": [], "support": [], "resources": {},
            "canonical_run_id": "production-run", "source": "canonical_database",
        }
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir()
            with patch("report_server.build_employee_capacity_utilization_payload", return_value=payload) as builder:
                client = create_report_server_app(root, "report_html").test_client()
                response = client.get("/api/employee-capacity-utilization/data")
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.get_json()["worklogs"][0]["hours_logged"], 7.5)
            self.assertEqual(response.get_json()["canonical_run_id"], "production-run")
            builder.assert_called_once_with(root / "assignee_hours_capacity.db")
