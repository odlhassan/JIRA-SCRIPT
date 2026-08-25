import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from generate_employee_capacity_utilization_report import _epic_for_issue, _html
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
        self.assertIn("Download Excel", html)
        self.assertIn("/api/employee-capacity-utilization/export", html)
        self.assertIn("metric==='employee'?'logged':metric", html)
        self.assertIn("'Epic name'", html)

    def test_epic_is_resolved_through_parent_hierarchy(self):
        items = {
            "SUB-1": {"issue_type": "Sub-task", "parent_issue_key": "STORY-1"},
            "STORY-1": {"issue_type": "Story", "parent_issue_key": "EPIC-1"},
            "EPIC-1": {"issue_type": "Epic", "summary": "Customer onboarding"},
        }
        self.assertEqual(_epic_for_issue("SUB-1", items), ("EPIC-1", "Customer onboarding"))

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

    def test_excel_export_contains_structured_sheets_and_epic_columns(self):
        payload = {
            "issues": [
                {"issue_key": "EPIC-1", "issue_type": "Epic", "summary": "Customer onboarding", "assignee": "Alice", "parent_issue_key": ""},
                {"issue_key": "STORY-1", "issue_type": "Story", "summary": "Portal", "assignee": "Alice", "parent_issue_key": "EPIC-1"},
                {"issue_key": "SUB-1", "issue_type": "Sub-task", "summary": "Build form", "assignee": "Alice", "parent_issue_key": "STORY-1", "start_date": "2026-08-01", "due_date": "2026-08-31", "original_estimate_hours": 16},
            ],
            "worklogs": [{"issue_id": "SUB-1", "worklog_author": "Alice", "issue_assignee": "Alice", "item_assignee": "Alice", "item_issue_type": "Sub-task", "item_summary": "Build form", "worklog_date": "2026-08-10", "hours_logged": 6, "epic_key": "EPIC-1", "epic_summary": "Customer onboarding"}],
            "leaves": [], "teams": [], "profiles": [], "support": [], "resources": {"Alice": {"resigned": False}},
            "canonical_run_id": "production-run", "generated_at": "2026-08-25T00:00:00Z", "jira_browse_base": "https://jira.example/browse",
        }
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir()
            with patch("report_server.build_employee_capacity_utilization_payload", return_value=payload):
                client = create_report_server_app(root, "report_html").test_client()
                response = client.post("/api/employee-capacity-utilization/export", json={"month": "2026-08", "scope": "any", "selected_teams": [], "display_resigned": False})
        self.assertEqual(response.status_code, 200)
        workbook = load_workbook(BytesIO(response.data), read_only=False, data_only=False)
        self.assertEqual(workbook.sheetnames, ["Export Info", "Summary", "Worklogs", "Booked Subtasks", "Leave Records", "Capacity Calendar", "Employees"])
        worklog_headers = [cell.value for cell in workbook["Worklogs"][1]]
        self.assertIn("Epic Name", worklog_headers)
        self.assertIn("Epic Jira Link", worklog_headers)
        self.assertEqual(workbook["Worklogs"]["F2"].value, "Customer onboarding")
        self.assertEqual(workbook["Worklogs"]["G2"].hyperlink.target, "https://jira.example/browse/EPIC-1")
