import unittest

from generate_employee_capacity_utilization_report import _html
from report_server import _base_page_catalog


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

    def test_report_is_available_to_page_categorization(self):
        pages = _base_page_catalog()
        self.assertTrue(any(page["page_key"] == "employee_capacity_utilization" for page in pages))
