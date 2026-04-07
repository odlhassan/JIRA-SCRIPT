from __future__ import annotations

import unittest

import report_server


class ReportServerCanonicalFieldsTests(unittest.TestCase):
    def test_canonical_issue_fields_include_leave_type_field(self):
        fields = report_server._canonical_issue_fields(
            start_field_id="customfield_10133",
            end_field_ids=["duedate"],
            fix_type_field_id="customfield_10115",
        )
        self.assertIn("customfield_10584", fields)

