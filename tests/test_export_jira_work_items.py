from __future__ import annotations

import unittest

from export_jira_work_items import _payload_has_required_detail_fields, _stable_resolved_since


class ExportJiraWorkItemsTests(unittest.TestCase):
    def test_payload_has_required_detail_fields_rejects_minimal_cached_payload(self):
        issue = {
            "key": "O2-1284",
            "fields": {
                "summary": "Testing By M.Abbas",
                "issuetype": {"name": "Sub-task"},
                "parent": {"key": "O2-790"},
                "timespent": 115200,
                "project": {"key": "O2"},
                "assignee": {"displayName": "Muhammad Abbas"},
                "updated": "2026-03-09T13:08:53.342+0500",
            },
        }
        self.assertFalse(
            _payload_has_required_detail_fields(
                issue,
                detail_fields=[
                    "project",
                    "summary",
                    "status",
                    "duedate",
                    "assignee",
                    "priority",
                    "timetracking",
                    "timeoriginalestimate",
                    "timespent",
                    "aggregatetimespent",
                    "issuetype",
                    "parent",
                    "customfield_10014",
                    "customfield_10584",
                    "created",
                    "updated",
                    "customfield_10211",
                ],
                start_date_field_id="customfield_10211",
                end_date_field_ids=["customfield_10216", "duedate"],
                fix_type_field_id="customfield_10115",
            )
        )

    def test_payload_has_required_detail_fields_accepts_full_payload_even_with_blank_dates(self):
        issue = {
            "key": "O2-1644",
            "fields": {
                "project": {"key": "O2"},
                "summary": "Product Mapping",
                "status": {"name": "To Do"},
                "duedate": None,
                "assignee": {"displayName": "Muhammad Abbas"},
                "priority": {"name": "Medium"},
                "timetracking": {},
                "timeoriginalestimate": 0,
                "timespent": 21600,
                "aggregatetimespent": 21600,
                "issuetype": {"name": "Sub-task"},
                "parent": {"key": "O2-1257"},
                "customfield_10014": None,
                "customfield_10584": None,
                "created": "2026-02-27T09:00:00.000+0500",
                "updated": "2026-03-09T13:12:50.004+0500",
                "customfield_10211": None,
                "customfield_10216": None,
                "customfield_10115": None,
            },
        }
        self.assertTrue(
            _payload_has_required_detail_fields(
                issue,
                detail_fields=[
                    "project",
                    "summary",
                    "status",
                    "duedate",
                    "assignee",
                    "priority",
                    "timetracking",
                    "timeoriginalestimate",
                    "timespent",
                    "aggregatetimespent",
                    "issuetype",
                    "parent",
                    "customfield_10014",
                    "customfield_10584",
                    "created",
                    "updated",
                    "customfield_10211",
                ],
                start_date_field_id="customfield_10211",
                end_date_field_ids=["customfield_10216", "duedate"],
                fix_type_field_id="customfield_10115",
            )
        )

    def test_payload_has_required_detail_fields_rejects_missing_leave_type_field_when_requested(self):
        issue = {
            "key": "RLT-249",
            "fields": {
                "project": {"key": "RLT"},
                "summary": "24-March-2026",
                "status": {"name": "Considered in Roadmap & Queued"},
                "duedate": "2026-03-24",
                "assignee": {"displayName": "Namra Zahid"},
                "priority": {"name": "Medium"},
                "timetracking": {},
                "timeoriginalestimate": 28800,
                "timespent": 28800,
                "aggregatetimespent": 28800,
                "issuetype": {"name": "Sub-task"},
                "parent": {"key": "RLT-15"},
                "customfield_10014": None,
                "created": "2026-04-01T11:04:10.369+0500",
                "updated": "2026-04-02T12:55:08.834+0500",
                "customfield_10133": "2026-03-24",
            },
        }
        self.assertFalse(
            _payload_has_required_detail_fields(
                issue,
                detail_fields=[
                    "project",
                    "summary",
                    "status",
                    "duedate",
                    "assignee",
                    "priority",
                    "timetracking",
                    "timeoriginalestimate",
                    "timespent",
                    "aggregatetimespent",
                    "issuetype",
                    "parent",
                    "customfield_10014",
                    "customfield_10584",
                    "created",
                    "updated",
                    "customfield_10133",
                ],
                start_date_field_id="customfield_10133",
                end_date_field_ids=["duedate"],
                fix_type_field_id=None,
            )
        )

    def test_stable_resolved_since_when_final_status_is_resolved(self):
        issue = {
            "changelog": {
                "histories": [
                    {
                        "created": "2026-02-01T10:00:00.000+0000",
                        "items": [{"field": "status", "fromString": "In Progress", "toString": "Resolved!"}],
                    }
                ]
            }
        }
        self.assertEqual(_stable_resolved_since(issue), "2026-02-01")

    def test_stable_resolved_since_resets_when_reopened(self):
        issue = {
            "changelog": {
                "histories": [
                    {
                        "created": "2026-02-01T10:00:00.000+0000",
                        "items": [{"field": "status", "fromString": "In Progress", "toString": "Resolved!"}],
                    },
                    {
                        "created": "2026-02-02T10:00:00.000+0000",
                        "items": [{"field": "status", "fromString": "Resolved!", "toString": "Reopened"}],
                    },
                    {
                        "created": "2026-02-03T10:00:00.000+0000",
                        "items": [{"field": "status", "fromString": "Reopened", "toString": "Resolved!"}],
                    },
                ]
            }
        }
        self.assertEqual(_stable_resolved_since(issue), "2026-02-03")

    def test_stable_resolved_since_empty_when_never_resolved(self):
        issue = {
            "changelog": {
                "histories": [
                    {
                        "created": "2026-02-01T10:00:00.000+0000",
                        "items": [{"field": "status", "fromString": "To Do", "toString": "In Progress"}],
                    }
                ]
            }
        }
        self.assertEqual(_stable_resolved_since(issue), "")


if __name__ == "__main__":
    unittest.main()
