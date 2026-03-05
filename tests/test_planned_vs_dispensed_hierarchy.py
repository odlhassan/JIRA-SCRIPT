from __future__ import annotations

from datetime import date
import unittest
from unittest.mock import patch

import report_server


class PlannedVsDispensedHierarchyTests(unittest.TestCase):
    def test_planned_dates_includes_epic_when_descendant_subtask_overlaps(self):
        epic_issue = {
            "key": "CRM-1",
            "fields": {
                "issuetype": {"name": "Epic"},
                "project": {"key": "CRM", "name": "CRM"},
                "summary": "CRM Epic",
                "status": {"name": "In Progress"},
                "assignee": {"displayName": "Alice"},
                "timeoriginalestimate": 36000,
                "custom_start": "2026-01-01",
                "duedate": "2026-01-15",
            },
        }
        story_issue = {
            "key": "CRM-2",
            "fields": {
                "issuetype": {"name": "Story"},
                "project": {"key": "CRM", "name": "CRM"},
                "summary": "CRM Story",
                "status": {"name": "In Progress"},
                "assignee": {"displayName": "Alice"},
                "customfield_10014": "CRM-1",
                "timeoriginalestimate": 0,
                "custom_start": "2026-01-20",
                "duedate": "2026-01-25",
            },
        }
        subtask_issue = {
            "key": "CRM-3",
            "fields": {
                "issuetype": {"name": "Sub-task"},
                "project": {"key": "CRM", "name": "CRM"},
                "parent": {"key": "CRM-2"},
                "summary": "CRM Subtask",
                "status": {"name": "In Progress"},
                "assignee": {"displayName": "Alice"},
                "timeoriginalestimate": 7200,
                "custom_start": "2026-02-10",
                "duedate": "2026-02-11",
            },
        }

        def _fake_fetch_jql(_session, jql, _fields):
            if "issuetype = Epic" in str(jql):
                return [epic_issue]
            return []

        with (
            patch("report_server.resolve_jira_start_date_field_id", return_value="custom_start"),
            patch("report_server.resolve_jira_end_date_field_ids", return_value=["duedate"]),
            patch("report_server._fetch_jira_issues_for_jql", side_effect=_fake_fetch_jql),
            patch("report_server._fetch_story_issues_for_epics", return_value=[story_issue]),
            patch("report_server._fetch_subtask_issues_for_stories", return_value=[subtask_issue]),
        ):
            out = report_server._load_planned_vs_dispensed_hierarchy(
                session=object(),
                from_date=date(2026, 2, 1),
                to_date=date(2026, 2, 28),
                mode="planned_dates",
                selected_projects={"CRM"},
            )

        self.assertEqual(len(out.get("epics", [])), 1)
        self.assertEqual(len(out.get("stories", [])), 1)
        self.assertEqual(len(out.get("subtasks", [])), 1)
        self.assertEqual(out["epics"][0]["issue_key"], "CRM-1")
        self.assertEqual(out["stories"][0]["epic_key"], "CRM-1")
        self.assertEqual(out["subtasks"][0]["epic_key"], "CRM-1")


if __name__ == "__main__":
    unittest.main()
