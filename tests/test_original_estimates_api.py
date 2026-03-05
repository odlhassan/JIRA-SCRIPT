from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from report_server import create_report_server_app


def _build_app(root: Path):
    (root / "report_html").mkdir(parents=True, exist_ok=True)
    (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
    wb = Workbook()
    ws = wb.active
    ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
    ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
    wb.save(root / "assignee_hours_report.xlsx")
    return create_report_server_app(base_dir=root, folder_raw="report_html")


def _hierarchy_fixture():
    return {
        "epics": [
            {
                "issue_key": "O2-EP1",
                "project_key": "O2",
                "project_name": "O2",
                "summary": "Epic One",
                "status": "In Progress",
                "assignee": "Alice",
                "estimate_hours": 20.0,
                "planned_start": "2026-02-01",
                "planned_due": "2026-02-20",
            },
            {
                "issue_key": "O2-EP2",
                "project_key": "O2",
                "project_name": "O2",
                "summary": "Epic Two",
                "status": "In Progress",
                "assignee": "Bob",
                "estimate_hours": 10.0,
                "planned_start": "2026-02-01",
                "planned_due": "2026-02-20",
            },
        ],
        "stories": [
            {
                "issue_key": "O2-ST1",
                "epic_key": "O2-EP1",
                "project_key": "O2",
                "project_name": "O2",
                "summary": "Story One",
                "status": "In Progress",
                "assignee": "Alice",
                "estimate_hours": 8.0,
                "planned_start": "2026-02-02",
                "planned_due": "2026-02-12",
            },
            {
                "issue_key": "O2-ST2",
                "epic_key": "O2-EP1",
                "project_key": "O2",
                "project_name": "O2",
                "summary": "Story Two",
                "status": "In Progress",
                "assignee": "Bob",
                "estimate_hours": 12.0,
                "planned_start": "2026-02-03",
                "planned_due": "2026-02-14",
            },
        ],
        "subtasks": [
            {
                "issue_key": "O2-SUB1",
                "story_key": "O2-ST1",
                "epic_key": "O2-EP1",
                "project_key": "O2",
                "project_name": "O2",
                "summary": "Subtask One",
                "status": "In Progress",
                "assignee": "Alice",
                "estimate_hours": 3.0,
                "planned_start": "2026-02-04",
                "planned_due": "2026-02-05",
            },
            {
                "issue_key": "O2-SUB2",
                "story_key": "O2-ST1",
                "epic_key": "O2-EP1",
                "project_key": "O2",
                "project_name": "O2",
                "summary": "Subtask Two",
                "status": "In Progress",
                "assignee": "Alice",
                "estimate_hours": 4.0,
                "planned_start": "2026-02-06",
                "planned_due": "2026-02-07",
            },
            {
                "issue_key": "O2-SUB3",
                "story_key": "O2-ST2",
                "epic_key": "O2-EP1",
                "project_key": "O2",
                "project_name": "O2",
                "summary": "Subtask Three",
                "status": "In Progress",
                "assignee": "Bob",
                "estimate_hours": 5.0,
                "planned_start": "2026-02-08",
                "planned_due": "2026-02-09",
            },
        ],
    }


def _make_issue(key: str, issue_type: str, summary: str, assignee: str, estimate_hours: float, start: str, due: str, project_key: str = "O2", parent_key: str = "", epic_link: str = ""):
    fields = {
        "summary": summary,
        "status": {"name": "In Progress"},
        "assignee": {"displayName": assignee},
        "issuetype": {"name": issue_type},
        "project": {"key": project_key, "name": project_key},
        "timeoriginalestimate": int(estimate_hours * 3600),
        "duedate": due,
        "customfield_20000": start,
    }
    if parent_key:
        fields["parent"] = {"key": parent_key}
    if epic_link:
        fields["customfield_10014"] = epic_link
    return {"key": key, "fields": fields}


class OriginalEstimatesApiTests(unittest.TestCase):
    def test_summary_rollups_and_filters(self):
        hierarchy = _hierarchy_fixture()
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            client = app.test_client()
            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server._load_planned_vs_dispensed_hierarchy", return_value=hierarchy),
            ):
                refresh_resp = client.post(
                    "/api/original-estimates/refresh",
                    json={"from": "2026-02-01", "to": "2026-02-28", "projects": ["O2"]},
                )
                self.assertEqual(refresh_resp.status_code, 200)
                self.assertTrue((refresh_resp.get_json() or {}).get("ok"))

            summary_resp = client.get("/api/original-estimates/summary?from=2026-02-01&to=2026-02-28&projects=O2")
            self.assertEqual(summary_resp.status_code, 200)
            summary = summary_resp.get_json() or {}
            self.assertTrue(summary.get("ok"))
            self.assertEqual(len(summary.get("epics") or []), 2)
            epic_one = next((item for item in (summary.get("epics") or []) if item.get("issue_key") == "O2-EP1"), None)
            self.assertIsNotNone(epic_one)
            self.assertEqual(float(epic_one.get("sum_original_estimate_hours") or 0.0), 20.0)
            story_one = next((item for item in (epic_one.get("stories") or []) if item.get("issue_key") == "O2-ST1"), None)
            self.assertIsNotNone(story_one)
            self.assertEqual(float(story_one.get("sum_original_estimate_hours") or 0.0), 7.0)

            filtered_resp = client.get(
                "/api/original-estimates/summary?from=2026-02-01&to=2026-02-28&projects=O2&assignees=alice"
            )
            self.assertEqual(filtered_resp.status_code, 200)
            filtered = filtered_resp.get_json() or {}
            filtered_epic = next((item for item in (filtered.get("epics") or []) if item.get("issue_key") == "O2-EP1"), None)
            self.assertIsNotNone(filtered_epic)
            filtered_stories = filtered_epic.get("stories") or []
            self.assertEqual(len(filtered_stories), 1)
            self.assertEqual(str(filtered_stories[0].get("issue_key")), "O2-ST1")
            self.assertEqual(float(filtered_stories[0].get("sum_original_estimate_hours") or 0.0), 7.0)

    def test_refresh_epic_updates_only_target_subtree(self):
        hierarchy = _hierarchy_fixture()
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            client = app.test_client()

            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server._load_planned_vs_dispensed_hierarchy", return_value=hierarchy),
            ):
                refresh_resp = client.post(
                    "/api/original-estimates/refresh",
                    json={"from": "2026-02-01", "to": "2026-02-28", "projects": ["O2"]},
                )
                self.assertEqual(refresh_resp.status_code, 200)

            epic_issue = _make_issue("O2-EP1", "Epic", "Epic One Updated", "Alice", 22.0, "2026-02-01", "2026-02-20")
            story_issue = _make_issue("O2-ST1", "Story", "Story One Updated", "Alice", 9.0, "2026-02-02", "2026-02-12", parent_key="", epic_link="O2-EP1")
            subtask_issue = _make_issue("O2-SUB1", "Sub-task", "Subtask One Updated", "Alice", 6.0, "2026-02-04", "2026-02-06", parent_key="O2-ST1")

            with (
                patch("report_server.get_session", return_value=object()),
                patch("report_server.resolve_jira_start_date_field_id", return_value="customfield_20000"),
                patch("report_server.resolve_jira_end_date_field_ids", return_value=["duedate"]),
                patch("report_server._fetch_jira_issues_by_keys", return_value=[epic_issue]),
                patch("report_server._fetch_story_issues_for_epics", return_value=[story_issue]),
                patch("report_server._fetch_subtask_issues_for_stories", return_value=[subtask_issue]),
            ):
                epic_refresh = client.post(
                    "/api/original-estimates/refresh-epic/O2-EP1",
                    json={"from": "2026-02-01", "to": "2026-02-28"},
                )
                self.assertEqual(epic_refresh.status_code, 200)
                self.assertTrue((epic_refresh.get_json() or {}).get("ok"))

            summary_resp = client.get("/api/original-estimates/summary?from=2026-02-01&to=2026-02-28&projects=O2")
            self.assertEqual(summary_resp.status_code, 200)
            payload = summary_resp.get_json() or {}
            epic_one = next((item for item in (payload.get("epics") or []) if item.get("issue_key") == "O2-EP1"), None)
            epic_two = next((item for item in (payload.get("epics") or []) if item.get("issue_key") == "O2-EP2"), None)
            self.assertIsNotNone(epic_one)
            self.assertIsNotNone(epic_two)
            self.assertEqual(str(epic_one.get("summary")), "Epic One Updated")
            self.assertEqual(float(epic_one.get("original_estimate_hours") or 0.0), 22.0)
            self.assertEqual(str(epic_two.get("summary")), "Epic Two")


if __name__ == "__main__":
    unittest.main()
