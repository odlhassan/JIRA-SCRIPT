import json
import os
import re
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

import generate_ipp_meeting_dashboard as dashboard_gen


def _create_epics_db(path: Path) -> None:
    conn = sqlite3.connect(path)
    try:
        conn.execute(
            """
            CREATE TABLE epics_management (
                epic_key TEXT PRIMARY KEY,
                project_key TEXT NOT NULL,
                project_name TEXT NOT NULL,
                product_category TEXT NOT NULL,
                epic_name TEXT NOT NULL,
                description TEXT NOT NULL DEFAULT '',
                originator TEXT NOT NULL DEFAULT '',
                priority TEXT NOT NULL DEFAULT 'Low',
                plan_status TEXT NOT NULL DEFAULT 'Plan',
                ipp_meeting_planned TEXT NOT NULL DEFAULT 'No',
                jira_url TEXT NOT NULL DEFAULT '',
                epic_plan_json TEXT NOT NULL DEFAULT '{}',
                research_urs_plan_json TEXT NOT NULL DEFAULT '{}',
                dds_plan_json TEXT NOT NULL DEFAULT '{}',
                development_plan_json TEXT NOT NULL DEFAULT '{}',
                sqa_plan_json TEXT NOT NULL DEFAULT '{}',
                production_plan_json TEXT NOT NULL DEFAULT '{}'
            )
            """
        )
        conn.execute(
            """
            INSERT INTO epics_management (
                epic_key, project_key, project_name, product_category, epic_name,
                description, originator, priority, plan_status, ipp_meeting_planned, jira_url,
                epic_plan_json, research_urs_plan_json, dds_plan_json,
                development_plan_json, sqa_plan_json, production_plan_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "O2-111",
                "O2",
                "O2 Project",
                "Payments",
                "Selected Epic",
                "Selected for dashboard",
                "Lead",
                "High",
                "Planned",
                "Yes",
                "https://jira.example.com/browse/O2-111",
                json.dumps({"man_days": 10, "start_date": "2026-02-01", "due_date": "2026-02-20"}),
                json.dumps({"man_days": 3, "start_date": "2026-02-01", "due_date": "2026-02-05"}),
                json.dumps({"man_days": 2, "start_date": "2026-02-06", "due_date": "2026-02-08"}),
                json.dumps({"man_days": 3, "start_date": "2026-02-09", "due_date": "2026-02-14"}),
                json.dumps({"man_days": 1, "start_date": "2026-02-15", "due_date": "2026-02-17"}),
                json.dumps({"man_days": 1, "start_date": "2026-02-18", "due_date": "2026-02-20"}),
            ),
        )
        conn.execute(
            """
            INSERT INTO epics_management (
                epic_key, project_key, project_name, product_category, epic_name,
                description, originator, priority, plan_status, ipp_meeting_planned, jira_url,
                epic_plan_json, research_urs_plan_json, dds_plan_json,
                development_plan_json, sqa_plan_json, production_plan_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '{}', '{}', '{}', '{}', '{}', '{}')
            """,
            (
                "O2-222",
                "O2",
                "O2 Project",
                "Payments",
                "Not Selected Epic",
                "",
                "",
                "Low",
                "Plan",
                "No",
                "https://jira.example.com/browse/O2-222",
            ),
        )
        conn.execute(
            """
            INSERT INTO epics_management (
                epic_key, project_key, project_name, product_category, epic_name,
                description, originator, priority, plan_status, ipp_meeting_planned, jira_url,
                epic_plan_json, research_urs_plan_json, dds_plan_json,
                development_plan_json, sqa_plan_json, production_plan_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '{}', '{}', '{}', '{}', '{}')
            """,
            (
                "O2-333",
                "O2",
                "O2 Project",
                "Core",
                "Selected Without Jira",
                "",
                "",
                "Medium",
                "Plan",
                "Yes",
                "",
                json.dumps({"man_days": 4, "start_date": "2026-03-01", "due_date": "2026-03-10"}),
            ),
        )
        conn.commit()
    finally:
        conn.close()


def _create_work_items_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "WorkItems"
    ws.append(
        [
            "project_key",
            "issue_key",
            "parent_issue_key",
            "epic_key",
            "work_item_id",
            "work_item_type",
            "jira_issue_type",
            "summary",
            "status",
            "start_date",
            "end_date",
            "actual_end_date",
            "original_estimate_hours",
            "assignee",
            "total_hours_logged",
            "jira_url",
            "IPP Actual Date (Production Date)",
            "IPP Remarks",
        ]
    )
    ws.append(
        [
            "O2",
            "O2-111",
            "",
            "O2-111",
            "1",
            "Epic",
            "Epic",
            "Selected Epic",
            "In Progress",
            "2026-02-02",
            "2026-02-21",
            "2026-02-18",
            20.0,
            "Alice",
            15.0,
            "https://jira.example.com/browse/O2-111",
            "2026-02-18",
            "On track",
        ]
    )
    ws.append(
        [
            "O2",
            "O2-222",
            "",
            "O2-222",
            "2",
            "Epic",
            "Epic",
            "Not Selected Epic",
            "Done",
            "2026-01-01",
            "2026-01-20",
            "2026-01-20",
            12.0,
            "Bob",
            12.0,
            "https://jira.example.com/browse/O2-222",
            "2026-01-20",
            "",
        ]
    )
    ws.append(
        [
            "O2",
            "O2-1111",
            "O2-111",
            "O2-111",
            "3",
            "Story",
            "Story",
            "Story from Excel",
            "In Progress",
            "2026-02-03",
            "2026-02-07",
            "",
            8.0,
            "Charlie",
            6.0,
            "https://jira.example.com/browse/O2-1111",
            "",
            "",
        ]
    )
    wb.save(path)


def _extract_payload_rows(output_html: Path) -> list[dict]:
    html = output_html.read_text(encoding="utf-8")
    match = re.search(
        r'<script id="ipp-phase-data" type="application/json">(.*?)</script>',
        html,
        flags=re.DOTALL,
    )
    if not match:
        raise AssertionError("Could not locate dashboard payload script tag.")
    payload_text = match.group(1).replace("<\\/script", "</script")
    payload = json.loads(payload_text)
    return payload.get("rows") or []


class IppMeetingDashboardMergeTests(unittest.TestCase):
    def test_selected_epics_only_with_jira_enrichment(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            work_items_path = root / "1_jira_work_items_export.xlsx"
            output_html = root / "ipp_meeting_dashboard.html"
            _create_epics_db(db_path)
            _create_work_items_xlsx(work_items_path)

            env = {
                "JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH": str(db_path),
                "JIRA_EXPORT_XLSX_PATH": str(work_items_path),
                "IPP_PHASE_DASHBOARD_HTML_PATH": str(output_html),
            }
            with patch.dict(os.environ, env, clear=False):
                with patch.object(sys, "argv", ["generate_ipp_meeting_dashboard.py"]):
                    dashboard_gen.main()

            rows = _extract_payload_rows(output_html)
            by_key = {str(item.get("epic_rmi") or "").upper(): item for item in rows}

            self.assertIn("O2-111", by_key)
            self.assertIn("O2-333", by_key)
            self.assertNotIn("O2-222", by_key)

            selected_with_jira = by_key["O2-111"]
            self.assertEqual(selected_with_jira.get("jira_status"), "In Progress")
            self.assertEqual(selected_with_jira.get("jira_assignee"), "Alice")
            self.assertEqual(selected_with_jira.get("jira_total_hours_logged"), 15.0)
            self.assertEqual(selected_with_jira.get("jira_original_estimate_hours"), 20.0)
            self.assertEqual(selected_with_jira.get("jira_progress_pct"), 75.0)
            self.assertEqual(selected_with_jira.get("epic_planned_start_date"), "2026-02-01")
            self.assertEqual(selected_with_jira.get("epic_planned_end_date"), "2026-02-20")
            self.assertEqual(selected_with_jira.get("epic_planned_start_date_db"), "2026-02-01")
            self.assertEqual(selected_with_jira.get("epic_planned_end_date_db"), "2026-02-20")
            self.assertEqual(selected_with_jira.get("epic_planned_start_date_jira"), "2026-02-02")
            self.assertEqual(selected_with_jira.get("epic_planned_end_date_jira"), "2026-02-21")
            self.assertEqual(selected_with_jira.get("epic_planned_hours_db"), 80.0)
            self.assertEqual(selected_with_jira.get("epic_planned_hours_jira"), 20.0)
            stories = selected_with_jira.get("stories") or []
            self.assertEqual(len(stories), 1)
            self.assertEqual(stories[0].get("story_key"), "O2-1111")
            self.assertEqual(stories[0].get("story_name"), "Story from Excel")
            self.assertEqual(stories[0].get("story_start_date"), "2026-02-03")
            self.assertEqual(stories[0].get("story_end_date"), "2026-02-07")
            self.assertEqual(stories[0].get("story_planned_hours"), 8.0)
            self.assertEqual(stories[0].get("story_logged_hours"), 6.0)
            self.assertEqual(stories[0].get("story_progress_pct"), 75.0)

            selected_without_jira = by_key["O2-333"]
            self.assertEqual(selected_without_jira.get("jira_status"), "")
            self.assertEqual(selected_without_jira.get("jira_assignee"), "")
            self.assertIsNone(selected_without_jira.get("jira_progress_pct"))
            self.assertEqual(selected_without_jira.get("epic_planned_start_date_db"), "2026-03-01")
            self.assertEqual(selected_without_jira.get("epic_planned_end_date_db"), "2026-03-10")
            self.assertEqual(selected_without_jira.get("epic_planned_start_date_jira"), "")
            self.assertEqual(selected_without_jira.get("epic_planned_end_date_jira"), "")
            self.assertEqual(selected_without_jira.get("epic_planned_hours_db"), 32.0)
            self.assertIsNone(selected_without_jira.get("epic_planned_hours_jira"))
            self.assertEqual(selected_without_jira.get("stories"), [])

    def test_returns_no_rows_when_no_selected_epics(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            work_items_path = root / "1_jira_work_items_export.xlsx"
            _create_epics_db(db_path)
            _create_work_items_xlsx(work_items_path)

            conn = sqlite3.connect(db_path)
            try:
                conn.execute("UPDATE epics_management SET ipp_meeting_planned='No'")
                conn.commit()
            finally:
                conn.close()

            env = {
                "JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH": str(db_path),
                "JIRA_EXPORT_XLSX_PATH": str(work_items_path),
            }
            with patch.dict(os.environ, env, clear=False):
                payload = dashboard_gen.build_payload_from_sources(base_dir=root)

            rows = payload.get("rows") or []
            self.assertEqual(payload.get("selection_mode"), "selected_only")
            self.assertEqual(rows, [])


if __name__ == "__main__":
    unittest.main()
