import json
import os
import re
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import generate_ipp_meeting_dashboard as dashboard_gen
from jira_export_db import ensure_schema
from jira_export_db import write_work_items


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
                component TEXT NOT NULL DEFAULT '',
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
                epic_key, project_key, project_name, product_category, component, epic_name,
                description, originator, priority, plan_status, ipp_meeting_planned, jira_url,
                epic_plan_json, research_urs_plan_json, dds_plan_json,
                development_plan_json, sqa_plan_json, production_plan_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "O2-111",
                "O2",
                "O2 Project",
                "Payments",
                "Dashboard",
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
                epic_key, project_key, project_name, product_category, component, epic_name,
                description, originator, priority, plan_status, ipp_meeting_planned, jira_url,
                epic_plan_json, research_urs_plan_json, dds_plan_json,
                development_plan_json, sqa_plan_json, production_plan_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '{}', '{}', '{}', '{}', '{}', '{}')
            """,
            (
                "O2-222",
                "O2",
                "O2 Project",
                "Payments",
                "",
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
                epic_key, project_key, project_name, product_category, component, epic_name,
                description, originator, priority, plan_status, ipp_meeting_planned, jira_url,
                epic_plan_json, research_urs_plan_json, dds_plan_json,
                development_plan_json, sqa_plan_json, production_plan_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '{}', '{}', '{}', '{}', '{}')
            """,
            (
                "O2-333",
                "O2",
                "O2 Project",
                "Core",
                "Enhancements",
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


def _work_item_row(
    project_key,
    issue_key,
    work_item_id,
    work_item_type,
    jira_issue_type,
    summary,
    status,
    start_date,
    end_date,
    actual_end_date,
    original_estimate_hours,
    assignee,
    total_hours_logged,
    parent_issue_key,
    jira_url,
    ipp_actual_date,
    ipp_remarks,
):
    """Build a row list in WORK_ITEMS_COLS order (28 cols)."""
    return [
        project_key,
        issue_key,
        str(work_item_id) if work_item_id else None,
        work_item_type,
        jira_issue_type,
        None,  # fix_type
        summary,
        status,
        start_date,
        end_date,
        None,  # actual_start_date
        actual_end_date,
        None,  # original_estimate
        original_estimate_hours,
        assignee,
        total_hours_logged,
        None,  # priority
        parent_issue_key or None,
        None,  # parent_work_item_id
        None,  # parent_jira_url
        jira_url,
        None,  # latest_ipp_meeting
        None,  # jira_ipp_rmi_dates_altered
        ipp_actual_date,
        ipp_remarks,
        None,  # ipp_actual_date_matches_jira_end_date
        None,  # created
        None,  # updated
    ]


def _create_work_items_db(exports_db_path: Path) -> None:
    """Create jira_exports.db with work_items table and test data (no Excel)."""
    exports_db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(exports_db_path))
    conn.row_factory = sqlite3.Row
    try:
        ensure_schema(conn)
        rows = [
            _work_item_row(
                "O2", "O2-111", "1", "Epic", "Epic", "Selected Epic", "In Progress",
                "2026-02-02", "2026-02-21", "2026-02-18", 20.0, "Alice", 15.0,
                "", "https://jira.example.com/browse/O2-111", "2026-02-18", "On track",
            ),
            _work_item_row(
                "O2", "O2-222", "2", "Epic", "Epic", "Not Selected Epic", "Done",
                "2026-01-01", "2026-01-20", "2026-01-20", 12.0, "Bob", 12.0,
                "", "https://jira.example.com/browse/O2-222", "2026-01-20", "",
            ),
            _work_item_row(
                "O2", "O2-1111", "3", "Story", "Story", "Story from Excel", "In Progress",
                "2026-02-03", "2026-02-07", "", 8.0, "Charlie", 6.0,
                "O2-111", "https://jira.example.com/browse/O2-1111", "", "",
            ),
        ]
        write_work_items(conn, rows)
    finally:
        conn.close()


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
            exports_db_path = root / "jira_exports.db"
            output_html = root / "ipp_meeting_dashboard.html"
            _create_epics_db(db_path)
            _create_work_items_db(exports_db_path)

            env = {
                "JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH": str(db_path),
                "JIRA_EXPORTS_DB_PATH": str(exports_db_path),
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
            self.assertEqual(stories[0].get("story_name"), "Story from Excel")  # from DB work_items
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
            exports_db_path = root / "jira_exports.db"
            _create_epics_db(db_path)
            _create_work_items_db(exports_db_path)

            conn = sqlite3.connect(db_path)
            try:
                conn.execute("UPDATE epics_management SET ipp_meeting_planned='No'")
                conn.commit()
            finally:
                conn.close()

            env = {
                "JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH": str(db_path),
                "JIRA_EXPORTS_DB_PATH": str(exports_db_path),
            }
            with patch.dict(os.environ, env, clear=False):
                payload = dashboard_gen.build_payload_from_sources(base_dir=root)

            rows = payload.get("rows") or []
            self.assertEqual(payload.get("selection_mode"), "selected_only")
            self.assertEqual(rows, [])

    def test_phase_jira_link_resolves_dates_from_work_items(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            exports_db_path = root / "jira_exports.db"

            from report_server import _init_epics_management_db

            _init_epics_management_db(db_path)
            conn = sqlite3.connect(db_path)
            try:
                conn.execute(
                    """
                    INSERT INTO epics_management (
                        id, epic_key, project_key, project_name, product_category, component, epic_name,
                        description, originator, priority, plan_status, ipp_meeting_planned, actual_production_date, delivery_status, remarks, jira_url,
                        epic_plan_json, research_urs_plan_json, dds_plan_json, development_plan_json, sqa_plan_json, user_manual_plan_json, production_plan_json, is_sealed
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, '', '', 'High', 'Planned', 'Yes', '', 'Yet to start', '', ?, '{}', '{}', '{}', '{}', '{}', '{}', '{}', 0)
                    """,
                    (
                        "row-phase",
                        "O2-999",
                        "O2",
                        "O2 Project",
                        "Payments",
                        "Scheduling",
                        "Phase Jira Epic",
                        "https://jira.example.com/browse/O2-999",
                    ),
                )
                now = "2026-04-01T00:00:00+00:00"
                conn.execute(
                    """
                    INSERT INTO epics_management_plan_values (
                        epic_row_id, epic_key, column_key, plan_json, created_at_utc, updated_at_utc
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        "row-phase",
                        "O2-999",
                        "development_plan",
                        json.dumps({
                            "most_likely_man_days": 2,
                            "man_days": 2,
                            "start_date": "",
                            "due_date": "",
                            "jira_url": "https://jira.example.com/browse/O2-444",
                        }),
                        now,
                        now,
                    ),
                )
                conn.commit()
            finally:
                conn.close()

            exports_db_path.parent.mkdir(parents=True, exist_ok=True)
            conn = sqlite3.connect(str(exports_db_path))
            conn.row_factory = sqlite3.Row
            try:
                ensure_schema(conn)
                write_work_items(
                    conn,
                    [
                        _work_item_row(
                            "O2", "O2-999", "1", "Epic", "Epic", "Phase Jira Epic", "In Progress",
                            "2026-05-01", "2026-05-10", "", 40.0, "Alice", 4.0,
                            "", "https://jira.example.com/browse/O2-999", "", "",
                        ),
                        _work_item_row(
                            "O2", "O2-444", "2", "Story", "Story", "Development", "To Do",
                            "2026-05-03", "2026-05-06", "", 16.0, "Bob", 8.0,
                            "O2-999", "https://jira.example.com/browse/O2-444", "", "",
                        ),
                    ],
                )
            finally:
                conn.close()

            env = {
                "JIRA_ASSIGNEE_HOURS_CAPACITY_DB_PATH": str(db_path),
                "JIRA_EXPORTS_DB_PATH": str(exports_db_path),
            }
            with patch.dict(os.environ, env, clear=False):
                payload = dashboard_gen.build_payload_from_sources(base_dir=root)

            rows = payload.get("rows") or []
            by_key = {str(item.get("epic_rmi") or "").upper(): item for item in rows}
            row = by_key["O2-999"]
            phase_values = list((row.get("phase_data") or {}).values())
            development = next(item for item in phase_values if item.get("plan_key") == "development_plan")
            self.assertEqual(development.get("jira_url"), "https://jira.example.com/browse/O2-444")
            self.assertEqual(development.get("linked_issue_key"), "O2-444")
            self.assertEqual(development.get("start"), "2026-05-03")
            self.assertEqual(development.get("end"), "2026-05-06")
            self.assertEqual(development.get("date_source"), "jira")
            self.assertEqual(development.get("linked_issue_status"), "To Do")
            self.assertEqual(development.get("linked_issue_assignee"), "Bob")
            self.assertTrue(row.get("mini_gantt", {}).get("has_dated_phases"))


def _make_meeting_epic(
    *,
    epic_key,
    epic_name,
    project_key="O2",
    project_name="O2 Project",
    item_kind="jira",
    issue_type="epic",
    source_tag="epics_planner",
    assignee_text="",
    plans=None,
    actual_production_date="",
):
    return {
        "epic_key": epic_key,
        "project_key": project_key,
        "project_name": project_name,
        "product_category": "",
        "component": "",
        "epic_name": epic_name,
        "description": "",
        "remarks": "",
        "originator": "",
        "priority": "",
        "plan_status": "",
        "jira_url": "",
        "ipp_meeting_planned": "Yes",
        "actual_production_date": actual_production_date,
        "delivery_status": "Yet to start",
        "plans": plans or {"epic_plan": {}},
        "_record_source": "IPP Meeting Planner",
        "item_kind": item_kind,
        "issue_type": issue_type,
        "source_tag": source_tag,
        "assignee_text": assignee_text,
    }


class IppMeetingDashboardWorkListTests(unittest.TestCase):
    """Tests for non-epic Work List items + custom items added directly to a meeting."""

    def test_jira_story_added_directly_pulls_dates_from_work_items(self):
        meeting_epics = [
            _make_meeting_epic(
                epic_key="O2-1111",
                epic_name="Direct Story",
                issue_type="story",
                source_tag="jira",
                plans={"epic_plan": {}},
            )
        ]
        rows_by_key = {
            "O2-1111": {
                "issue_key": "O2-1111",
                "issue_type": "Story",
                "summary": "Direct Story",
                "status": "In Progress",
                "assignee": "Alice",
                "jira_url": "https://jira.example.com/browse/O2-1111",
                "start_date": "2026-04-01",
                "end_date": "2026-04-10",
                "actual_end_date": "",
                "parent_issue_key": "O2-111",
                "original_estimate_hours": 16.0,
                "total_hours_logged": 4.0,
                "progress_pct": 25.0,
            }
        }
        records = dashboard_gen._build_records(meeting_epics, {}, {}, rows_by_key, {"O2-1111": "O2-111"})
        rows = dashboard_gen._rows_for_payload(records)
        self.assertEqual(len(rows), 1)
        r = rows[0]
        self.assertEqual(r["record_kind"], "jira_non_epic")
        self.assertEqual(r["epic_planned_start_date"], "2026-04-01")
        self.assertEqual(r["epic_planned_end_date"], "2026-04-10")
        # Epic-specific (DB) and (Jira Excel) columns must be blank for non-epics.
        self.assertEqual(r["epic_planned_start_date_db"], "")
        self.assertEqual(r["epic_planned_end_date_db"], "")
        self.assertEqual(r["epic_planned_start_date_jira"], "")
        self.assertEqual(r["epic_planned_end_date_jira"], "")
        # Item-level columns populated.
        self.assertEqual(r["item_planned_start_date"], "2026-04-01")
        self.assertEqual(r["item_planned_end_date"], "2026-04-10")
        # man_days = original_estimate_hours / 8 = 2.0
        self.assertEqual(r["db_epic_planned_mandays"], 2.0)
        self.assertEqual(r["epic_planned_hours_db"], 16.0)
        self.assertEqual(r["jira_status"], "In Progress")
        self.assertEqual(r["jira_assignee"], "Alice")
        self.assertEqual(r["parent_key"], "O2-111")
        # Gantt-relevant flag.
        # has_valid_epic_plan from internal record; surfaced as roadmap.valid in payload row.
        self.assertTrue(r["roadmap"]["valid"])
        # Phase plan must be marked absent for non-epic Jira items.
        self.assertFalse(r["has_phase_plan"])

    def test_subtask_with_empty_dates_inherits_from_parent_story(self):
        meeting_epics = [
            _make_meeting_epic(
                epic_key="O2-2222",
                epic_name="Empty-Dates Subtask",
                issue_type="subtask",
                source_tag="jira",
            )
        ]
        rows_by_key = {
            "O2-2222": {
                "issue_key": "O2-2222",
                "issue_type": "Sub-task",
                "summary": "Empty-Dates Subtask",
                "status": "To Do",
                "assignee": "Bob",
                "jira_url": "",
                "start_date": "",
                "end_date": "",
                "actual_end_date": "",
                "parent_issue_key": "O2-1111",
                "original_estimate_hours": 4.0,
                "total_hours_logged": None,
                "progress_pct": None,
            },
            "O2-1111": {
                "issue_key": "O2-1111",
                "issue_type": "Story",
                "summary": "Parent Story",
                "status": "In Progress",
                "assignee": "Alice",
                "jira_url": "",
                "start_date": "2026-05-01",
                "end_date": "2026-05-08",
                "actual_end_date": "",
                "parent_issue_key": "O2-111",
                "original_estimate_hours": None,
                "total_hours_logged": None,
                "progress_pct": None,
            },
        }
        parent_by_key = {"O2-2222": "O2-1111", "O2-1111": "O2-111"}
        records = dashboard_gen._build_records(meeting_epics, {}, {}, rows_by_key, parent_by_key)
        rows = dashboard_gen._rows_for_payload(records)
        r = rows[0]
        self.assertEqual(r["record_kind"], "jira_non_epic")
        # Inherited from parent Story.
        self.assertEqual(r["epic_planned_start_date"], "2026-05-01")
        self.assertEqual(r["epic_planned_end_date"], "2026-05-08")
        self.assertEqual(r["item_planned_start_date"], "2026-05-01")
        self.assertEqual(r["item_planned_end_date"], "2026-05-08")
        self.assertEqual(r["date_inherited_from"], "O2-1111")
        self.assertTrue(r["roadmap"]["valid"])

    def test_custom_item_uses_assignee_text_and_blank_jira_link(self):
        meeting_epics = [
            _make_meeting_epic(
                epic_key="CUSTOM-9-1",
                epic_name="Chairperson note",
                item_kind="custom",
                issue_type="custom",
                source_tag="custom",
                assignee_text="Aniqa",
                plans={"epic_plan": {"start_date": "2026-06-01", "due_date": "2026-06-05", "man_days": 1}},
            )
        ]
        records = dashboard_gen._build_records(meeting_epics, {}, {}, {}, {})
        rows = dashboard_gen._rows_for_payload(records)
        r = rows[0]
        self.assertEqual(r["record_kind"], "custom")
        self.assertEqual(r["jira_link"], "")
        self.assertEqual(r["jira_assignee"], "Aniqa")
        self.assertEqual(r["epic_planned_start_date"], "2026-06-01")
        self.assertEqual(r["epic_planned_end_date"], "2026-06-05")
        self.assertEqual(r["epic_planned_start_date_db"], "")
        self.assertEqual(r["epic_planned_end_date_db"], "")
        self.assertEqual(r["item_planned_start_date"], "2026-06-01")
        self.assertEqual(r["item_planned_end_date"], "2026-06-05")
        self.assertTrue(r["roadmap"]["valid"])
        self.assertFalse(r["has_phase_plan"])


if __name__ == "__main__":
    unittest.main()
