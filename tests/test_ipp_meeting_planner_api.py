"""Tests for IPP Meeting Planner APIs (current meeting, history, complete, meeting epics)."""
from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

from report_server import (
    IPP_MEETING_PLANNER_SETTINGS_ROUTE,
    create_report_server_app,
    _init_epics_management_db,
    _ipp_meeting_planner_get_current_meeting,
    _ipp_meeting_planner_list_meetings,
    _ipp_meeting_planner_complete_meeting,
    _ipp_meeting_planner_add_epic,
    _ipp_meeting_planner_get_meeting_with_epics,
    _ipp_meeting_planner_remove_epic,
    _ipp_meeting_planner_search_work_items,
    _ipp_meeting_planner_add_custom_item,
)


def _create_capacity_db_with_epics(db_path: Path) -> None:
    _init_epics_management_db(db_path)
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            INSERT OR REPLACE INTO epics_management (
                epic_key, project_key, project_name, product_category, component, epic_name,
                description, originator, priority, plan_status, ipp_meeting_planned, actual_production_date, delivery_status, remarks, jira_url,
                epic_plan_json, research_urs_plan_json, dds_plan_json,
                development_plan_json, sqa_plan_json, production_plan_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'No', '', 'Yet to start', '', ?, ?, '{}', '{}', '{}', '{}', '{}')
            """,
            (
                "O2-100",
                "O2",
                "O2 Project",
                "Payments",
                "",
                "Test Epic",
                "",
                "",
                "Low",
                "Planned",
                "",
                '{"man_days": 2, "start_date": "2026-03-10", "due_date": "2026-03-20"}',
            ),
        )
        conn.commit()
    finally:
        conn.close()


class IppMeetingPlannerApiTests(unittest.TestCase):
    def test_init_epics_management_db_adds_and_backfills_is_tk_epic(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            _init_epics_management_db(db_path)
            conn = sqlite3.connect(db_path)
            try:
                conn.execute(
                    """
                    INSERT INTO epics_management (
                        id, epic_key, project_key, project_name, product_category, component, epic_name,
                        description, originator, priority, plan_status, ipp_meeting_planned, actual_production_date, delivery_status, remarks, jira_url,
                        epic_plan_json, research_urs_plan_json, dds_plan_json, development_plan_json, sqa_plan_json, user_manual_plan_json, production_plan_json, is_sealed
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, '', '', 'Low', 'Planned', 'No', '', 'Yet to start', '', '', '{}', '{}', '{}', '{}', '{}', '{}', '{}', 0)
                    """,
                    ("row-1", "O2-100", "O2", "O2 Project", "Payments", "", "Test Epic"),
                )
                conn.execute(
                    """
                    INSERT INTO epics_management_story_sync (
                        story_key, epic_row_id, epic_key, project_key, story_name, story_status, jira_url, start_date, due_date, estimate_hours, payload_json, synced_at_utc
                    ) VALUES (?, ?, ?, ?, '', '', '', '', '', 0, '{}', '')
                    """,
                    ("O2-101", "row-1", "O2-100", "O2"),
                )
                conn.execute("DELETE FROM epics_management_meta WHERE meta_key='is_tk_epic_backfilled'")
                conn.commit()
            finally:
                conn.close()

            _init_epics_management_db(db_path)

            conn = sqlite3.connect(db_path)
            try:
                cols = [row[1] for row in conn.execute("PRAGMA table_info(epics_management)").fetchall()]
                self.assertIn("is_tk_epic", cols)
                val = conn.execute("SELECT is_tk_epic FROM epics_management WHERE id='row-1'").fetchone()
                self.assertEqual(int(val[0]), 1)
            finally:
                conn.close()

    def test_current_meeting_exists_after_bootstrap(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            _init_epics_management_db(db_path)
            meeting = _ipp_meeting_planner_get_current_meeting(db_path)
            self.assertIsNotNone(meeting)
            self.assertEqual(meeting.get("status"), "Scheduled")
            self.assertIn("id", meeting)

    def test_list_meetings_includes_scheduled(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            _init_epics_management_db(db_path)
            meetings = _ipp_meeting_planner_list_meetings(db_path, limit=10)
            self.assertGreaterEqual(len(meetings), 1)
            scheduled = [m for m in meetings if m.get("status") == "Scheduled"]
            self.assertGreaterEqual(len(scheduled), 1)

    def test_add_epic_and_get_meeting_with_epics(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            _create_capacity_db_with_epics(db_path)
            meeting = _ipp_meeting_planner_get_current_meeting(db_path)
            self.assertIsNotNone(meeting)
            meeting_id = meeting["id"]
            row = _ipp_meeting_planner_add_epic(
                db_path,
                meeting_id,
                "O2-100",
                "O2",
                project_name="O2 Project",
                epic_name="Test Epic",
                delivery_status="On-track",
            )
            self.assertIsNotNone(row)
            self.assertEqual(row.get("epic_key"), "O2-100")
            data = _ipp_meeting_planner_get_meeting_with_epics(db_path, meeting_id)
            self.assertIsNotNone(data)
            epics = data.get("epics") or []
            self.assertEqual(len(epics), 1)
            self.assertEqual(epics[0].get("epic_key"), "O2-100")
            self.assertEqual(epics[0].get("delivery_status"), "On-track")
            self.assertEqual(epics[0].get("start_date"), "2026-03-10")
            self.assertEqual(epics[0].get("due_date"), "2026-03-20")

    def test_meeting_epic_remarks_fall_back_when_rich_text_is_visually_empty(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            _create_capacity_db_with_epics(db_path)
            planner_remarks = '<ul><li><font color="#b91c1c">Ali Qumail (18 sessions)</font></li></ul>'
            conn = sqlite3.connect(db_path)
            try:
                conn.execute(
                    "UPDATE epics_management SET remarks = ? WHERE UPPER(epic_key) = 'O2-100'",
                    (planner_remarks,),
                )
                conn.commit()
            finally:
                conn.close()
            meeting = _ipp_meeting_planner_get_current_meeting(db_path)
            meeting_id = meeting["id"]
            _ipp_meeting_planner_add_epic(
                db_path,
                meeting_id,
                "O2-100",
                "O2",
                project_name="O2 Project",
                epic_name="Test Epic",
                remarks_rich_text="<div><br></div>",
            )
            data = _ipp_meeting_planner_get_meeting_with_epics(db_path, meeting_id)
            self.assertIsNotNone(data)
            epics = data.get("epics") or []
            self.assertEqual(len(epics), 1)
            self.assertEqual(epics[0].get("remarks_rich_text"), planner_remarks)

    def test_meeting_epic_delivery_status_falls_back_from_planner(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            _create_capacity_db_with_epics(db_path)
            conn = sqlite3.connect(db_path)
            try:
                conn.execute(
                    "UPDATE epics_management SET delivery_status = 'Late' WHERE UPPER(epic_key) = 'O2-100'"
                )
                conn.commit()
            finally:
                conn.close()
            meeting = _ipp_meeting_planner_get_current_meeting(db_path)
            meeting_id = meeting["id"]
            _ipp_meeting_planner_add_epic(
                db_path,
                meeting_id,
                "O2-100",
                "O2",
                project_name="O2 Project",
                epic_name="Test Epic",
                delivery_status="Yet to start",
            )
            data = _ipp_meeting_planner_get_meeting_with_epics(db_path, meeting_id)
            self.assertIsNotNone(data)
            epics = data.get("epics") or []
            self.assertEqual(len(epics), 1)
            self.assertEqual(epics[0].get("delivery_status"), "Late")

    def test_remove_epic(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            _create_capacity_db_with_epics(db_path)
            meeting = _ipp_meeting_planner_get_current_meeting(db_path)
            meeting_id = meeting["id"]
            _ipp_meeting_planner_add_epic(db_path, meeting_id, "O2-100", "O2", epic_name="Test")
            removed = _ipp_meeting_planner_remove_epic(db_path, meeting_id, "O2-100")
            self.assertTrue(removed)
            data = _ipp_meeting_planner_get_meeting_with_epics(db_path, meeting_id)
            self.assertEqual(len(data.get("epics") or []), 0)

    def test_complete_meeting_creates_next_scheduled(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            _init_epics_management_db(db_path)
            meeting = _ipp_meeting_planner_get_current_meeting(db_path)
            meeting_id = meeting["id"]
            result = _ipp_meeting_planner_complete_meeting(db_path, meeting_id, "2026-04-15")
            self.assertIn("completed_meeting", result)
            self.assertIn("next_meeting", result)
            self.assertEqual(result["completed_meeting"].get("status"), "Completed")
            self.assertEqual(result["next_meeting"].get("status"), "Scheduled")
            self.assertEqual(result["next_meeting"].get("intended_date"), "2026-04-15")
            current_after = _ipp_meeting_planner_get_current_meeting(db_path)
            self.assertIsNotNone(current_after)
            self.assertEqual(current_after["id"], result["next_meeting"]["id"])

    def test_ipp_meeting_planner_settings_page_returns_html(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir(parents=True, exist_ok=True)
            (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            with app.test_client() as client:
                r = client.get(IPP_MEETING_PLANNER_SETTINGS_ROUTE)
                self.assertEqual(r.status_code, 200)
                self.assertIn(b"IPP Meeting Planner", r.data)
                self.assertIn(b"IPP Builder", r.data)
                self.assertIn(b"Work List", r.data)
                self.assertIn(b"History", r.data)

    def test_search_work_items_returns_epics_planner_epic(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            _create_capacity_db_with_epics(db_path)
            data = _ipp_meeting_planner_search_work_items(db_path, query="O2-100")
            self.assertGreaterEqual(data.get("total_seen", 0), 1)
            keys = [it["key"] for it in data.get("items") or []]
            self.assertIn("O2-100", keys)
            for item in data["items"]:
                if item["key"] == "O2-100":
                    self.assertEqual(item["issue_type"], "epic")
                    self.assertTrue(item["in_epics_planner"])
                    self.assertIn(item["source_tag"], ("epics_planner", "tk"))

    def test_search_work_items_filters_by_type_and_source(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            _create_capacity_db_with_epics(db_path)
            # Asking for stories only must not include the epic
            data = _ipp_meeting_planner_search_work_items(db_path, query="O2", types=["story"])
            keys = [it["key"] for it in data.get("items") or []]
            self.assertNotIn("O2-100", keys)
            # Asking for epics_planner source includes the epic
            data2 = _ipp_meeting_planner_search_work_items(db_path, query="O2", sources=["epics_planner", "tk"])
            keys2 = [it["key"] for it in data2.get("items") or []]
            self.assertIn("O2-100", keys2)

    def test_search_work_items_excludes_keys(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            _create_capacity_db_with_epics(db_path)
            data = _ipp_meeting_planner_search_work_items(db_path, query="O2", exclude_keys=["O2-100"])
            keys = [it["key"] for it in data.get("items") or []]
            self.assertNotIn("O2-100", keys)

    def test_add_custom_item_creates_row_with_custom_kind(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            _init_epics_management_db(db_path)
            meeting = _ipp_meeting_planner_get_current_meeting(db_path)
            self.assertIsNotNone(meeting)
            row = _ipp_meeting_planner_add_custom_item(
                db_path,
                meeting["id"],
                title="Chairperson follow-up",
                assignee_text="Ali Qumail",
                start_date="2026-05-01",
                due_date="2026-05-15",
                delivery_status="On-track",
                remarks_rich_text="<p>Tracked manually</p>",
            )
            self.assertIsNotNone(row)
            self.assertEqual(row.get("item_kind"), "custom")
            self.assertEqual(row.get("issue_type"), "custom")
            self.assertEqual(row.get("source_tag"), "custom")
            self.assertEqual(row.get("epic_name"), "Chairperson follow-up")
            self.assertEqual(row.get("assignee_text"), "Ali Qumail")
            self.assertTrue(str(row.get("epic_key", "")).startswith("CUSTOM-"))
            data = _ipp_meeting_planner_get_meeting_with_epics(db_path, meeting["id"])
            self.assertIsNotNone(data)
            epics = data.get("epics") or []
            self.assertEqual(len(epics), 1)
            self.assertEqual(epics[0]["item_kind"], "custom")

    def test_meeting_date_save_endpoint_returns_updated(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir(parents=True, exist_ok=True)
            (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            db_path = root / "assignee_hours_capacity.db"
            _init_epics_management_db(db_path)
            meeting = _ipp_meeting_planner_get_current_meeting(db_path)
            self.assertIsNotNone(meeting)
            with app.test_client() as client:
                r = client.patch(
                    "/api/ipp-meeting-planner/meetings/{0}".format(meeting["id"]),
                    json={"meeting_date": "2026-05-12"},
                )
                self.assertEqual(r.status_code, 200)
                data = r.get_json()
                self.assertEqual(data.get("meeting_date"), "2026-05-12")


if __name__ == "__main__":
    unittest.main()
