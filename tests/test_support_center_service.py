from __future__ import annotations

import sqlite3
import tempfile
import unittest
from datetime import date
from pathlib import Path

from support_center_service import (
    build_support_center_overview,
    build_support_center_project_detail,
    is_booking_story,
)

RUN_ID = "run-1"


def _create_canonical_db(db_path: Path) -> None:
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE canonical_refresh_state (
                id INTEGER PRIMARY KEY,
                active_run_id TEXT NOT NULL DEFAULT '',
                last_success_run_id TEXT NOT NULL DEFAULT '',
                updated_at_utc TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.execute(
            "INSERT INTO canonical_refresh_state(id, active_run_id, last_success_run_id, updated_at_utc) VALUES (1, ?, ?, ?)",
            (RUN_ID, RUN_ID, "2026-06-30T00:00:00Z"),
        )
        conn.execute(
            """
            CREATE TABLE canonical_issues (
                run_id TEXT NOT NULL,
                issue_id TEXT NOT NULL DEFAULT '',
                issue_key TEXT NOT NULL,
                project_key TEXT NOT NULL DEFAULT '',
                issue_type TEXT NOT NULL DEFAULT '',
                summary TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT '',
                assignee TEXT NOT NULL DEFAULT '',
                start_date TEXT NOT NULL DEFAULT '',
                due_date TEXT NOT NULL DEFAULT '',
                created_utc TEXT NOT NULL DEFAULT '',
                updated_utc TEXT NOT NULL DEFAULT '',
                resolved_stable_since_date TEXT NOT NULL DEFAULT '',
                original_estimate_hours REAL NOT NULL DEFAULT 0,
                total_hours_logged REAL NOT NULL DEFAULT 0,
                fix_type TEXT NOT NULL DEFAULT '',
                parent_issue_key TEXT NOT NULL DEFAULT '',
                story_key TEXT NOT NULL DEFAULT '',
                epic_key TEXT NOT NULL DEFAULT '',
                raw_payload_json TEXT NOT NULL DEFAULT '{}',
                PRIMARY KEY (run_id, issue_key)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE canonical_worklogs (
                run_id TEXT NOT NULL,
                worklog_id TEXT NOT NULL,
                issue_key TEXT NOT NULL DEFAULT '',
                project_key TEXT NOT NULL DEFAULT '',
                worklog_author TEXT NOT NULL DEFAULT '',
                issue_assignee TEXT NOT NULL DEFAULT '',
                started_utc TEXT NOT NULL DEFAULT '',
                started_date TEXT NOT NULL DEFAULT '',
                updated_utc TEXT NOT NULL DEFAULT '',
                hours_logged REAL NOT NULL DEFAULT 0,
                PRIMARY KEY (run_id, worklog_id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE canonical_issue_actuals (
                run_id TEXT NOT NULL,
                issue_key TEXT NOT NULL,
                first_worklog_date TEXT NOT NULL DEFAULT '',
                last_worklog_date TEXT NOT NULL DEFAULT '',
                actual_complete_date TEXT NOT NULL DEFAULT '',
                total_worklog_hours REAL NOT NULL DEFAULT 0,
                worklog_count INTEGER NOT NULL DEFAULT 0,
                PRIMARY KEY (run_id, issue_key)
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def _ins_issue(conn, key, itype, summary, *, status="In Progress", assignee="",
               story_key="", epic_key="", project_key="O2", start="", due="",
               total_logged=0.0):
    conn.execute(
        """
        INSERT INTO canonical_issues(
            run_id, issue_key, project_key, issue_type, summary, status, assignee,
            start_date, due_date, total_hours_logged, story_key, epic_key
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (RUN_ID, key, project_key, itype, summary, status, assignee, start, due,
         total_logged, story_key, epic_key),
    )


def _ins_actual(conn, key, *, complete_date="", total_hours=0.0, last_worklog=""):
    conn.execute(
        """
        INSERT INTO canonical_issue_actuals(
            run_id, issue_key, last_worklog_date, actual_complete_date, total_worklog_hours, worklog_count
        ) VALUES (?, ?, ?, ?, ?, ?)
        """,
        (RUN_ID, key, last_worklog, complete_date, total_hours, 1),
    )


def _create_support_db(db_path: Path, keys: list[tuple[str, str, str, str]]) -> None:
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE support_issues (
                issue_key TEXT PRIMARY KEY,
                project_key TEXT NOT NULL DEFAULT '',
                issue_type TEXT NOT NULL DEFAULT '',
                summary TEXT NOT NULL DEFAULT '',
                work_type_value TEXT NOT NULL DEFAULT '',
                synced_at_utc TEXT NOT NULL DEFAULT ''
            )
            """
        )
        conn.executemany(
            """
            INSERT INTO support_issues(issue_key, project_key, issue_type, summary, work_type_value, synced_at_utc)
            VALUES (?, ?, ?, ?, 'Support', '2026-06-01T00:00:00Z')
            """,
            keys,
        )
        conn.commit()
    finally:
        conn.close()


class SupportCenterServiceTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        base = Path(self._tmp.name)
        self.canonical = base / "assignee_hours_capacity.db"
        self.support = base / "support_center.db"
        _create_canonical_db(self.canonical)

        conn = sqlite3.connect(self.canonical)
        try:
            # Support epic (June 2026)
            _ins_issue(conn, "O2-1", "Epic", "Support June 2026", epic_key="O2-1")
            # Booking story (roster only)
            _ins_issue(conn, "O2-2", "Story", "Support by Nadeem (June 2026)",
                       assignee="Nadeem", epic_key="O2-1")
            # Actual support story #1 — DONE, completes in range
            _ins_issue(conn, "O2-3", "Story", "Investigate flaky login",
                       status="Done", assignee="Ali", story_key="O2-3", epic_key="O2-1")
            _ins_actual(conn, "O2-3", complete_date="2026-06-15", total_hours=0.0)
            # subtasks under O2-3
            _ins_issue(conn, "O2-3-1", "Sub-task", "Repro", status="Done",
                       assignee="Ali", story_key="O2-3", epic_key="O2-1")
            _ins_actual(conn, "O2-3-1", complete_date="2026-06-12", total_hours=5.0)
            _ins_issue(conn, "O2-3-2", "Bug Subtask", "Fix", status="In Progress",
                       assignee="Sara", story_key="O2-3", epic_key="O2-1")
            _ins_actual(conn, "O2-3-2", complete_date="2026-06-14", total_hours=3.0)
            # Actual support story #2 — open, completes OUT of range (July)
            _ins_issue(conn, "O2-4", "Story", "Performance tuning",
                       status="In Progress", story_key="O2-4", epic_key="O2-1",
                       due="2026-07-10")
            _ins_issue(conn, "O2-4-1", "Sub-task", "Profile", story_key="O2-4", epic_key="O2-1")
            _ins_actual(conn, "O2-4-1", complete_date="2026-07-05", total_hours=9.0)
            conn.commit()
        finally:
            conn.close()

        _create_support_db(self.support, [
            ("O2-2", "O2", "Story", "Support by Nadeem (June 2026)"),
            ("O2-3", "O2", "Story", "Investigate flaky login"),
            ("O2-3-1", "O2", "Sub-task", "Repro"),
            ("O2-3-2", "O2", "Bug Subtask", "Fix"),
            ("O2-4", "O2", "Story", "Performance tuning"),
        ])

    def tearDown(self):
        self._tmp.cleanup()

    def test_booking_regex(self):
        self.assertTrue(is_booking_story("Support by Nadeem (June 2026)"))
        self.assertTrue(is_booking_story("  support by Team A (July 2026) "))
        self.assertFalse(is_booking_story("Investigate flaky login"))
        self.assertFalse(is_booking_story("Support escalation handling"))

    def test_overview_june_range(self):
        payload = build_support_center_overview(
            self.canonical, self.support, RUN_ID,
            date(2026, 6, 1), date(2026, 6, 30),
        )
        be = payload["birds_eye"]
        # Only O2-3 (June) counts; O2-4 completes in July, excluded.
        self.assertEqual(be["support_story_count"], 1)
        self.assertEqual(be["resolved_support_stories"], 1)
        # invested = 5 + 3 from O2-3 subtasks
        self.assertAlmostEqual(be["invested_hours"], 8.0, places=2)
        # available hours fixture absent -> 0 (capacity model swallowed)
        self.assertEqual(be["available_hours"], 0.0)

        self.assertEqual(len(payload["by_project"]), 1)
        proj = payload["by_project"][0]
        self.assertEqual(proj["project_key"], "O2")
        self.assertEqual(proj["story_count"], 1)
        self.assertAlmostEqual(proj["invested_hours"], 8.0, places=2)
        self.assertEqual(proj["subtask_count"], 2)

        # roster from booking story
        self.assertEqual(len(payload["roster"]), 1)
        self.assertEqual(payload["roster"][0]["assignee"], "Nadeem")
        self.assertEqual(payload["roster"][0]["booked_for"], "June 2026")

    def test_overview_july_range_picks_other_story(self):
        payload = build_support_center_overview(
            self.canonical, self.support, RUN_ID,
            date(2026, 7, 1), date(2026, 7, 31),
        )
        be = payload["birds_eye"]
        self.assertEqual(be["support_story_count"], 1)
        self.assertEqual(be["resolved_support_stories"], 0)  # O2-4 still open
        self.assertAlmostEqual(be["invested_hours"], 9.0, places=2)

    def test_project_detail(self):
        payload = build_support_center_project_detail(
            self.canonical, self.support, RUN_ID, "O2",
            date(2026, 6, 1), date(2026, 6, 30),
        )
        self.assertEqual(payload["project_key"], "O2")
        self.assertEqual(payload["summary"]["story_count"], 1)
        self.assertEqual(payload["summary"]["resolved_count"], 1)
        self.assertAlmostEqual(payload["summary"]["invested_hours"], 8.0, places=2)
        self.assertEqual(len(payload["stories"]), 1)
        story = payload["stories"][0]
        self.assertEqual(story["issue_key"], "O2-3")
        self.assertEqual(len(story["subtasks"]), 2)
        sub_keys = {s["issue_key"] for s in story["subtasks"]}
        self.assertEqual(sub_keys, {"O2-3-1", "O2-3-2"})
        self.assertEqual(len(payload["roster"]), 1)


if __name__ == "__main__":
    unittest.main()
