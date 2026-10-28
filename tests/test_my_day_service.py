from datetime import date
from pathlib import Path
from contextlib import closing
import sqlite3
import tempfile
import unittest
from unittest.mock import patch

from my_day_service import build_my_day, day_status, expected_hours, list_people

RUN = "run-1"
EMPLOYEE = "Ayesha Khan"


def make_db(path: Path) -> None:
    with closing(sqlite3.connect(path)) as conn, conn:
        conn.executescript("""
            CREATE TABLE canonical_refresh_state (id INTEGER PRIMARY KEY, active_run_id TEXT, last_success_run_id TEXT, updated_at_utc TEXT);
            CREATE TABLE canonical_issues (run_id TEXT, issue_key TEXT, project_key TEXT, issue_type TEXT, summary TEXT, status TEXT,
                assignee TEXT, start_date TEXT, due_date TEXT, original_estimate_hours REAL, total_hours_logged REAL, epic_key TEXT,
                parent_issue_key TEXT DEFAULT '', updated_utc TEXT DEFAULT '');
            CREATE TABLE canonical_worklogs (run_id TEXT, worklog_id TEXT, issue_key TEXT, project_key TEXT, worklog_author TEXT,
                issue_assignee TEXT, started_utc TEXT, started_date TEXT, updated_utc TEXT, hours_logged REAL);
            CREATE TABLE assignee_capacity_settings (from_date TEXT, to_date TEXT, employee_count INTEGER, standard_hours_per_day REAL,
                ramadan_start_date TEXT, ramadan_end_date TEXT, ramadan_hours_per_day REAL, holiday_dates_json TEXT, updated_at_utc TEXT);
        """)
        conn.execute("INSERT INTO canonical_refresh_state VALUES (1, ?, ?, '2026-10-21T05:00:00+00:00')", (RUN, RUN))
        conn.execute("INSERT INTO assignee_capacity_settings VALUES ('2026-01-01','2026-12-31',10,8,'','',0,'[\"2026-10-23\"]','x')")
        issues = [
            ("PLAT-1", "PLAT", "Epic", "Jira Incremental Sync", "In-Progress", "", "", "", 0, 0, ""),
            ("PLAT-2", "PLAT", "Sub-task", "Refactor sync retry", "In-Progress", EMPLOYEE, "2026-10-12", "2026-10-30", 6, 3.5, "PLAT-1"),
            ("PLAT-3", "PLAT", "Sub-task", "Old finished task", "Resolved", EMPLOYEE, "2026-09-01", "2026-09-05", 2, 2, "PLAT-1"),
            ("PLAT-4", "PLAT", "Bug Subtask", "Fix cache bug", "To Do", EMPLOYEE, "2026-10-01", "2026-10-10", 3, 0, "PLAT-1"),
            ("PLAT-5", "PLAT", "Sub-task", "Future work", "To Do", EMPLOYEE, "2026-11-10", "2026-11-20", 4, 0, "PLAT-1"),
            ("PLAT-6", "PLAT", "Sub-task", "Someone else's", "In-Progress", "Bilal Ahmed", "2026-10-01", "2026-10-30", 4, 0, "PLAT-1"),
        ]
        conn.executemany("INSERT INTO canonical_issues (run_id, issue_key, project_key, issue_type, summary, status, assignee,"
                         " start_date, due_date, original_estimate_hours, total_hours_logged, epic_key) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                         [(RUN, *row) for row in issues])
        logs = [("w1", "PLAT-2", "2026-10-19", 8), ("w2", "PLAT-2", "2026-10-20", 7.5), ("w3", "PLAT-2", "2026-10-21", 3),
                ("w4", "PLAT-2", "2026-10-12", 8), ("w5", "PLAT-2", "2026-10-13", 8), ("w6", "PLAT-2", "2026-10-14", 8),
                ("w7", "PLAT-2", "2026-10-15", 8), ("w8", "PLAT-2", "2026-10-16", 8)]
        conn.executemany("INSERT INTO canonical_worklogs VALUES (?, ?, ?, 'PLAT', ?, ?, '', ?, '', ?)",
                         [(RUN, wid, key, EMPLOYEE, EMPLOYEE, day, hours) for wid, key, day, hours in logs])


class MyDayServiceTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(self.tmp.cleanup)
        self.db = Path(self.tmp.name) / "capacity.db"
        make_db(self.db)
        leave = patch("my_day_service.build_rlt_leave_snapshot", return_value={"daily": [
            {"assignee": EMPLOYEE, "period_day": "2026-10-28", "planned_not_taken_hours": 8, "jira_task_ids": "RLT-9"}]})
        leave.start()
        self.addCleanup(leave.stop)

    def test_expected_hours_respects_weekends_and_holidays(self):
        profiles = [{"from_date": "2026-01-01", "to_date": "2026-12-31", "standard_hours_per_day": 8, "holiday_dates": ["2026-10-23"]}]
        self.assertEqual(expected_hours(profiles, date(2026, 10, 24)), (0.0, "weekend"))
        self.assertEqual(expected_hours(profiles, date(2026, 10, 23)), (0.0, "holiday"))
        self.assertEqual(expected_hours(profiles, date(2026, 10, 22)), (8.0, "workday"))

    def test_day_status_grace_and_sync_boundary(self):
        today = date(2026, 10, 22)
        self.assertEqual(day_status(date(2026, 10, 20), today, "2026-10-21", 8, 7.5, 0)[0], "complete")
        self.assertEqual(day_status(date(2026, 10, 21), today, "2026-10-21", 8, 3, 0), ("missing", 5.0))
        self.assertEqual(day_status(date(2026, 10, 21), today, "2026-10-20", 8, 3, 0)[0], "unsynced")
        self.assertEqual(day_status(today, today, "2026-10-21", 8, 2, 0)[0], "today")

    def test_build_my_day_summarizes_tasks_week_gaps_and_streak(self):
        payload = build_my_day(self.db, "ayesha khan", date(2026, 10, 22))
        self.assertEqual(payload["missed"]["date"], "2026-10-21")
        self.assertEqual(payload["missed"]["missing"], 5.0)
        self.assertEqual([d["status"] for d in payload["week"]], ["complete", "complete", "missing", "today", "future"])
        self.assertEqual(payload["streak"]["weeks"], 1)
        keys = [t["issue_key"] for t in payload["tasks"]]
        self.assertEqual(keys, ["PLAT-4", "PLAT-2"])
        self.assertTrue(payload["tasks"][0]["overdue"])
        self.assertEqual(payload["tasks"][1]["epic_name"], "Jira Incremental Sync")
        self.assertEqual(payload["epics"][0]["epic_key"], "PLAT-1")
        self.assertEqual(payload["epics"][0]["month_hours"], 58.5)
        self.assertEqual(payload["upcoming"][0], {"date": "2026-10-23", "kind": "holiday", "title": "Official holiday", "detail": "No hours expected"})
        self.assertEqual(payload["upcoming"][1]["kind"], "leave")
        self.assertFalse(payload["live"])

    def test_live_worklogs_replace_this_weeks_canonical_hours(self):
        live = [{"issue_key": "PLAT-2", "date": "2026-10-22", "hours": 6, "summary": "Refactor sync retry"},
                {"issue_key": "PLAT-2", "date": "2026-10-21", "hours": 8, "summary": "Refactor sync retry"}]
        payload = build_my_day(self.db, EMPLOYEE, date(2026, 10, 22), live)
        self.assertTrue(payload["live"])
        self.assertEqual(payload["today"]["logged"], 6)
        self.assertIsNone(payload["missed"])
        self.assertEqual([d["logged"] for d in payload["week"][:2]], [0, 0])

    def test_list_people_and_missing_run(self):
        self.assertEqual(list_people(self.db), [EMPLOYEE])
        empty = Path(self.tmp.name) / "empty.db"
        with self.assertRaises(RuntimeError):
            build_my_day(empty, EMPLOYEE, date(2026, 10, 22))


if __name__ == "__main__":
    unittest.main()
