from __future__ import annotations

import json
import sqlite3
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from generate_phase_rmi_gantt_html import _build_html, _load_team_rmi_payload
from sync_team_rmi_gantt_sqlite import sync_team_rmi_gantt_to_sqlite


def _write_work_items_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(
        [
            "project_key",
            "issue_key",
            "work_item_id",
            "work_item_type",
            "jira_issue_type",
            "fix_type",
            "summary",
            "status",
            "resolved_stable_since_date",
            "start_date",
            "end_date",
            "actual_start_date",
            "actual_end_date",
            "original_estimate",
            "original_estimate_hours",
            "assignee",
            "total_hours_logged",
            "priority",
            "parent_issue_key",
            "parent_work_item_id",
            "parent_jira_url",
            "jira_url",
            "Latest IPP Meeting",
            "Jira IPP RMI Dates Altered",
            "IPP Actual Date (Production Date)",
            "IPP Remarks",
            "IPP Actual Date Matches Jira End Date",
            "created",
            "updated",
        ]
    )

    # Epic master rows.
    ws.append(["P1", "P1-100", "", "", "Epic", "", "Epic Alpha", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "https://jira.example/browse/P1-100", "", "", "", "", "", "", ""])
    ws.append(["P1", "P1-200", "", "", "Epic", "", "Epic Beta", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""])

    # Included: mapped team + epic.
    ws.append(["P1", "P1-101", "", "", "Story", "", "Story A", "", "", "2026-02-01", "2026-02-10", "", "", "", 16, "Alice Smith", "", "", "P1-100", "", "", "", "", "", "", "", "", "", ""])
    # Included: same team + same epic.
    ws.append(["P1", "P1-102", "", "", "Story", "", "Story B", "", "", "2026-02-05", "2026-02-20", "", "", "", 8, " Alice   Smith ", "", "", "P1-100", "", "", "", "", "", "", "", "", "", ""])
    # Included: unmapped lane.
    ws.append(["P1", "P1-103", "", "", "Story", "", "Story C", "", "", "2026-03-01", "2026-03-05", "", "", "", 4, "Unknown Person", "", "", "P1-200", "", "", "", "", "", "", "", "", "", ""])

    # Excluded: missing epic.
    ws.append(["P1", "P1-104", "", "", "Story", "", "Story D", "", "", "2026-03-01", "2026-03-05", "", "", "", 4, "Alice Smith", "", "", "", "", "", "", "", "", "", "", "", "", ""])
    # Excluded: missing dates.
    ws.append(["P1", "P1-105", "", "", "Story", "", "Story E", "", "", "", "", "", "", "", 5, "Alice Smith", "", "", "P1-100", "", "", "", "", "", "", "", "", "", ""])
    # Excluded: missing estimate.
    ws.append(["P1", "P1-106", "", "", "Story", "", "Story F", "", "", "2026-03-01", "2026-03-06", "", "", "", 0, "Alice Smith", "", "", "P1-100", "", "", "", "", "", "", "", "", "", ""])

    wb.save(path)


def _seed_performance_teams(db_path: Path) -> None:
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS performance_teams (
                team_name TEXT PRIMARY KEY,
                team_leader TEXT NOT NULL DEFAULT '',
                assignees_json TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            INSERT INTO performance_teams (team_name, team_leader, assignees_json, updated_at)
            VALUES (?, ?, ?, ?)
            """,
            ("Technical Writing", "Alice Smith", json.dumps(["Alice Smith"]), "2026-03-02 00:00:00"),
        )
        conn.commit()
    finally:
        conn.close()


class TeamRmiGanttSyncTests(unittest.TestCase):
    def test_sync_builds_sqlite_snapshot_and_applies_rules(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            xlsx_path = root / "1_jira_work_items_export.xlsx"
            _seed_performance_teams(db_path)
            _write_work_items_xlsx(xlsx_path)

            snapshot = sync_team_rmi_gantt_to_sqlite(xlsx_path, db_path)

            self.assertEqual(snapshot["total_story_rows"], 6)
            self.assertEqual(snapshot["included_story_rows"], 3)
            self.assertEqual(snapshot["excluded_missing_epic"], 1)
            self.assertEqual(snapshot["excluded_missing_dates"], 1)
            self.assertEqual(snapshot["excluded_missing_estimate"], 1)

            items = snapshot["items"]
            self.assertEqual(len(items), 2)

            tw = next(i for i in items if i["team_name"] == "Technical Writing")
            self.assertEqual(tw["epic_key"], "P1-100")
            self.assertEqual(tw["planned_start"], "2026-02-01")
            self.assertEqual(tw["planned_end"], "2026-02-20")
            self.assertEqual(tw["planned_hours"], 24.0)
            self.assertEqual(tw["planned_man_days"], 3.0)
            self.assertEqual(tw["story_count"], 2)
            self.assertEqual(tw["epic_url"], "https://jira.example/browse/P1-100")

            unmapped = next(i for i in items if i["team_name"] == "Unmapped Team")
            self.assertEqual(unmapped["epic_key"], "P1-200")
            self.assertEqual(unmapped["is_unmapped_team"], 1)
            self.assertIn("/browse/P1-200", unmapped["epic_url"])

            conn = sqlite3.connect(db_path)
            conn.row_factory = sqlite3.Row
            try:
                db_items = conn.execute("SELECT * FROM team_rmi_gantt_items").fetchall()
                self.assertEqual(len(db_items), 2)
                meta = conn.execute("SELECT * FROM team_rmi_gantt_snapshot_meta WHERE id = 1").fetchone()
                self.assertIsNotNone(meta)
                self.assertEqual(int(meta["included_story_rows"]), 3)
            finally:
                conn.close()

    def test_html_loader_and_render_include_team_payload(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            xlsx_path = root / "1_jira_work_items_export.xlsx"
            _seed_performance_teams(db_path)
            _write_work_items_xlsx(xlsx_path)
            sync_team_rmi_gantt_to_sqlite(xlsx_path, db_path)

            payload = _load_team_rmi_payload(db_path)
            self.assertEqual(payload["team_names"][0], "Technical Writing")
            self.assertIn("Unmapped Team", payload["team_names"])
            self.assertEqual(len(payload["items"]), 2)

            html = _build_html(
                {
                    "generated_at": "2026-03-02 00:00 UTC",
                    "source_file": payload["source_file"],
                    "team_names": payload["team_names"],
                    "items": payload["items"],
                    "snapshot_meta": payload["snapshot_meta"],
                }
            )
            self.assertIn("Team Owner RMI Gantt", html)
            self.assertIn("Technical Writing", html)
            self.assertIn("Unmapped Team", html)
            self.assertIn("Cards open Jira epic links", html)
            self.assertIn("href=", html)
            self.assertIn("function defaultRange()", html)
            self.assertIn("return { from: startOfYear(today), to: endOfYear(today) };", html)

    def test_loader_handles_missing_snapshot_tables(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            db_path = root / "assignee_hours_capacity.db"
            conn = sqlite3.connect(db_path)
            conn.close()
            payload = _load_team_rmi_payload(db_path)
            self.assertEqual(payload["team_names"], [])
            self.assertEqual(payload["items"], [])


if __name__ == "__main__":
    unittest.main()
