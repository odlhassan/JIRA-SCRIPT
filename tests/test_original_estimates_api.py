from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from report_server import _resolve_capacity_runtime_paths, create_report_server_app

OE_FIXTURE_RUN_ID = "test-original-estimates-run"
CANON_STALE_AT = "2020-01-01T00:00:00+00:00"


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


def _seed_canonical_for_original_estimates(root: Path, hierarchy: dict, *, run_id: str = OE_FIXTURE_RUN_ID) -> None:
    """Populate canonical_issues + last-success metadata so the OEH summary reads live hierarchy from SQLite."""
    db_path = _resolve_capacity_runtime_paths(root)["db_path"]
    rows: list[tuple] = []
    for epic in hierarchy["epics"]:
        ek = str(epic["issue_key"])
        rows.append(
            (
                run_id,
                "",
                ek,
                str(epic["project_key"]),
                "Epic",
                str(epic["summary"]),
                str(epic["status"]),
                str(epic["assignee"]),
                str(epic["planned_start"]),
                str(epic["planned_due"]),
                "",
                "",
                "",
                float(epic["estimate_hours"]),
                0.0,
                "",
                "",
                "",
                ek,
                "{}",
            )
        )
    for story in hierarchy["stories"]:
        sk = str(story["issue_key"])
        ek = str(story["epic_key"])
        rows.append(
            (
                run_id,
                "",
                sk,
                str(story["project_key"]),
                "Story",
                str(story["summary"]),
                str(story["status"]),
                str(story["assignee"]),
                str(story["planned_start"]),
                str(story["planned_due"]),
                "",
                "",
                "",
                float(story["estimate_hours"]),
                0.0,
                "",
                "",
                sk,
                ek,
                "{}",
            )
        )
    for sub in hierarchy["subtasks"]:
        k = str(sub["issue_key"])
        sk = str(sub["story_key"])
        ek = str(sub["epic_key"])
        rows.append(
            (
                run_id,
                "",
                k,
                str(sub["project_key"]),
                "Sub-task",
                str(sub["summary"]),
                str(sub["status"]),
                str(sub["assignee"]),
                str(sub["planned_start"]),
                str(sub["planned_due"]),
                "",
                "",
                "",
                float(sub["estimate_hours"]),
                0.0,
                "",
                "",
                sk,
                ek,
                "{}",
            )
        )
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT OR REPLACE INTO canonical_refresh_runs(
                run_id, scope_year, managed_project_keys_json, started_at_utc, ended_at_utc,
                status, trigger_source, error_message, stats_json,
                progress_step, progress_pct, cancel_requested, updated_at_utc
            ) VALUES (?, 2026, '["O2"]', ?, ?, 'success', 'test', '', '{}', 'done', 100, 0, ?)
            """,
            (run_id, CANON_STALE_AT, CANON_STALE_AT, CANON_STALE_AT),
        )
        conn.execute(
            "UPDATE canonical_refresh_state SET active_run_id=?, last_success_run_id=?, updated_at_utc=? WHERE id=1",
            ("", run_id, CANON_STALE_AT),
        )
        conn.executemany(
            """
            INSERT OR REPLACE INTO canonical_issues(
                run_id, issue_id, issue_key, project_key, issue_type, summary, status, assignee,
                start_date, due_date, created_utc, updated_utc, resolved_stable_since_date,
                original_estimate_hours, total_hours_logged, fix_type, parent_issue_key,
                story_key, epic_key, raw_payload_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            rows,
        )
        conn.commit()


class OriginalEstimatesApiTests(unittest.TestCase):
    def test_summary_rollups_and_filters(self):
        hierarchy = _hierarchy_fixture()
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            app = _build_app(root)
            _seed_canonical_for_original_estimates(root, hierarchy)
            client = app.test_client()
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
            self.assertEqual(summary.get("source"), "canonical_hierarchy")
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
            _seed_canonical_for_original_estimates(root, hierarchy)
            client = app.test_client()

            refresh_resp = client.post(
                "/api/original-estimates/refresh",
                json={"from": "2026-02-01", "to": "2026-02-28", "projects": ["O2"]},
            )
            self.assertEqual(refresh_resp.status_code, 200)

            before = client.get("/api/original-estimates/summary?from=2026-02-01&to=2026-02-28&projects=O2").get_json() or {}
            epic_one_before = next((item for item in (before.get("epics") or []) if item.get("issue_key") == "O2-EP1"), None)
            self.assertIsNotNone(epic_one_before)

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
            self.assertEqual(str(epic_one.get("summary")), str(epic_one_before.get("summary")))
            self.assertEqual(float(epic_one.get("original_estimate_hours") or 0.0), float(epic_one_before.get("original_estimate_hours") or 0.0))
            self.assertEqual(str(epic_two.get("summary")), "Epic Two")


if __name__ == "__main__":
    unittest.main()
