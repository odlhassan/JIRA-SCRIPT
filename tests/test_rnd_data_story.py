from __future__ import annotations

import unittest

from generate_rnd_data_story import (
    _build_html,
    aggregate_epic_logged_hours,
    classify_status,
    epic_in_range,
    pending_hours,
    planned_committed_hours,
)


class RndDataStoryTests(unittest.TestCase):
    def test_epic_in_range_start_or_end(self):
        self.assertTrue(epic_in_range("2026-02-10", "2026-03-10", "2026-02-01", "2026-02-28"))
        self.assertTrue(epic_in_range("2026-01-01", "2026-02-20", "2026-02-01", "2026-02-28"))
        self.assertFalse(epic_in_range("", "", "2026-02-01", "2026-02-28"))
        self.assertTrue(epic_in_range("2026-02-01", "", "2026-02-01", "2026-02-28"))
        self.assertTrue(epic_in_range("", "2026-02-28", "2026-02-01", "2026-02-28"))

    def test_pending_hours_clamps_to_zero(self):
        self.assertEqual(pending_hours(10, 4), 6)
        self.assertEqual(pending_hours(8, 8), 0)
        self.assertEqual(pending_hours(5, 7), 0)

    def test_aggregate_epic_logged_hours(self):
        rows = [
            {"parent_epic_id": "O2-1", "hours_logged": 2.5},
            {"parent_epic_id": "O2-1", "hours_logged": 1.5},
            {"parent_epic_id": "FF-9", "hours_logged": 3},
            {"parent_epic_id": "", "hours_logged": 4},
        ]
        result = aggregate_epic_logged_hours(rows)
        self.assertEqual(result["O2-1"], 4.0)
        self.assertEqual(result["FF-9"], 3.0)
        self.assertNotIn("", result)

    def test_status_classification(self):
        self.assertEqual(classify_status("Resolved"), "closed")
        self.assertEqual(classify_status("resolved!"), "closed")
        self.assertEqual(classify_status("In Progress"), "open")
        self.assertEqual(classify_status("In-Progress"), "open")
        self.assertEqual(classify_status("IN PROGRESS"), "open")
        self.assertEqual(classify_status("To Do"), "other")

    def test_investable_arithmetic_consistency(self):
        available_capacity_hours = 120
        planned_taken = 10
        planned_not_taken = 5
        unplanned_taken = 5
        work_on_plate = 60
        capacity_after_leaves = available_capacity_hours - (planned_taken + planned_not_taken + unplanned_taken)
        investable_more = capacity_after_leaves - work_on_plate
        self.assertEqual(capacity_after_leaves, 100)
        self.assertEqual(investable_more, 40)

    def test_planned_committed_hours_uses_epic_estimate_date_rule_and_excludes_rlt(self):
        epics = [
            {"project_key": "O2", "start_date": "2026-02-10", "end_date": "", "original_estimate_hours": 100},
            {"project_key": "FF", "start_date": "", "end_date": "2026-02-20", "original_estimate_hours": 200},
            {"project_key": "O2", "start_date": "2026-01-05", "end_date": "2026-03-01", "original_estimate_hours": 300},
            {"project_key": "RLT", "start_date": "2026-02-11", "end_date": "", "original_estimate_hours": 400},
        ]
        total = planned_committed_hours(epics, "2026-02-01", "2026-02-28")
        self.assertEqual(total, 300.0)

    def test_html_uses_scoped_subtasks_endpoint_for_shared_scope(self):
        html = _build_html(
            {
                "department_name": "Research and Development (RnD)",
                "generated_at": "2026-03-12 00:00 UTC",
                "source_files": {},
                "defaults": {"from_date": "2026-02-01", "to_date": "2026-02-28"},
                "default_employee_count": 0,
                "epics": [],
                "epic_logged_hours_by_key": {},
                "worklog_rows": [],
                "planned_epic_rows": [],
                "project_actual_rows": [],
                "page1_dataset": {},
                "capacity_profiles": [],
                "leave_daily_rows": [],
            }
        )
        self.assertIn('const SCOPED_SUBTASKS_ENDPOINT="/api/scoped-subtasks";', html)
        self.assertIn("async function loadScopedSubtasks(fromIso,toIso,mode)", html)
        self.assertIn("function buildEpicRowsFromScopedPayload(scopedPayload)", html)


if __name__ == "__main__":
    unittest.main()
