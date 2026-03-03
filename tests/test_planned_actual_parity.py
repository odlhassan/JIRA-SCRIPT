from __future__ import annotations

import unittest

from planned_actual_table_view_service import build_snapshot_payload


class PlannedActualParityTests(unittest.TestCase):
    def test_fixed_fixture_totals_and_row_contract_parity(self):
        hierarchy = {
            "epics": [
                {
                    "issue_key": "O2-EP1",
                    "project_key": "O2",
                    "project_name": "OmniConnect",
                    "summary": "Epic One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 40.0,
                    "planned_start": "2026-02-01",
                    "planned_due": "2026-02-20",
                },
                {
                    "issue_key": "O3-EP1",
                    "project_key": "O3",
                    "project_name": "Orbit",
                    "summary": "Epic Two",
                    "status": "Resolved",
                    "assignee": "Bob",
                    "estimate_hours": 24.0,
                    "planned_start": "2026-02-02",
                    "planned_due": "2026-02-18",
                },
            ],
            "stories": [
                {
                    "issue_key": "O2-ST1",
                    "project_key": "O2",
                    "project_name": "OmniConnect",
                    "epic_key": "O2-EP1",
                    "summary": "Story One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 16.0,
                    "planned_start": "2026-02-02",
                    "planned_due": "2026-02-10",
                },
                {
                    "issue_key": "O3-ST1",
                    "project_key": "O3",
                    "project_name": "Orbit",
                    "epic_key": "O3-EP1",
                    "summary": "Story Two",
                    "status": "Resolved",
                    "assignee": "Bob",
                    "estimate_hours": 12.0,
                    "planned_start": "2026-02-03",
                    "planned_due": "2026-02-12",
                },
            ],
            "subtasks": [
                {
                    "issue_key": "O2-SUB1",
                    "project_key": "O2",
                    "project_name": "OmniConnect",
                    "epic_key": "O2-EP1",
                    "story_key": "O2-ST1",
                    "summary": "Sub One",
                    "status": "In Progress",
                    "assignee": "Alice",
                    "estimate_hours": 8.0,
                    "planned_start": "2026-02-03",
                    "planned_due": "2026-02-04",
                },
                {
                    "issue_key": "O3-SUB1",
                    "project_key": "O3",
                    "project_name": "Orbit",
                    "epic_key": "O3-EP1",
                    "story_key": "O3-ST1",
                    "summary": "Sub Two",
                    "status": "Resolved",
                    "assignee": "Bob",
                    "estimate_hours": 6.0,
                    "planned_start": "2026-02-05",
                    "planned_due": "2026-02-06",
                },
            ],
        }
        actual_hours_by_subtask = {
            "O2-SUB1": 6.5,
            "O3-SUB1": 10.0,
        }

        rows, totals, _options = build_snapshot_payload(
            hierarchy=hierarchy,
            actual_hours_by_subtask=actual_hours_by_subtask,
            selected_projects={"O2", "O3"},
            selected_statuses={"in progress", "resolved"},
            selected_assignees={"alice", "bob"},
        )

        # Parity baseline for fixed fixture: strict total-hours contract.
        self.assertEqual(float(totals.get("planned_hours") or 0.0), 64.0)
        self.assertEqual(float(totals.get("actual_hours") or 0.0), 16.5)
        self.assertEqual(float(totals.get("variance_hours") or 0.0), 47.5)
        self.assertEqual(int(totals.get("project_count") or 0), 2)
        self.assertEqual(int(totals.get("epic_count") or 0), 2)
        self.assertEqual(int(totals.get("story_count") or 0), 2)
        self.assertEqual(int(totals.get("subtask_count") or 0), 2)

        row_types = [str(item.get("row_type") or "") for item in rows]
        self.assertEqual(row_types.count("project"), 2)
        self.assertEqual(row_types.count("epic"), 2)

        by_project = {
            str(item.get("project_key")): item
            for item in rows
            if str(item.get("row_type")) == "project"
        }
        self.assertAlmostEqual(float(by_project["O2"]["planned_hours"]), 40.0, places=6)
        self.assertAlmostEqual(float(by_project["O2"]["actual_hours"]), 6.5, places=6)
        self.assertAlmostEqual(float(by_project["O3"]["planned_hours"]), 24.0, places=6)
        self.assertAlmostEqual(float(by_project["O3"]["actual_hours"]), 10.0, places=6)


if __name__ == "__main__":
    unittest.main()
