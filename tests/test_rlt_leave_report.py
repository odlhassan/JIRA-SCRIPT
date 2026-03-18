from __future__ import annotations

import os
import sys
import unittest
from datetime import date
from unittest.mock import patch

from generate_rlt_leave_report import (
    DEFAULT_SOURCE,
    SubtaskRow,
    WorklogRow,
    _compute_aggregates,
    _parse_args,
    _redistribute_continuous_leave_subtasks,
    _redistribute_continuous_leave_worklogs,
    classify_leave,
    infer_date_range_from_summary,
    is_clubbed_leave,
    is_defective_no_entry,
    normalize_subtask_dates,
    resolve_window_range,
)


class RltLeaveReportTests(unittest.TestCase):
    def _profile(self) -> dict[str, float | str]:
        return {
            "standard_hours_per_day": 8.0,
            "ramadan_hours_per_day": 6.5,
            "ramadan_start_date": "",
            "ramadan_end_date": "",
        }

    def _subtask(
        self,
        issue_key: str = "RLT-1",
        start_date: str = "2026-01-06",
        due_date: str = "2026-01-10",
        estimate: float = 24.0,
        leave_classification: str = "Planned",
    ) -> SubtaskRow:
        return SubtaskRow(
            issue_key=issue_key,
            issue_id="1",
            summary="Continuous leave",
            status="Planned Leave",
            assignee="Alice",
            parent_task_key="RLT-100",
            parent_task_assignee="Alice",
            created="",
            updated="",
            start_date=start_date,
            due_date=due_date,
            original_estimate_hours=estimate,
            timespent_hours=0.0,
            leave_type_raw="Planned Leave",
            leave_classification=leave_classification,
            total_worklog_hours=24.0,
            planned_date_for_bucket=start_date,
            clubbed_leave="Yes",
            no_entry_flag="No",
        )

    def test_classification_precedence_leave_type_over_status(self):
        self.assertEqual(
            classify_leave("Unplanned Leave", "Planned Leave", "planned leave"),
            "Unplanned",
        )

    def test_blank_type_fallback_mapping(self):
        self.assertEqual(classify_leave("", "Sick Leave", "Leave"), "Unplanned")
        self.assertEqual(classify_leave("", "Considered in Roadmap & Queued", "Annual leave"), "Planned")

    def test_defective_no_entry_rule(self):
        self.assertTrue(is_defective_no_entry("Planned", 0.0, 0.0, "", ""))
        self.assertFalse(is_defective_no_entry("Planned", 8.0, 0.0, "", ""))

    def test_clubbed_leave_rule(self):
        self.assertTrue(is_clubbed_leave(9.0, 0.0, "", ""))
        self.assertTrue(is_clubbed_leave(0.0, 9.0, "", ""))
        self.assertTrue(is_clubbed_leave(0.0, 0.0, "2026-01-01", "2026-01-02"))
        self.assertFalse(is_clubbed_leave(8.0, 8.0, "2026-01-01", "2026-01-01"))

    def test_window_resolver_prev_current_next(self):
        f, t = resolve_window_range(today=date(2026, 2, 20))
        self.assertEqual(f, "2026-01-01")
        self.assertEqual(t, "2026-03-31")

    def test_parse_args_defaults_to_canonical_source(self):
        with patch.object(sys, "argv", ["generate_rlt_leave_report.py"]), patch.dict(os.environ, {}, clear=False):
            args = _parse_args()
        self.assertEqual(args.source, DEFAULT_SOURCE)

    def test_parse_args_allows_explicit_legacy_jira_source(self):
        with patch.object(sys, "argv", ["generate_rlt_leave_report.py", "--source", "jira"]):
            args = _parse_args()
        self.assertEqual(args.source, "jira")

    def test_infer_date_range_from_summary(self):
        self.assertEqual(
            infer_date_range_from_summary("Maternity Leaves 1 Jan - 31 March 2026"),
            ("2026-01-01", "2026-03-31"),
        )
        self.assertEqual(
            infer_date_range_from_summary("Annual Leaves 1 April - 30 April 2026"),
            ("2026-04-01", "2026-04-30"),
        )

    def test_normalize_subtask_dates_only_fills_missing(self):
        self.assertEqual(
            normalize_subtask_dates("", "", "Maternity Leaves 1 Jan - 31 March 2026"),
            ("2026-01-01", "2026-03-31"),
        )
        self.assertEqual(
            normalize_subtask_dates("2026-01-15", "", "Maternity Leaves 1 Jan - 31 March 2026"),
            ("2026-01-15", "2026-03-31"),
        )
        self.assertEqual(
            normalize_subtask_dates("2026-01-07", "2026-01-10", "Maternity Leaves 1 Jan - 31 March 2026"),
            ("2026-01-07", "2026-01-10"),
        )

    def test_weekly_monthly_bucketing_with_missing_dates(self):
        subtasks = [
            SubtaskRow(
                issue_key="RLT-1",
                issue_id="1",
                summary="Planned leave",
                status="Planned Leave",
                assignee="Alice",
                parent_task_key="RLT-100",
                parent_task_assignee="Alice",
                created="",
                updated="",
                start_date="2026-02-10",
                due_date="2026-02-10",
                original_estimate_hours=8.0,
                timespent_hours=0.0,
                leave_type_raw="Planned Leave",
                leave_classification="Planned",
                total_worklog_hours=0.0,
                planned_date_for_bucket="2026-02-10",
                clubbed_leave="No",
                no_entry_flag="No",
            ),
            SubtaskRow(
                issue_key="RLT-2",
                issue_id="2",
                summary="Planned leave no entry",
                status="Planned Leave",
                assignee="Alice",
                parent_task_key="RLT-100",
                parent_task_assignee="Alice",
                created="",
                updated="",
                start_date="",
                due_date="",
                original_estimate_hours=0.0,
                timespent_hours=0.0,
                leave_type_raw="Planned Leave",
                leave_classification="Planned",
                total_worklog_hours=0.0,
                planned_date_for_bucket="",
                clubbed_leave="No",
                no_entry_flag="Yes",
            ),
            SubtaskRow(
                issue_key="RLT-3",
                issue_id="3",
                summary="Sick leave",
                status="Sick Leave",
                assignee="Alice",
                parent_task_key="RLT-100",
                parent_task_assignee="Alice",
                created="",
                updated="",
                start_date="",
                due_date="",
                original_estimate_hours=0.0,
                timespent_hours=8.0,
                leave_type_raw="",
                leave_classification="Unplanned",
                total_worklog_hours=8.0,
                planned_date_for_bucket="",
                clubbed_leave="No",
                no_entry_flag="No",
            ),
        ]
        worklogs = [
            WorklogRow(issue_key="RLT-3", started_raw="", started_date="2026-02-12", author="Alice", hours_logged=8.0),
        ]

        out = _compute_aggregates(
            subtasks,
            worklogs,
            "2026-01-01",
            "2026-03-31",
            {
                "standard_hours_per_day": 8.0,
                "ramadan_hours_per_day": 6.5,
                "ramadan_start_date": "",
                "ramadan_end_date": "",
            },
        )
        self.assertEqual(len(out["daily"]), 2)
        self.assertEqual(len(out["weekly"]), 1)
        self.assertEqual(len(out["monthly"]), 1)
        summary = out["assignee_summary"][0]
        self.assertEqual(summary["planned_not_taken_hours"], 8.0)
        self.assertEqual(summary["planned_not_taken_no_entry_count"], 1)
        self.assertEqual(summary["unplanned_taken_hours"], 8.0)

    def test_redistribute_continuous_leave_even_split_weekdays(self):
        subtasks = [self._subtask(start_date="2026-01-05", due_date="2026-01-09", estimate=40.0)]
        worklogs = [WorklogRow(issue_key="RLT-1", started_raw="raw", started_date="2026-01-09", author="Alice", hours_logged=40.0)]
        out = _redistribute_continuous_leave_worklogs(subtasks, worklogs, self._profile())
        self.assertEqual(len(out), 5)
        self.assertEqual([r.started_date for r in out], ["2026-01-05", "2026-01-06", "2026-01-07", "2026-01-08", "2026-01-09"])
        self.assertTrue(all(date.fromisoformat(r.started_date).weekday() < 5 for r in out))
        self.assertTrue(all(abs(r.hours_logged - 8.0) < 1e-9 for r in out))
        self.assertEqual(round(sum(r.hours_logged for r in out), 2), 40.0)

    def test_redistribute_preserves_total_with_rounding(self):
        subtasks = [self._subtask(start_date="2026-01-05", due_date="2026-01-07", estimate=24.0)]
        worklogs = [WorklogRow(issue_key="RLT-1", started_raw="raw", started_date="2026-01-07", author="Alice", hours_logged=10.0)]
        out = _redistribute_continuous_leave_worklogs(subtasks, worklogs, self._profile())
        self.assertEqual(len(out), 2)
        self.assertEqual([r.hours_logged for r in out], [8.0, 2.0])
        self.assertEqual(round(sum(r.hours_logged for r in out), 2), 10.0)

    def test_redistribute_fallback_for_single_or_zero_weekday_range(self):
        subtasks = [
            self._subtask(issue_key="RLT-1", start_date="2026-01-10", due_date="2026-01-11", estimate=16.0),
            self._subtask(issue_key="RLT-2", start_date="2026-01-09", due_date="2026-01-09", estimate=8.0),
        ]
        worklogs = [
            WorklogRow(issue_key="RLT-1", started_raw="raw", started_date="2026-01-10", author="Alice", hours_logged=16.0),
            WorklogRow(issue_key="RLT-2", started_raw="raw", started_date="2026-01-09", author="Alice", hours_logged=8.0),
        ]
        out = _redistribute_continuous_leave_worklogs(subtasks, worklogs, self._profile())
        self.assertEqual([(r.issue_key, r.started_date, r.hours_logged) for r in out], [
            ("RLT-2", "2026-01-09", 8.0),
            ("RLT-1", "2026-01-10", 16.0),
        ])

    def test_redistribute_worklog_discards_overflow_past_date_range(self):
        subtasks = [self._subtask(start_date="2026-01-05", due_date="2026-01-07", estimate=24.0)]
        worklogs = [WorklogRow(issue_key="RLT-1", started_raw="raw", started_date="2026-01-07", author="Alice", hours_logged=40.0)]
        out = _redistribute_continuous_leave_worklogs(subtasks, worklogs, self._profile())
        self.assertEqual(len(out), 3)
        self.assertEqual([r.hours_logged for r in out], [8.0, 8.0, 8.0])
        self.assertEqual(round(sum(r.hours_logged for r in out), 2), 24.0)

    def test_nonqualifying_subtask_not_redistributed(self):
        subtasks = [self._subtask(start_date="", due_date="", estimate=4.0)]
        worklogs = [WorklogRow(issue_key="RLT-1", started_raw="raw", started_date="2026-01-09", author="Alice", hours_logged=8.0)]
        out = _redistribute_continuous_leave_worklogs(subtasks, worklogs, self._profile())
        self.assertEqual(len(out), 1)
        self.assertEqual(out[0].started_date, "2026-01-09")
        self.assertEqual(out[0].hours_logged, 8.0)

    def test_redistribute_subtask_estimate_and_dates(self):
        subtasks = [self._subtask(start_date="2026-01-05", due_date="2026-01-07", estimate=24.0)]
        out = _redistribute_continuous_leave_subtasks(subtasks, self._profile())
        self.assertEqual(len(out), 3)
        self.assertEqual([(r.start_date, r.due_date, r.original_estimate_hours) for r in out], [
            ("2026-01-05", "2026-01-05", 8.0),
            ("2026-01-06", "2026-01-06", 8.0),
            ("2026-01-07", "2026-01-07", 8.0),
        ])

    def test_redistribute_subtask_estimate_discards_overflow_past_due_date(self):
        subtasks = [self._subtask(start_date="2026-02-01", due_date="2026-02-03", estimate=48.0)]
        out = _redistribute_continuous_leave_subtasks(subtasks, self._profile())
        self.assertEqual(len(out), 3)
        self.assertEqual([r.original_estimate_hours for r in out], [8.0, 8.0, 8.0])
        self.assertEqual(round(sum(r.original_estimate_hours for r in out), 2), 24.0)

    def test_planned_not_taken_estimate_distributes_across_weekdays(self):
        subtasks = [
            SubtaskRow(
                issue_key="RLT-10",
                issue_id="10",
                summary="Planned leave",
                status="Planned Leave",
                assignee="Alice",
                parent_task_key="RLT-100",
                parent_task_assignee="Alice",
                created="",
                updated="",
                start_date="2026-01-08",
                due_date="2026-01-12",
                original_estimate_hours=16.0,
                timespent_hours=0.0,
                leave_type_raw="Planned Leave",
                leave_classification="Planned",
                total_worklog_hours=0.0,
                planned_date_for_bucket="2026-01-08",
                clubbed_leave="Yes",
                no_entry_flag="No",
            ),
        ]
        out = _compute_aggregates(
            subtasks,
            [],
            "2026-01-01",
            "2026-01-31",
            self._profile(),
        )
        daily = {(r["assignee"], r["period_day"]): r["planned_not_taken_hours"] for r in out["daily"]}
        self.assertEqual(daily[("Alice", "2026-01-08")], 8.0)
        self.assertEqual(daily[("Alice", "2026-01-09")], 8.0)
        self.assertEqual(daily.get(("Alice", "2026-01-10"), 0.0), 0.0)
        self.assertEqual(daily.get(("Alice", "2026-01-11"), 0.0), 0.0)
        self.assertEqual(daily.get(("Alice", "2026-01-12"), 0.0), 0.0)
        summary = out["assignee_summary"][0]
        self.assertEqual(summary["planned_not_taken_hours"], 16.0)


if __name__ == "__main__":
    unittest.main()
