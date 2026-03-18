from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from generate_assignee_hours_report import _build_html as build_assignee_html
from generate_employee_performance_report import _build_html as build_employee_perf_html
from generate_nested_view_html import _build_html as build_nested_html
from generate_phase_rmi_gantt_html import _build_html as build_team_rmi_gantt_html
from generate_planned_rmis_html import _build_html as build_planned_rmis_html
from generate_rnd_data_story import _build_html as build_rnd_story_html
from openpyxl import Workbook
from report_server import create_report_server_app


class ReportUiSmokeTests(unittest.TestCase):
    def test_assignee_header_and_drawer_controls_exist(self):
        payload = {
            "rows": [],
            "projects": [],
            "default_from": "2026-01-01",
            "default_to": "2026-01-31",
            "capacity_profiles": [],
            "leave_daily_rows": [],
            "generated_at": "2026-02-21 00:00 UTC",
        }
        html = build_assignee_html(payload)
        self.assertIn('class="enterprise-header"', html)
        self.assertIn('id="open-capacity-settings"', html)
        self.assertIn('id="settings-drawer"', html)
        self.assertIn('id="settings-drawer-overlay"', html)
        self.assertIn('id="capacity-employees"', html)
        self.assertIn('id="capacity-profile-select"', html)
        self.assertIn('id="capacity-profile-apply"', html)
        self.assertIn('id="actual-hours-mode"', html)
        self.assertIn("/api/capacity/profiles", html)
        self.assertIn("/api/capacity/calculate", html)
        self.assertIn("/api/actual-hours/aggregate", html)

    def test_employee_performance_simple_score_drawer_controls_exist(self):
        payload = {
            "worklogs": [],
            "work_items": [],
            "leave_rows": [],
            "settings": {},
            "teams": [],
            "projects": [],
            "default_from": "2026-01-01",
            "default_to": "2026-01-31",
            "leave_hours_per_day": 8,
            "entities_catalog": [],
            "managed_fields": [],
            "capacity_profiles": [],
            "simple_scoring": [],
            "jira_browse_base": "https://example.atlassian.net/browse",
            "generated_at": "2026-02-21 00:00 UTC",
        }
        html = build_employee_perf_html(payload)
        self.assertIn('id="score-detail-drawer"', html)
        self.assertIn('id="score-detail-drawer-close"', html)
        self.assertIn('id="score-detail-drawer-body"', html)
        self.assertIn('id="header-average-performance-value"', html)
        self.assertIn("Average Performance", html)
        self.assertIn("Simple Score Details", html)
        self.assertIn("Planned Due Date", html)
        self.assertIn("Last Logged Date", html)
        self.assertIn("Actual Complete Date", html)

    def test_nested_view_options_and_profile_controls_exist(self):
        payload = {
            "generated_at": "2026-02-21 00:00 UTC",
            "source_file": "nested view.xlsx",
            "rows": [],
            "capacity_profiles": [],
            "leave_daily_rows": [],
            "leave_subtask_rows": [],
        }
        html = build_nested_html(payload)
        self.assertIn('class="scorecards"', html)
        self.assertIn('id="view-options"', html)
        self.assertIn('id="view-options-toggle"', html)
        self.assertIn('id="view-options-menu"', html)
        self.assertIn('id="theme-toggle"', html)
        self.assertIn('id="toggle-density"', html)
        self.assertIn('id="toggle-no-entry"', html)
        self.assertIn('id="toggle-product"', html)
        self.assertIn('id="date-filter-from"', html)
        self.assertIn('id="date-filter-to"', html)
        self.assertIn('id="planned-hours-source"', html)
        self.assertIn('id="extended-actual-hours-toggle"', html)
        self.assertIn('id="project-filter-progress"', html)
        self.assertIn('id="team-filter-progress"', html)
        self.assertIn('id="actual-hours-mode"', html)
        self.assertIn('href="/settings/capacity"', html)
        self.assertIn('id="score-total-capacity-formula"', html)
        self.assertIn('id="score-total-capacity-formula-hours"', html)
        self.assertIn('id="score-total-leaves-planned-formula"', html)
        self.assertIn("Availability", html)
        self.assertIn("Total Capacity (Hours) - Total Leaves Planned", html)
        self.assertIn("function subtaskMatchesActualHoursMode(row, bounds)", html)
        self.assertIn(
            "Sum(All Logged Hours for subtasks whose planned Start OR Due date is within selected range)",
            html,
        )
        self.assertIn(
            "Sum(Logged Hours in selected date range for subtasks with worklog dates in selected range)",
            html,
        )

    def test_nested_capacity_endpoints_present(self):
        payload = {
            "generated_at": "2026-02-21 00:00 UTC",
            "source_file": "nested view.xlsx",
            "rows": [],
            "capacity_profiles": [],
            "leave_daily_rows": [],
            "leave_subtask_rows": [],
        }
        html = build_nested_html(payload)
        self.assertIn("/api/capacity/profiles", html)
        self.assertIn("/api/nested-view/actual-hours", html)
        self.assertIn("/api/actual-hours/aggregate", html)
        self.assertIn("/api/manage-fields?include_inactive=0", html)
        self.assertIn("hasCapacityApi", html)
        self.assertIn("hasManagedFieldsApi", html)
        self.assertIn("refreshManagedFieldsFromApi", html)
        self.assertIn("evaluateManagedField", html)
        self.assertIn("leave_subtask_rows", html)

    def test_nested_date_filter_uses_active_selection_bounds(self):
        payload = {
            "generated_at": "2026-02-21 00:00 UTC",
            "source_file": "nested view.xlsx",
            "rows": [],
            "capacity_profiles": [],
            "leave_daily_rows": [],
            "leave_subtask_rows": [],
        }
        html = build_nested_html(payload)
        self.assertIn(
            "const bounds = getDateFilterBoundsFor(activeSelection.dateFrom, activeSelection.dateTo);",
            html,
        )
        self.assertIn("function matchesDateFilter(row, selection)", html)
        self.assertIn("const activeSelection = buildScorecardSelectionSnapshot(selection);", html)

    def test_rnd_story_controls_exist(self):
        payload = {
            "department_name": "Research and Development (RnD)",
            "generated_at": "2026-02-21 00:00 UTC",
            "source_files": {},
            "defaults": {"from_date": "2026-02-01", "to_date": "2026-02-28"},
            "epics": [],
            "epic_logged_hours_by_key": {},
            "worklog_rows": [],
            "capacity_profiles": [],
            "leave_daily_rows": [],
        }
        html = build_rnd_story_html(payload)
        self.assertIn('id="from-date"', html)
        self.assertIn('id="to-date"', html)
        self.assertIn('id="capacity-profile-select"', html)
        self.assertIn('id="apply-profile-btn"', html)
        self.assertIn('id="actual-hours-mode"', html)
        self.assertIn('id="kpi-capacity-after-leaves"', html)
        self.assertIn('id="kpi-hours-required-projects"', html)
        self.assertIn("funnel-hours-required-track", html)
        self.assertIn("funnel-hours-required-val", html)
        self.assertIn("/api/capacity?from=", html)
        self.assertIn("/api/actual-hours/aggregate", html)
        self.assertIn("/api/scoped-subtasks", html)
        self.assertIn("/api/manage-fields?include_inactive=0", html)
        self.assertIn("evaluateManagedField", html)
        self.assertIn("managedFieldFormulaText", html)

    def test_planned_rmis_actual_mode_controls_exist(self):
        payload = {
            "rows": [],
            "generated_at": "2026-02-24 00:00 UTC",
            "source_file": "nested view.xlsx",
            "default_from": "2026-02-01",
            "default_to": "2026-02-28",
        }
        html = build_planned_rmis_html(payload)
        self.assertIn('id=\'actual-hours-mode\'', html)
        self.assertIn('id=\'actual-hours-status\'', html)
        self.assertIn("/api/actual-hours/aggregate", html)

    def test_team_rmi_gantt_contains_team_lanes_and_clickable_epic_links(self):
        payload = {
            "generated_at": "2026-03-02 00:00 UTC",
            "source_file": "1_jira_work_items_export.xlsx",
            "team_names": ["Technical Writing", "Unmapped Team"],
            "items": [
                {
                    "team_name": "Technical Writing",
                    "epic_key": "P1-100",
                    "epic_name": "Epic Alpha",
                    "epic_url": "https://jira.example/browse/P1-100",
                    "project_key": "P1",
                    "planned_start": "2026-02-01",
                    "planned_end": "2026-02-20",
                    "planned_hours": 24.0,
                    "planned_man_days": 3.0,
                    "story_count": 2,
                    "is_unmapped_team": 0,
                    "snapshot_utc": "2026-03-02 00:00:00",
                }
            ],
            "snapshot_meta": {
                "snapshot_utc": "2026-03-02 00:00:00",
                "source_work_items_path": "1_jira_work_items_export.xlsx",
                "total_story_rows": 6,
                "included_story_rows": 3,
                "excluded_missing_epic": 1,
                "excluded_missing_dates": 1,
                "excluded_missing_estimate": 1,
            },
        }
        html = build_team_rmi_gantt_html(payload)
        self.assertIn("Team Owner RMI Gantt", html)
        self.assertIn("Technical Writing", html)
        self.assertIn("Unmapped Team", html)
        self.assertIn("Cards open Jira epic links", html)
        self.assertIn("target=\"_blank\"", html)
        self.assertIn("team_names", html)

    def test_employee_performance_controls_exist(self):
        payload = {
            "worklogs": [],
            "leave_rows": [],
            "projects": [],
            "default_from": "2026-02-01",
            "default_to": "2026-02-28",
            "settings": {
                "base_score": 100,
                "min_score": 0,
                "max_score": 100,
                "points_per_bug_hour": 0.5,
                "points_per_bug_late_hour": 1.5,
                "points_per_unplanned_leave_hour": 0.75,
                "points_per_subtask_late_hour": 1.0,
                "points_per_estimate_overrun_hour": 1.25,
            },
            "generated_at": "2026-02-22 00:00 UTC",
        }
        html = build_employee_perf_html(payload)
        self.assertIn('id="from"', html)
        self.assertIn('id="to"', html)
        self.assertIn('id="projects"', html)
        self.assertIn('id="leaderboard"', html)
        self.assertIn("/settings/performance", html)
        self.assertIn('id="shortcut-current-month"', html)
        self.assertIn('id="shortcut-previous-month"', html)
        self.assertIn('id="shortcut-last-30-days"', html)
        self.assertIn('id="shortcut-quarter-to-date"', html)
        self.assertIn('id="employee-refresh-btn"', html)
        self.assertIn('id="employee-refresh-cancel-btn"', html)
        self.assertIn('id="assignee-extended-actuals-toggle"', html)
        self.assertIn("/api/employee-performance/refresh", html)
        self.assertIn("/api/employee-performance/cancel", html)
        self.assertIn('data-score-drawer-accordion="rules"', html)
        self.assertIn('id="score-drawer-rules-content" class="score-drawer-section-content" hidden', html)
        self.assertIn('aria-expanded="false"', html)
        self.assertIn('id="score-subtask-epic-filter"', html)
        self.assertIn('id="score-subtask-project-filter"', html)
        self.assertIn('id="score-subtask-table-body"', html)
        self.assertIn("Actual Completed Date", html)
        self.assertIn("due-status-pill", html)

    def test_report_entities_formula_editor_controls_exist(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir(parents=True, exist_ok=True)
            (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            wb = Workbook()
            ws = wb.active
            ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
            ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
            wb.save(root / "assignee_hours_report.xlsx")
            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()
            resp = client.get("/settings/report-entities")
            self.assertEqual(resp.status_code, 200)
            html = resp.get_data(as_text=True)
            self.assertIn('id="e-formula-expression"', html)
            self.assertIn('id="formula-suggestions"', html)
            self.assertIn('id="formula-validation"', html)
            self.assertIn('id="formula-quick-insert"', html)

    def test_canonical_refresh_settings_formats_timestamps_for_display(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir(parents=True, exist_ok=True)
            (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            resp = client.get("/settings/canonical-refresh")
            self.assertEqual(resp.status_code, 200)
            html = resp.get_data(as_text=True)

            self.assertIn("formatFriendlyTimestamp", html)
            self.assertIn("formatTimestampHtml", html)
            self.assertIn(".timestamp-display", html)
            self.assertIn('metaStarted.innerHTML = formatTimestampHtml(item.started_at_utc);', html)
            self.assertIn('title="UTC: ${esc(raw)}"', html)

    def test_manage_fields_page_and_settings_links_exist(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir(parents=True, exist_ok=True)
            (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            wb = Workbook()
            ws = wb.active
            ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
            ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
            wb.save(root / "assignee_hours_report.xlsx")
            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            manage = client.get("/settings/manage-fields")
            self.assertEqual(manage.status_code, 200)
            manage_html = manage.get_data(as_text=True)
            self.assertIn("Manage Fields", manage_html)
            self.assertIn('id="mf-formula-expression"', manage_html)
            self.assertIn('id="mf-formula-suggestions"', manage_html)
            self.assertIn('id="mf-formula-quick-insert"', manage_html)
            self.assertIn('id="mf-formula-validation"', manage_html)
            self.assertIn('id="mf-field-key" class="mono" readonly', manage_html)
            self.assertIn("Auto-generated from Label", manage_html)
            self.assertIn("updateFormulaMetaFromReferences", manage_html)
            self.assertIn("meta.references", manage_html)
            self.assertIn("/api/manage-fields", manage_html)

            capacity_html = client.get("/settings/capacity").get_data(as_text=True)
            self.assertNotIn('href="/settings/manage-fields"', capacity_html)
            perf_html = client.get("/settings/performance").get_data(as_text=True)
            self.assertNotIn('href="/settings/manage-fields"', perf_html)
            self.assertIn('id="team-unassigned-list"', perf_html)
            self.assertIn("Assignees Not in Any Team", perf_html)
            entities_html = client.get("/settings/report-entities").get_data(as_text=True)
            self.assertNotIn('href="/settings/manage-fields"', entities_html)
            self.assertNotIn('href="/settings/projects"', entities_html)
            self.assertNotIn('href="/settings/epics-dropdown-options"', entities_html)
            self.assertNotIn('href="/settings/epic-phases"', entities_html)
            self.assertNotIn('href="/settings/epics-management"', entities_html)

            dropdowns = client.get("/settings/epics-dropdown-options")
            self.assertEqual(dropdowns.status_code, 200)
            dropdowns_html = dropdowns.get_data(as_text=True)
            self.assertIn("Epic Dropdown Options", dropdowns_html)
            self.assertIn('id="product-category-options"', dropdowns_html)
            self.assertIn('id="component-options"', dropdowns_html)
            self.assertIn("/api/epics-management/dropdown-options", dropdowns_html)

            projects = client.get("/settings/projects")
            self.assertEqual(projects.status_code, 200)
            projects_html = projects.get_data(as_text=True)
            self.assertIn("Managed Projects", projects_html)
            self.assertIn('id="jira-search"', projects_html)
            self.assertIn('id="jira-search-results"', projects_html)
            self.assertIn('id="project-key"', projects_html)
            self.assertIn('id="project-name"', projects_html)
            self.assertIn('id="display-name"', projects_html)
            self.assertIn('id="color-hex"', projects_html)
            self.assertIn("/api/projects", projects_html)
            self.assertIn("/api/jira/projects/search", projects_html)

            epic_phases = client.get("/settings/epic-phases")
            self.assertEqual(epic_phases.status_code, 200)
            epic_phases_html = epic_phases.get_data(as_text=True)
            self.assertIn("Manage Epic Phases", epic_phases_html)
            self.assertIn("Epic Plan Columns are managed here as Epic Phases", epic_phases_html)
            self.assertIn('id="phase-name"', epic_phases_html)
            self.assertIn('id="phase-position"', epic_phases_html)
            self.assertIn('id="phase-jira-enabled"', epic_phases_html)
            self.assertIn('id="add-phase-btn"', epic_phases_html)
            self.assertIn('id="tab-active"', epic_phases_html)
            self.assertIn('id="tab-deleted"', epic_phases_html)
            self.assertIn('id="phases-tbody"', epic_phases_html)
            self.assertIn("/api/epics-management/plan-columns", epic_phases_html)
            self.assertIn('data-rename-phase', epic_phases_html)
            self.assertIn("/api/epics-management/plan-columns/order", epic_phases_html)
            self.assertIn("/restore", epic_phases_html)

            epics = client.get("/settings/epics-management")
            self.assertEqual(epics.status_code, 200)
            epics_html = epics.get_data(as_text=True)
            self.assertIn("Epics Planner", epics_html)
            self.assertIn("Quick add epic", epics_html)
            self.assertIn("<kbd>Shift</kbd>", epics_html)
            self.assertIn('id="epics-tbody"', epics_html)
            self.assertIn("Project/Product Categorization/Component groups", epics_html)
            self.assertIn("IPP Meeting Planner", epics_html)
            self.assertIn("User Manual Plan", epics_html)
            self.assertIn('id="plan-dialog"', epics_html)
            self.assertIn('id="add-epic-btn"', epics_html)
            self.assertIn('id="add-plan-column-btn"', epics_html)
            self.assertIn('id="manage-plan-columns-btn"', epics_html)
            self.assertIn('id="epic-dialog"', epics_html)
            self.assertIn('id="epic-ipp-meeting-planned"', epics_html)
            self.assertIn("/api/epics-management/dropdown-options", epics_html)
            self.assertIn("/api/epics-management/plan-columns", epics_html)
            self.assertIn("/api/epics-management/plan-columns/order", epics_html)
            self.assertIn("/restore", epics_html)
            self.assertIn('id="manage-plan-columns-btn"', epics_html)
            self.assertIn('id="plan-column-dialog"', epics_html)
            self.assertIn('id="manage-columns-dialog"', epics_html)
            self.assertIn('id="plan-column-name"', epics_html)
            self.assertIn('id="plan-column-position"', epics_html)
            self.assertIn('id="plan-column-jira-enabled"', epics_html)
            self.assertIn('id="plan-column-restore-hint"', epics_html)
            self.assertIn('id="manage-columns-dialog"', epics_html)
            self.assertIn("data-sync-epic-row", epics_html)
            self.assertIn('id="epic-research-urs-plan-jira-url"', epics_html)
            self.assertIn('id="epic-dds-plan-jira-url"', epics_html)
            self.assertIn('id="epic-development-plan-jira-url"', epics_html)
            self.assertIn('id="epic-sqa-plan-jira-url"', epics_html)
            self.assertIn('id="epic-user-manual-plan-jira-url"', epics_html)
            self.assertIn('id="epic-production-plan-jira-url"', epics_html)
            self.assertIn('id="dynamic-plan-fields"', epics_html)

    def test_page_categories_page_contains_report_display_name_controls(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir(parents=True, exist_ok=True)
            (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            wb = Workbook()
            ws = wb.active
            ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
            ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
            wb.save(root / "assignee_hours_report.xlsx")
            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            resp = client.get("/settings/page-categories")
            self.assertEqual(resp.status_code, 200)
            html = resp.get_data(as_text=True)
            self.assertIn("Report display names are editable here; slugs stay fixed.", html)
            self.assertIn("page_overrides", html)
            self.assertIn("data-page-display-name", html)

    def test_epics_management_create_and_update_api(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir(parents=True, exist_ok=True)
            (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            wb = Workbook()
            ws = wb.active
            ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
            ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
            wb.save(root / "assignee_hours_report.xlsx")
            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            columns_resp = client.get("/api/epics-management/plan-columns")
            self.assertEqual(columns_resp.status_code, 200)
            columns_body = columns_resp.get_json()
            seeded = {str(item.get("key")) for item in columns_body.get("columns", [])}
            self.assertIn("epic_plan", seeded)
            self.assertIn("production_plan", seeded)

            add_column_resp = client.post(
                "/api/epics-management/plan-columns",
                json={"label": "Security Plan", "jira_link_enabled": True, "insert_position": 2},
            )
            self.assertEqual(add_column_resp.status_code, 201)
            add_column_body = add_column_resp.get_json()
            security_key = add_column_body["column"]["key"]
            self.assertTrue(security_key.startswith("security_plan"))

            rename_column_resp = client.put(
                f"/api/epics-management/plan-columns/{security_key}",
                json={"label": "Security Review Plan"},
            )
            self.assertEqual(rename_column_resp.status_code, 200)
            renamed_column = (rename_column_resp.get_json() or {}).get("column") or {}
            self.assertEqual(str(renamed_column.get("key")), security_key)
            self.assertEqual(str(renamed_column.get("label")), "Security Review Plan")

            reordered_resp = client.put(
                "/api/epics-management/plan-columns/order",
                json={
                    "ordered_keys": [
                        "research_urs_plan",
                        "epic_plan",
                        security_key,
                        "dds_plan",
                        "development_plan",
                        "sqa_plan",
                        "user_manual_plan",
                        "production_plan",
                    ]
                },
            )
            self.assertEqual(reordered_resp.status_code, 200)
            reordered_body = reordered_resp.get_json()
            reordered_keys = [str(item.get("key")) for item in reordered_body.get("columns", [])]
            self.assertEqual(
                reordered_keys,
                [
                    "research_urs_plan",
                    "epic_plan",
                    security_key,
                    "dds_plan",
                    "development_plan",
                    "sqa_plan",
                    "user_manual_plan",
                    "production_plan",
                ],
            )

            delete_default_resp = client.delete("/api/epics-management/plan-columns/epic_plan")
            self.assertEqual(delete_default_resp.status_code, 400)

            create_resp = client.post(
                "/api/epics-management/rows",
                json={
                    "epic_key": "O2-999",
                    "project_key": "O2",
                    "project_name": "O2 Project",
                    "product_category": "Core",
                    "epic_name": "Payments Revamp",
                    "description": "Initial epic",
                    "originator": "Lead A",
                    "priority": "High",
                    "ipp_meeting_planned": "Yes",
                    "jira_url": "https://jira.example.com/browse/O2-999",
                    "plans": {
                        "epic_plan": {"man_days": 8, "start_date": "2026-02-01", "due_date": "2026-02-10"},
                        "research_urs_plan": {
                            "man_days": 2,
                            "start_date": "2026-02-01",
                            "due_date": "2026-02-03",
                            "jira_url": "https://jira.example.com/browse/O2-1101",
                        },
                        "user_manual_plan": {
                            "jira_url": "https://jira.example.com/browse/O2-1199",
                        },
                        security_key: {
                            "jira_url": "https://jira.example.com/browse/O2-1205",
                            "man_days": 1.5,
                            "start_date": "2026-02-05",
                            "due_date": "2026-02-06",
                        },
                    },
                },
            )
            self.assertEqual(create_resp.status_code, 201)
            create_body = create_resp.get_json()
            self.assertEqual(create_body["row"]["epic_key"], "O2-999")
            self.assertEqual(create_body["row"]["priority"], "High")
            self.assertEqual(create_body["row"]["ipp_meeting_planned"], "Yes")
            self.assertEqual(
                create_body["row"]["plans"]["research_urs_plan"]["jira_url"],
                "https://jira.example.com/browse/O2-1101",
            )
            self.assertEqual(
                create_body["row"]["plans"]["user_manual_plan"]["jira_url"],
                "https://jira.example.com/browse/O2-1199",
            )
            self.assertEqual(
                create_body["row"]["plans"][security_key]["jira_url"],
                "https://jira.example.com/browse/O2-1205",
            )
            self.assertEqual(create_body["row"]["plans"][security_key]["man_days"], 1.5)

            update_resp = client.put(
                "/api/epics-management/rows/O2-999",
                json={
                    "description": "Updated epic",
                    "priority": "Highest",
                    "ipp_meeting_planned": "No",
                    "plans": {
                        "epic_plan": {"man_days": 10, "start_date": "2026-02-01", "due_date": "2026-02-12"},
                        "dds_plan": {"jira_url": "https://jira.example.com/browse/O2-1201"},
                        security_key: {"man_days": 2, "start_date": "2026-02-07", "due_date": "2026-02-09"},
                    },
                },
            )
            self.assertEqual(update_resp.status_code, 200)
            update_body = update_resp.get_json()
            self.assertEqual(update_body["row"]["description"], "Updated epic")
            self.assertEqual(update_body["row"]["priority"], "Highest")
            self.assertEqual(update_body["row"]["ipp_meeting_planned"], "No")
            self.assertEqual(update_body["row"]["plans"]["epic_plan"]["man_days"], 10)
            self.assertEqual(
                update_body["row"]["plans"]["dds_plan"]["jira_url"],
                "https://jira.example.com/browse/O2-1201",
            )
            self.assertEqual(update_body["row"]["plans"][security_key]["man_days"], 2)

            planner_columns_after_rename = client.get("/api/epics-management/plan-columns")
            self.assertEqual(planner_columns_after_rename.status_code, 200)
            planner_columns = planner_columns_after_rename.get_json().get("columns", [])
            security_columns = [item for item in planner_columns if str(item.get("key")) == security_key]
            self.assertEqual(len(security_columns), 1)
            self.assertEqual(str(security_columns[0].get("label")), "Security Review Plan")

            delete_dynamic_resp = client.delete(f"/api/epics-management/plan-columns/{security_key}")
            self.assertEqual(delete_dynamic_resp.status_code, 200)
            delete_dynamic_body = delete_dynamic_resp.get_json()
            keys_after_delete = {str(item.get("key")) for item in delete_dynamic_body.get("columns", [])}
            self.assertNotIn(security_key, keys_after_delete)

            restore_resp = client.post(f"/api/epics-management/plan-columns/{security_key}/restore")
            self.assertEqual(restore_resp.status_code, 200)
            restored_column = (restore_resp.get_json() or {}).get("column") or {}
            self.assertEqual(str(restored_column.get("key")), security_key)

            create_default_resp = client.post(
                "/api/epics-management/rows",
                json={
                    "epic_key": "O2-1000",
                    "project_key": "O2",
                    "project_name": "O2 Project",
                    "product_category": "Core",
                    "epic_name": "Default Planner Flag",
                },
            )
            self.assertEqual(create_default_resp.status_code, 201)
            create_default_body = create_default_resp.get_json()
            self.assertEqual(create_default_body["row"]["ipp_meeting_planned"], "No")

            add_delete_candidate_resp = client.post(
                "/api/epics-management/plan-columns",
                json={"label": "Deprecation Plan", "jira_link_enabled": False},
            )
            self.assertEqual(add_delete_candidate_resp.status_code, 201)
            delete_candidate_key = str(add_delete_candidate_resp.get_json()["column"]["key"])

            delete_column_resp = client.delete(f"/api/epics-management/plan-columns/{delete_candidate_key}")
            self.assertEqual(delete_column_resp.status_code, 200)
            delete_column_body = delete_column_resp.get_json()
            remaining_keys = [str(item.get("key")) for item in delete_column_body.get("columns", [])]
            self.assertNotIn(delete_candidate_key, remaining_keys)

            delete_default_resp = client.delete("/api/epics-management/plan-columns/epic_plan")
            self.assertEqual(delete_default_resp.status_code, 400)

    def test_epics_dropdown_options_api(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir(parents=True, exist_ok=True)
            (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            wb = Workbook()
            ws = wb.active
            ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
            ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
            wb.save(root / "assignee_hours_report.xlsx")
            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            get_initial = client.get("/api/epics-management/dropdown-options")
            self.assertEqual(get_initial.status_code, 200)
            self.assertEqual(get_initial.get_json()["product_category_options"], [])
            self.assertEqual(get_initial.get_json()["component_options"], [])
            self.assertEqual(get_initial.get_json()["plan_status_options"], [])

            save_resp = client.put(
                "/api/epics-management/dropdown-options",
                json={
                    "product_category": ["Core", "Payments", "core"],
                    "components": ["Checkout API", "Portal"],
                    "plan_statuses": ["Planned", "Not Planned Yet", "planned"],
                },
            )
            self.assertEqual(save_resp.status_code, 200)
            body = save_resp.get_json()
            self.assertEqual(body["product_category_options"], ["Core", "Payments"])
            self.assertEqual(body["component_options"], ["Checkout API", "Portal"])
            self.assertEqual(body["plan_status_options"], ["Planned", "Not Planned Yet"])

            get_saved = client.get("/api/epics-management/dropdown-options")
            self.assertEqual(get_saved.status_code, 200)
            saved_body = get_saved.get_json()
            self.assertEqual(saved_body["product_category_options"], ["Core", "Payments"])
            self.assertEqual(saved_body["component_options"], ["Checkout API", "Portal"])
            self.assertEqual(saved_body["plan_status_options"], ["Planned", "Not Planned Yet"])

    def test_epics_management_tmp_key_orphan_create_and_jira_key_promotion(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir(parents=True, exist_ok=True)
            (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            wb = Workbook()
            ws = wb.active
            ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
            ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
            wb.save(root / "assignee_hours_report.xlsx")
            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            create_resp = client.post(
                "/api/epics-management/rows",
                json={"epic_name": "Planning Draft Epic"},
            )
            self.assertEqual(create_resp.status_code, 201)
            body = create_resp.get_json()
            tmp_key = str(body["row"]["epic_key"])
            self.assertRegex(tmp_key, r"^TMP-\d{8}T\d{6}Z-[A-Z0-9]{6}$")
            self.assertEqual(body["row"]["project_key"], "ORPHAN")
            self.assertEqual(body["row"]["project_name"], "Orphan")
            self.assertEqual(body["row"]["product_category"], "Orphan")

            update_resp = client.put(
                f"/api/epics-management/rows/{tmp_key}",
                json={"jira_url": "https://jira.example.com/browse/O2-4242"},
            )
            self.assertEqual(update_resp.status_code, 200)
            update_body = update_resp.get_json()
            self.assertEqual(update_body["row"]["epic_key"], "O2-4242")
            self.assertEqual(update_body["row"]["jira_url"], "https://jira.example.com/browse/O2-4242")

            rows_resp = client.get("/api/epics-management/rows")
            self.assertEqual(rows_resp.status_code, 200)
            keys = {str(item.get("epic_key")) for item in rows_resp.get_json().get("rows", [])}
            self.assertIn("O2-4242", keys)
            self.assertNotIn(tmp_key, keys)

    def test_epics_management_tmp_key_conflict_offers_vacant_key_reuse(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir(parents=True, exist_ok=True)
            (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            wb = Workbook()
            ws = wb.active
            ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
            ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
            wb.save(root / "assignee_hours_report.xlsx")
            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            vacant_tmp_key = "TMP-20260101T000000Z-ABC123"
            seed_resp = client.post(
                "/api/epics-management/rows",
                json={
                    "epic_key": vacant_tmp_key,
                    "epic_name": vacant_tmp_key,
                    "project_key": "ORPHAN",
                    "project_name": "Orphan",
                    "product_category": "Orphan",
                    "component": "",
                    "description": "",
                    "originator": "",
                    "jira_url": "",
                    "plans": {},
                },
            )
            self.assertEqual(seed_resp.status_code, 201)

            conflict_resp = client.post(
                "/api/epics-management/rows",
                json={
                    "epic_key": vacant_tmp_key,
                    "epic_name": "New Planned Epic",
                },
            )
            self.assertEqual(conflict_resp.status_code, 409)
            conflict_body = conflict_resp.get_json() or {}
            self.assertEqual(conflict_body.get("code"), "epic_key_exists")
            self.assertEqual(conflict_body.get("vacant_tmp_key"), vacant_tmp_key)
            self.assertTrue(conflict_body.get("can_reuse_vacant_tmp_key"))
            # Error message must not expose backend epic keys to the user
            self.assertNotIn("TMP-", conflict_body.get("error", ""))

            reuse_resp = client.put(
                f"/api/epics-management/rows/{vacant_tmp_key}",
                json={
                    "epic_name": "New Planned Epic",
                    "description": "Saved by reusing vacant TMP key",
                },
            )
            self.assertEqual(reuse_resp.status_code, 200)
            reuse_body = reuse_resp.get_json() or {}
            self.assertEqual(reuse_body.get("row", {}).get("epic_key"), vacant_tmp_key)
            self.assertEqual(reuse_body.get("row", {}).get("epic_name"), "New Planned Epic")

    @patch("report_server._fetch_jira_issues_for_jql")
    @patch("report_server.resolve_jira_end_date_field_ids")
    @patch("report_server.resolve_jira_start_date_field_id")
    @patch("report_server.get_session")
    @patch("report_server.extract_jira_key_from_url")
    def test_epics_management_sync_persists_epic_and_story_rows(
        self,
        mock_extract_key,
        mock_get_session,
        mock_start_field,
        mock_end_fields,
        mock_fetch_jql,
    ):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root = Path(td)
            (root / "report_html").mkdir(parents=True, exist_ok=True)
            (root / "report_html" / "dashboard.html").write_text("<html><body>ok</body></html>", encoding="utf-8")
            wb = Workbook()
            ws = wb.active
            ws.append(["project_key", "worklog_date", "period_day", "period_week", "period_month", "issue_assignee", "hours_logged"])
            ws.append(["O2", "2026-02-01", "2026-02-01", "2026-W05", "2026-02", "Alice", 1.0])
            wb.save(root / "assignee_hours_report.xlsx")
            app = create_report_server_app(base_dir=root, folder_raw="report_html")
            client = app.test_client()

            create_resp = client.post(
                "/api/epics-management/rows",
                json={
                    "epic_key": "O2-321",
                    "project_key": "O2",
                    "project_name": "O2 Project",
                    "product_category": "Core",
                    "epic_name": "Existing Epic",
                    "jira_url": "https://jira.example.com/browse/O2-321",
                    "plans": {
                        "research_urs_plan": {"jira_url": "https://jira.example.com/browse/O2-101"},
                        "dds_plan": {"jira_url": "https://jira.example.com/browse/O2-102"},
                    },
                },
            )
            self.assertEqual(create_resp.status_code, 201)

            mock_extract_key.side_effect = lambda url: str(url or "").rstrip("/").split("/")[-1]
            mock_get_session.return_value = object()
            mock_start_field.return_value = "customfield_start"
            mock_end_fields.return_value = ["customfield_end"]
            mock_fetch_jql.side_effect = [
                [
                    {
                        "key": "O2-321",
                        "fields": {
                            "issuetype": {"name": "Epic"},
                            "summary": "Jira Epic Summary",
                            "description": {
                                "type": "doc",
                                "content": [
                                    {
                                        "type": "paragraph",
                                        "content": [{"type": "text", "text": "Epic desc from Jira"}],
                                    }
                                ],
                            },
                            "timeoriginalestimate": 28800,
                            "customfield_start": "2026-02-03",
                            "customfield_end": "2026-02-20",
                        },
                    }
                ],
                [
                    {
                        "key": "O2-101",
                        "fields": {
                            "issuetype": {"name": "Story"},
                            "customfield_10014": "O2-321",
                            "summary": "Story 101",
                            "status": {"name": "In Progress"},
                            "timeoriginalestimate": 14400,
                            "customfield_start": "2026-02-01",
                            "customfield_end": "2026-02-12",
                        },
                    },
                    {
                        "key": "O2-102",
                        "fields": {
                            "issuetype": {"name": "Story"},
                            "customfield_10014": "O2-321",
                            "summary": "Story 102",
                            "status": {"name": "To Do"},
                            "timeoriginalestimate": 7200,
                            "customfield_start": "2026-02-05",
                            "customfield_end": "2026-02-25",
                        },
                    },
                    {
                        "key": "O2-103",
                        "fields": {
                            "issuetype": {"name": "Sub-task"},
                            "customfield_10014": "O2-321",
                            "summary": "Subtask 103",
                            "status": {"name": "Done"},
                            "timeoriginalestimate": 3600,
                            "customfield_start": "2026-02-02",
                            "customfield_end": "2026-02-10",
                        },
                    },
                ],
            ]

            sync_resp = client.post(
                "/api/epics-management/rows/O2-321/sync-jira-plan",
                json={"jira_url": "https://jira.example.com/browse/O2-321"},
            )
            self.assertEqual(sync_resp.status_code, 200)
            body = sync_resp.get_json()
            self.assertEqual(body["synced_story_count"], 2)
            self.assertEqual(body["row"]["epic_name"], "Jira Epic Summary")
            self.assertIn("Epic desc from Jira", body["row"]["description"])
            self.assertEqual(body["row"]["plans"]["research_urs_plan"]["man_days"], 0.5)
            self.assertEqual(body["row"]["plans"]["research_urs_plan"]["start_date"], "2026-02-01")
            self.assertEqual(body["row"]["plans"]["research_urs_plan"]["due_date"], "2026-02-12")
            self.assertEqual(body["row"]["plans"]["dds_plan"]["man_days"], 0.25)
            self.assertEqual(body["row"]["plans"]["dds_plan"]["start_date"], "2026-02-05")
            self.assertEqual(body["row"]["plans"]["dds_plan"]["due_date"], "2026-02-25")

            db_path = root / "assignee_hours_capacity.db"
            conn = sqlite3.connect(db_path)
            try:
                row = conn.execute(
                    "SELECT epic_name, description FROM epics_management WHERE epic_key=?",
                    ("O2-321",),
                ).fetchone()
                self.assertIsNotNone(row)
                self.assertEqual(row[0], "Jira Epic Summary")
                self.assertIn("Epic desc from Jira", row[1])

                story_rows = conn.execute(
                    "SELECT story_key, epic_key, story_name, story_status FROM epics_management_story_sync WHERE epic_key=? ORDER BY story_key",
                    ("O2-321",),
                ).fetchall()
                self.assertEqual(len(story_rows), 2)
                self.assertEqual(story_rows[0][0], "O2-101")
                self.assertEqual(story_rows[0][1], "O2-321")
                self.assertEqual(story_rows[0][2], "Story 101")
                self.assertEqual(story_rows[1][0], "O2-102")
            finally:
                conn.close()

    def test_dashboard_template_uses_planner_validation_alerts(self):
        template_path = Path(__file__).resolve().parents[1] / "dashboard_template.html"
        html = template_path.read_text(encoding="utf-8")
        self.assertIn("Planner Validation:", html)
        self.assertIn("Planner Dates:", html)
        self.assertIn("Planner Hours:", html)
        self.assertIn("kind === 'story'", html)
        self.assertIn("storyPlannerStartCell", html)
        self.assertIn("storyPlannerEndCell", html)
        self.assertIn("Jira planned dates/hours differ from Epics Planner epic plan.", html)
        self.assertIn("mismatch-planner-btn", html)
        self.assertIn("/settings/epics-management?epic_key=", html)
        self.assertIn("reason=planner_mismatch", html)
        self.assertNotIn("Alert: Jira planned dates differ from IPP meeting dates.", html)

    def test_planned_vs_dispensed_page_controls_exist(self):
        html_path = Path(__file__).resolve().parents[1] / "report_html" / "planned_vs_dispensed_report.html"
        self.assertTrue(html_path.exists())
        html = html_path.read_text(encoding="utf-8")
        self.assertIn('id="date-filter-from"', html)
        self.assertIn('id="date-filter-to"', html)
        self.assertIn('id="date-filter-apply"', html)
        self.assertIn('id="date-filter-reset"', html)
        self.assertIn('id="adv-filter-menu"', html)
        self.assertIn('id="planned-hours-source"', html)
        self.assertIn('id="plan-source"', html)
        self.assertIn('id="projects-trigger"', html)
        self.assertIn('id="projects-menu"', html)
        self.assertIn('id="projects-select-all"', html)
        self.assertIn('id="projects-clear-all"', html)
        self.assertIn('id="projects-options"', html)
        self.assertIn("By Log Date", html)
        self.assertIn("By Planned Date", html)
        self.assertIn("col-resize-handle", html)
        self.assertIn("/api/approved-vs-planned-hours/ui-settings", html)
        self.assertIn("Approved vs Planned Hours Report", html)
        self.assertIn("Total Approved Hours", html)
        self.assertIn("Total Planned Hours", html)
        self.assertIn("ACTUAL HOURS", html)
        self.assertIn('id="pvd-total-actual-hours"', html)
        self.assertIn("Planned Hours (Subtask Original Estimates)", html)
        self.assertIn("Actual Hours (Subtask and Bug Subtask Worklogs)", html)
        self.assertIn("Epic Drill-down", html)
        self.assertNotIn("details.epic", html)
        self.assertNotIn("details.story", html)
        self.assertNotIn('<details class="epic"', html)
        self.assertNotIn('<details class="story"', html)
        self.assertIn('id="pvd-comparison-chart"', html)
        self.assertIn('id="pvd-detail-root"', html)
        self.assertIn("/api/approved-vs-planned-hours/summary", html)
        self.assertIn("/api/approved-vs-planned-hours/details", html)

    def test_planned_actual_table_view_page_controls_exist(self):
        html_path = Path(__file__).resolve().parents[1] / "report_html" / "planned_actual_table_view.html"
        self.assertTrue(html_path.exists())
        html = html_path.read_text(encoding="utf-8")
        self.assertIn('id="from-date"', html)
        self.assertIn('id="to-date"', html)
        self.assertIn('id="mode"', html)
        self.assertIn('id="projects"', html)
        self.assertIn('id="statuses"', html)
        self.assertIn('id="assignees"', html)
        self.assertIn('id="load-btn"', html)
        self.assertIn('id="fetch-btn"', html)
        self.assertIn("/api/planned-actual-table-view/summary", html)
        self.assertIn("/api/planned-actual-table-view/refresh", html)
        self.assertIn("/api/planned-actual-table-view/filter-options", html)
        self.assertIn("/api/planned-actual-table-view/queue", html)
        self.assertIn("/api/planned-actual-table-view/cancel", html)
        self.assertIn("/api/planned-actual-table-view/history", html)
        self.assertIn("/api/planned-actual-table-view/diff", html)
        self.assertIn("/api/planned-actual-table-view/export", html)
        self.assertIn("Fetch Queue", html)
        self.assertIn("Cancel and Rollback", html)

    def test_original_estimates_hierarchy_page_controls_exist(self):
        html_path = Path(__file__).resolve().parents[1] / "report_html" / "original_estimates_hierarchy_report.html"
        self.assertTrue(html_path.exists())
        html = html_path.read_text(encoding="utf-8")
        self.assertIn('id="from-date"', html)
        self.assertIn('id="to-date"', html)
        self.assertIn('id="projects"', html)
        self.assertIn('id="statuses"', html)
        self.assertIn('id="assignees"', html)
        self.assertIn('id="search-anything"', html)
        self.assertIn('id="apply-btn"', html)
        self.assertIn('id="reset-btn"', html)
        self.assertIn('id="fetch-btn"', html)
        self.assertIn('id="table-body"', html)
        self.assertIn("/api/original-estimates/filter-options", html)
        self.assertIn("/api/original-estimates/summary", html)
        self.assertIn("/api/original-estimates/refresh", html)
        self.assertIn("/api/original-estimates/refresh-epic/", html)


if __name__ == "__main__":
    unittest.main()
