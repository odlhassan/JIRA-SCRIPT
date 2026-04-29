# Report Intelligence Overview

## Purpose

This module documents report logic and the unified click-to-open `i` drawer content used across all leadership reports.

## Source Of Truth

- Code logic is authoritative where older markdown differs.
- `window.reportInfoCatalog` is the normalized frontend contract for drawer content.
- All `i` interactions use right-side drawer behavior (click based).

## Drawer Contract

Each info item contains:

- `id`
- `label`
- `report`
- `ui_targets`
- `definition`
- `formula`
- `ingredients`
- `business_validations`
- `field_linkages`
- `cross_report_linkages`
- `data_sources`
- `leadership_interpretation`

## Screen Docs

- `docs/report-user-guide/screens/00-introduction-epr-tool.md`
- `docs/report-user-guide/screens/01-assignee-hours-report.md`
- `docs/report-user-guide/screens/02-dashboard-report.md`
- `docs/report-user-guide/screens/03-employee-performance-report.md`
- `docs/report-user-guide/screens/04-gantt-chart-report.md`
- `docs/report-user-guide/screens/05-ipp-meeting-dashboard.md`
- `docs/report-user-guide/screens/06-leaves-planned-calendar.md`
- `docs/report-user-guide/screens/07-missed-entries-report.md`
- `docs/report-user-guide/screens/08-nested-view-report.md`
- `docs/report-user-guide/screens/09-phase-rmi-gantt-report.md`
- `docs/report-user-guide/screens/10-rlt-leave-report.md`
- `docs/report-user-guide/screens/11-rnd-data-story-report.md`
- `docs/report-user-guide/screens/12-epics-planner-tk-estimates.md`
- `docs/report-user-guide/screens/13-rmi-jira-gantt-report.md`
