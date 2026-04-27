# Epics Planner TK Estimate Upgrade

## Summary
Update `E:\JIRA SCRIPT` so Epics Planner has separate Most Likely and TK Budgeted phase columns. Users enter Most Likely estimates/dates; TK Budgeted estimates are computed from the reference workbook logic and shown read-only. Existing Epic Phases are reused wherever possible.

## Key Changes
- Copy all previously identified reference files into:
  - `E:\JIRA SCRIPT\Reference TK Estimates Folder`
- Extend Epic Phase metadata so each phase can define:
  - base phase name/key
  - whether it has a Most Likely input column
  - whether it has a TK Budgeted computed column
  - formula role/percentage/fixed days
- Use this mapping:
  - `research_urs_plan` -> `R/URS`
  - `dds_plan` -> `R/DDS`
  - `development_plan` -> `Dev`
  - `sqa_plan` -> `SQA`
  - `user_manual_plan` -> `Doc / User Manual`
  - `qa_handover` -> formula-managed `Handover`
  - `bug_fixing` -> formula-managed `Bug Fixing`
  - `production_plan` -> formula-managed `Release`
  - add missing phases if absent: `Process Design`, `Process QA Testing`, `Regression SQA Testing`
- Store each phase plan with both layers:
  - `most_likely_man_days`, `start_date`, `due_date`, `jira_url`
  - `tk_budgeted_man_days`, copied `tk_budgeted_start_date`, copied `tk_budgeted_due_date`
  - keep `man_days = tk_budgeted_man_days` for compatibility with existing reports.
- Compute RMI totals:
  - `most_likely_total = sum(Most Likely input phase estimates)`
  - `optimistic = most_likely_total - 50%`
  - `pessimistic = most_likely_total + 10%`
  - `calculated = (optimistic + 4 * most_likely_total + pessimistic) / 6`
  - `tk_approved = calculated / 2`
- Compute TK Budgeted phases:
  - `R/URS = tk_approved * 5%` if Most Likely R/URS has input
  - `R/DDS = tk_approved * 10%` if Most Likely R/DDS has input
  - `Bug Fixing = tk_approved * 15%`
  - `Doc / User Manual = tk_approved * 5%` if Most Likely Doc/User Manual has input
  - `Regression SQA Testing = tk_approved * 10%` if Most Likely Regression SQA has input
  - `Handover = 0.5` if Dev has input, else `0`
  - `Release = 2` if `tk_approved > 0`, else `0`
  - Dev/SQA split remaining TK Approved budget using the reference `40:15` weighting when both exist.
- Update Epics Planner UI:
  - render separate Most Likely and TK Budgeted columns
  - make Most Likely estimates/dates editable
  - make TK Budgeted estimates read-only
  - copy manual dates into TK Budgeted date display
  - show Optimistic, Pessimistic, Calculated, and TK Approved on the Epic Plan summary.
- Update seal behavior:
  - recompute all TK values before sealing
  - seal locks every editable column
  - backend rejects update/delete/sync for sealed epics until RE-BUDGET.

## Test Plan
- Add/update tests for:
  - phase metadata migration and mapping
  - missing phase creation
  - Most Likely to TK Budgeted calculations
  - copied TK Budgeted dates
  - sealed row backend protection
  - RE-BUDGET unlock behavior
  - reference folder file copy presence
- Update UI smoke tests for:
  - separate Most Likely/TK Budgeted phase columns
  - formula-managed phases displayed as read-only
  - Epic Phases manager showing phase roles.

## Assumptions
- `JIRA PROJECT` means `E:\JIRA SCRIPT`.
- Existing phase data is preserved and migrated into the Most Likely layer.
- TK Budgeted dates copy from the matching Most Likely phase dates.
- Formula-managed phases do not require user-entered Most Likely estimates.

## How to test locally
After implementation:
```powershell
cd "E:\JIRA SCRIPT"
python init_epics_management_db.py
python -m pytest tests/test_report_ui_smoke.py tests/test_ipp_meeting_dashboard_merge.py tests/test_fetch_jira_dashboard_planner_validation.py
python run_server.py --no-sync
```

Manual URLs:
- `http://127.0.0.1:3000/settings/epic-phases`
- `http://127.0.0.1:3000/settings/epics-management`


---
# Reference TK Estimates Inspect Report

## Purpose

This report records the reference files found for the TK estimate upgrade work. Future agents can use it to understand which upstream files informed the Epics Planner phase mapping, RMI estimate formulas, and Jira/Gantt reference behavior.

## Source Locations

- Reference parent folder: `C:\Users\hmalik\OneDrive - Octopus Digital\Documents\00 - Documentation via Code`
- Direct reference project folder: `C:\Users\hmalik\OneDrive - Octopus Digital\Documents\00 - Documentation via Code\IPP Meeting Reports`
- Target copy folder in this repo: `E:\JIRA SCRIPT\Reference TK Estimates Folder`

## Reference Parent Project Files

| File | Purpose |
| --- | --- |
| `extract_rmi_jira_to_sqlite.py` | Extracts RMI/Jira data into SQLite for downstream reporting and Gantt generation. |
| `generate_rmi_gantt_html.py` | Generates the RMI Jira Gantt HTML report from extracted Jira data. |
| `run_rmi_pipeline.py` | Orchestrates the reference RMI extraction and report generation pipeline. |
| `populate_ipp_from_jira.py` | Populates IPP-related data from Jira references. |
| `tk_approved_for_ipp.py` | Contains TK approved estimate logic used as the main reference for budgeted man-day calculations. |
| `tests\test_extract_rmi_jira_to_sqlite.py` | Regression coverage for Jira-to-SQLite extraction behavior. |
| `tests\test_generate_rmi_gantt_html.py` | Regression coverage for RMI Gantt HTML generation behavior. |
| `tests\test_run_rmi_pipeline.py` | Regression coverage for the reference pipeline orchestration. |

## Direct Reference Project Files

| File | Purpose |
| --- | --- |
| `Epic Estimates Approved Plan.xlsx` | Primary approved estimate workbook used as the reference source for TK budget logic. |
| `Epic Estimates Approved Plan - backup before bug fixing update.xlsx` | Workbook backup before bug-fixing estimate changes. Useful for comparing formula changes. |
| `Epic Estimates Approved Plan - backup before handover update.xlsx` | Workbook backup before handover estimate changes. Useful for comparing formula changes. |
| `Epic Estimates Approved Plan.codex-copy.xlsx` | Working copy of the approved estimate workbook for code-assisted inspection. |
| `rmi_jira_extract.db` | SQLite extract produced by the reference RMI/Jira pipeline. |
| `rmi_jira_gantt.html` | Generated RMI Jira Gantt report from the reference project. |
| `rmi_jira_gantt.backup.html` | Backup copy of the generated RMI Jira Gantt report. |
| `RMI Jira Gantt Report Documentation.md` | Documentation for the reference RMI Jira Gantt report and related process. |

## Notes For Future Agents

- The workbook and `tk_approved_for_ipp.py` are the most important references for the TK Budgeted calculation behavior.
- The Gantt generator and SQLite extract are useful for understanding how Jira phases and RMI data are represented outside the Epics Planner.
- The files should be copied into `Reference TK Estimates Folder` before considering the TK estimate upgrade complete.
