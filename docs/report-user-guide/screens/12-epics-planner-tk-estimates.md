# Epics Planner TK Estimates

## Purpose

Epics Planner manages RMI phase budgets with two layers:

| Layer | User behavior | System behavior |
| --- | --- | --- |
| Most Likely | Editable man-days, start date, due date, and Jira URL where enabled. | Used as the input for RMI totals and TK Budgeted calculations. |
| TK Budgeted | Read-only computed man-days and copied dates. | Stored as `tk_budgeted_man_days`; also copied to `man_days` for report compatibility. |

## Entry Points

| Area | Route |
| --- | --- |
| Epics Planner | `/settings/epics-management` |
| Epic Phases Manager | `/settings/epic-phases` |
| Planner rows API | `/api/epics-management/rows` |
| Phase metadata API | `/api/epics-management/plan-columns` |
| Approved plan import page | `/settings/epics-management/import` |
| Approved plan import preview API | `/api/epics-management/import/preview` |
| Approved plan import submit API | `/api/epics-management/import/submit` |
| Seal API | `/api/epics-management/seal` |
| Re-budget API | `/api/epics-management/rows/<epic_key>/re-budget` |

## Phase Metadata

Epic phases are managed in the Epic Phases Manager. Default phases are seeded or migrated with metadata for role, Most Likely availability, TK Budgeted availability, formula behavior, and lock state. The default/dynamic type is informational; delete protection is controlled by the editable **Locked** checkbox. A user can unlock and save a phase before deleting it, or lock a dynamic phase to protect it from deletion.

Editable cells save atomically. The phase name commits when the user presses Enter or leaves the field; Jira URL support and Locked toggles commit immediately when changed. These cell-level saves update only the changed field instead of requiring a row-level Save action.

| Phase key | Label | Role |
| --- | --- | --- |
| `research_urs_plan` | R/URS | Most Likely input; TK Budgeted is 5% of TK Approved when input exists. |
| `dds_plan` | R/DDS | Most Likely input; TK Budgeted is 10% of TK Approved when input exists. |
| `development_plan` | Dev | Most Likely input; shares remaining TK Approved budget with SQA by 40:15 weighting. |
| `sqa_plan` | SQA | Most Likely input; shares remaining TK Approved budget with Dev by 40:15 weighting. |
| `user_manual_plan` | Doc / User Manual | Most Likely input; TK Budgeted is 5% of TK Approved when input exists. |
| `qa_handover` | Handover | Formula-managed; TK Budgeted is 0.5 days when Dev has input. |
| `bug_fixing` | Bug Fixing | Formula-managed; TK Budgeted is 15% of TK Approved. |
| `production_plan` | Release | Formula-managed; TK Budgeted is 2 days when TK Approved is greater than zero. |
| `process_design` | Process Design | Most Likely input; direct TK Budgeted pass-through. |
| `process_qa_testing` | Process QA Testing | Most Likely input; direct TK Budgeted pass-through. |
| `regression_sqa_testing` | Regression SQA Testing | Most Likely input; TK Budgeted is 10% of TK Approved when input exists. |

## Calculation Rules

The system computes the Epic Plan summary from Most Likely phase inputs:

| Field | Formula |
| --- | --- |
| Most Likely total | Sum of Most Likely input phase estimates. |
| Optimistic | Most Likely total * 50%. |
| Pessimistic | Most Likely total * 110%. |
| Calculated | `(optimistic + 4 * most_likely_total + pessimistic) / 6`. |
| TK Approved | `calculated / 2`. |

For each phase, TK Budgeted dates copy from the corresponding Most Likely start and due dates. Formula-managed phases do not require user-entered Most Likely estimates.

## Approved Plan Import

The import page reads `Epic Estimates Approved Plan.xlsx` from `EPICS_PLANNER_IMPORT_XLSX_PATH` or the default OneDrive source path. It processes only worksheet names containing `RMI` and derives the project display name from the sheet name before `RMI`.

The parser uses the row 1 merged `RnD Most likely` group and row 2 headers to map the current workbook layout:

| Workbook field | Planner behavior |
| --- | --- |
| `B` Category | Forward-filled across merged cells into Product Categorization. |
| `C` Components | Forward-filled across merged cells into Component. |
| `E` Jira ID | Provides the epic key and Jira metadata lookup. Rows without a valid Jira URL are shown but not imported. |
| `F` Originator | Saved to Originator. |
| `M` Work Status | Displayed on the review page only. |
| `N:V` RnD Most Likely phases | Saved as Most Likely phase estimates. |
| `W` Man Days | Used only to flag total mismatches against `sum(N:V)`. |

The review page fetches Jira epic summary/description and child issues, then suggests phase Jira links and dates. Suggested phase links are editable and optional; rejected or blank links do not block submit. Submit creates a timestamped backup of `assignee_hours_capacity.db`, auto re-budgets sealed existing epics, updates or inserts rows, sets Priority to `High`, and sets Plan Status to `Planned` for successfully written rows.

## Planner Display

The Epics Planner shows user-managed phase labels from the Epic Phases Manager. Most Likely cells use a light orange background, TK Budgeted cells use a light grass green background, and paired phase layers share stronger left/right borders so each phase and its TK instance read as one visual group.

The page also shows an executive summary report above the planner grid. It groups epics by product/project and displays counts for **Total Planned**, **Unplanned**, and **Onhold**. The default product order is OmniConnect, Fintech Fuel, Digital Log, Subscription, and OmniChat; any additional products are appended alphabetically. Planned rows are counted from Plan Status `Planned`, unplanned rows are all other plan statuses, and Onhold rows are counted when a plan or delivery status contains a hold/onhold value.

## Seal Behavior

Before sealing, the system recomputes TK values and stores the computed planner snapshot. A sealed epic locks all editable columns in the UI, and backend update, delete, and Jira sync requests return a lock error until the user clicks **RE-BUDGET**.

## Related Code

| File | Responsibility |
| --- | --- |
| `report_server.py` | Routes, database schema/migration, phase metadata, calculations, seal/re-budget enforcement, Epics Planner UI, and approved-plan import workflow. |
| `tests/test_report_ui_smoke.py` | API and UI smoke coverage for phase metadata, TK calculations, seal protection, approved-plan import parsing, preview, and submit behavior. |
| `EPICS_PLANNER_SEAL_GUIDE.md` | User-facing guide for sealing and re-budgeting. |
| `REFERENCE_TK_ESTIMATES_INSPECT_REPORT.md` | Reference file inventory for future agents. |
| `Reference TK Estimates Folder/` | Local copy of the reference scripts, tests, workbook, database, and generated reports. |
