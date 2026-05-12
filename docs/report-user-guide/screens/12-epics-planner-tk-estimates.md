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
| Jira sync API | `/api/epics-management/rows/<epic_key>/sync-jira-plan` |
| Jira publish preview API | `/api/epics-management/populate-jira/preflight` |
| Jira publish execute API | `/api/epics-management/populate-jira/execute` |
| Jira publish report history API | `/api/epics-management/populate-jira/reports` |
| Jira publish report detail API | `/api/epics-management/populate-jira/reports/<report_id>` |
| Jira publish failed-item retry API | `/api/epics-management/populate-jira/reports/<report_id>/retry` |
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
| `qa_handover` | Handover | Formula-managed for man-days; planned start/due dates are user-editable; TK Budgeted is 0.5 days when Dev has input. |
| `bug_fixing` | Bug Fixing | Formula-managed for man-days; planned start/due dates are user-editable; TK Budgeted is 15% of TK Approved. |
| `production_plan` | Release | Formula-managed for man-days; planned start/due dates are user-editable; TK Budgeted is 2 days when TK Approved is greater than zero. |
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

For each phase, TK Budgeted dates copy from the phase planned start and due dates. Formula-managed phases do not require user-entered Most Likely estimates, but Handover, Bug Fixing, and Release do allow direct planned date entry while their man-days remain computed.

## Jira Sync Scope

The **Sync Jira Epic** action opens a modal before any Jira data is applied. The modal has separate checkboxes for man-days and planned dates so the user can choose whether to refresh:

- epic values only
- epic values plus linked phase values

When only epic options are selected, linked phase plans keep their current man-days and dates. When linked phase options are selected, only phases with configured Jira URLs are updated from Jira. If no sync option is selected, the request is rejected. Sealed epics still reject Jira sync until the user clicks **RE-BUDGET**.

## Jira Publish Reports

The **POPULATE JIRA** action publishes selected planner epics into Jira. Before launch, the modal shows whether each epic will create a new Jira structure or update an already linked Jira epic. Duplicate-name warnings must be acknowledged before execution.

Every publish attempt writes a persistent report to `epics_management_jira_publish_reports`. The report records:

| Report area | Contents |
| --- | --- |
| Summary | Status, start/completion timestamps, requested epic count, total work items, succeeded count, failed count, and skipped count. |
| Work item rows | Epic key, issue level (`epic`, `story`, or `subtask`), phase/month context, create/update action, Jira issue link, status, and error text. |
| Retry context | Failed epic and story rows are marked retryable. Failed subtasks are shown with the error but are not individually retried from the report UI. |

After a publish completes or partially fails, the modal replaces the preview with the detailed report. Failed rows show their Jira/API error text, and the **Retry Failed Work Items** button launches a new publish attempt for retryable failed epics and stories. Story retries use update mode for the parent epic and limit the publish request to the failed phase keys so successful stories are not recreated unnecessarily.

The **Publish History** button opens the latest saved reports. Each history row can be opened with **View Details**, which reloads the saved report from the detail API.

## Approved Plan Import

The import page reads `Epic Estimates Approved Plan.xlsx` from `EPICS_PLANNER_IMPORT_XLSX_PATH` or the default OneDrive source path. It processes only worksheet names containing `RMI` and derives the project display name from the sheet name before `RMI`.

The parser uses the row 1 merged `RnD Most likely` group and row 2 headers to map the current workbook layout:

| Workbook field | Planner behavior |
| --- | --- |
| `B` Category | Forward-filled across merged cells into Product Categorization. |
| `C` Components | Forward-filled across merged cells into Component. |
| `E` Jira ID | Provides the epic key and Jira metadata lookup. When the cell contains rendered link text, the import reads the cell hyperlink target instead of the display text. Rows without a valid Jira URL are shown but not imported. |
| `F` Originator | Saved to Originator. |
| `M` Work Status | Displayed on the review page only. |
| `N:V` RnD Most Likely phases | Saved as Most Likely phase estimates. |
| `W` Man Days | Used only to flag total mismatches against `sum(N:V)`. |

The review page fetches Jira epic summary/description and child issues, then suggests phase Jira links and dates. Suggested phase links are editable and optional; rejected or blank links do not block submit. Submit creates a timestamped backup of `assignee_hours_capacity.db`, auto re-budgets sealed existing epics, updates or inserts rows, sets Priority to `High`, and sets Plan Status to `Planned` for successfully written rows.

## Planner Display

The Epics Planner shows user-managed phase labels from the Epic Phases Manager. Most Likely cells use a light orange background, TK Budgeted cells use a light grass green background, and paired phase layers share stronger left/right borders so each phase and its TK instance read as one visual group.

Epic rows stay compact in the main planner table. The **Plan Overview** cell shows a two-part split summary:

- left side in orange for **Most Likely**
- right side in green for **TK Approved**

Hovering each half shows the metric name. Click an epic row to expand an accordion panel that shows:

- a summary strip for Most Likely, Optimistic, Pessimistic, Calculated, and TK Approved totals
- a three-column phase matrix with **Phase**, **Most likely**, and **TK Budgeted**

The matrix keeps phase editing in-context: editable Most Likely cells open the phase planner dialog, formula-managed phases still allow planned-date edits from the TK Budgeted side, and phase Jira open/edit actions stay attached to each phase row. Phase rows are intentionally compact and show the phase name without an extra subtitle line.

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
