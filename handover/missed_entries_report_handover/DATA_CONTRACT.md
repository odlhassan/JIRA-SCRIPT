# Input Data Contract (Excel)

## Expected Workbook
- Single workbook input.
- Active sheet is read.
- Row 1 must be header row.

## Required Columns
- `issue_key`

## Optional Columns Used (first match wins)
- `jira_issue_type` or `issue_type`
- `assignee`
- `summary`
- Start date candidates:
  - `start_date`
  - `planned start date`
  - `planned_start_date`
- Due date candidates:
  - `end_date`
  - `planned end date`
  - `planned_end_date`
  - `duedate`
- `original_estimate`
- Logged hours candidates:
  - `total_hours_logged`
  - `total hours_logged`
  - `hours_logged`
- `jira_url` (optional; fallback URL is built if missing)

## Date Parsing Rules
Accepted as planned start/due dates:
- `YYYY-MM-DD`
- `DD-Mon-YYYY`
- `DD-Month-YYYY`
- `MM/DD/YYYY`
- `DD/MM/YYYY`

Dates are normalized to `YYYY-MM-DD`.

## Output Row Shape Embedded in HTML
- `issue_key`
- `issue_type`
- `assignee`
- `summary`
- `jira_start_date`
- `jira_due_date`
- `original_estimate`
- `resource_logged_hours` (`Yes`/`No`)
- `jira_url`
