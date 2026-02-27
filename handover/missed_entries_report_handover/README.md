# Missed Entries Report Handover

## Package Contents
- `generate_missed_entries_html.py`: Source generator script.
- `missed_entries.html`: Latest generated report output.
- `shared-nav.css`: Shared navigation stylesheet referenced by HTML.
- `shared-nav.js`: Shared navigation script referenced by HTML.
- `DATA_CONTRACT.md`: Required input workbook schema.
- `requirements_missed_entries.txt`: Minimal Python dependency list.
- `RUNBOOK.md`: Regeneration and deployment steps.

## What This Module Does
- Reads Jira work-item export Excel (`1_jira_work_items_export.xlsx` by default).
- Detects missing values for:
  - Jira Planned Start Date
  - Jira Planned Due Date
  - Original Estimates
- Builds interactive HTML report with:
  - Month range filters
  - Missing-field toggles
  - Assignee summary and drilldown
  - `Export Excel` button for filtered data

## Runtime Prerequisites
- Python `3.10+`
- Package: `openpyxl`
- Browser access to CDN for Excel export from UI:
  - `https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js`

## Key Environment Variables
- `JIRA_EXPORT_XLSX_PATH`:
  - Input Excel path.
  - Default: `1_jira_work_items_export.xlsx`
- `JIRA_MISSED_ENTRIES_HTML_PATH`:
  - Output HTML path.
  - Default: `missed_entries.html`
- `JIRA_SITE`:
  - Jira tenant subdomain used to build fallback issue links.
  - Default: `octopusdtlsupport`

## Hand-off Notes for .NET Integration
- Fastest path:
  - Keep this Python generator as a build step and serve the generated HTML from .NET static files.
- Alternative path:
  - Port `_load_rows` and filter logic from `generate_missed_entries_html.py` into C# backend API and keep frontend as static page.
- If navigation shell is not needed:
  - Remove `shared-nav.css` and `shared-nav.js` references from HTML.
- If internet is blocked:
  - Vendor the `xlsx` browser library locally and update script tag in HTML.

## Verification Targets
- After generation, confirm these rows have due date values from export source:
  - `ODL-55`
  - `ODL-65`
- Confirm UI counts are unique-work-item based:
  - `Total Missed` is not a sum of missing-column counts.

## Ownership Scope
This package is limited to Missed Entries report only.
Other reports and pipeline scripts are intentionally excluded.
