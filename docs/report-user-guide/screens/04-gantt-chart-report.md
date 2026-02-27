# Gantt Chart Report

Report ID: `gantt_chart`

INFO_IDS: `gantt.timeline_window`

## Key Fields

| Field | Definition | Formula / Logic | Ingredients | Business Validations | Cross-Report Linkage |
| --- | --- | --- | --- | --- | --- |
| Timeline Window | Visible range and gantt positioning behavior. | Min/max scoped planned dates with fit/reset/zoom controls. | planned start/end, zoom state | Missing date rows flagged as no-range states. | Phase RMI weekly load and Nested planned totals. |

## Drawer Notes

- Drawer explains timeline fit behavior and how planning dates drive visible schedule bars.
