# Phase RMI Gantt Report

Report ID: `phase_rmi_gantt`

INFO_IDS: `phase_rmi.weekly_load`

## Key Fields

| Field | Definition | Formula / Logic | Ingredients | Business Validations | Cross-Report Linkage |
| --- | --- | --- | --- | --- | --- |
| Weekly Load Chips | Weekly phase-lane intensity summary. | Aggregate phase man-days over overlapping week buckets. | phase start/end, phase man-days | Only phases with valid date ranges participate in weekly load. | Gantt timeline spread and IPP roadmap phase pressure. |

## Drawer Notes

- Drawer explains overload chips and how they should be interpreted for weekly planning risk.
