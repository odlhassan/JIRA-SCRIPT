# Missed Entries Report

Report ID: `missed_entries`

INFO_IDS: `missed_entries.total_missed`

## Key Fields

| Field | Definition | Formula / Logic | Ingredients | Business Validations | Cross-Report Linkage |
| --- | --- | --- | --- | --- | --- |
| Total Missed Entries | Total records failing selected required planning fields. | Count rows where any selected missing condition is true. | start date, due date, original estimate, selected missing filters | Date and field filter settings both affect result count. | Employee penalty risk and dashboard planning confidence. |

## Drawer Notes

- Drawer details missing-field semantics and why data quality affects leadership planning decisions.
