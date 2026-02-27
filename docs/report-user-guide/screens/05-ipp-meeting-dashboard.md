# IPP Meeting Dashboard

Report ID: `ipp_meeting`

INFO_IDS: `ipp.roadmap_geometry`

## Key Fields

| Field | Definition | Formula / Logic | Ingredients | Business Validations | Cross-Report Linkage |
| --- | --- | --- | --- | --- | --- |
| Roadmap Geometry | Visual schedule geometry for roadmap and mini-gantt. | Uses precomputed transformed workbook geometry fields. | roadmap axis span, bar offsets, phase geometry JSON | Invalid rows remain visible with warnings and fallback handling. | Phase RMI gantt and dashboard delivery context. |

## Drawer Notes

- Drawer explains that rendering geometry is computed upstream in transformed IPP workbook columns.
