# IPP Phase Transform Logic

Last Updated: 2026-02-17

## Purpose

This pipeline transforms IPP Meeting phase range data into a normalized, auditable structure and computes all dashboard geometry server-side.

Primary outputs:
- `ipp_phase_breakdown_YYYY-MM-DD.xlsx` (timestamped archive snapshot)
- `ipp_meeting_dashboard.html` (render-only dashboard generated from transformed data)

## Source Workbook and Sheet Selection

Source workbook path resolution order:
1. CLI `--input-xlsx`
2. `IPP_MEETING_XLSX_PATH`
3. Default path: `C:\Users\hmalik\OneDrive - Octopus Digital\ALL DOCS\IPP Meeting\all ipp meetings.xlsx`

Source sheet selection order:
1. CLI `--sheet` or env `IPP_PHASE_SOURCE_SHEET` (exact match)
2. Latest sheet name matching `Latest - DD Mon YYYY` (highest parsed date)
3. Latest sheet name matching `DD Mon YYYY` (highest parsed date)
4. First sheet in workbook

## Required Input Columns

Case-insensitive header matching is used.

Required base columns:
- `Product`
- `Epic/RMI`
- `Jira Task ID`
- `Planned Start Date`
- `Planned End Date`
- `Actual Date (Production Date)`

Optional base column:
- `Remarks` (blank if missing)

Required phase columns:
- `Research/URS - Planned Range`
- `DDS - Planned Range`
- `Development - Planned Range`
- `SQA - Planned Range`
- `User Manual - Planned Range`
- `Production - Planned Range`

## Parsing Grammar and Rules

Expected phase text shape:
- `[STARTDATE] <> [ENDDATE] | [N] mandays`

Accepted date formats:
- `DD-Mon-YYYY`
- `D-Mon-YYYY`
- `DD/Mon/YYYY`
- `YYYY-MM-DD`
- permissive 2-digit year variants (`DD-Mon-YY`, `DD/Mon/YY`)

Man-days parsing:
- Numeric value before `manday` or `mandays`
- Supports integer and decimal values

Status semantics:
- `Skipped` -> `skipped`
- `Not planned` (case-insensitive variants) -> `not_planned`
- Empty/null -> `no_entry`
- Parsed values with no validation issues -> `planned`
- Parsed values with validation issues -> `invalid`

Special case for `not_planned` rows:
- Parsed dates are kept if present.
- Man-days can remain blank.

Invalid range policy:
- If parsed `start > end`, values are kept as parsed.
- Parse warning includes `start_after_end`.
- Planning state becomes `invalid` for planned-style records with warnings.

## Parse Warning Catalog

Possible warnings per phase:
- `missing_or_invalid_date_range`
- `missing_or_non_numeric_mandays`
- `start_after_end`

Warnings are comma-separated in `<Phase> Parse Warning`.

## Output Workbook Schema

### Sheet: `IPP Phase Breakdown`

Base columns:
1. `Source Sheet`
2. `Row Number`
3. `Product`
4. `Epic/RMI`
5. `Epic/RMI Jira Link`
6. `Epic Planned Start Date`
7. `Epic Planned End Date`
8. `Epic Actual Date (Production Date)`
9. `Remarks`

Per phase columns (for each of 6 phases):
1. `<Phase> Planned Start Date`
2. `<Phase> Planned End Date`
3. `<Phase> Planned Man-days`
4. `<Phase> Raw Planned Range`
5. `<Phase> Planning State`
6. `<Phase> Parse Warning`

Row-level validation flags:
1. `Any Phase Parse Warning`
2. `Any No Entry`
3. `Any Not Planned`
4. `Any Skipped`

Computed base extensions:
1. `Computed Total Phase Man-days`
2. `Computed Has Valid Epic Plan`
3. `Computed Epic Planned Start ISO`
4. `Computed Epic Planned End ISO`
5. `Computed Epic Actual ISO`

### Sheet: `IPP Dashboard Computed`

One row per RMI. This is the dashboard rendering contract.

Columns:
1. `Source Sheet`
2. `Row Number`
3. `Product`
4. `Epic/RMI`
5. `Epic/RMI Jira Link`
6. `Epic Planned Start Date`
7. `Epic Planned End Date`
8. `Epic Actual Date (Production Date)`
9. `Remarks`
10. `Computed Total Phase Man-days`
11. `Computed Roadmap Valid`
12. `Computed Roadmap Axis Start ISO`
13. `Computed Roadmap Axis End ISO`
14. `Computed Roadmap Axis Span Days`
15. `Computed Roadmap Today In Range`
16. `Computed Roadmap Today Left Pct`
17. `Computed Roadmap Bar Left Pct`
18. `Computed Roadmap Bar Width Pct`
19. `Computed Roadmap Actual Left Pct`
20. `Computed Roadmap Week Ticks JSON`
21. `Computed MiniGantt Has Dated Phases`
22. `Computed MiniGantt Axis Start ISO`
23. `Computed MiniGantt Axis End ISO`
24. `Computed MiniGantt Axis Span Days`
25. `Computed MiniGantt Timeline Width Px`
26. `Computed MiniGantt Scroll Enabled`
27. `Computed MiniGantt Week Ticks JSON`
28. `Computed MiniGantt Today In Range`
29. `Computed MiniGantt Today Left Pct`
30. `Computed Phase Geometry JSON`

## Computation Rules (Server-Side)

All dashboard geometry is computed in Python, not in HTML/JS.

Roadmap rules:
- Axis window: min valid epic start to max valid epic end with `+/-10` day padding.
- Tick bucket: bi-weekly (`14` days).
- Today marker: computed at export time (UTC date).
- Bar positions: left/width percentages clamped to `[0, 100]`, minimum width `0.8%`.

Mini-gantt rules:
- Axis window per RMI: min valid phase start to max valid phase end with `+/-10` day padding.
- Tick bucket: weekly (`7` days).
- Horizontal scroll: enabled when raw phase span (`max_end - min_start + 1`) exceeds `7` days.
- Timeline width: `max(520, axis_span_days * 24)` when scroll enabled; else `520`.
- Thickness benchmark: one global max phase man-days across dataset.
- Thickness mapping: `8px` to `22px`, with `15px` fallback midpoint when benchmark is non-positive.
- Track top offset: `(30 - thickness) / 2` where track baseline is `30px`.

`Computed Phase Geometry JSON` fields per phase:
- `state`, `state_label`, `warning`
- `raw`
- `start_iso`, `end_iso`
- `mandays_text`, `mandays_num`
- `valid`
- `bar_left_pct`, `bar_width_pct`
- `bar_thickness_px`, `bar_top_offset_px`
- `start_label`, `end_label`, `bar_label`
- `show_no_bar`

## Metadata Sheet

Sheet: `Metadata`

Includes:
- `Generated At`
- `Source Workbook`
- `Source Sheet`
- `Output Mode`
- `Data Row Count`
- `Phase Count`
- `Computed Sheet Name`
- `Global Max Man-days Benchmark`
- `Computation Timestamp UTC`
- `Computation Rules Version`

## Jira Link Normalization

`Epic/RMI Jira Link` behavior:
- Keep URL values unchanged.
- If value contains an issue key like `ABC-123`, normalize to `https://<JIRA_SITE>.atlassian.net/browse/ABC-123`.
- Otherwise store original text.

## Dashboard Contract and Behavior

`generate_ipp_meeting_dashboard.py` reads:
- Base sheet `IPP Phase Breakdown` for compatibility checks.
- Computed sheet `IPP Dashboard Computed` for rendering payload.

Payload includes:
- Base display fields
- `roadmap` precomputed structure
- `mini_gantt` precomputed structure
- `phase_data` mapped from computed phase JSON

HTML contract:
- `ipp_meeting_dashboard_template.html` is render-only.
- No client-side date parsing, tick generation, bar geometry, or man-day aggregation.

## CLI and Environment Interface

### `export_ipp_phase_breakdown.py`
CLI:
- `--input-xlsx`
- `--sheet`
- `--output-dir`
- `--output-mode` (`timestamped`, `fixed`, `both`)
- `--fixed-output-name`

Env:
- `IPP_MEETING_XLSX_PATH`
- `IPP_PHASE_SOURCE_SHEET`
- `IPP_PHASE_OUTPUT_DIR`
- `IPP_PHASE_OUTPUT_MODE`
- `IPP_PHASE_OUTPUT_FIXED_NAME`

### `generate_ipp_meeting_dashboard.py`
CLI:
- `--input-xlsx`
- `--output-html`
- `--output-dir`

Env:
- `IPP_PHASE_OUTPUT_DIR`
- `IPP_PHASE_DASHBOARD_HTML_PATH`

## Run Commands

Generate transformed workbook:
- `python export_ipp_phase_breakdown.py`

Generate dashboard from latest transformed workbook:
- `python generate_ipp_meeting_dashboard.py`

Generate dashboard from explicit workbook:
- `python generate_ipp_meeting_dashboard.py --input-xlsx ipp_phase_breakdown_2026-02-17.xlsx`
