# RMI Jira Gantt Report — Technical & Business Documentation

> **Report file:** `IPP Meeting Reports/rmi_jira_gantt.html`
> **Generator script:** `generate_rmi_gantt_html.py`
> **Data extractor:** `extract_rmi_jira_to_sqlite.py`
> **Pipeline runner:** `run_rmi_pipeline.py`
> **Test suite:** `tests/test_generate_rmi_gantt_html.py`

---

## 1. Purpose

The RMI Jira Gantt Report is a self-contained, single-file HTML dashboard that visualises Roadmap Management Initiative (RMI) data from Jira and an Excel workbook. It provides:

- The report in the live app (`/rmi_jira_gantt_report.html`) now scopes to **TK Epics only** (`is_tk_epic = 1` in Epics Planner).
- TK Epic tagging is driven by Epics Planner import (`/settings/epics-management/import`) where users now **upload a workbook file** instead of relying on a fixed local default path.

- A **Gantt chart** of epic timelines grouped by product.
- A **hierarchical table** of epics → stories → subtasks/bugs with estimates, worklogs, and statuses.
- **Metric cards** summarising estimates, TK Approved targets, and idle capacity.
- **Month analysis toggles** to scope data by delivery, start, or inclusive "Through" month.
- A **capacity calculator** to compare team availability against approved work.
- **Product filtering** and **full-text search** across all epics.
- **Interactive drawers** showing detailed breakdowns of scoped epics, stories, and subtasks.

---

## 2. Data Sources

### 2.1 Excel Workbook (Source of Truth for Estimates)

The pipeline reads an Excel file where each sheet represents a product (e.g., "OmniConnect RMI", "Fintech Fuel RMI"). Relevant columns:

| Column | Field | Description |
|--------|-------|-------------|
| D | `roadmap_item` | User-defined epic title |
| E | `jira_id` | Jira epic key (e.g., `O2-793`) |
| V | `man_days` / `man_days_value` | Most Likely estimate (man-days) |
| W | `optimistic_50` / `optimistic_50_value` | Optimistic 50th percentile estimate |
| X | `pessimistic_10` / `pessimistic_10_value` | Pessimistic 10th percentile estimate |
| Y | `est_formula` / `est_formula_value` | Weighted formula estimate |
| Z | `tk_target` / `tk_target_value` | TK Approved estimate (man-days) |

### 2.2 Jira API

For each epic key found in the workbook, the extractor fetches from Jira:

- **Epic metadata** — summary, status, priority, start date, due date, original estimate, aggregate estimate.
- **Stories** — child issues under the epic with their own dates and estimates.
- **Story descendants** — subtasks and bugs under each story with dates, estimates, and aggregated worklogs.
- **Worklogs** — individual time entries (author, date, seconds) on descendant issues.

### 2.3 SQLite Database

All extracted data is stored in `IPP Meeting Reports/rmi_jira_extract.db` with six tables:

| Table | Purpose | Key Columns |
|-------|---------|-------------|
| `source_rmi_rows` | Excel workbook rows mapped to Jira IDs | `sheet_name`, `jira_id`, `roadmap_item`, `man_days_value`, `tk_target_value` |
| `epics` | Jira epic-level metadata | `epic_key` (PK), `summary`, `status`, `priority`, `jira_start_date`, `jira_due_date`, `jira_original_estimate_seconds` |
| `stories` | Jira stories under each epic | `story_key` (PK), `epic_key` (FK), `jira_original_estimate_seconds`, dates |
| `story_descendants` | Subtasks and bugs under stories | `issue_key` (PK), `parent_story_key` (FK), `is_subtask`, `jira_original_estimate_seconds`, dates |
| `worklogs` | Time entries on descendants | `worklog_id` (PK), `issue_key` (FK), `time_spent_seconds`, `author_display_name`, `started` |
| `run_errors` | Extraction/processing errors | `error_scope`, `issue_key`, `message` |

**Joins used by the generator:**

```
source_rmi_rows LEFT JOIN epics ON jira_id = epic_key
stories → story_descendants ON story_key = parent_story_key
story_descendants LEFT JOIN worklogs ON issue_key = issue_key
```

---

## 3. Related Files

| File | Role |
|------|------|
| `extract_rmi_jira_to_sqlite.py` | Reads Excel workbook + calls Jira REST API → writes SQLite DB |
| `generate_rmi_gantt_html.py` | Reads SQLite DB → generates self-contained HTML report |
| `run_rmi_pipeline.py` | Orchestrates extract → generate in a single command |
| `tests/test_generate_rmi_gantt_html.py` | Unit tests for generator functions (unittest) |
| `IPP Meeting Reports/rmi_jira_extract.db` | SQLite database (generated artifact) |
| `IPP Meeting Reports/rmi_jira_gantt.html` | HTML report (generated artifact) |
| `.env` / `env.example.txt` | Jira credentials (`JIRA_SITE`, `JIRA_EMAIL`, `JIRA_API_TOKEN`) |

### Running the Pipeline

```bash
# Full pipeline (extract + generate)
python run_rmi_pipeline.py

# Generate report only (from existing DB)
python generate_rmi_gantt_html.py

# With custom paths
python run_rmi_pipeline.py --workbook path/to/file.xlsx --db path/to/db.sqlite --html path/to/output.html

# Dry-run with row limit
python run_rmi_pipeline.py --limit 5
```

---

## 4. Data Flow

```
Excel Workbook + Jira API
        │
        ▼
 extract_rmi_jira_to_sqlite.py
        │
        ▼
   SQLite Database (6 tables)
        │
        ▼
 generate_rmi_gantt_html.py
   ├── load_report_data()          → nested Python dicts
   ├── build_summary()             → aggregate counts
   ├── build_epic_metric_summary() → per-product metric totals
   ├── build_epic_detail_records() → lightweight JSON for client
   ├── render_html()               → HTML + embedded CSS + JS
   │     ├── Embeds epicMetrics as JSON
   │     ├── Embeds epicDetails as JSON (with nested stories/subtasks)
   │     ├── Embeds storyDetails, subtaskDetails as JSON
   │     └── Embeds capacityMonths as JSON
   └── generate_html_report()      → writes .html file
        │
        ▼
   rmi_jira_gantt.html (self-contained)
        │
        ▼
   Browser: JS reads embedded JSON → renders UI → responds to user interactions
```

---

## 5. UI Structure

### 5.1 Report Sections (top to bottom)

| Section | Default State | Description |
|---------|---------------|-------------|
| **Header** | Visible | Title, database path, generation note |
| **Capacity Calculator** | Visible | Inputs for employees, month, leaves → capacity/availability outputs |
| **Metric Cards Grid** | Visible | 9 metric cards (epic count, estimate range, TK Approved, idle capacity, estimates) |
| **Product Summary Cards** | Visible | Clickable cards for All Products + each product, showing TK Approved totals |
| **TK Month Toolbar** | Visible | 3 toggle checkboxes + month dropdown + status text |
| **Month Story Analysis Panel** | Hidden | Appears when any toggle is active (or all off to show all-months) — cards, bar chart, exclusion table |
| **Search Toolbar** | Visible | Full-text search input with clear button |
| **Unit Toolbar** | Visible | Hours / Days toggle buttons |
| **View Toolbar** | Visible | Gantt View / Table View toggle buttons |
| **Gantt View** | Hidden | SVG timeline chart grouped by product |
| **Table View** | Visible | Hierarchical expandable table (epics → stories → subtasks) |
| **Run Errors** | Visible | Table of extraction/processing errors |
| **Drawer Modal** | Hidden | Slide-in panel for detailed epic/story/subtask breakdowns |

### 5.2 Metric Cards

| Card | Color | Type | Source | Clickable |
|------|-------|------|--------|-----------|
| **Epic Count** | Teal | Count | `epic_count` from metric summary | No |
| **Optimistic (50%)** | Light Blue | Duration | `optimistic_seconds` (50th percentile × 28800) | No |
| **Most Likely** | Blue | Duration | `most_likely_seconds` (man_days × 28800) | No |
| **Pessimistic (10%)** | Dark Blue | Duration | `pessimistic_seconds` (10th percentile × 28800) | No |
| **Calculated** | Darkest Blue | Duration | `calculated_seconds` (formula × 28800) | No |
| **TK Approved** | Emerald (hero, 2× width) | Duration | `tk_approved_seconds` (tk_target × 28800) | Yes → opens in-scope epics drawer |
| **Idle Hours/Days** | Cyan | Duration | `availability − tk_approved` (computed in JS) | No |
| **Epic Estimates** | Indigo | Duration | `jira_original_estimate_seconds` | Yes → opens epic estimate drawer |
| **Story Estimates** | Slate | Duration | `story_estimate_seconds` (sum of story originals) | Yes → opens story estimate drawer |
| **Subtask Estimates** | Violet | Duration | `subtask_estimate_seconds` (sum of subtask originals) | Yes → opens subtask estimate drawer |

### 5.3 Product Summary Cards

One card for "All Products" plus one per product found in the data. Each shows the TK Approved total for that product. Clicking a card toggles product filtering across the entire page (metrics, table, gantt, analysis).

- **Multi-select:** Clicking individual products toggles them; clicking "All Products" resets to all.
- **Active state:** Selected cards have highlighted borders and backgrounds.

### 5.4 RMI Estimation & Scheduling Table

The RMI Estimation & Scheduling panel is written into the HTML with an initial populated table body and footer during report generation, then enhanced by client-side JavaScript for year switching, Jira-only filtering, product cards, and unit changes.

- **Initial year selection:** The generator prefers the current year only when the schedule dataset actually contains dates in that year.
- **Fallback behavior:** If the current calendar year has no schedule dates, the table starts on the latest year found in epic/story/subtask schedule data instead of opening as an empty grid.
- **First paint:** Because the first table body is rendered server-side, rerunning `python generate_rmi_gantt_html.py` produces a populated table even before browser-side interactions run.

---

## 6. Business Logic & Formulas

### 6.1 Estimate Conversion

All workbook values are in **man-days**. The report converts to seconds for internal calculations:

```
seconds = man_days × 28800   (1 man-day = 8 hours = 28800 seconds)
```

Display toggles between hours and days:

```
hours = seconds / 3600
days  = seconds / 28800
```

### 6.2 Capacity Calculator

| Field | Input Type | Description |
|-------|-----------|-------------|
| **No. of Employees** | Number (min 0, step 1) | Headcount |
| **Month of 2026** | Select | Pre-populated with Jan–Dec 2026 and working day counts |
| **Total Leaves** | Number (hours or days, follows unit toggle) | Leave deductions |

**Formulas:**

```
Total Capacity   = Employees × Working Days × 28800
Total Availability = Total Capacity − Leaves (in seconds)
Idle Capacity    = Total Availability − TK Approved (current scope)
```

**Working days** are Mon–Fri counts per calendar month. Public holidays are **not** subtracted — the leaves field is intended for that adjustment.

### 6.3 TK Approved Scoping

The TK Approved card value changes based on active month filters:

| Filter State | TK Approved Shows |
|-------------|-------------------|
| All toggles OFF | Sum of `tk_target_value × 28800` across all product-scoped epics |
| Started toggle ON | Sum for epics where `start_date` month = selected month |
| Delivered toggle ON | Sum for epics where `due_date` month = selected month |
| Started + Delivered ON | Sum for epics where both start and due = selected month |
| Through toggle ON | Sum for epics where `start_date_month ≤ selected_month ≤ due_date_month` |

### 6.4 Idle Capacity

```
Idle Capacity = Total Availability (from calculator) − TK Approved (current scope)
```

- Positive value → "Remaining availability after TK Approved"
- Negative value → "TK Approved exceeds total availability"

### 6.5 Story Estimate Rollups

| Metric | Formula |
|--------|---------|
| **Epic Original Estimate** | `epic.jira_original_estimate_seconds` (from Jira) |
| **Story Estimate Total** | `SUM(story.jira_original_estimate_seconds)` for all stories under epic |
| **Subtask Estimate Total** | `SUM(subtask.jira_original_estimate_seconds)` for all subtasks under all stories |
| **Total Logged** | `SUM(worklog.time_spent_seconds)` across all descendants |

### 6.6 TK Target Comparison Notes

Each epic row in the table shows a comparison badge (hover tooltip):

```
delta_hours = tk_target_value − (epic_original_estimate_seconds / 3600)
```

| State | Condition | Badge |
|-------|-----------|-------|
| **match** | `|delta| < 0.01` | Green ℹ — "Matches Jira original estimate" |
| **over** | `delta > 0` | Red ℹ — "+X.XX h above Jira original estimate" |
| **under** | `delta < 0` | Amber ℹ — "X.XX h below Jira original estimate" |
| **neutral** | Missing TK target or estimate | Grey ℹ — "No numeric TK target available" |

### 6.7 Story Rollup Comparison Notes

Similar comparison between the story-level rollup total and the epic's aggregate original estimate:

```
delta_seconds = story_rollup − epic_original_estimate_seconds
```

Same match/over/under/neutral states as TK Target comparison.

---

## 7. Month Toggle System

### 7.1 Toggle Definitions

| Toggle | Label | Filter Logic |
|--------|-------|-------------|
| **Started** | "For epics started in" | `parseMonthKey(epic.start_date) === selectedMonth` |
| **Delivered** | "For epics delivered in" | `parseMonthKey(epic.due_date) === selectedMonth` |
| **Through** | "Any Work Done Through" | `startMonth ≤ selectedMonth ≤ dueMonth` (inclusive overlap) |

### 7.2 Toggle Combination Matrix

| Started | Delivered | Through | Behaviour |
|---------|-----------|---------|-----------|
| OFF | OFF | OFF | All epics shown; analysis panel shows **all available months** chart |
| ON | OFF | OFF | Epics whose start_date falls in selected month |
| OFF | ON | OFF | Epics whose due_date falls in selected month |
| ON | ON | OFF | Epics whose start_date **AND** due_date both fall in selected month |
| OFF | OFF | ON | Epics whose date range **includes** the selected month (inclusive overlap) |
| ON | — | ON | ❌ Not possible — Through disables Started |
| — | ON | ON | ❌ Not possible — Through disables Delivered |

### 7.3 Through Mode Specifics

When the "Any Work Done Through" toggle is enabled:

1. **Started and Delivered toggles are disabled** (unchecked and greyed out).
2. **Month dropdown** is populated with every month from every epic's `start_date` through `due_date` (inclusive range), not just the start/due months.
3. **Epic matching** uses inclusive overlap: an epic is included if `start_month ≤ selected_month ≤ due_month`.
4. **Examples** (selected month = Feb 2026):
   - Epic start=Jan, due=Mar → **Included** (Jan ≤ Feb ≤ Mar)
   - Epic start=Feb, due=Feb → **Included** (Feb ≤ Feb ≤ Feb)
   - Epic start=Jan, due=Feb → **Included** (Jan ≤ Feb ≤ Feb)
   - Epic start=Mar, due=Apr → **Excluded** (Mar > Feb)

### 7.4 Month Dropdown Population

| Mode | Months Shown |
|------|-------------|
| **Through ON** | Every month from `min(start_date)` through `max(due_date)` across all epics (inclusive ranges expanded) |
| **Started/Delivered** | Only months that appear as a start_date or due_date month in any epic |
| **All toggles OFF** | Same as Started/Delivered (union of start + due months) |

The dropdown defaults to the current calendar month if available, otherwise the first available month.

---

## 8. Month Story Estimate Analysis

### 8.1 Analysis Modes

| Mode | Trigger | Entries | Scope |
|------|---------|---------|-------|
| **Scoped** | Any toggle ON | 3 entries: Previous Month, Selected Month, Next Month | Epics matching the active month filter |
| **All** | All toggles OFF | One entry per available month | All product-scoped epics |

### 8.2 Allocation Algorithm

For each in-scope epic, the analysis allocates story and subtask estimates into month buckets:

```
FOR each epic in scope:
  FOR each story under the epic:
    IF story fits in a single month:
      → Add story.estimate_seconds to that month's bucket (if month is in allowed set)
    ELSE IF story spans multiple months:
      → Try to allocate individual subtask estimates by their dates
      → Each subtask is bucketed by: same-month → that month; else due_date month; else start_date month
      IF no subtasks have usable estimates:
        → EXCLUDE the epic with reason "Story {KEY} spans {START} to {DUE}
           but has no usable subtask estimates."
  IF no story/subtask contributed any estimate:
    → EXCLUDE the epic with reason "No story or subtask estimates fall
       within the month window."
```

### 8.3 Analysis Panel UI

| Element | Description |
|---------|-------------|
| **Header** | Title + footnote explaining methodology |
| **Status text** | Product scope, epic count, included/excluded counts |
| **"See Epics" button** | Opens drawer with included and excluded epic lists |
| **Analysis Cards** | Dynamic cards — in "scoped" mode: Previous/Selected/Next; in "all" mode: All Available Months total, Months In Chart count, Peak Month value |
| **Bar Chart** | Vertical bars per month; height scaled to max value; highlighted bar = selected month (scoped) or peak month (all) |
| **Summary Pills** | "X included" / "Y excluded" counts |
| **Exclusion Table** | Columns: Epic Key (clickable), Product, Reason. Clicking an epic key opens its exclusion details in a drawer. |

### 8.4 Exclusion Reasons

| Reason | When |
|--------|------|
| "No stories available for month estimate analysis." | Epic has zero stories |
| "Story {KEY} spans {START} to {DUE} but has no usable subtask estimates." | Cross-month story with no subtask estimates > 0 in allowed months |
| "No story or subtask estimates fall within the previous, selected, or next month window." | No contributions from any story/subtask (catch-all) |

---

## 9. Interactive Drawers

The report includes a slide-in drawer modal triggered by various UI elements:

| Trigger | Drawer Content |
|---------|---------------|
| **Click TK Approved card** | Lists all epics in the current product scope with TK Approved values |
| **Click Epic Estimates card** | Lists all epics sorted by Jira original estimate (descending) |
| **Click Story Estimates card** | Lists all stories sorted by original estimate (descending) |
| **Click Subtask Estimates card** | Lists all subtasks sorted by original estimate (descending) |
| **Click "See Epics" button** | Shows included and excluded epics from the current month analysis |
| **Click excluded epic key** | Shows the specific epic's details, exclusion reasons, stories, and subtasks |

Each drawer card displays: Jira key (linked), product, title, start/due dates, status, priority, story count, and the relevant estimate value.

---

## 10. Gantt Chart View

- **Structure:** One SVG section per product, containing horizontal bars for each epic.
- **Bar position:** Calculated from `start_date` to `due_date` relative to the timeline's overall min/max range.
- **Grid lines:** Monthly and weekly divisions.
- **Bar metadata:** Hover/label shows status, TK Approved hours, and date range.
- **Product colours:**
  - OmniConnect: `#0f766e` (teal)
  - Fintech Fuel: `#b45309` (amber)
  - OmniChat: `#2563eb` (blue)
  - Digital Log: `#7c3aed` (violet)
- **Filtering:** Respects both product filter and search query.
- **Requirement:** Epics must have both `start_date` and `due_date` to appear in the Gantt chart.

---

## 11. Table View

### 11.1 Hierarchical Structure

```
Epic Row (purple background)
  ├── Story Row (blue background)     [expand/collapse]
  │     ├── Subtask Row (green)
  │     ├── Bug Subtask Row (red)
  │     └── Worklog Table (nested)
  └── Story Row ...
```

### 11.2 Table Columns

| # | Column | Description |
|---|--------|-------------|
| 1 | Toggle (+/−) | Expand/collapse child rows |
| 2 | Jira Link | "Go" button linking to Jira issue |
| 3 | Title Block | Epic summary, roadmap item, status badge, priority, logged hours |
| 4 | Start Date | `jira_start_date` formatted as DD-Mon-YYYY |
| 5 | Due Date | `jira_due_date` formatted as DD-Mon-YYYY |
| 6 | TK Approved | `tk_target_value` in hours/days with comparison note |
| 7 | Jira Estimate | `jira_original_estimate_seconds` formatted |
| 8 | Aggregate Estimate | `jira_aggregate_original_estimate_seconds` |
| 9 | Story Estimate | Sum of story-level original estimates with comparison note |
| 10 | Subtask Estimate | Sum of subtask-level original estimates |
| 11 | Source Values | Workbook formula values (man_days, optimistic, pessimistic, est_formula, tk_target) |

### 11.3 Row Colours

| Row Type | Background | Indicator |
|----------|-----------|-----------|
| Epic | Light purple | Left border purple |
| Story | Light blue | Left border blue |
| Subtask | Light green | Left border green |
| Bug Subtask | Light red | Left border red |

---

## 12. Search & Filtering

### 12.1 Full-Text Search

The search input filters epics by matching against a concatenated text of:
- Jira key
- Epic summary
- Roadmap item
- Product name
- Epic status
- Epic priority

Matching is case-insensitive. The table and gantt views are filtered simultaneously. A status indicator shows the count of visible results.

### 12.2 Product Filter

- **All Products** — shows everything (default).
- **Individual products** — click one or more product cards/buttons to filter.
- **Multi-select** — clicking additional products adds them; clicking again deselects.
- **Reset** — clicking "All Products" returns to showing everything.
- **Scope** — affects metrics, table, gantt, month analysis, and drawer contents.

---

## 13. Unit Toggle (Hours / Days)

All duration values in the report can switch between hours and days display:

| Unit | Conversion | Format (standard) | Format (compact) |
|------|-----------|-------------------|------------------|
| Hours | `seconds / 3600` | `1,234.56 h` | `1,235 h` |
| Days | `seconds / 28800` | `154.32 d` | `154 d` |

The toggle updates all `duration-value` spans, metric cards, analysis panels, drawers, and the capacity calculator leaves field simultaneously.

---

## 14. Validations & Warnings

### 14.1 Data Extraction Warnings

Errors encountered during Jira API extraction are stored in the `run_errors` table and displayed in the **Run Errors** section at the bottom of the report:

| Column | Description |
|--------|-------------|
| Scope | Error category (e.g., `worklog_lookup`, `story_fetch`) |
| Issue Key | Affected Jira issue (if applicable) |
| Sheet | Source workbook sheet |
| Row | Workbook row number |
| Message | Human-readable error description |

### 14.2 Report-Level Validations

| Validation | Handling |
|-----------|----------|
| Epic missing start_date or due_date | Excluded from Gantt chart; Through mode may not match correctly |
| Story/subtask with estimate ≤ 0 | Skipped during month analysis allocation |
| Cross-month story with no usable subtasks | Epic excluded from analysis; reason shown in exclusion table |
| TK target or Jira estimate missing | Comparison note shows "neutral" state |
| No epics match current month filter | TK Approved card shows base total; status shows "no months available" |
| Negative idle capacity | Meta text changes to "TK Approved exceeds total availability" |

### 14.3 Capacity Calculator Constraints

| Input | Min | Max | Step |
|-------|-----|-----|------|
| Employees | 0 | — | 1 |
| Leaves (hours mode) | 0 | — | 1 |
| Leaves (days mode) | 0 | — | 0.5 |

---

## 15. CSS Styling Architecture

The report embeds approximately 2,500 lines of CSS in a `<style>` block. Key design patterns:

| Pattern | Description |
|---------|-------------|
| **CSS Variables** | Product colours defined as constants (`#0f766e`, `#b45309`, etc.) |
| **Sticky Columns** | First 3 table columns (toggle, link, title) are position:sticky for horizontal scroll |
| **Responsive Grid** | Metric cards use CSS grid with `auto-fill` and minmax for responsive layout |
| **Colour-coded Rows** | `.row-epic`, `.row-story`, `.row-subtask`, `.row-bug-subtask` for hierarchy |
| **Toggle Track/Thumb** | Custom checkbox styling mimicking iOS-style toggle switches |
| **Drawer Overlay** | Fixed positioning with backdrop blur and slide-in animation |
| **Bar Chart** | CSS flex + percentage heights for dynamic bar scaling |
| **Print Styles** | Hidden interactive elements in `@media print` |

---

## 16. Embedded JSON Data

The HTML report embeds four JSON datasets in `<script>` tags:

| Variable | Source | Content |
|----------|--------|---------|
| `epicMetrics` | `build_epic_metric_summary()` | Per-product + "all" totals for all metric keys |
| `epicDetails` | `build_epic_detail_records()` | Array of epics with nested `stories[].subtasks[]` — used for month analysis and drawers |
| `storyDetails` | `build_story_detail_records()` | Flat array of all stories with estimates — used for story estimate drawer |
| `subtaskDetails` | `build_subtask_detail_records()` | Flat array of all subtasks with estimates — used for subtask estimate drawer |
| `capacityMonths` | `working_days_2026()` | Array of 12 month objects with `key`, `label`, `working_days` |

---

## 17. Core Logic Summary

### 17.1 Python Generation Pipeline

```
load_report_data(db_path)
  └─ SQL queries → nested dicts: epic → stories → descendants → worklogs

build_summary(source_rows, errors)
  └─ Counts: epics, stories, descendants, worklogs, errors

build_epic_metric_summary(source_rows)
  └─ Per-product sums: estimates × 28800, TK approved, story/subtask rollups

build_epic_detail_records(source_rows)
  └─ Lightweight JSON: epic + nested stories + nested subtasks (no worklogs)

render_html(report_data, db_path)
  └─ Assembles complete HTML with:
       ├─ render_capacity_calculator()    → capacity section
       ├─ render_metric_cards()           → metric card grid
       ├─ render_product_tk_cards()       → product summary cards
       ├─ render_tk_month_toolbar()       → 3 toggles + month select
       ├─ render_tk_month_story_analysis_panel() → analysis section
       ├─ render_search_toolbar()         → search input
       ├─ render_gantt()                  → SVG gantt chart
       ├─ render_epic_table_view()        → hierarchical table
       ├─ render_error_table()            → run errors
       ├─ render_drawer_modal()           → drawer overlay
       └─ Embedded <script> with all JS logic
```

### 17.2 JavaScript Client-Side Flow

```
Page Load
  ├─ Parse embedded JSON (epicMetrics, epicDetails, etc.)
  ├─ populateMonthOptions()       → fill month dropdown
  ├─ populateCapacityMonths()     → fill capacity month dropdown
  ├─ updateCapacityResult()       → compute initial capacity
  ├─ setUnits('hours')            → default unit display
  ├─ setProduct('all')            → default product scope
  ├─ setView('table-view')        → default view
  └─ syncTkMonthUi()              → sync toggle states
       ├─ Read toggle checkboxes
       ├─ If Through ON → disable Started/Delivered
       ├─ populateMonthOptions()
       ├─ Enable/disable month select
       ├─ updateMetricCards()     → refresh all metric values
       └─ updateTkMonthCards()    → refresh TK Approved + analysis
            ├─ Compute scoped TK total
            ├─ renderTkMonthEstimateAnalysis()
            │    ├─ buildTkMonthEstimateAnalysis()  → allocate estimates to months
            │    ├─ renderMonthAnalysisCards()       → generate card HTML
            │    └─ renderMonthAnalysisBars()        → generate bar chart HTML
            └─ updateIdleCapacityCard()              → recompute idle capacity
```

---

## 18. Test Coverage

The test suite (`tests/test_generate_rmi_gantt_html.py`) uses `unittest` with an in-memory SQLite database:

| Test | Validates |
|------|----------|
| `test_build_summary_counts_nested_records` | Summary counts for epics, stories, descendants, worklogs |
| `test_build_epic_metric_summary_rolls_up_by_product` | Per-product metric aggregation including story/subtask rollups |
| `test_epic_search_text_includes_key_title_and_status_fields` | Full-text search string composition |
| `test_generate_html_report_writes_gantt_and_drilldown` | End-to-end HTML generation, Gantt and table presence, Month Analysis panel |
| `test_build_epic_detail_records_include_nested_story_and_subtask_details` | Nested JSON payload structure with stories and subtasks |
