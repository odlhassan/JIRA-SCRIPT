# UI Upgrade Plan: Monthly Epic Plan vs Actual Report

**Source**: Design critique based on `monthly_epic_plan_progress_report.html`
**Target files**: `monthly_epic_plan_progress_report.html`, `monthly_epic_plan_progress_service.py`
**Date**: 2026-05-19

## Changelog

- **2026-06-09**: Added **Logged hours scope** toggle (TK dates / Subtask dates) in the header filters. TK dates scopes by epic TK approved dates (default for TK EPICS mode), Subtask dates scopes by subtask's own planned dates (default for ALL JIRA EPICS mode). Both options available when epic scope = ALL EPICS. Executive card and drawer now use the same scoping logic, fixing the mismatch between the executive summary's "Logged this month" and the drawer's "Logged in range".
- **2026-06-09**: Replaced the Month selector (`<input type="month">`) in the header controls with a **Date Range** selector (From/To `<input type="date">` pair). Removed the separate "Or use a custom date range" section below the SHOW EPICS toggles. The date range now initializes to the current month boundaries. API backward-compatible via a `_monthShim` property that derives `YYYY-MM` from the From date.

---

## Summary

The report is visually polished with a consistent blue-slate palette and a well-used semantic color system. The primary opportunity for improvement is the **controls panel**, which is too heavy and overwhelms users before any data is visible. Secondary issues span accessibility, consistency, and visual hierarchy.

---

## Priority 1 — Critical (Do First)

### 1.1 Collapse Secondary Filters into "Advanced" Row

**Problem**: The controls area has 3 separate rows (month/project/unit/overdue/toggle/scope/apply → employee → date-filter), presenting 7+ choices to configure before any data is visible.

**Action**:
- Keep primary row: **Date Range (From/To)**, **Project**, **Apply**.
- Move to a collapsible "Advanced Filters" section (hidden by default, toggleable via a `▾ Advanced` button):
  - Overdue lookback input (`overdue-threshold-input`)
  - On Hold toggle
  - Date filter range toggles
- Employee filter row stays visible (frequently used), but its hint text is shortened (see 2.2).

**Files to change**:
- `monthly_epic_plan_progress_report.html` — restructure `.controls-panel` HTML and add collapse JS.
- `monthly_epic_plan_progress_service.py` — if controls HTML is generated here, update the template section.

**Expected impact**: ~40% reduction in perceived page complexity on first load.

---

## Priority 2 — Moderate (Do Second)

### 2.1 Fix Minimum Font Size Floor at 0.72rem

**Problem**: Multiple text classes fall below readable threshold:
| Selector | Current size |
|---|---|
| `.emp-dd-res` | `0.62rem` |
| `.epic-kpi-label` | `0.67rem` |
| `.card .label` | `0.72rem` |
| `th` | `0.78rem` |

**Action**:
- Set hard floor: **no visible text below `0.72rem`**.
- `.emp-dd-res` → raise to `0.72rem` or replace with icon.
- `.epic-kpi-label` → raise to `0.72rem`.
- Sub-`0.70rem` usage allowed **only** for purely decorative badges with no information content.

**Files to change**: CSS block in `monthly_epic_plan_progress_report.html`.

---

### 2.2 Shorten Employee Filter Hint Text

**Problem**: `<p class="field-hint">` contains multi-sentence developer-style instructions ("Filter capacity and leave by team… Resigned labels use Performance resource records."). Reads as internal docs, not UX copy.

**Action**:
- Replace with a single ≤12-word hint line, e.g., *"Filter capacity and leave data by team member."*
- Move the longer explanation to an `ⓘ` tooltip (hover/click shows full detail).

**Files to change**: `monthly_epic_plan_progress_report.html` (or service template).

---

### 2.3 Fix `overdue-threshold-input` Visual Inconsistency

**Problem**: The overdue threshold input uses inline styles (`border-radius: 4px`) inconsistent with `.control input` styles (`border-radius: 10px`, `min-height: 40px`).

**Action**:
- Remove inline `style` attribute from `overdue-threshold-input`.
- Ensure it inherits `.control input` CSS rules.

**Files to change**: `monthly_epic_plan_progress_report.html`.

---

### 2.4 Controls Grid Responsive Wrapping

**Problem**: The controls grid is fixed at 5 columns with 7+ items — items wrap and misalign at narrower viewports.

**Action**:
- Change controls grid CSS to: `grid-template-columns: repeat(auto-fit, minmax(160px, 1fr))`.
- Test at 1280px and 1920px widths.

**Files to change**: CSS in `monthly_epic_plan_progress_report.html`.

---

### 2.5 Reduce Button Variants from 5+ to 3

**Problem**: Currently 5+ button styles:
- `.btn` (primary)
- `.btn.alt` (secondary)
- `.excl-btn` (orange-bordered)
- `.epic-detail-close` (square icon)
- `.row-toggle-btn` (bordered square)

**Action**:
- Standardize to **3 types only**:
  1. **Primary action** — `.btn`
  2. **Secondary / ghost** — `.btn.ghost`
  3. **Icon-only** — `.btn.icon`
- Map existing variants: `.excl-btn` → `.btn.ghost` with orange accent; `.epic-detail-close` and `.row-toggle-btn` → `.btn.icon`.

**Files to change**: CSS + HTML in `monthly_epic_plan_progress_report.html`.

---

### 2.6 Unify Binary Toggle Pattern

**Problem**: Effort unit uses `<input type="radio">` with plain labels; On Hold and date filter toggles use custom `.df-toggle` components. Same function, two different visual patterns.

**Action**:
- Choose **one** pattern for all binary on/off controls — recommended: use the existing `.df-toggle` component consistently.
- Replace the radio-based effort unit toggle with a `.df-toggle` equivalent.

**Files to change**: HTML + CSS in `monthly_epic_plan_progress_report.html`.

---

## Priority 3 — Accessibility (Do in Parallel with Priority 2)

### 3.1 Fix Focus Ring Opacity

**Problem**: `control input:focus` uses `box-shadow: 0 0 0 3px rgba(29,95,143,0.14)` — 14% opacity is invisible for low-vision and keyboard users.

**Action**:
- Change all `:focus` `box-shadow` alpha values from `0.14` → `0.40` minimum.
- Alternatively add a solid `outline: 2px solid #1d5f8f` as fallback.
- Global find in CSS: `rgba(29,95,143,0.14)` → `rgba(29,95,143,0.40)`.

**Files to change**: CSS in `monthly_epic_plan_progress_report.html`.

---

### 3.2 Add `<fieldset>/<legend>` to Effort Unit Radio Group

**Problem**: `<input type="radio">` buttons for Hours/Days have no `<fieldset>/<legend>` wrapper — screen readers cannot identify the group purpose.

**Action**:
```html
<fieldset class="toggle-wrap">
  <legend class="sr-only">Effort unit</legend>
  <!-- existing radio inputs -->
</fieldset>
```
Add `.sr-only` utility class if not present: `position:absolute; width:1px; height:1px; overflow:hidden; clip:rect(0,0,0,0)`.

**Files to change**: HTML in `monthly_epic_plan_progress_report.html`.

---

### 3.3 Increase Touch Target Sizes

**Problem**:
- `.row-toggle-btn` is `22×22px` — below 44×44px recommended minimum.
- `.emp-dd-row` padding `3px 6px` — too tight for touch.

**Action**:
- `.row-toggle-btn`: set `min-width: 44px; min-height: 44px` (use padding to expand hit area without changing visual size, e.g., `padding: 11px`).
- `.emp-dd-row`: increase padding to `8px 12px`.

**Files to change**: CSS in `monthly_epic_plan_progress_report.html`.

---

### 3.4 Improve Color Contrast for Small Text

**Problem**: `--muted: #5c6b7a` on `--bg: #eef4fb` ≈ 4.2:1 — barely passes WCAG AA for normal text but **fails** for text at `0.62rem` (requires ≥4.5:1 at small sizes).

**Action**:
- After raising font floors (2.1), re-check contrast.
- If any muted text remains below `0.75rem`, darken `--muted` to `#4a5a6a` (~5.1:1).

**Files to change**: CSS variables in `monthly_epic_plan_progress_report.html`.

---

## Priority 4 — Minor / Polish

### 4.1 Narrow Epic Detail Drawer Width

**Problem**: Epic detail drawer opens at 72% viewport width — feels like a full-page takeover, loses table context.

**Action**:
- Set drawer width to fixed `480px` (or `min(560px, 90vw)` for responsiveness).
- Transition stays as-is.

**Files to change**: CSS in `monthly_epic_plan_progress_report.html`.

---

### 4.2 Strengthen Section Separators

**Problem**: All `<section class="panel">` elements use identical border and shadow — high-priority sections (KPI summary) are visually indistinct from lower-priority sections (workforce capacity).

**Action**:
- Add a stronger visual weight to the KPI/Executive Summary panel: slightly larger border-top accent (`4px solid var(--accent)`).
- Lower-priority panels remain as-is.

**Files to change**: CSS in `monthly_epic_plan_progress_report.html`.

---

### 4.3 Auto-Apply or Prominent "Pending" State for Filters

**Problem**: Filters require an explicit Apply button click — users can change a filter and not notice data is stale.

**Action** (choose one):
- **Option A** (preferred): Add debounced auto-apply (`400ms`) on filter change. Remove or hide the Apply button.
- **Option B**: Style the Apply button with a pulsing highlight (e.g., `box-shadow` animation) when filters are "dirty" (changed but not applied).

**Files to change**: JS in `monthly_epic_plan_progress_report.html`.

---

## What to Preserve (Do Not Change)

The following features are working well and should not be altered:

| Feature | Why it works |
|---|---|
| Semantic color system (`--good`, `--warn`, `--bad`) | Used consistently across pills, dots, headers, badges — at-a-glance status reading excellent |
| Sticky table headers + expandable child rows | Correct pattern for dense epic data |
| Progress bar CSS transitions (`0.4s ease`) + toggle spring animation | Polished micro-interactions |
| Gantt chart side-by-side with table | Good power-user dual-view without page navigation |
| Project chips with color-coded left border (`.project-chip`) | Immediately scannable in dense table |
| `excl-btn` "See excluded epics" drawer | Surfaces excluded data without cluttering main view |

---

## Implementation Order

```
Phase 1 (Critical)
  └── 1.1 Collapse secondary filters

Phase 2 (Moderate — can be done in parallel)
  ├── 2.1 Font size floor
  ├── 2.2 Employee hint text
  ├── 2.3 overdue-threshold-input fix
  ├── 2.4 Controls grid responsive
  ├── 2.5 Button variant consolidation
  └── 2.6 Unified toggle pattern

Phase 3 (Accessibility — can be done in parallel with Phase 2)
  ├── 3.1 Focus ring opacity
  ├── 3.2 Fieldset for radio group
  ├── 3.3 Touch targets
  └── 3.4 Contrast for small text

Phase 4 (Polish — after Phase 2 & 3)
  ├── 4.1 Drawer width
  ├── 4.2 Section separators
  └── 4.3 Auto-apply filters
```

---

## Files Affected

| File | Sections touched |
|---|---|
| `monthly_epic_plan_progress_report.html` | CSS variables, controls HTML, focus ring, font sizes, button classes, drawer width, toggle markup |
| `monthly_epic_plan_progress_service.py` | Template/generator sections if controls HTML is generated here |

---

## Change Notes

- 2026-06-08: Added a report-level **Include Bug Subtasks** checkbox to the Monthly Epic Plan Progress controls. It is checked by default and re-fetches `/api/monthly-epic-plan-progress/summary` with `include_bug_subtasks=1`; when unchecked, the API receives `include_bug_subtasks=0` and excludes Bug Subtask issue keys before planned-hour, logged-hour, estimate-rollup, child-row, and Gantt calculations. The existing Story Overrun drawer bug toggle remains a drawer-local diagnostic switch inside the current report scope.
- 2026-05-20: Fixed the report table renderer so the Start Signal and End Signal cells read the selected month inside `rowToHtml`. Without that local month value, the API response loaded successfully but rendering failed with `monthYm is not defined`, leaving the page in the "Could not load report" state and clearing employee data.
- 2026-05-20: Added **Resource Planning** summary panel (section `#res-summary-panel`) inserted right before the "By project" section. Displays Total resources (Head Count + Capacity), Dev Resources (Head Count + Capacity + Leaves + Availability), and Support Resources (same 4 metrics, only shown when a support team is configured). Values are derived client-side from existing `workforce` and `support_team` payloads — no backend change required. Cards use the same blue palette as the capacity screenshot reference. JS: new `_resSummaryState` object and `renderResourceSummary()` function; called from both `renderWorkforce()` and `renderSupportTeam()`.
- 2026-05-20: Added **Process team auto-exclusion** from the employee dropdown. On first load (when `assignee_filter_active` is `false`), all team members are checked except those belonging to a team whose `aria-label` contains "process" (case-insensitive). After auto-excluding, `loadSummary` is re-triggered via `setTimeout(loadSummary, 0)` so capacity, leave, and availability stats immediately reflect the dev-only headcount. Subsequent renders (where the server returns `assignee_filter_active: true`) preserve the user's manual selection. Tests added in `tests/test_monthly_epic_plan_progress.py`.
