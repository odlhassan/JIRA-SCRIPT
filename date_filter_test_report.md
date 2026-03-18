# Date Filter Testing Report
## Nested View Report - Date Filter Functionality Test

**Test Date:** 2026-03-11
**Test URL:** http://127.0.0.1:3000/nested_view_report.html
**Browser:** Playwright Chromium

---

## Executive Summary

**FINDING:** The date filter UI is working correctly, but the scorecard values DO NOT update when the date range is changed.

**ROOT CAUSE:** The scorecard calculations are based on statically embedded data in the HTML file (`reportData.rows`), not on the API responses fetched when the date filter changes.

---

## Test Results

### Initial State

**Date Range:**
- From Date: 2026-01-31
- To Date: 2026-02-27

**Initial Scorecard Values:**
- Total Capacity: 4900h
- Total Planned Hours: 4602.91h
- Total Leaves Planned: 224h
- Total Actual Hours: 4117.01h
- Availability: 4676h

### After Changing to January 2026 (2026-01-01 to 2026-01-31)

**New Date Range:**
- From Date: 2026-01-01
- To Date: 2026-01-31

**Final Scorecard Values:**
- Total Capacity: 4900h (NO CHANGE)
- Total Planned Hours: 4602.91h (NO CHANGE)
- Total Leaves Planned: 224h (NO CHANGE)
- Total Actual Hours: 4117.01h (NO CHANGE)
- Availability: 4676h (NO CHANGE)

**Result:** Values DID NOT CHANGE

### After Changing to February 2026 (2026-02-01 to 2026-02-28)

**Scorecard Values:**
- Total Capacity: 4900h (NO CHANGE)
- Total Planned Hours: 4602.91h (NO CHANGE)
- Total Leaves Planned: 224h (NO CHANGE)
- Total Actual Hours: 4117.01h (NO CHANGE)
- Availability: 4676h (NO CHANGE)

**Result:** Values DID NOT CHANGE

---

## Technical Findings

### 1. UI Behavior

✓ **Date input fields are working correctly**
  - Successfully changed From date to 2026-01-01
  - Successfully changed To date to 2026-01-31
  - Date values are reflected in the input fields

✓ **Apply button is working correctly**
  - Button ID: `date-filter-apply`
  - Button is clickable
  - Click event is triggered successfully

### 2. Network Activity

✓ **API calls ARE being made when Apply is clicked**

API calls observed:
```
- /api/nested-view/actual-hours?from=2026-01-01&to=2026-01-31&mode=log_date&report=nested_view (200 OK)
- /api/report-date-filter (200 OK)
- /api/nested-view/actual-hours?from=2026-01-31&to=2026-02-28&mode=log_date&report=nested_view (200 OK)
```

✓ **Date parameters are being sent correctly in API requests**
  - The from/to dates in the URL match the selected dates
  - API responses return 200 OK status

### 3. JavaScript State

**Date filter status message:**
- "Apply recomputes the report even for the current date range."

**Table rendering:**
- Total table rows: 0
- No data is being displayed in the nested table

**JavaScript application state:**
```json
{
  "hasApplyFunction": false,
  "hasState": false,
  "dateFrom": "2026-01-31",
  "dateTo": "2026-02-28"
}
```

### 4. Console Errors

**Errors found:** 1
- "Failed to load resource: the server responded with a status of 404 (NOT FOUND)"
  - This appears to be a font file (material-symbols-outlined.woff2)
  - NOT related to the date filter issue

**No JavaScript errors** related to date filtering or scorecard calculation

### 5. Code Analysis

**Data Source:**
The page uses embedded static data:
```javascript
const reportData = {
  "generated_at": "2026-03-11 10:26 UTC",
  "source_file": "canonical_db",
  "rows": [...]  // Static data embedded in HTML
}
```

**Scorecard Calculation:**
```javascript
// Line 3424-3425
let totalPlannedHours = 0;
// ... calculation based on rows ...

// Line 5276
scorecardSourceRows = scorecardFilteredByDate;

// Line 5894
void updateScoreCards(scorecardSourceRows, requestVersion, renderSelection);
```

**Date Filter Apply Function:**
```javascript
// Line 4933
async function applyPendingDateRange() {
  // ... fetch new data from API ...
  const payload = await fetchActualHoursForDateRange(nextFrom, nextTo, nextMode, selectedTeamAssignees);
  applyFetchedActualHours(payload);  // Updates actual hours only
  // ... 
  rerender(true, { selection: {...} });  // Rerender with same source data
}
```

**Key Issue:**
The `scorecardSourceRows` variable is populated from `reportData.rows` which is static embedded data. When the date filter changes:
1. API calls are made to fetch updated actual hours
2. The actual hours are applied to existing rows
3. BUT the base data (`reportData.rows`) remains unchanged
4. Scorecards are recalculated using the same static source data

---

## Screenshots

Generated screenshots:
- `screenshot_1_initial.png` - Initial state with dates 2026-01-31 to 2026-02-27
- `screenshot_2_after_filter.png` - After changing to 2026-01-01 to 2026-01-31
- `screenshot_final_test.png` - Final state after all tests

---

## Recommendations

### Issue Explanation

The nested view report appears to be designed as a **static HTML export** with client-side date filtering capabilities for actual hours only. The planned hours, capacity, and other metrics are calculated from the embedded dataset which represents a specific time period (the period when the HTML was generated).

### Possible Solutions

**Option 1: Expected Behavior (No Fix Needed)**
If this is the intended design - i.e., the HTML is a static export for a specific date range and the date filter only adjusts actual hours for comparison purposes - then no fix is needed. The user should regenerate the report HTML for different date ranges.

**Option 2: Dynamic Data Loading (Requires Backend API)**
If scorecards should update with date changes, the solution would be:
1. Create API endpoints to provide filtered data based on date range
2. Modify `applyPendingDateRange()` to fetch complete filtered dataset
3. Update `reportData.rows` with the new data
4. Trigger scorecard recalculation

**Option 3: Client-Side Filtering (Limited)**
Filter the existing `reportData.rows` based on date range. This would only work if all the necessary raw data (with dates) is already embedded in the HTML.

---

## Conclusion

**Status:** Date filter UI is functional, API calls are successful, but scorecard values do not update.

**Cause:** Static data architecture - scorecards are calculated from embedded data that doesn't change when date filter is applied.

**Impact:** Users cannot dynamically filter the report by different date ranges to see how metrics change over time.

**Next Steps:** Clarify the intended behavior with the product owner/developer to determine if this is:
1. A bug that needs fixing
2. Expected behavior (static export design)
3. A feature that needs to be implemented

---

## Test Artifacts

- Test script: `test_date_filter.py`
- Detailed test script: `test_date_filter_detailed.py`
- Screenshots: `screenshot_1_initial.png`, `screenshot_2_after_filter.png`, `screenshot_final_test.png`
- This report: `date_filter_test_report.md`

---

**Tested by:** Automated Playwright Test Suite
**Report Generated:** 2026-03-11
