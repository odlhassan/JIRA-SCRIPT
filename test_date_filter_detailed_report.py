from playwright.sync_api import sync_playwright
import time
import json
import sys
import io

# Fix console encoding for Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def capture_scorecard_values(page):
    """Capture all scorecard values from the page"""
    scorecard_selectors = {
        'Total Capacity': 'score-total-capacity',
        'Total Planned Hours': 'score-total-planned',
        'Total Leaves Planned': 'score-total-leaves-planned',
        'Total Actual Hours': 'score-total-logged',
        'Availability': 'score-total-capacity-planned-leaves-adjusted'
    }
    
    values = {}
    for label, elem_id in scorecard_selectors.items():
        try:
            element = page.locator(f'#{elem_id}')
            if element.count() > 0:
                value = element.text_content().strip()
                values[label] = value
            else:
                values[label] = "NOT FOUND"
        except Exception as e:
            values[label] = f"ERROR: {str(e)}"
    
    return values

def print_scorecard_values(values, title):
    """Print scorecard values in a formatted way"""
    print("\n" + "="*70)
    print(f"{title}")
    print("="*70)
    for label, value in values.items():
        print(f"{label:30s}: {value}")

def count_table_rows(page):
    """Count visible rows in the nested table"""
    try:
        rows = page.locator('.nested-table tbody tr').all()
        return len(rows)
    except:
        return 0

def get_date_filter_status(page):
    """Get the status message from the date filter"""
    try:
        status_element = page.locator('#date-filter-status')
        if status_element.count() > 0:
            return status_element.text_content().strip()
        return "Status element not found"
    except Exception as e:
        return f"ERROR: {str(e)}"

def test_date_filter_comprehensive():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        
        # Navigate to the page
        print("="*70)
        print("STARTING DATE FILTER TEST")
        print("="*70)
        print("Navigating to http://127.0.0.1:3000/nested_view_report.html")
        page.goto('http://127.0.0.1:3000/nested_view_report.html')
        
        # Wait for page to load fully
        print("Waiting for page to load fully...")
        time.sleep(3)
        
        # ===== TEST 1: INITIAL STATE =====
        print("\n" + "="*70)
        print("TEST 1: INITIAL STATE (Page Load)")
        print("="*70)
        
        # Get initial date range
        try:
            from_date_initial = page.locator('#date-filter-from').input_value()
            to_date_initial = page.locator('#date-filter-to').input_value()
            print(f"Initial Date Range: {from_date_initial} to {to_date_initial}")
        except Exception as e:
            print(f"ERROR getting initial dates: {str(e)}")
            from_date_initial = "ERROR"
            to_date_initial = "ERROR"
        
        # Capture initial scorecard values
        initial_values = capture_scorecard_values(page)
        print_scorecard_values(initial_values, "INITIAL SCORECARD VALUES")
        
        # Count initial table rows
        initial_rows = count_table_rows(page)
        print(f"\nTable rows: {initial_rows}")
        
        # Get initial status
        initial_status = get_date_filter_status(page)
        print(f"Status message: {initial_status}")
        
        # Take initial screenshot
        page.screenshot(path='screenshot_01_initial.png')
        print("\n[OK] Screenshot saved: screenshot_01_initial.png")
        
        # ===== TEST 2: JANUARY 2025 =====
        print("\n" + "="*70)
        print("TEST 2: JANUARY 2025 (2025-01-01 to 2025-01-31)")
        print("="*70)
        
        # Change dates
        print("Changing dates...")
        page.locator('#date-filter-from').fill('2025-01-01')
        print("  Set From Date: 2025-01-01")
        page.locator('#date-filter-to').fill('2025-01-31')
        print("  Set To Date: 2025-01-31")
        
        # Click Apply
        print("Clicking Apply button...")
        page.locator('#date-filter-apply').click()
        print("  [OK] Apply button clicked")
        
        # Wait 5 seconds
        print("Waiting 5 seconds for everything to complete...")
        time.sleep(5)
        
        # Take screenshot
        page.screenshot(path='screenshot_02_jan2025.png')
        print("[OK] Screenshot saved: screenshot_02_jan2025.png")
        
        # Capture January scorecard values
        jan_values = capture_scorecard_values(page)
        print_scorecard_values(jan_values, "JANUARY 2025 SCORECARD VALUES")
        
        # Count January table rows
        jan_rows = count_table_rows(page)
        print(f"\nTable rows: {jan_rows}")
        
        # Get January status
        jan_status = get_date_filter_status(page)
        print(f"Status message: {jan_status}")
        
        # ===== TEST 3: JUNE 2025 =====
        print("\n" + "="*70)
        print("TEST 3: JUNE 2025 (2025-06-01 to 2025-06-30)")
        print("="*70)
        
        # Change dates
        print("Changing dates...")
        page.locator('#date-filter-from').fill('2025-06-01')
        print("  Set From Date: 2025-06-01")
        page.locator('#date-filter-to').fill('2025-06-30')
        print("  Set To Date: 2025-06-30")
        
        # Click Apply
        print("Clicking Apply button...")
        page.locator('#date-filter-apply').click()
        print("  [OK] Apply button clicked")
        
        # Wait 5 seconds
        print("Waiting 5 seconds for everything to complete...")
        time.sleep(5)
        
        # Take screenshot
        page.screenshot(path='screenshot_03_jun2025.png')
        print("[OK] Screenshot saved: screenshot_03_jun2025.png")
        
        # Capture June scorecard values
        jun_values = capture_scorecard_values(page)
        print_scorecard_values(jun_values, "JUNE 2025 SCORECARD VALUES")
        
        # Count June table rows
        jun_rows = count_table_rows(page)
        print(f"\nTable rows: {jun_rows}")
        
        # Get June status
        jun_status = get_date_filter_status(page)
        print(f"Status message: {jun_status}")
        
        # ===== FINAL REPORT =====
        print("\n" + "="*70)
        print("FINAL COMPREHENSIVE REPORT")
        print("="*70)
        
        print("\n--- DATE RANGES TESTED ---")
        print(f"Initial:     {from_date_initial} to {to_date_initial}")
        print(f"January:     2025-01-01 to 2025-01-31")
        print(f"June:        2025-06-01 to 2025-06-30")
        
        print("\n--- SCORECARD VALUES COMPARISON ---")
        all_labels = list(initial_values.keys())
        
        # Create comparison table
        print(f"\n{'Metric':<30} | {'Initial':<15} | {'January':<15} | {'June':<15}")
        print("-" * 85)
        for label in all_labels:
            initial = initial_values.get(label, "N/A")
            jan = jan_values.get(label, "N/A")
            jun = jun_values.get(label, "N/A")
            print(f"{label:<30} | {initial:<15} | {jan:<15} | {jun:<15}")
        
        print("\n--- CHANGES DETECTED ---")
        changes_detected = False
        for label in all_labels:
            initial = initial_values.get(label, "N/A")
            jan = jan_values.get(label, "N/A")
            jun = jun_values.get(label, "N/A")
            
            if initial != jan or jan != jun or initial != jun:
                print(f"✓ {label}: Values changed across tests")
                changes_detected = True
            else:
                print(f"✗ {label}: Values stayed the same ({initial})")
        
        if not changes_detected:
            print("\n⚠ WARNING: No scorecard values changed between any tests!")
        else:
            print("\n✓ At least some scorecard values changed between tests")
        
        print("\n--- TABLE ROWS COMPARISON ---")
        print(f"Initial:  {initial_rows} rows")
        print(f"January:  {jan_rows} rows")
        print(f"June:     {jun_rows} rows")
        
        if initial_rows != jan_rows or jan_rows != jun_rows or initial_rows != jun_rows:
            print("✓ Table rows changed between tests")
        else:
            print("✗ Table rows stayed the same")
        
        print("\n--- STATUS MESSAGES ---")
        print(f"Initial:  {initial_status}")
        print(f"January:  {jan_status}")
        print(f"June:     {jun_status}")
        
        print("\n--- SCREENSHOTS SAVED ---")
        print("screenshot_01_initial.png  - Initial page load")
        print("screenshot_02_jan2025.png  - After filtering to January 2025")
        print("screenshot_03_jun2025.png  - After filtering to June 2025")
        
        print("\n" + "="*70)
        print("TEST COMPLETE")
        print("="*70)
        
        # Keep browser open for a moment to review
        print("\nKeeping browser open for 3 seconds for review...")
        time.sleep(3)
        
        browser.close()

if __name__ == "__main__":
    test_date_filter_comprehensive()
