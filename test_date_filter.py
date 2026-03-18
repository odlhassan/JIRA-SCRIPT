from playwright.sync_api import sync_playwright
import time
import json
import sys
import io

# Fix console encoding for Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def test_date_filter():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        
        # Enable console logging
        console_messages = []
        page.on("console", lambda msg: console_messages.append({
            "type": msg.type,
            "text": msg.text
        }))
        
        # Navigate to the page
        print("Navigating to http://127.0.0.1:3000/nested_view_report.html")
        page.goto('http://127.0.0.1:3000/nested_view_report.html')
        
        # Wait for page to load
        print("Waiting for page to load...")
        time.sleep(3)
        
        # Take initial screenshot
        page.screenshot(path='screenshot_1_initial.png')
        print("[OK] Initial screenshot taken")
        
        # Get initial scorecard values
        print("\n" + "="*60)
        print("INITIAL SCORECARD VALUES")
        print("="*60)
        initial_values = {}
        
        scorecard_selectors = {
            'Total Capacity': 'score-total-capacity',
            'Total Planned Hours': 'score-total-planned',
            'Total Leaves Planned': 'score-total-leaves-planned',
            'Total Actual Hours': 'score-total-logged',
            'Availability': 'score-total-capacity-planned-leaves-adjusted'
        }
        
        for label, elem_id in scorecard_selectors.items():
            try:
                element = page.locator(f'#{elem_id}')
                if element.count() > 0:
                    value = element.text_content().strip()
                    initial_values[label] = value
                    print(f"{label:30s}: {value}")
                else:
                    initial_values[label] = "NOT FOUND"
                    print(f"{label:30s}: NOT FOUND")
            except Exception as e:
                initial_values[label] = f"ERROR: {str(e)}"
                print(f"{label:30s}: ERROR - {str(e)}")
        
        # Get current date filter values
        print("\n" + "="*60)
        print("CURRENT DATE FILTER VALUES")
        print("="*60)
        try:
            from_date = page.locator('#date-filter-from').input_value()
            to_date = page.locator('#date-filter-to').input_value()
            print(f"From Date: {from_date}")
            print(f"To Date:   {to_date}")
        except Exception as e:
            print(f"ERROR getting date values: {str(e)}")
            from_date = "ERROR"
            to_date = "ERROR"
        
        # Change the dates
        print("\n" + "="*60)
        print("CHANGING DATES")
        print("="*60)
        try:
            # Clear and fill From date
            page.locator('#date-filter-from').fill('2026-01-01')
            print("[OK] Set From Date to: 2026-01-01")
            
            # Clear and fill To date
            page.locator('#date-filter-to').fill('2026-01-31')
            print("[OK] Set To Date to: 2026-01-31")
            
            # Click Apply button (correct ID: date-filter-apply)
            apply_btn = page.locator('#date-filter-apply')
            if apply_btn.count() > 0:
                apply_btn.click()
                print("[OK] Clicked Apply button")
            else:
                print("[ERROR] Apply button not found!")
            
        except Exception as e:
            print(f"[ERROR] changing dates: {str(e)}")
        
        # Wait for API calls to complete
        print("\n[WAIT] Waiting 3 seconds for API calls...")
        time.sleep(3)
        
        # Take screenshot after applying filter
        page.screenshot(path='screenshot_2_after_filter.png')
        print("[OK] After-filter screenshot taken")
        
        # Get final scorecard values
        print("\n" + "="*60)
        print("FINAL SCORECARD VALUES")
        print("="*60)
        final_values = {}
        
        for label, elem_id in scorecard_selectors.items():
            try:
                element = page.locator(f'#{elem_id}')
                if element.count() > 0:
                    value = element.text_content().strip()
                    final_values[label] = value
                    print(f"{label:30s}: {value}")
                else:
                    final_values[label] = "NOT FOUND"
                    print(f"{label:30s}: NOT FOUND")
            except Exception as e:
                final_values[label] = f"ERROR: {str(e)}"
                print(f"{label:30s}: ERROR - {str(e)}")
        
        # Check for changes
        print("\n" + "="*60)
        print("COMPARISON - BEFORE vs AFTER")
        print("="*60)
        changes_detected = False
        for label in scorecard_selectors.keys():
            initial = initial_values.get(label, "N/A")
            final = final_values.get(label, "N/A")
            if initial != final:
                print(f"[CHANGED] {label:30s}: {initial} -> {final}")
                changes_detected = True
            else:
                print(f"[NO CHANGE] {label:30s}: {initial}")
        
        # Check for error messages in date filter status
        print("\n" + "="*60)
        print("DATE FILTER STATUS")
        print("="*60)
        try:
            status_element = page.locator('#date-filter-status')
            if status_element.count() > 0:
                status_text = status_element.text_content().strip()
                print(f"Status: {status_text}")
            else:
                print("Status element not found")
        except Exception as e:
            print(f"ERROR checking status: {str(e)}")
        
        # Check console errors
        print("\n" + "="*60)
        print("CONSOLE MESSAGES")
        print("="*60)
        errors = [msg for msg in console_messages if msg['type'] == 'error']
        if errors:
            print(f"Found {len(errors)} console error(s):")
            for err in errors:
                print(f"  [ERROR] {err['text']}")
        else:
            print("[OK] No console errors found")
        
        if console_messages:
            warnings = [msg for msg in console_messages if msg['type'] == 'warning']
            if warnings:
                print(f"\nFound {len(warnings)} console warning(s):")
                for warn in warnings[:5]:  # Show first 5 warnings
                    print(f"  [WARN] {warn['text']}")
        
        # Summary
        print("\n" + "="*60)
        print("SUMMARY")
        print("="*60)
        print(f"Initial Date Range:  {from_date} to {to_date}")
        print(f"New Date Range:      2026-01-01 to 2026-01-31")
        print(f"Changes Detected:    {'YES' if changes_detected else 'NO'}")
        print(f"Console Errors:      {len(errors)}")
        print("="*60)
        
        browser.close()

if __name__ == "__main__":
    test_date_filter()
