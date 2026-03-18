from playwright.sync_api import sync_playwright
import time
import json
import sys
import io

# Fix console encoding for Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def test_date_filter_with_network():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        
        # Track network requests
        api_calls = []
        def handle_request(request):
            if '/api/' in request.url or 'fetch' in request.url.lower():
                api_calls.append({
                    'url': request.url,
                    'method': request.method
                })
        
        def handle_response(response):
            if '/api/' in response.url or 'fetch' in response.url.lower():
                print(f"[API RESPONSE] {response.status} - {response.url}")
        
        page.on("request", handle_request)
        page.on("response", handle_response)
        
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
        print("Waiting for initial page load...")
        time.sleep(3)
        
        print("\n" + "="*60)
        print("TEST 1: Check if date inputs are working")
        print("="*60)
        
        # Get initial values
        from_date_initial = page.locator('#date-filter-from').input_value()
        to_date_initial = page.locator('#date-filter-to').input_value()
        print(f"Initial From Date: {from_date_initial}")
        print(f"Initial To Date:   {to_date_initial}")
        
        # Get initial scorecard value
        initial_planned = page.locator('#score-total-planned').text_content().strip()
        print(f"Initial Total Planned Hours: {initial_planned}")
        
        print("\n" + "="*60)
        print("TEST 2: Change dates to January 2026")
        print("="*60)
        
        # Change dates
        page.locator('#date-filter-from').fill('2026-01-01')
        page.locator('#date-filter-to').fill('2026-01-31')
        
        # Verify dates changed in UI
        from_date_new = page.locator('#date-filter-from').input_value()
        to_date_new = page.locator('#date-filter-to').input_value()
        print(f"After fill - From Date: {from_date_new}")
        print(f"After fill - To Date:   {to_date_new}")
        
        # Check status message
        status_before_apply = page.locator('#date-filter-status').text_content().strip()
        print(f"Status before apply: {status_before_apply}")
        
        # Click Apply
        print("\nClicking Apply button...")
        page.locator('#date-filter-apply').click()
        time.sleep(1)
        
        # Check status after apply
        status_after_apply = page.locator('#date-filter-status').text_content().strip()
        print(f"Status after apply: {status_after_apply}")
        
        # Wait for updates
        print("Waiting 3 seconds for any updates...")
        time.sleep(3)
        
        # Get new scorecard value
        final_planned = page.locator('#score-total-planned').text_content().strip()
        print(f"Final Total Planned Hours: {final_planned}")
        
        if initial_planned != final_planned:
            print(f"\n[SUCCESS] Values CHANGED: {initial_planned} -> {final_planned}")
        else:
            print(f"\n[ISSUE] Values DID NOT CHANGE: {initial_planned}")
        
        print("\n" + "="*60)
        print("TEST 3: Try a different date range")
        print("="*60)
        
        # Try February only
        page.locator('#date-filter-from').fill('2026-02-01')
        page.locator('#date-filter-to').fill('2026-02-28')
        print("Changed dates to February 2026")
        
        page.locator('#date-filter-apply').click()
        print("Clicked Apply")
        time.sleep(3)
        
        feb_planned = page.locator('#score-total-planned').text_content().strip()
        print(f"February Total Planned Hours: {feb_planned}")
        
        print("\n" + "="*60)
        print("TEST 4: Check the HTML table for any data")
        print("="*60)
        
        # Check if there's data in the table
        rows = page.locator('.nested-table tbody tr').all()
        print(f"Total table rows: {len(rows)}")
        
        if rows:
            # Get first few rows
            for i, row in enumerate(rows[:5]):
                try:
                    text = row.text_content()[:100]
                    print(f"Row {i}: {text}")
                except:
                    pass
        
        print("\n" + "="*60)
        print("TEST 5: Check JavaScript console for clues")
        print("="*60)
        
        errors = [msg for msg in console_messages if msg['type'] == 'error']
        print(f"Console Errors: {len(errors)}")
        for err in errors:
            print(f"  {err['text']}")
        
        logs = [msg for msg in console_messages if msg['type'] == 'log']
        print(f"\nConsole Logs: {len(logs)}")
        for log in logs[-10:]:  # Last 10 logs
            print(f"  {log['text']}")
        
        print("\n" + "="*60)
        print("TEST 6: Execute JavaScript to check app state")
        print("="*60)
        
        # Try to get app state if available
        try:
            app_state = page.evaluate("""() => {
                return {
                    hasApplyFunction: typeof window.applyFilters === 'function',
                    hasState: typeof window.appState !== 'undefined',
                    dateFrom: document.getElementById('date-filter-from')?.value,
                    dateTo: document.getElementById('date-filter-to')?.value
                }
            }""")
            print(f"App State Check: {json.dumps(app_state, indent=2)}")
        except Exception as e:
            print(f"Could not check app state: {e}")
        
        print("\n" + "="*60)
        print("SUMMARY")
        print("="*60)
        print(f"Initial Values: From={from_date_initial}, To={to_date_initial}, Planned={initial_planned}")
        print(f"January Values: From=2026-01-01, To=2026-01-31, Planned={final_planned}")
        print(f"February Values: From=2026-02-01, To=2026-02-28, Planned={feb_planned}")
        print(f"Values Changed: {initial_planned != final_planned or initial_planned != feb_planned}")
        
        page.screenshot(path='screenshot_final_test.png')
        browser.close()

if __name__ == "__main__":
    test_date_filter_with_network()
