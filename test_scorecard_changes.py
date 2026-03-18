"""
Test script to verify scorecard values change when dates are changed
on the nested_view_report.html page
"""
import asyncio
from playwright.async_api import async_playwright
import time

async def test_scorecard_changes():
    results = []
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
        
        # Enable console logging
        console_messages = []
        errors = []
        
        page.on("console", lambda msg: console_messages.append(f"{msg.type}: {msg.text}"))
        page.on("pageerror", lambda err: errors.append(str(err)))
        
        try:
            print("Step 1: Navigating to http://127.0.0.1:3000/nested_view_report.html")
            await page.goto("http://127.0.0.1:3000/nested_view_report.html")
            
            print("Step 2: Waiting 5 seconds for page to load...")
            await asyncio.sleep(5)
            
            # Take screenshot 1
            await page.screenshot(path="screenshot_1_initial.png")
            print("Step 3: Screenshot 1 saved")
            
            # Get initial scorecard values
            print("\nStep 4: Reading initial scorecard values...")
            scorecard_values_1 = await get_scorecard_values(page)
            print(f"Initial Scorecard Values: {scorecard_values_1}")
            results.append(f"TEST 1 - Initial Load:\n{scorecard_values_1}\n")
            
            # Get current date range
            from_date = await page.input_value("#date-filter-from")
            to_date = await page.input_value("#date-filter-to")
            print(f"Current date range: From={from_date}, To={to_date}")
            results.append(f"Date Range: {from_date} to {to_date}\n")
            
            # TEST 2: Change dates to Jan 2026
            print("\n" + "="*60)
            print("TEST 2: Changing dates to January 2026")
            print("="*60)
            
            # Clear and fill from date
            await page.fill("#date-filter-from", "2026-01-01")
            print("Step 5: Set From date to 2026-01-01")
            
            # Clear and fill to date
            await page.fill("#date-filter-to", "2026-01-31")
            print("Step 6: Set To date to 2026-01-31")
            
            # Click Apply button
            await page.click("#date-filter-apply")
            print("Step 7: Clicked Apply button")
            
            # Wait 5 seconds
            print("Step 8: Waiting 5 seconds...")
            await asyncio.sleep(5)
            
            # Take screenshot 2
            await page.screenshot(path="screenshot_2_jan2026.png")
            print("Step 9: Screenshot 2 saved")
            
            # Get scorecard values after first change
            print("\nStep 10: Reading scorecard values after date change...")
            scorecard_values_2 = await get_scorecard_values(page)
            print(f"Jan 2026 Scorecard Values: {scorecard_values_2}")
            results.append(f"\nTEST 2 - January 2026 (2026-01-01 to 2026-01-31):\n{scorecard_values_2}\n")
            
            # TEST 3: Change dates to Feb 2026
            print("\n" + "="*60)
            print("TEST 3: Changing dates to February 2026")
            print("="*60)
            
            # Clear and fill from date
            await page.fill("#date-filter-from", "2026-02-01")
            print("Step 11: Set From date to 2026-02-01")
            
            # Clear and fill to date
            await page.fill("#date-filter-to", "2026-02-28")
            print("Step 12: Set To date to 2026-02-28")
            
            # Click Apply button
            await page.click("#date-filter-apply")
            print("Step 13: Clicked Apply button")
            
            # Wait 5 seconds
            print("Step 14: Waiting 5 seconds...")
            await asyncio.sleep(5)
            
            # Take screenshot 3
            await page.screenshot(path="screenshot_3_feb2026.png")
            print("Step 15: Screenshot 3 saved")
            
            # Get scorecard values after second change
            print("\nStep 16: Reading scorecard values after second date change...")
            scorecard_values_3 = await get_scorecard_values(page)
            print(f"Feb 2026 Scorecard Values: {scorecard_values_3}")
            results.append(f"\nTEST 3 - February 2026 (2026-02-01 to 2026-02-28):\n{scorecard_values_3}\n")
            
            # Check for console errors
            print("\n" + "="*60)
            print("CONSOLE ERRORS CHECK")
            print("="*60)
            if errors:
                print(f"Found {len(errors)} JavaScript errors:")
                for error in errors:
                    print(f"  - {error}")
                results.append(f"\nJavaScript Errors Found: {len(errors)}\n")
                for error in errors:
                    results.append(f"  - {error}\n")
            else:
                print("No JavaScript errors found!")
                results.append("\nNo JavaScript errors found.\n")
            
            # Check for ReferenceError in console
            reference_errors = [msg for msg in console_messages if "ReferenceError" in msg or "is not defined" in msg]
            if reference_errors:
                print(f"\nFound {len(reference_errors)} reference errors in console:")
                for err in reference_errors:
                    print(f"  - {err}")
                results.append(f"\nReference Errors in Console: {len(reference_errors)}\n")
                for err in reference_errors:
                    results.append(f"  - {err}\n")
            
            # Analysis
            print("\n" + "="*60)
            print("ANALYSIS")
            print("="*60)
            
            values_changed_test2 = scorecard_values_1 != scorecard_values_2
            values_changed_test3 = scorecard_values_2 != scorecard_values_3
            
            print(f"\nDid values change from Test 1 to Test 2? {values_changed_test2}")
            print(f"Did values change from Test 2 to Test 3? {values_changed_test3}")
            
            results.append(f"\n{'='*60}\n")
            results.append("ANALYSIS SUMMARY\n")
            results.append(f"{'='*60}\n")
            results.append(f"Values changed from Initial to Jan 2026: {values_changed_test2}\n")
            results.append(f"Values changed from Jan 2026 to Feb 2026: {values_changed_test3}\n")
            
            if not values_changed_test2 and not values_changed_test3:
                results.append("\n[WARNING] Scorecard values DID NOT CHANGE when dates were changed!\n")
                print("\n[WARNING] Scorecard values DID NOT CHANGE when dates were changed!")
            elif values_changed_test2 and values_changed_test3:
                results.append("\n[SUCCESS] Scorecard values changed correctly with date changes!\n")
                print("\n[SUCCESS] Scorecard values changed correctly with date changes!")
            else:
                results.append("\n[PARTIAL] Some values changed but not consistently.\n")
                print("\n[PARTIAL] Some values changed but not consistently.")
            
        except Exception as e:
            print(f"Error during test: {str(e)}")
            results.append(f"\nError during test: {str(e)}\n")
            import traceback
            traceback.print_exc()
        
        finally:
            # Keep browser open for 5 seconds to see final state
            await asyncio.sleep(5)
            await browser.close()
    
    # Write results to file
    with open("test_results.txt", "w", encoding="utf-8") as f:
        f.writelines(results)
    
    print("\n" + "="*60)
    print("Test complete! Results saved to test_results.txt")
    print("Screenshots saved as: screenshot_1_initial.png, screenshot_2_jan2026.png, screenshot_3_feb2026.png")
    print("="*60)

async def get_scorecard_values(page):
    """Extract all scorecard values from the page"""
    try:
        scorecard_data = {}
        
        # Extract specific scorecard values by ID
        scorecard_ids = {
            "score-total-capacity": "Total Capacity",
            "score-total-planned": "Total Planned Hours",
            "score-total-leaves-planned": "Total Leaves Planned",
            "score-total-logged": "Total Actual Hours",
            "score-total-capacity-planned-leaves-adjusted": "Availability (Capacity - Leaves)",
            "score-total-leaves": "Total Leaves Taken",
        }
        
        for elem_id, label in scorecard_ids.items():
            try:
                element = await page.query_selector(f"#{elem_id}")
                if element:
                    text = await element.text_content()
                    scorecard_data[label] = text.strip() if text else "N/A"
                else:
                    scorecard_data[label] = "Element not found"
            except Exception as e:
                scorecard_data[label] = f"Error: {str(e)}"
        
        # Get row count from visible table rows
        try:
            rows = await page.query_selector_all("tbody tr:not([hidden])")
            scorecard_data["Row Count"] = f"{len(rows)} rows visible"
        except Exception as e:
            scorecard_data["Row Count"] = f"Error getting row count: {str(e)}"
        
        # Format output
        result_lines = []
        for key, value in scorecard_data.items():
            result_lines.append(f"{key}: {value}")
        
        return "\n".join(result_lines)
        
    except Exception as e:
        return f"Error extracting scorecard values: {str(e)}"

if __name__ == "__main__":
    asyncio.run(test_scorecard_changes())
