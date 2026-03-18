from playwright.sync_api import sync_playwright
import time

def inspect_page():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        
        print("Navigating to page...")
        page.goto('http://127.0.0.1:3000/nested_view_report.html')
        time.sleep(3)
        
        print("\n=== Searching for date filter elements ===")
        
        # Find date filter inputs
        from_input = page.locator('input[type="date"]').first
        if from_input.count() > 0:
            print(f"Found date input (first): id={from_input.get_attribute('id')}")
        
        all_date_inputs = page.locator('input[type="date"]').all()
        print(f"Total date inputs found: {len(all_date_inputs)}")
        for i, inp in enumerate(all_date_inputs):
            try:
                input_id = inp.get_attribute('id')
                input_value = inp.input_value()
                print(f"  Date input {i}: id='{input_id}', value='{input_value}'")
            except:
                print(f"  Date input {i}: Could not get details")
        
        # Find buttons
        print("\n=== Searching for buttons ===")
        all_buttons = page.locator('button').all()
        print(f"Total buttons found: {len(all_buttons)}")
        for i, btn in enumerate(all_buttons):
            try:
                btn_id = btn.get_attribute('id')
                btn_text = btn.text_content()
                btn_class = btn.get_attribute('class')
                print(f"  Button {i}: id='{btn_id}', text='{btn_text}', class='{btn_class}'")
            except:
                print(f"  Button {i}: Could not get details")
        
        # Find elements with 'scorecard' in class
        print("\n=== Searching for scorecard elements ===")
        scorecards = page.locator('[class*="scorecard"]').all()
        print(f"Total scorecard elements found: {len(scorecards)}")
        for i, card in enumerate(scorecards[:5]):  # Show first 5
            try:
                card_class = card.get_attribute('class')
                card_id = card.get_attribute('id')
                card_text = card.text_content()[:100]
                print(f"  Scorecard {i}: id='{card_id}', class='{card_class}'")
                print(f"    Text: {card_text}")
            except:
                print(f"  Scorecard {i}: Could not get details")
        
        # Find elements with numbers that might be scorecards
        print("\n=== Searching for elements with 'capacity' or 'hours' text ===")
        capacity_elements = page.get_by_text("Capacity", exact=False).all()
        print(f"Elements with 'Capacity': {len(capacity_elements)}")
        
        hours_elements = page.get_by_text("Hours", exact=False).all()
        print(f"Elements with 'Hours': {len(hours_elements)}")
        
        # Try to get all divs and look for scorecard-like content
        print("\n=== Looking for div elements with scorecard-like structure ===")
        page.screenshot(path='inspect_screenshot.png')
        
        # Get the outer HTML of the body to see structure
        print("\n=== Getting page structure (first 5000 chars) ===")
        body_html = page.locator('body').inner_html()
        print(body_html[:5000])
        
        browser.close()

if __name__ == "__main__":
    inspect_page()
