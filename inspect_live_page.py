#!/usr/bin/env python3
"""
Inspect live page at http://127.0.0.1:XXXX/team_capacity_planner.html
Uses playwright (installed via pip) or requests+selenium as fallback.
"""
import sys
import time
import json
import subprocess
import re

try:
    from playwright.sync_api import sync_playwright
    HAS_PLAYWRIGHT = True
except ImportError:
    HAS_PLAYWRIGHT = False
    print("⚠️  Playwright not available, trying alternative methods...")

def inspect_with_playwright():
    """Use Playwright to inspect the page."""
    ports = [3000, 3001, 4173, 5000, 8000, 8080]
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        
        console_logs = []
        page_errors = []
        network_errors = []
        
        def handle_console(msg):
            console_logs.append({
                'type': msg.type,
                'text': msg.text
            })
        
        def handle_error(err):
            page_errors.append({
                'message': str(err)
            })
        
        page.on("console", handle_console)
        page.on("pageerror", handle_error)
        
        url = None
        for port in ports:
            try:
                test_url = f"http://127.0.0.1:{port}/team_capacity_planner.html"
                try:
                    response = page.goto(test_url, wait_until="networkidle", timeout=5000)
                    if response and response.ok:
                        url = test_url
                        print(f"✓ Connected to {url}\n")
                        break
                except:
                    pass
            except:
                pass
        
        if not url:
            print("✗ Could not connect to any port (3000, 3001, 4173, 5000, 8000, 8080)")
            browser.close()
            return False
        
        # Wait for page to load
        page.wait_for_timeout(2000)
        
        # Evaluate page structure
        findings = page.evaluate("""() => {
            const result = {
                wiContainerExists: !!document.getElementById('wi-container'),
                epicRowsCount: 0,
                firstRowsInfo: [],
                epicRowStyles: {},
                epicRowHeadStyles: {},
                boundingBoxes: {}
            };
            
            const wiContainer = document.getElementById('wi-container');
            if (wiContainer) {
                const epicRows = wiContainer.querySelectorAll('.epic-row');
                result.epicRowsCount = epicRows.length;
                
                for (let i = 0; i < Math.min(3, epicRows.length); i++) {
                    const row = epicRows[i];
                    const rowHead = row.querySelector('.epic-row-head');
                    const rowText = row.textContent || '';
                    const headText = rowHead ? (rowHead.textContent || '') : '';
                    
                    result.firstRowsInfo.push({
                        rowIndex: i,
                        textContent: rowText.substring(0, 150),
                        headTextContent: headText.substring(0, 100),
                        rowHeight: row.offsetHeight,
                        rowWidth: row.offsetWidth,
                        headHeight: rowHead ? rowHead.offsetHeight : null,
                        display: window.getComputedStyle(row).display,
                        visibility: window.getComputedStyle(row).visibility,
                        height: window.getComputedStyle(row).height,
                        minHeight: window.getComputedStyle(row).minHeight,
                        maxHeight: window.getComputedStyle(row).maxHeight,
                        overflow: window.getComputedStyle(row).overflow
                    });
                }
                
                if (epicRows.length > 0) {
                    const firstRow = epicRows[0];
                    const styles = window.getComputedStyle(firstRow);
                    result.epicRowStyles = {
                        display: styles.display,
                        visibility: styles.visibility,
                        height: styles.height,
                        minHeight: styles.minHeight,
                        maxHeight: styles.maxHeight,
                        overflow: styles.overflow,
                        opacity: styles.opacity,
                        lineHeight: styles.lineHeight,
                        fontSize: styles.fontSize,
                        backgroundColor: styles.backgroundColor
                    };
                    
                    result.boundingBoxes.epicRow = {
                        width: firstRow.getBoundingClientRect().width,
                        height: firstRow.getBoundingClientRect().height
                    };
                    
                    const rowHead = firstRow.querySelector('.epic-row-head');
                    if (rowHead) {
                        const headStyles = window.getComputedStyle(rowHead);
                        result.epicRowHeadStyles = {
                            display: headStyles.display,
                            height: headStyles.height,
                            minHeight: headStyles.minHeight,
                            maxHeight: headStyles.maxHeight,
                            overflow: headStyles.overflow,
                            visibility: headStyles.visibility,
                            opacity: headStyles.opacity,
                            fontSize: headStyles.fontSize,
                            lineHeight: headStyles.lineHeight
                        };
                        
                        result.boundingBoxes.epicRowHead = {
                            width: rowHead.getBoundingClientRect().width,
                            height: rowHead.getBoundingClientRect().height
                        };
                    }
                }
            }
            
            return result;
        }""")
        
        # Print findings
        print("=" * 50)
        print("PAGE INSPECTION FINDINGS")
        print("=" * 50)
        
        print("\n1. WI-CONTAINER & EPIC-ROWS:")
        print(f"   - #wi-container exists: {findings['wiContainerExists']}")
        print(f"   - .epic-row count: {findings['epicRowsCount']}")
        print(f"   - Contains epic-row nodes: {'YES' if findings['epicRowsCount'] > 0 else 'NO'}")
        
        if findings['firstRowsInfo']:
            print("\n2. FIRST 1-3 ROWS CONTENT:")
            for row in findings['firstRowsInfo']:
                print(f"\n   Row {row['rowIndex']}:")
                text = row['textContent']
                if text:
                    print(f"   - textContent: \"{text[:80]}{'...' if len(text) > 80 else ''}\"")
                head_text = row['headTextContent']
                if head_text:
                    print(f"   - headTextContent: \"{head_text[:80]}{'...' if len(head_text) > 80 else ''}\"")
                print(f"   - Dimensions: {row['rowWidth']}x{row['rowHeight']}px (head height: {row['headHeight']}px)")
        
        print("\n3. COMPUTED STYLES FOR .epic-row:")
        for key, val in findings['epicRowStyles'].items():
            if val and val not in ['auto', 'none', '']:
                print(f"   - {key}: {val}")
        
        print("\n4. COMPUTED STYLES FOR .epic-row-head:")
        for key, val in findings['epicRowHeadStyles'].items():
            if val and val not in ['auto', 'none', '']:
                print(f"   - {key}: {val}")
        
        print("\n5. BOUNDING BOXES:")
        if findings['boundingBoxes'].get('epicRow'):
            bb = findings['boundingBoxes']['epicRow']
            print(f"   - .epic-row: {bb['width']:.0f}x{bb['height']:.0f}px")
        if findings['boundingBoxes'].get('epicRowHead'):
            bb = findings['boundingBoxes']['epicRowHead']
            print(f"   - .epic-row-head: {bb['width']:.0f}x{bb['height']:.0f}px")
        
        print("\n6. CONSOLE ERRORS/WARNINGS:")
        errors = [m for m in console_logs if m['type'] in ['error', 'warning']]
        if errors:
            for msg in errors[:5]:
                print(f"   [{msg['type']}] {msg['text']}")
        else:
            print("   (none)")
        
        print("\n7. PAGE ERRORS:")
        if page_errors:
            for err in page_errors[:5]:
                print(f"   {err['message']}")
        else:
            print("   (none)")
        
        browser.close()
        return True


if __name__ == '__main__':
    if HAS_PLAYWRIGHT:
        success = inspect_with_playwright()
        sys.exit(0 if success else 1)
    else:
        print("ERROR: Playwright is required but not installed.")
        print("Install with: pip install playwright")
        sys.exit(1)
