const { chromium } = require('playwright');

async function testNestedViewReport() {
    const browser = await chromium.launch({ 
        headless: false,
        devtools: true
    });
    
    const context = await browser.newContext();
    const page = await context.newPage();
    
    const consoleMessages = [];
    const debugMessages = [];
    const errorMessages = [];
    
    page.on('console', msg => {
        const text = msg.text();
        consoleMessages.push({
            type: msg.type(),
            text: text,
            timestamp: new Date().toISOString()
        });
        
        if (text.includes('[DEBUG')) {
            debugMessages.push(text);
            console.log('🔍 DEBUG:', text);
        }
        
        if (msg.type() === 'error') {
            errorMessages.push(text);
            console.log('❌ ERROR:', text);
        }
    });
    
    page.on('pageerror', error => {
        errorMessages.push(error.message);
        console.log('❌ PAGE ERROR:', error.message);
    });
    
    console.log('='.repeat(80));
    console.log('STEP 1: Navigating to http://127.0.0.1:3000/nested_view_report.html');
    console.log('='.repeat(80));
    
    try {
        await page.goto('http://127.0.0.1:3000/nested_view_report.html', {
            waitUntil: 'networkidle',
            timeout: 30000
        });
        
        console.log('✓ Page loaded successfully');
        
        console.log('\n' + '='.repeat(80));
        console.log('STEP 2: Waiting 5 seconds for full page load');
        console.log('='.repeat(80));
        
        await page.waitForTimeout(5000);
        
        console.log('\n' + '='.repeat(80));
        console.log('STEP 3: Reading initial console messages');
        console.log('='.repeat(80));
        
        console.log(`\nTotal console messages so far: ${consoleMessages.length}`);
        console.log(`DEBUG messages found: ${debugMessages.length}`);
        
        if (debugMessages.length > 0) {
            console.log('\n[DEBUG MESSAGES - INITIAL]:');
            debugMessages.forEach((msg, idx) => {
                console.log(`  ${idx + 1}. ${msg}`);
            });
        }
        
        console.log('\n' + '='.repeat(80));
        console.log('STEP 4: Reading initial scorecard values');
        console.log('='.repeat(80));
        
        const initialScorecard = await page.evaluate(() => {
            const getValue = (id) => {
                const elem = document.getElementById(id);
                return elem ? elem.textContent.trim() : 'NOT FOUND';
            };
            
            return {
                totalCapacity: getValue('total-capacity'),
                totalPlannedHours: getValue('total-planned-hours'),
                totalActualHours: getValue('total-actual-hours'),
                utilizationPercent: getValue('utilization-percent'),
                totalIssues: getValue('total-issues'),
                completedIssues: getValue('completed-issues'),
                inProgressIssues: getValue('in-progress-issues'),
                todoIssues: getValue('todo-issues')
            };
        });
        
        console.log('\nINITIAL SCORECARD VALUES:');
        Object.entries(initialScorecard).forEach(([key, value]) => {
            console.log(`  ${key}: ${value}`);
        });
        
        const initialRowCount = await page.locator('table tbody tr').count();
        console.log(`\nInitial row count: ${initialRowCount}`);
        
        console.log('\n' + '='.repeat(80));
        console.log('STEP 5: Changing "From" date to 2025-06-01');
        console.log('='.repeat(80));
        
        const fromDateInput = page.locator('#date-filter-from');
        await fromDateInput.waitFor({ state: 'visible' });
        await fromDateInput.clear();
        await fromDateInput.fill('2025-06-01');
        console.log('✓ From date set to: 2025-06-01');
        
        console.log('\n' + '='.repeat(80));
        console.log('STEP 6: Changing "To" date to 2025-06-30');
        console.log('='.repeat(80));
        
        const toDateInput = page.locator('#date-filter-to');
        await toDateInput.waitFor({ state: 'visible' });
        await toDateInput.clear();
        await toDateInput.fill('2025-06-30');
        console.log('✓ To date set to: 2025-06-30');
        
        console.log('\n' + '='.repeat(80));
        console.log('STEP 7: Clicking Apply button');
        console.log('='.repeat(80));
        
        const initialDebugCount = debugMessages.length;
        
        const applyButton = page.locator('#date-filter-apply');
        await applyButton.waitFor({ state: 'visible' });
        await applyButton.click();
        console.log('✓ Apply button clicked');
        
        console.log('\n' + '='.repeat(80));
        console.log('STEP 8: Waiting 5 seconds after apply');
        console.log('='.repeat(80));
        
        await page.waitForTimeout(5000);
        
        console.log('\n' + '='.repeat(80));
        console.log('STEP 9: Reading NEW console messages after apply');
        console.log('='.repeat(80));
        
        const newDebugMessages = debugMessages.slice(initialDebugCount);
        console.log(`\nNEW DEBUG messages after apply: ${newDebugMessages.length}`);
        
        if (newDebugMessages.length > 0) {
            console.log('\n[DEBUG MESSAGES - AFTER APPLY]:');
            newDebugMessages.forEach((msg, idx) => {
                console.log(`  ${idx + 1}. ${msg}`);
            });
        } else {
            console.log('\n⚠️  No new DEBUG messages appeared after clicking Apply');
        }
        
        console.log('\n' + '='.repeat(80));
        console.log('STEP 10: Reading scorecard values after apply');
        console.log('='.repeat(80));
        
        const afterScorecard = await page.evaluate(() => {
            const getValue = (id) => {
                const elem = document.getElementById(id);
                return elem ? elem.textContent.trim() : 'NOT FOUND';
            };
            
            return {
                totalCapacity: getValue('total-capacity'),
                totalPlannedHours: getValue('total-planned-hours'),
                totalActualHours: getValue('total-actual-hours'),
                utilizationPercent: getValue('utilization-percent'),
                totalIssues: getValue('total-issues'),
                completedIssues: getValue('completed-issues'),
                inProgressIssues: getValue('in-progress-issues'),
                todoIssues: getValue('todo-issues')
            };
        });
        
        console.log('\nSCORECARD VALUES AFTER APPLY:');
        Object.entries(afterScorecard).forEach(([key, value]) => {
            console.log(`  ${key}: ${value}`);
        });
        
        const afterRowCount = await page.locator('table tbody tr').count();
        console.log(`\nRow count after apply: ${afterRowCount}`);
        console.log(`Row count changed: ${afterRowCount !== initialRowCount ? 'YES' : 'NO'} (${initialRowCount} → ${afterRowCount})`);
        
        console.log('\n' + '='.repeat(80));
        console.log('STEP 11: Checking for errors');
        console.log('='.repeat(80));
        
        if (errorMessages.length > 0) {
            console.log(`\n❌ ${errorMessages.length} ERROR(S) FOUND:`);
            errorMessages.forEach((msg, idx) => {
                console.log(`  ${idx + 1}. ${msg}`);
            });
        } else {
            console.log('\n✓ No errors found in console');
        }
        
        console.log('\n' + '='.repeat(80));
        console.log('DETAILED REPORT SUMMARY');
        console.log('='.repeat(80));
        
        console.log('\n📊 ALL DEBUG MESSAGES (in order):');
        if (debugMessages.length > 0) {
            debugMessages.forEach((msg, idx) => {
                console.log(`  ${idx + 1}. ${msg}`);
            });
        } else {
            console.log('  ⚠️  No [DEBUG messages found');
        }
        
        console.log('\n📈 SCORECARD COMPARISON:');
        console.log('  Field                  | Initial           | After Apply       | Changed?');
        console.log('  ' + '-'.repeat(75));
        Object.keys(initialScorecard).forEach(key => {
            const initial = initialScorecard[key];
            const after = afterScorecard[key];
            const changed = initial !== after ? '✓ YES' : 'NO';
            console.log(`  ${key.padEnd(22)} | ${String(initial).padEnd(17)} | ${String(after).padEnd(17)} | ${changed}`);
        });
        
        console.log('\n📋 ROW COUNT:');
        console.log(`  Initial: ${initialRowCount}`);
        console.log(`  After:   ${afterRowCount}`);
        console.log(`  Changed: ${afterRowCount !== initialRowCount ? '✓ YES' : 'NO'}`);
        
        console.log('\n🐛 ERRORS:');
        if (errorMessages.length > 0) {
            errorMessages.forEach((msg, idx) => {
                console.log(`  ${idx + 1}. ${msg}`);
            });
        } else {
            console.log('  ✓ None');
        }
        
        console.log('\n' + '='.repeat(80));
        console.log('TEST COMPLETE - Browser will remain open for 30 seconds');
        console.log('='.repeat(80));
        
        await page.waitForTimeout(30000);
        
    } catch (error) {
        console.error('\n❌ TEST FAILED:', error.message);
        console.error(error.stack);
    } finally {
        await browser.close();
    }
}

testNestedViewReport().catch(console.error);
