const playwright = require('playwright');

(async () => {
  const ports = [3000, 3001, 4173, 5000, 8000, 8080];
  let browser, context, page;
  let success = false;
  let targetUrl = null;

  try {
    // Try Playwright
    browser = await playwright.chromium.launch({ headless: true });
    context = await browser.newContext();
    page = await context.newPage();

    // Try to connect to the page
    for (const port of ports) {
      const url = `http://127.0.0.1:${port}/team_capacity_planner.html`;
      try {
        console.log(`\n=== Attempting ${url} ===`);
        const response = await page.goto(url, { waitUntil: 'networkidle', timeout: 10000 });
        if (response && response.ok()) {
          targetUrl = url;
          success = true;
          console.log(`✓ Connected to ${url}`);
          break;
        }
      } catch (e) {
        console.log(`✗ Port ${port} failed: ${e.message}`);
        continue;
      }
    }

    if (!success) {
      console.error('\n❌ Could not connect to any port');
      process.exit(1);
    }

    // Wait for page to settle
    await page.waitForTimeout(2000);

    // Collect findings
    const findings = {};

    // 1. Count of #wi-container .epic-row
    const epicRowCount = await page.evaluate(() => {
      const container = document.querySelector('#wi-container');
      if (!container) return 0;
      return container.querySelectorAll('.epic-row').length;
    });
    findings.epicRowCount = epicRowCount;
    console.log(`\n✓ Epic row count: ${epicRowCount}`);

    // 2. Text/HTML of first 1-3 rows
    const firstRows = await page.evaluate(() => {
      const container = document.querySelector('#wi-container');
      if (!container) return [];
      const rows = container.querySelectorAll('.epic-row');
      const result = [];
      for (let i = 0; i < Math.min(3, rows.length); i++) {
        result.push({
          index: i,
          textContent: rows[i].textContent.trim(),
          html: rows[i].outerHTML.substring(0, 500) + (rows[i].outerHTML.length > 500 ? '...' : '')
        });
      }
      return result;
    });
    findings.firstRows = firstRows;
    console.log(`✓ First 1-3 rows collected (${firstRows.length} rows)`);

    // 3. getBoundingClientRect for first .epic-row and .epic-row-head
    const boundingRects = await page.evaluate(() => {
      const firstEpicRow = document.querySelector('.epic-row');
      const firstRowHead = document.querySelector('.epic-row-head');
      const result = {};
      
      if (firstEpicRow) {
        const rect = firstEpicRow.getBoundingClientRect();
        result.epicRow = {
          top: rect.top,
          left: rect.left,
          bottom: rect.bottom,
          right: rect.right,
          width: rect.width,
          height: rect.height
        };
      }
      
      if (firstRowHead) {
        const rect = firstRowHead.getBoundingClientRect();
        result.epicRowHead = {
          top: rect.top,
          left: rect.left,
          bottom: rect.bottom,
          right: rect.right,
          width: rect.width,
          height: rect.height
        };
      }
      
      return result;
    });
    findings.boundingRects = boundingRects;
    console.log(`✓ Bounding rectangles collected`);

    // 4. Computed styles for first .epic-row and .epic-row-head
    const computedStyles = await page.evaluate(() => {
      const firstEpicRow = document.querySelector('.epic-row');
      const firstRowHead = document.querySelector('.epic-row-head');
      const result = {};
      
      if (firstEpicRow) {
        const style = window.getComputedStyle(firstEpicRow);
        result.epicRow = {
          display: style.display,
          position: style.position,
          width: style.width,
          height: style.height,
          backgroundColor: style.backgroundColor,
          color: style.color,
          padding: style.padding,
          margin: style.margin,
          border: style.border
        };
      }
      
      if (firstRowHead) {
        const style = window.getComputedStyle(firstRowHead);
        result.epicRowHead = {
          display: style.display,
          position: style.position,
          width: style.width,
          height: style.height,
          backgroundColor: style.backgroundColor,
          color: style.color,
          padding: style.padding,
          margin: style.margin,
          border: style.border
        };
      }
      
      return result;
    });
    findings.computedStyles = computedStyles;
    console.log(`✓ Computed styles collected`);

    // 5. Console and page errors
    const errors = { console: [], page: [] };
    page.on('console', msg => {
      if (msg.type() === 'error' || msg.type() === 'warning') {
        errors.console.push({
          type: msg.type(),
          text: msg.text(),
          location: msg.location()
        });
      }
    });
    
    page.on('pageerror', err => {
      errors.page.push({
        name: err.name,
        message: err.message,
        stack: err.stack ? err.stack.substring(0, 200) : ''
      });
    });

    // Give a moment for any errors to fire
    await page.waitForTimeout(1000);
    findings.errors = errors;
    console.log(`✓ Error monitoring completed (${errors.console.length} console messages, ${errors.page.length} page errors)`);

    // Output all findings as JSON
    console.log('\n=== LIVE FINDINGS ===');
    console.log(JSON.stringify(findings, null, 2));

    await browser.close();
    process.exit(0);
  } catch (err) {
    console.error('\n❌ Browser operation failed:', err.message);
    if (browser) await browser.close();
    process.exit(1);
  }
})();
