const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.createContext();
  const page = await context.newPage();

  // Listen for console messages
  const consoleLogs = [];
  page.on('console', msg => {
    consoleLogs.push({
      type: msg.type(),
      text: msg.text(),
      location: msg.location()
    });
  });

  // Listen for page errors
  const pageErrors = [];
  page.on('pageerror', err => {
    pageErrors.push({
      message: err.message,
      stack: err.stack
    });
  });

  // Listen for network errors
  const networkErrors = [];
  page.on('requestfailed', req => {
    networkErrors.push({
      url: req.url(),
      error: req.failure().errorText
    });
  });

  const ports = [3000, 3001, 4173, 5000, 8000, 8080];
  let connected = false;
  let url = '';

  for (const port of ports) {
    try {
      url = `http://127.0.0.1:${port}/team_capacity_planner.html`;
      const response = await page.goto(url, { waitUntil: 'networkidle', timeout: 5000 }).catch(() => null);
      if (response && response.ok()) {
        console.log(`✓ Connected to ${url}`);
        connected = true;
        break;
      }
    } catch (e) {
      // Try next port
    }
  }

  if (!connected) {
    console.log('✗ Could not connect to any port');
    await browser.close();
    process.exit(1);
  }

  // Wait a bit for page to fully load
  await page.waitForTimeout(2000);

  // Inspect the page
  const findings = await page.evaluate(() => {
    const result = {
      wiContainerExists: !!document.getElementById('wi-container'),
      epicRowsInContainer: [],
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

      // Get first 1-3 rows content
      for (let i = 0; i < Math.min(3, epicRows.length); i++) {
        const row = epicRows[i];
        const rowHead = row.querySelector('.epic-row-head');
        result.firstRowsInfo.push({
          rowIndex: i,
          innerHTML: row.innerHTML.substring(0, 200) + (row.innerHTML.length > 200 ? '...' : ''),
          textContent: (row.textContent || '').substring(0, 150),
          classList: Array.from(row.classList),
          headInnerHTML: rowHead ? rowHead.innerHTML.substring(0, 150) : null,
          headTextContent: rowHead ? (rowHead.textContent || '').substring(0, 100) : null,
          rowHeight: row.offsetHeight,
          rowWidth: row.offsetWidth,
          headHeight: rowHead ? rowHead.offsetHeight : null,
          headWidth: rowHead ? rowHead.offsetWidth : null,
          display: window.getComputedStyle(row).display,
          visibility: window.getComputedStyle(row).visibility,
          height: window.getComputedStyle(row).height,
          minHeight: window.getComputedStyle(row).minHeight,
          maxHeight: window.getComputedStyle(row).maxHeight,
          overflow: window.getComputedStyle(row).overflow,
          headDisplay: rowHead ? window.getComputedStyle(rowHead).display : null,
          headHeight_computed: rowHead ? window.getComputedStyle(rowHead).height : null
        });
      }

      // Get computed styles for first epic-row
      if (epicRows.length > 0) {
        const firstRow = epicRows[0];
        const styles = window.getComputedStyle(firstRow);
        result.epicRowStyles = {
          display: styles.display,
          visibility: styles.visibility,
          height: styles.height,
          minHeight: styles.minHeight,
          maxHeight: styles.maxHeight,
          width: styles.width,
          overflow: styles.overflow,
          opacity: styles.opacity,
          pointerEvents: styles.pointerEvents,
          position: styles.position,
          top: styles.top,
          left: styles.left,
          backgroundColor: styles.backgroundColor,
          padding: styles.padding,
          margin: styles.margin,
          borderWidth: styles.borderWidth,
          lineHeight: styles.lineHeight,
          fontSize: styles.fontSize
        };

        // Get bounding boxes
        result.boundingBoxes.epicRow = {
          top: firstRow.getBoundingClientRect().top,
          left: firstRow.getBoundingClientRect().left,
          width: firstRow.getBoundingClientRect().width,
          height: firstRow.getBoundingClientRect().height,
          bottom: firstRow.getBoundingClientRect().bottom,
          right: firstRow.getBoundingClientRect().right
        };

        const rowHead = firstRow.querySelector('.epic-row-head');
        if (rowHead) {
          const headStyles = window.getComputedStyle(rowHead);
          result.epicRowHeadStyles = {
            display: headStyles.display,
            height: headStyles.height,
            minHeight: headStyles.minHeight,
            maxHeight: headStyles.maxHeight,
            width: headStyles.width,
            visibility: headStyles.visibility,
            opacity: headStyles.opacity,
            overflow: headStyles.overflow,
            backgroundColor: headStyles.backgroundColor,
            padding: headStyles.padding,
            margin: headStyles.margin,
            fontSize: headStyles.fontSize,
            lineHeight: headStyles.lineHeight
          };

          result.boundingBoxes.epicRowHead = {
            top: rowHead.getBoundingClientRect().top,
            left: rowHead.getBoundingClientRect().left,
            width: rowHead.getBoundingClientRect().width,
            height: rowHead.getBoundingClientRect().height,
            bottom: rowHead.getBoundingClientRect().bottom,
            right: rowHead.getBoundingClientRect().right
          };
        }
      }
    }

    return result;
  });

  console.log('\n=== PAGE INSPECTION FINDINGS ===\n');
  console.log('1. WI-CONTAINER & EPIC-ROWS:');
  console.log(`   - #wi-container exists: ${findings.wiContainerExists}`);
  console.log(`   - .epic-row count: ${findings.epicRowsCount}`);
  console.log(`   - Contains epic-row nodes: ${findings.epicRowsCount > 0 ? 'YES' : 'NO'}`);

  if (findings.firstRowsInfo.length > 0) {
    console.log('\n2. FIRST 1-3 ROWS CONTENT:');
    findings.firstRowsInfo.forEach((row, idx) => {
      console.log(`\n   Row ${idx}:`);
      console.log(`   - textContent: "${row.textContent.substring(0, 80)}${row.textContent.length > 80 ? '...' : ''}"`);
      console.log(`   - headTextContent: "${row.headTextContent ? row.headTextContent.substring(0, 80) : 'N/A'}"`);
      console.log(`   - Dimensions: ${row.rowWidth}x${row.rowHeight}px (head: ${row.headWidth || 'N/A'}x${row.headHeight || 'N/A'}px)`);
    });
  }

  if (Object.keys(findings.epicRowStyles).length > 0) {
    console.log('\n3. COMPUTED STYLES FOR .epic-row:');
    Object.entries(findings.epicRowStyles).forEach(([key, val]) => {
      if (val !== 'auto' && val !== '' && val !== 'none') {
        console.log(`   - ${key}: ${val}`);
      }
    });
  }

  if (Object.keys(findings.epicRowHeadStyles).length > 0) {
    console.log('\n4. COMPUTED STYLES FOR .epic-row-head:');
    Object.entries(findings.epicRowHeadStyles).forEach(([key, val]) => {
      if (val !== 'auto' && val !== '' && val !== 'none') {
        console.log(`   - ${key}: ${val}`);
      }
    });
  }

  if (findings.boundingBoxes.epicRow) {
    console.log('\n5. BOUNDING BOXES:');
    console.log(`   - .epic-row: width=${findings.boundingBoxes.epicRow.width.toFixed(2)}px, height=${findings.boundingBoxes.epicRow.height.toFixed(2)}px`);
    if (findings.boundingBoxes.epicRowHead) {
      console.log(`   - .epic-row-head: width=${findings.boundingBoxes.epicRowHead.width.toFixed(2)}px, height=${findings.boundingBoxes.epicRowHead.height.toFixed(2)}px`);
    }
  }

  console.log('\n6. CONSOLE ERRORS:');
  if (consoleLogs.length > 0) {
    consoleLogs.slice(0, 10).forEach(log => {
      if (log.type === 'error' || log.type === 'warning') {
        console.log(`   [${log.type}] ${log.text}`);
      }
    });
    if (consoleLogs.filter(l => l.type === 'error' || l.type === 'warning').length === 0) {
      console.log('   (none)');
    }
  } else {
    console.log('   (none)');
  }

  console.log('\n7. PAGE ERRORS:');
  if (pageErrors.length > 0) {
    pageErrors.slice(0, 10).forEach(err => {
      console.log(`   ${err.message}`);
    });
  } else {
    console.log('   (none)');
  }

  console.log('\n8. NETWORK ERRORS:');
  if (networkErrors.length > 0) {
    networkErrors.slice(0, 10).forEach(err => {
      console.log(`   ${err.url}: ${err.error}`);
    });
  } else {
    console.log('   (none)');
  }

  // Check for CSS rules that might cause thin stripes
  const cssAnalysis = await page.evaluate(() => {
    const analysis = {
      heightLimitingRules: [],
      hiddenElements: [],
      scripts: []
    };

    // Check all stylesheets
    for (const sheet of document.styleSheets) {
      try {
        for (const rule of sheet.cssRules || []) {
          if (rule.style && rule.selectorText) {
            if ((rule.selectorText.includes('epic-row') || rule.selectorText.includes('wi-container')) &&
                (rule.style.height || rule.style.minHeight || rule.style.maxHeight || rule.style.overflow)) {
              analysis.heightLimitingRules.push({
                selector: rule.selectorText,
                height: rule.style.height,
                minHeight: rule.style.minHeight,
                maxHeight: rule.style.maxHeight,
                overflow: rule.style.overflow
              });
            }
          }
        }
      } catch (e) {
        // CORS or other restrictions on stylesheet
      }
    }

    // Check for hidden/visibility issues
    const epicRows = document.querySelectorAll('.epic-row');
    epicRows.forEach((row, idx) => {
      const style = window.getComputedStyle(row);
      if (style.display === 'none' || style.visibility === 'hidden' || style.height === '0px') {
        analysis.hiddenElements.push({
          index: idx,
          reason: `display: ${style.display}, visibility: ${style.visibility}, height: ${style.height}`
        });
      }
    });

    // List scripts
    document.querySelectorAll('script').forEach((script, idx) => {
      analysis.scripts.push({
        index: idx,
        src: script.src,
        type: script.type,
        inline: !!script.textContent
      });
    });

    return analysis;
  });

  console.log('\n9. CSS/SCRIPT ANALYSIS:');
  if (cssAnalysis.heightLimitingRules.length > 0) {
    console.log('\n   Height-limiting CSS rules for .epic-row/.wi-container:');
    cssAnalysis.heightLimitingRules.forEach(rule => {
      console.log(`   - ${rule.selector}`);
      console.log(`     height: ${rule.height}, minHeight: ${rule.minHeight}, maxHeight: ${rule.maxHeight}, overflow: ${rule.overflow}`);
    });
  }

  if (cssAnalysis.hiddenElements.length > 0) {
    console.log('\n   Hidden .epic-row elements:');
    cssAnalysis.hiddenElements.forEach(elem => {
      console.log(`   - Row ${elem.index}: ${elem.reason}`);
    });
  }

  await browser.close();
  process.exit(0);
})();
