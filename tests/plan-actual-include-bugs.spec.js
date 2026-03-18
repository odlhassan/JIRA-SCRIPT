// Playwright test for Plan vs Actual drawer behavior
const { test, expect } = require('@playwright/test');

// NOTE: Requires the local report server to be running and serving employee_performance_report.html
// Example: `python report_server.py` from the project root.

test('Plan vs Actual drawer opens with actual subtasks table', async ({ page }) => {
  await page.goto('http://localhost:8000/report_html/employee_performance_report.html', {
    waitUntil: 'networkidle',
  });

  await page.getByRole('heading', { name: 'Executive Scorecards' }).waitFor();

  const toggleButton = page.locator('#header-plan-actual-toggle');
  await expect(toggleButton).toBeVisible();
  await toggleButton.click();

  const drawer = page.locator('#score-detail-drawer');
  await expect(drawer).toBeVisible();

  await expect(page.locator('#plan-actual-drawer-include-bugs')).toHaveCount(0);
  await expect(page.locator('#score-detail-drawer-body')).toContainText('Actual Hours Subtasks');
  const renderedRows = await page.locator('#score-detail-drawer-body table tbody tr').count();
  expect(renderedRows).toBeGreaterThan(0);
});

