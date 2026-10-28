const { test, expect } = require('@playwright/test');
const fs = require('node:fs');
const html = fs.readFileSync('employee_capacity_utilization_report.html', 'utf8');
const payload = {
  jira_browse_base: 'https://jira.example/browse', resources: { Alice: {}, Bob: {} }, teams: [], support: [],
  profiles: [{ from_date: '2026-01-01', to_date: '2026-12-31', standard_hours_per_day: 8, holiday_dates: ['2026-10-09'] }],
  issues: [{ issue_key: 'APP-1', issue_type: 'Sub-task', summary: 'Build calendar', assignee: 'Bob', start_date: '2026-10-01', due_date: '2026-10-20', original_estimate_hours: 16 }],
  worklogs: [
    { issue_id: 'APP-1', item_issue_type: 'Sub-task', worklog_author: 'Alice', item_assignee: 'Bob', worklog_date: '2026-10-05', hours_logged: 2 },
    { issue_id: 'APP-1', item_issue_type: 'Sub-task', worklog_author: 'Alice', item_assignee: 'Bob', worklog_date: '2026-10-05', hours_logged: 3.5 },
    { issue_id: 'STORY-1', item_issue_type: 'Story', worklog_author: 'Alice', worklog_date: '2026-10-05', hours_logged: 9 },
    { issue_id: 'RLT-1', project_key: 'RLT', item_issue_type: 'Sub-task', worklog_author: 'Alice', worklog_date: '2026-10-05', hours_logged: 2 },
    { issue_id: 'APP-1', item_issue_type: 'Sub-task', worklog_author: 'Bob', worklog_date: '2026-10-05', hours_logged: 8 },
    { issue_id: 'APP-1', item_issue_type: 'Sub-task', worklog_author: 'Alice', worklog_date: '2026-09-05', hours_logged: 8 },
  ],
  leaves: [ { assignee: 'Alice', period_day: '2026-10-05', planned_taken_hours: 2, unplanned_taken_hours: 1, jira_task_ids:'RLT-1, RLT-2', jira_task_links:'https://jira.example/browse/RLT-1, https://jira.example/browse/RLT-2' }, { assignee: 'Alice', period_day: '2026-10-07', planned_taken_hours: 8 } ],
};
test.beforeEach(async ({page}) => {
  await page.route('http://calendar.test/**', route => {
    const url = route.request().url();
    if (url.endsWith('/api/employee-capacity-utilization/data')) return route.fulfill({json:payload});
    if (url.endsWith('/report.html')) return route.fulfill({body:html, contentType:'text/html'});
    return route.fulfill({body:'', contentType:url.endsWith('.css')?'text/css':'text/javascript'});
  });
  await page.goto('http://calendar.test/report.html');
  await page.locator('#month').fill('2026-10');
  await expect(page.locator('#table tbody tr').first()).toContainText('Alice');
});
test('daily calendar keeps work, leave, grouping, and employee boundaries correct', async ({page}) => {
  await page.locator('#scope').selectOption('assigned');
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await expect(page.locator('#ecu-drawer-title')).toHaveText('Availability Details');
  await expect(page.locator('.availability-day[data-calendar-date]')).toHaveCount(31);
  await expect(page.locator('.availability-day')).toHaveCount(35);
  const mixed = page.locator('[data-calendar-date="2026-10-05"]');
  await expect(mixed).toContainText('5.5h logged');
  await expect(mixed).toContainText('3h leave');
  await expect(page.locator('[data-calendar-date="2026-10-07"]')).toContainText('0h logged');
  await expect(page.locator('[data-calendar-date="2026-10-07"]')).toContainText('8h leave');
  const sizes = await page.locator('.availability-day').evaluateAll(nodes => nodes.map(node => [node.offsetWidth,node.offsetHeight]));
  expect(new Set(sizes.map(x=>x[1])).size).toBe(1);
  expect(Math.max(...sizes.map(x=>x[0]))-Math.min(...sizes.map(x=>x[0]))).toBeLessThanOrEqual(1);
  const width = await page.locator('#ecu-drawer').evaluate(el=>el.getBoundingClientRect().width);
  expect(width).toBeCloseTo(page.viewportSize().width*.8,0);
  await mixed.locator('[data-work-date]').click();
  await expect(page.locator('#availability-day-details tbody tr')).toHaveCount(1);
  await expect(page.locator('#availability-day-details')).toContainText('Build calendar');
  await expect(page.locator('#availability-day-details')).toContainText('16h');
  await expect(page.locator('#availability-day-details')).toContainText('2026-10-20');
  await expect(page.locator('#availability-day-details a')).toHaveAttribute('href','https://jira.example/browse/APP-1');
  await page.keyboard.press('Escape');
  await expect(page.locator('#ecu-drawer')).toHaveAttribute('aria-hidden','true');
  await expect(page.locator('#table tbody tr').first().locator('td').nth(4)).toBeFocused();
});
test('menus paint above sticky report header and mobile calendar retains uniform cells', async ({page}) => {
  await page.setViewportSize({width:1280,height:900});
  for (const selector of ['.team-picker','.color-picker']) {
    await page.locator(selector+' summary').click();
    const above = await page.evaluate(selector=>{
      const menu=document.querySelector(selector+' > div'),head=document.querySelector('#table thead');
      menu.style.top=(head.getBoundingClientRect().top-menu.parentElement.getBoundingClientRect().top)+'px';
      const r=menu.getBoundingClientRect();
      return menu.contains(document.elementFromPoint(r.left+30,head.getBoundingClientRect().top+10));
    },selector);
    expect(above).toBe(true);
    await page.keyboard.press('Escape');
  }
  await page.setViewportSize({width:390,height:844});
  await page.evaluate(()=>window.openEcuDrawer('Alice','availability',document.querySelector('#month')));
  expect(await page.locator('#ecu-drawer').evaluate(el=>el.getBoundingClientRect().width)).toBeLessThanOrEqual(390);
  await expect(page.locator('[data-work-date]')).toHaveCount(1);
});
test('month boundaries support leap day and six-week months', async ({page}) => {
  for (const [month,count,slots] of [['2024-02',29,35],['2026-03',31,42]]) {
    await page.locator('#month').fill(month);
    await page.locator('#table tbody tr').first().locator('td').nth(4).click();
    await expect(page.locator('.availability-day[data-calendar-date]')).toHaveCount(count);
    await expect(page.locator('.availability-day')).toHaveCount(slots);
    await page.keyboard.press('Escape');
  }
});
test('today follows local calendar date, stays distinct after selection, and spacing is compact', async ({page}) => {
  // UTC is still October 27; Karachi has crossed midnight into October 28.
  await page.clock.install({time:new Date('2026-10-27T19:30:00Z')});
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  const today=page.locator('[data-calendar-date="2026-10-28"]');
  await expect(today).toHaveAttribute('aria-current','date');
  await expect(today).toContainText('Today');
  await expect(page.locator('.availability-day.today')).toHaveCount(1);
  await page.locator('[data-work-date]').first().click();
  await expect(today).toHaveClass(/today/);
  const dimensions=await page.locator('.availability-day').evaluateAll(nodes=>nodes.map(el=>el.getBoundingClientRect().height));
  expect(new Set(dimensions).size).toBe(1);
  expect(dimensions[0]).toBeLessThanOrEqual(103);
  await page.keyboard.press('Escape');
  await page.locator('#month').fill('2026-09');
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await expect(page.locator('[aria-current="date"]')).toHaveCount(0);
});
test('leave badges open daily source rows and every Jira link; work and leave switch independently', async ({page}) => {
  await page.locator('#scope').selectOption('assigned');
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  const leave=page.locator('[data-leave-date="2026-10-05"]');
  await leave.focus();
  await page.keyboard.press('Enter');
  const details=page.locator('#availability-day-details');
  await expect(details).toContainText('3h leave taken');
  await expect(details.locator('tbody tr')).toHaveCount(1);
  await expect(details.locator('a')).toHaveCount(2);
  await expect(details.locator('a').nth(1)).toHaveAttribute('href','https://jira.example/browse/RLT-2');
  await expect(details).toContainText('Planned taken');
  await expect(details).toContainText('2h');
  await expect(details).toContainText('1h');
  await expect(leave).toHaveAttribute('aria-pressed','true');
  await page.locator('[data-work-date="2026-10-05"]').click();
  await expect(details).toContainText('Build calendar');
  await expect(leave).toHaveAttribute('aria-pressed','false');
  await page.locator('[data-leave-date="2026-10-07"]').click();
  await expect(details).toContainText('8h leave taken');
  await expect(details).not.toContainText('Build calendar');
});
