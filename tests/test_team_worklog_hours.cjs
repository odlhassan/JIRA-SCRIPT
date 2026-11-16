const {test,expect}=require('@playwright/test');
const fs=require('fs');
const html=fs.readFileSync('team_worklog_hours_report.html','utf8');
const payload={resources:{Alice:{},Bob:{}},issues:[],teams:[],profiles:[],support:[],worklogs:[
 {issue_id:'APP-1',item_issue_type:'Sub-task',item_assignee:'Alice',worklog_author:'Alice',worklog_date:'2026-10-05',hours_logged:3},
 {issue_id:'APP-1',item_issue_type:'Sub-task',item_assignee:'Alice',worklog_author:'Alice',worklog_date:'2026-10-03',hours_logged:2},
 {issue_id:'STORY-1',item_issue_type:'Story',worklog_author:'Alice',worklog_date:'2026-10-05',hours_logged:2},
 {issue_id:'RLT-1',project_key:'RLT',item_issue_type:'Sub-task',item_assignee:'Alice',worklog_author:'Alice',worklog_date:'2026-10-05',hours_logged:1}
],leaves:[{assignee:'Alice',period_day:'2026-10-05',planned_taken_hours:1,unplanned_taken_hours:1}]};
test('matrix respects scope and leave inclusion and exports visible values',async({page})=>{
 await page.route('**/*',route=>route.request().url().endsWith('/api/employee-capacity-utilization/data')?route.fulfill({json:payload}):route.request().url().endsWith('/team_worklog_hours_report.html')?route.fulfill({contentType:'text/html',body:html}):route.fulfill({status:404,body:''}));
 await page.goto('http://example.test/team_worklog_hours_report.html');
 await expect(page.locator('#table tbody tr')).toHaveCount(2);
 await page.locator('#month').fill('2026-10');
 const alice=page.locator('#table tbody tr').filter({hasText:'Alice'});
 const fifth=alice.locator('td').nth(5);
 await expect(fifth).toContainText('5.00h');await expect(fifth).toContainText('Leave 2.00h');
 await expect(alice.locator('td').nth(6)).toContainText('0.00h');
 const saturday=alice.locator('td').nth(3), sunday=alice.locator('td').nth(4);
 await expect(saturday).toHaveClass(/weekend/);await expect(saturday).toContainText('2.00h');
 await expect(sunday).toHaveClass(/weekend/);await expect(sunday).toContainText('0.00h');
 await expect(page.locator('.matrix th').nth(3)).toHaveClass(/weekend/);
 await expect(saturday).toHaveCSS('background-color','rgb(232, 235, 239)');
 await expect(sunday).toHaveCSS('background-color','rgb(232, 235, 239)');
 await expect(fifth).not.toHaveClass(/weekend/);
 await page.locator('#scope').selectOption('assigned');await expect(fifth).toContainText('3.00h');
 await page.locator('#include-leaves-logged').check();await expect(fifth).toContainText('4.00h');
 const matrix=await page.evaluate(()=>matrixExportRows);expect(matrix[1][9]).toBe(4);expect(matrix[1][10]).toBe(2);
});

test('capacity report links directly to the team matrix',async({page})=>{
 const capacity=fs.readFileSync('employee_capacity_utilization_report.html','utf8');
 await page.route('**/*',route=>route.request().url().endsWith('/api/employee-capacity-utilization/data')?route.fulfill({json:payload}):route.request().url().endsWith('/employee_capacity_utilization_report.html')?route.fulfill({contentType:'text/html',body:capacity}):route.fulfill({status:404,body:''}));
 await page.goto('http://example.test/employee_capacity_utilization_report.html');
 const link=page.getByRole('link',{name:'Team Worklog Hours'});await expect(link).toHaveAttribute('href','/team_worklog_hours_report.html');
});

test('today highlights the current local date, including on a weekend',async({page})=>{
 await page.clock.setFixedTime(new Date('2026-10-03T10:00:00+05:00'));
 await page.route('**/*',route=>route.request().url().endsWith('/api/employee-capacity-utilization/data')?route.fulfill({json:payload}):route.request().url().endsWith('/team_worklog_hours_report.html')?route.fulfill({contentType:'text/html',body:html}):route.fulfill({status:404,body:''}));
 await page.goto('http://example.test/team_worklog_hours_report.html');
 await expect(page.locator('#table tbody tr')).toHaveCount(2);
 const header=page.locator('.matrix th').nth(3),cell=page.locator('#table tbody tr').first().locator('td').nth(3);
 await expect(header).toHaveClass(/today/);await expect(header).toHaveClass(/weekend/);await expect(header).toContainText('Today');
 await expect(cell).toHaveClass(/today/);await expect(cell).toHaveCSS('background-color','rgb(232, 235, 239)');
 await expect(page.locator('.matrix th.today')).toHaveCount(1);
 await page.locator('#month').fill('2026-09');await expect(page.locator('.matrix .today')).toHaveCount(0);
});

test('matrix cells open and restore the full availability drawer',async({page})=>{
 let tallyRequests=0;
 await page.route('**/*',route=>{const url=new URL(route.request().url());
  if(url.pathname==='/api/employee-capacity-utilization/data')return route.fulfill({json:payload});
  if(url.pathname==='/api/employee-capacity-utilization/tally'){tallyRequests++;return route.fulfill({json:{employee:url.searchParams.get('employee'),month:url.searchParams.get('month'),entries:[],leaves:[],sync:{mode:'fresh'}}});}
  if(url.pathname==='/api/employee-capacity-utilization/worklog/identity')return route.fulfill({json:{author:'Alice',signed_in:true,write_enabled:true}});
  if(url.pathname==='/api/jira-auth/status')return route.fulfill({json:{configured:true,signed_in:true}});
  if(url.pathname==='/team_worklog_hours_report.html')return route.fulfill({contentType:'text/html',body:html});
  return route.fulfill({status:404,body:''});
 });
 await page.goto('http://example.test/team_worklog_hours_report.html');
 await expect(page.locator('#table tbody tr')).toHaveCount(2);await page.locator('#month').fill('2026-10');
 const row=page.locator('#table tbody tr').filter({hasText:'Alice'}),cell=row.locator('td').nth(5);
 await cell.hover();await expect(cell).toHaveAttribute('title',/Click to log work for Alice on 2026-10-05/);await cell.click();
 const dialog=page.locator('#ecu-drawer');await expect(dialog).toHaveClass(/open/);await expect(dialog).toContainText('Alice · 2026-10');
 await expect(dialog.getByRole('button',{name:'New subtask'}).first()).toBeVisible();await expect(dialog.getByRole('button',{name:'Manage my logged hours'})).toBeVisible();
 await expect(dialog.getByRole('button',{name:'Download Jira tally PDF'})).toBeVisible();await expect(dialog.getByRole('button',{name:'Refresh Jira worklogs'})).toBeVisible();
 await expect(dialog.getByRole('button',{name:'Copy link to details'})).toBeVisible();await expect(dialog.locator('#calendar-worklog-form')).toBeVisible();await expect(dialog.locator('.worklog-date-chip')).toContainText('2026-10-05');
 const shared=page.url();expect(shared).toContain('modal=availability');expect(shared).toContain('employee=Alice');expect(shared).toContain('day=2026-10-05');expect(shared).toContain('panel=log');
 await page.getByRole('button',{name:'Close details'}).click();await row.locator('td').first().focus();await page.keyboard.press('Enter');await expect(dialog).toHaveClass(/open/);await expect(dialog.locator('#calendar-worklog-form')).toHaveCount(0);
 await dialog.getByRole('button',{name:'Previous month'}).click();await expect(dialog).toContainText('Alice · 2026-09');await expect(page.locator('#month')).toHaveValue('2026-09');
 await page.goto(shared);await expect(dialog).toHaveClass(/open/);await expect(dialog).toContainText('Alice · 2026-10');await expect(dialog.locator('.worklog-date-chip')).toContainText('2026-10-05');
 await page.getByRole('button',{name:'Close details'}).click();const other=page.locator('#table tbody tr').filter({hasText:'Alice'}).locator('td').nth(6);await other.focus();await page.keyboard.press('Space');await expect(dialog.locator('.worklog-date-chip')).toContainText('2026-10-06');
});
