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

test('live worklogs and leave update separately and failures preserve visible values',async({page})=>{
 await page.clock.setFixedTime(new Date('2026-10-05T10:00:00+05:00'));
 let workRequests=0;let releaseLeave;const leaveWait=new Promise(resolve=>{releaseLeave=resolve});
 await page.route('**/*',async route=>{const u=new URL(route.request().url());
  if(u.pathname==='/api/employee-capacity-utilization/data')return route.fulfill({json:payload});
  if(u.pathname==='/team_worklog_hours_report.html')return route.fulfill({contentType:'text/html',body:html});
  if(u.pathname==='/api/team-worklog-hours/live'){
   const employee=u.searchParams.get('employee'),month=u.searchParams.get('month'),kind=u.searchParams.get('kind');
   if(employee!=='Alice')return route.fulfill({status:502,json:{error:'Unavailable'}});
   if(kind==='worklogs'){workRequests++;if(workRequests>1)return route.fulfill({status:502,json:{error:'Unavailable'}});
    return route.fulfill({json:{employee,month,kind,checked_at:'2026-10-05T10:00:00Z',values:[{issue_key:'APP-1',issue_type:'Sub-task',assigned_to_author:true,date:month+'-05',seconds:25200}]}});}
   await leaveWait;return route.fulfill({json:{employee,month,kind,checked_at:'2026-10-05T10:00:00Z',values:[{assignee:'Alice',period_day:month+'-05',planned_taken_hours:4,unplanned_taken_hours:0}]}});
  }
  return route.fulfill({status:404,body:''});
 });
 await page.goto('http://example.test/team_worklog_hours_report.html');await expect(page.locator('#table tbody tr')).toHaveCount(2);await page.locator('#month').fill('2026-10');
 const alice=page.locator('#table tbody tr').filter({hasText:'Alice'}),day=alice.locator('td').nth(5);
 await expect(day).toContainText('7.00h');await expect(day).toContainText('Leave 2.00h');await expect(alice).toContainText('Leave: updating');
 releaseLeave();await expect(day).toContainText('Leave 4.00h');await expect(alice).toContainText('Work: live');await expect(alice).toContainText('Leave: live');
 await page.getByRole('button',{name:'Refresh live hours and leave'}).click();await expect(alice).toContainText('Work: update failed · last live');await expect(day).toContainText('7.00h');
});

test('filter changes immediately update the shareable URL and restore on reload',async({page})=>{
 const filtered={...payload,teams:[{team_name:'Product',assignees:['Alice']},{team_name:'Process Team',assignees:['Bob']}],profiles:[{from_date:'2026-01-01',to_date:'2026-12-31',standard_hours_per_day:8}]};
 await page.route('**/*',route=>{const path=new URL(route.request().url()).pathname;if(path==='/api/employee-capacity-utilization/data')return route.fulfill({json:filtered});if(path==='/team_worklog_hours_report.html')return route.fulfill({contentType:'text/html',body:html});return route.fulfill({status:404,body:''})});
 await page.goto('http://example.test/team_worklog_hours_report.html');await expect(page.locator('#table tbody tr')).toHaveCount(2);await expect(page.locator('.color-picker')).toBeHidden();
 await page.locator('#month').fill('2026-10');await page.locator('#profile').selectOption('0');await page.locator('#scope').selectOption('assigned');
 await page.locator('#include-leaves-logged').check();await page.locator('#resigned').check();await page.locator('#support').uncheck();
 await page.locator('.team-picker summary').click();await page.locator('#teams input[value="Process Team"]').check();await page.locator('#employee-search').fill('Alice');
 let params=new URLSearchParams(new URL(page.url()).hash.slice(1));expect(params.get('month')).toBe('2026-10');expect(params.get('profile')).toBe('0');expect(params.get('scope')).toBe('assigned');expect(params.get('leaves')).toBe('1');expect(params.get('resigned')).toBe('1');expect(params.get('support')).toBe('0');expect(JSON.parse(params.get('teams'))).toEqual(['Product','Process Team']);expect(params.get('employee_search')).toBe('Alice');
 const shared=page.url();await page.goto(shared);await expect(page.locator('#month')).toHaveValue('2026-10');await expect(page.locator('#profile')).toHaveValue('0');await expect(page.locator('#scope')).toHaveValue('assigned');await expect(page.locator('#include-leaves-logged')).toBeChecked();await expect(page.locator('#resigned')).toBeChecked();await expect(page.locator('#support')).not.toBeChecked();await expect(page.locator('#teams input[value="Process Team"]')).toBeChecked();await expect(page.locator('#employee-search')).toHaveValue('Alice');
 await page.locator('#profile').selectOption('');params=new URLSearchParams(new URL(page.url()).hash.slice(1));expect(params.has('profile')).toBe(false);
});

test('three rolling worklog requests finish before leaves and preserve other rows',async({page})=>{
 await page.clock.setFixedTime(new Date('2026-10-05T10:00:00+05:00'));
 const names=['Alice','Bob','Carol','Dave','Eve'],pending=new Map(),requested=[],leaveStarted=[];let inFlight=0,peak=0,settled=0;
 await page.route('**/*',route=>{const u=new URL(route.request().url()),path=u.pathname;
  if(path==='/api/employee-capacity-utilization/data')return route.fulfill({json:{...payload,resources:Object.fromEntries(names.map(name=>[name,{}]))}});
  if(path==='/team_worklog_hours_report.html')return route.fulfill({contentType:'text/html',body:html});
  if(path==='/api/team-worklog-hours/live'){const name=u.searchParams.get('employee'),kind=u.searchParams.get('kind');
   if(kind==='worklogs'){inFlight++;peak=Math.max(peak,inFlight);requested.push(name);pending.set(name,route);return;}
   leaveStarted.push(name);expect(settled).toBe(5);return route.fulfill({json:{employee:name,month:'2026-10',kind,values:[],checked_at:'2026-10-05T10:00:00Z'}});
  }return route.fulfill({status:404,body:''});
 });
 await page.goto('http://example.test/team_worklog_hours_report.html');await expect.poll(()=>requested.length).toBe(3);expect(leaveStarted).toHaveLength(0);
 await page.evaluate(()=>{window.originalBob=[...document.querySelectorAll('#table tbody tr')].find(row=>row.firstElementChild.firstChild.textContent==='Bob')});
 async function finish(name,fail=false){inFlight--;settled++;const route=pending.get(name);await route.fulfill(fail?{status:502,json:{error:'Unavailable'}}:{json:{employee:name,month:'2026-10',kind:'worklogs',values:[],checked_at:'2026-10-05T10:00:00Z'}});}
 await finish('Alice');await expect.poll(()=>requested.length).toBe(4);expect(leaveStarted).toHaveLength(0);
 expect(await page.evaluate(()=>window.originalBob.isConnected)).toBe(true);
 await finish('Dave',true);await expect.poll(()=>requested.length).toBe(5);expect(leaveStarted).toHaveLength(0);
 await finish('Bob');await finish('Carol');expect(leaveStarted).toHaveLength(0);await finish('Eve');
 await expect.poll(()=>leaveStarted.length).toBe(5);expect(peak).toBe(3);
 await expect(page.locator('.matrix-progress')).toContainText('Hours: 5/5 checked');await expect(page.locator('.matrix-progress')).toContainText('Leaves: 5/5 checked');await expect(page.locator('.matrix-progress')).toContainText('1 failed');await expect(page.locator('#status')).toContainText('5 employees');
});

test('busy live reads retry and recover without replacing snapshot values on the throttle',async({page})=>{
 let attempts=0;
 await page.route('**/*',route=>{const u=new URL(route.request().url());
  if(u.pathname==='/api/employee-capacity-utilization/data')return route.fulfill({json:{...payload,resources:{Alice:{}}}});
  if(u.pathname==='/team_worklog_hours_report.html')return route.fulfill({contentType:'text/html',body:html});
  if(u.pathname==='/api/team-worklog-hours/live'){const kind=u.searchParams.get('kind');
   if(kind==='worklogs'&&++attempts===1)return route.fulfill({status:429,headers:{'Retry-After':'0'},json:{error:'Busy'}});
   return route.fulfill({json:{employee:'Alice',month:u.searchParams.get('month'),kind,values:[],checked_at:'2026-10-05T10:00:00Z'}});
  }return route.fulfill({status:404,body:''});
 });
 await page.goto('http://example.test/team_worklog_hours_report.html');
 await expect(page.locator('.matrix-live-status')).toContainText('Work: live');await expect(page.locator('.matrix-live-status')).toContainText('Leave: live');expect(attempts).toBe(2);await expect(page.locator('.matrix-progress')).not.toContainText('failed');
});

for(const report of ['employee_capacity_utilization_report.html','team_worklog_hours_report.html']) test(`${report} shows roster-only employees and defaults to all teams`,async({page})=>{
 const body=fs.readFileSync(report,'utf8');
 const roster={...payload,issues:[],worklogs:[],leaves:[],resources:{'Usama Tariq':{resigned:false},'Former Member':{resigned:true}},teams:[{team_name:'DE Team',assignees:['Usama Tariq','Former Member']},{team_name:'Process Team',assignees:[]}]};
 await page.route('**/*',route=>{const path=new URL(route.request().url()).pathname;if(path==='/api/employee-capacity-utilization/data')return route.fulfill({json:roster});if(path==='/'+report)return route.fulfill({contentType:'text/html',body});return route.fulfill({status:404,body:''})});
 await page.goto('http://example.test/'+report);await page.locator('.team-picker summary').click();
 const process=page.locator('#teams input[value="Process Team"]');await expect(process).toBeChecked();
 await process.uncheck();await page.locator('[data-team-action="default"]').click();await expect(process).toBeChecked();
 await page.locator('#employee-search').fill('usama');await expect(page.locator('#table tbody tr').filter({hasText:'Usama Tariq'})).toHaveCount(1);
 await page.locator('#employee-search').fill('Former');await expect(page.locator('#table tbody tr').filter({hasText:'Former Member'})).toHaveCount(0);
 await page.locator('#resigned').check();await expect(page.locator('#table tbody tr').filter({hasText:'Former Member'})).toHaveCount(1);
});

for(const report of ['employee_capacity_utilization_report.html','team_worklog_hours_report.html']) test(`${report} excludes the unassigned placeholder`,async({page})=>{
 const body=fs.readFileSync(report,'utf8');
 const snapshot={...payload,resources:{Alice:{},'  UNASSIGNED  ':{}},issues:[{assignee:'Unassigned',issue_type:'Sub-task',original_estimate_hours:12,start_date:'2026-10-01',due_date:'2026-10-31'}],worklogs:[],leaves:[]};
 await page.route('**/*',route=>{const path=new URL(route.request().url()).pathname;if(path==='/api/employee-capacity-utilization/data')return route.fulfill({json:snapshot});if(path==='/'+report)return route.fulfill({contentType:'text/html',body});return route.fulfill({status:404,body:''})});
 await page.goto('http://example.test/'+report);await expect(page.locator('#status')).toContainText('1 employees');
 await expect(page.locator('#table tbody')).toContainText('Alice');await expect(page.locator('#table tbody')).not.toContainText('UNASSIGNED');
 if(report.startsWith('employee_')){const rows=page.locator('#table tbody tr');await expect(rows).toHaveCount(2);expect(await rows.nth(0).locator('td').nth(1).innerText()).toBe(await rows.nth(1).locator('td').nth(1).innerText())}
});

test('capacity columns sort numerically and keep grand total last',async({page})=>{
 const body=fs.readFileSync('employee_capacity_utilization_report.html','utf8');
 const snapshot={...payload,resources:{Alice:{},Bob:{},Carol:{}},issues:[],leaves:[],teams:[],worklogs:[{worklog_author:'Alice',worklog_date:'2026-10-05',hours_logged:2},{worklog_author:'Bob',worklog_date:'2026-10-05',hours_logged:12},{worklog_author:'Carol',worklog_date:'2026-10-05',hours_logged:5}]};
 await page.route('**/*',route=>{const path=new URL(route.request().url()).pathname;if(path==='/api/employee-capacity-utilization/data')return route.fulfill({json:snapshot});if(path==='/employee_capacity_utilization_report.html')return route.fulfill({contentType:'text/html',body});return route.fulfill({status:404,body:''})});
 await page.goto('http://example.test/employee_capacity_utilization_report.html');await page.locator('#month').fill('2026-10');
 const labels=()=>page.locator('#table tbody tr td:first-child').allTextContents();
 await page.getByRole('button',{name:'Logged hours (all work)'}).click();expect(await labels()).toEqual(['Bob','Carol','Alice','Grand Total']);
 await page.getByRole('button',{name:'Logged hours (all work)'}).click();expect(await labels()).toEqual(['Alice','Carol','Bob','Grand Total']);
 await expect(page.locator('th').filter({has:page.getByRole('button',{name:'Logged hours (all work)'})})).toHaveAttribute('aria-sort','ascending');
 await page.getByRole('button',{name:'Employee Name'}).click();expect(await labels()).toEqual(['Alice','Bob','Carol','Grand Total']);
});
