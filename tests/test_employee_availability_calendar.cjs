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

async function mockSubtaskJira(page){
  const state={reviews:0,writes:0,body:null};
  await page.route('**/worklog/identity',route=>route.fulfill({json:{author:'Alice',signed_in:true,write_enabled:true}}));
  await page.route('**/api/jira-auth/status',route=>route.fulfill({json:{configured:true,signed_in:true}}));
  await page.route('**/worklog/subtask/options?*',route=>{
    const params=new URL(route.request().url()).searchParams,kind=params.get('kind');
    let options=kind==='projects'?[{key:'NEW',summary:'Customer delight'}]:kind==='epics'?[{key:'NEW-1',summary:'Better customer experience'}]:[{key:'NEW-2',summary:'Find my customers'}];
    if(params.get('q')==='missing')options=[];
    if(params.get('cursor'))options=[{key:'MORE',summary:'Another space'}];
    return route.fulfill({json:{options,next_cursor:kind==='projects'&&!params.get('cursor')&&params.get('q')!=='missing'?'50':null,author:'Alice',today:'2026-11-11',first:'2026-11-01',last:'2026-11-30'}});
  });
  await page.route('**/worklog/subtask/review',route=>{
    state.reviews++;state.body=route.request().postDataJSON();
    return route.fulfill({json:{token:'subtask-token',review:{...state.body,author:'Alice',epic_summary:'Better customer experience',story_summary:'Find my customers'}}});
  });
  await page.route('**/worklog/subtask/submit',route=>{
    state.writes++;expect(route.request().postDataJSON()).toEqual({token:'subtask-token',confirmed:true});
    return route.fulfill({json:{issue_key:'NEW-3',summary:state.body.summary,assignee:'Alice',issue_type:'Sub-task',parent_issue_key:'NEW-2',start_date:state.body.start_date,due_date:state.body.due_date,original_estimate_hours:8,verified:true}});
  });
  return state;
}
async function choosePicker(dialog,name,key){
  await dialog.getByRole('combobox',{name,exact:true}).click();
  await dialog.getByRole('option',{name:new RegExp('^'+key+'(?: |$)')}).click();
}
async function chooseSubtaskHome(page){
  const dialog=page.getByRole('dialog',{name:'Create a new subtask'});
  await choosePicker(dialog,'Jira space','NEW');
  await choosePicker(dialog,'Epic','NEW-1');
  await choosePicker(dialog,'User story','NEW-2');
  await dialog.getByLabel('Subtask title').fill('Build a friendly search filter');
  return dialog;
}

test('new subtask guides live hierarchy, defaults estimate, reviews once and can log immediately',async({page})=>{
  const state=await mockSubtaskJira(page);
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'New subtask',exact:true}).click();
  const dialog=await chooseSubtaskHome(page);
  await expect(dialog.getByLabel('Original estimate',{exact:true})).toHaveValue('1d');
  await expect(dialog.getByLabel('Due date',{exact:true})).toHaveValue('2026-11-11');
  await expect(dialog.getByLabel('Due date',{exact:true})).toHaveAttribute('max','2026-11-30');
  for(const value of ['30m','4h','1w','1d']){
    await dialog.getByRole('button',{name:value,exact:true}).click();
    await expect(dialog.getByLabel('Original estimate',{exact:true})).toHaveValue(value);
  }
  await dialog.getByRole('button',{name:'Review subtask'}).click();
  await expect(dialog.locator('.subtask-check')).toContainText('NEW-2');
  expect(state.writes).toBe(0);
  await dialog.getByRole('button',{name:'Edit details'}).click();
  await expect(dialog.getByLabel('Subtask title')).toHaveValue('Build a friendly search filter');
  await dialog.getByRole('button',{name:'Review subtask'}).click();
  await dialog.getByRole('button',{name:'Create subtask in Jira'}).dblclick();
  await expect(dialog.locator('.subtask-success')).toContainText('NEW-3');
  expect(state.writes).toBe(1);expect(state.body.estimate).toBe('1d');
  await dialog.getByRole('button',{name:'Log work on this subtask'}).click();
  await expect(page.locator('#worklog-issue')).toHaveValue('NEW-3');
  await expect(page.locator('#worklog-selection')).toContainText('Build a friendly search filter');
  await expect(page.locator('#month')).toHaveValue('2026-11');
  await expect(page.locator('.availability-stat.work strong')).toHaveText('0h');
});

test('subtask dates reject other months and reversed dates without any request; mobile fits',async({page})=>{
  const state=await mockSubtaskJira(page);
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'New subtask',exact:true}).click();
  const dialog=await chooseSubtaskHome(page),due=dialog.getByLabel('Due date',{exact:true});
  for(const value of ['2026-10-31','2026-12-01','2026-11-10']){
    await due.fill(value);await dialog.getByRole('button',{name:'Review subtask'}).click();
    expect(state.reviews).toBe(0);expect(await due.evaluate(el=>el.checkValidity())).toBe(false);
  }
  await due.fill('2026-11-30');
  await dialog.getByRole('button',{name:'1w',exact:true}).click();
  await dialog.screenshot({path:'tmp/subtask-creator-desktop.png'});
  await page.setViewportSize({width:390,height:844});
  expect(await dialog.evaluate(el=>el.scrollWidth-el.clientWidth)).toBeLessThanOrEqual(1);
  await dialog.screenshot({path:'tmp/subtask-creator-mobile.png'});
  await page.keyboard.press('Escape');
  await expect(dialog).toHaveCount(0);
  await expect(page.locator('#ecu-drawer')).toHaveClass(/open/);
});

test('subtask creator opens over log form, preserves draft and adds created task to its chooser',async({page})=>{
  await mockSubtaskJira(page);
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'Log work on 2026-10-05'}).click();
  await page.locator('#worklog-comment').fill('Keep this work description');
  await page.locator('#availability-day-details').getByRole('button',{name:'New subtask',exact:true}).click();
  const dialog=await chooseSubtaskHome(page);
  await dialog.getByRole('button',{name:'Review subtask'}).click();
  await dialog.getByRole('button',{name:'Create subtask in Jira'}).click();
  await expect(dialog.locator('.subtask-success')).toContainText('NEW-3');
  await dialog.getByRole('button',{name:'Close subtask creator'}).click();
  await expect(page.locator('#worklog-comment')).toHaveValue('Keep this work description');
  await expect(page.locator('#worklog-issue')).toHaveValue('NEW-3');
});

test('live parent dropdown search preserves selections and changing a parent clears descendants',async({page})=>{
  await mockSubtaskJira(page);
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'New subtask',exact:true}).click();
  const dialog=await chooseSubtaskHome(page);
  await dialog.getByRole('combobox',{name:'Jira space',exact:true}).click();
  await dialog.locator('[data-more="projects"]').click();
  await expect(dialog.locator('[data-options="projects"] [role="option"]')).toHaveCount(2);
  await dialog.getByRole('searchbox',{name:'Search jira space',exact:true}).fill('missing');
  await expect(dialog.locator('[data-menu-status="projects"]')).toContainText('No matches');
  await expect(dialog.locator('[name="projects"]')).toHaveValue('NEW');
  await expect(dialog.locator('[name="epics"]')).toHaveValue('NEW-1');
  await expect(dialog.locator('[name="stories"]')).toHaveValue('NEW-2');
  await dialog.getByRole('searchbox',{name:'Search jira space',exact:true}).fill('');
  await dialog.locator('[data-more="projects"]').click();
  await dialog.getByRole('option',{name:/^MORE /}).click();
  await expect(dialog.getByRole('combobox',{name:'User story',exact:true})).toBeDisabled();
  await expect(dialog.locator('[name="epics"]')).toHaveValue('');
  await expect(dialog.locator('[name="stories"]')).toHaveValue('');
});

test('uncertain Jira creation is never silently retried and sign-in failures are visible',async({page})=>{
  const state=await mockSubtaskJira(page);
  await page.route('**/worklog/subtask/submit',route=>{state.writes++;return route.fulfill({status:502,json:{error:'Creation outcome needs verification in Jira.'}})});
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'New subtask',exact:true}).click();
  const dialog=await chooseSubtaskHome(page);
  await dialog.getByRole('button',{name:'Review subtask'}).click();
  await dialog.getByRole('button',{name:'Create subtask in Jira'}).click();
  await expect(dialog.locator('[data-status]')).toContainText('Check the selected user story in Jira');
  await expect(dialog.getByRole('button',{name:'Create subtask in Jira'})).toBeDisabled();
  expect(state.writes).toBe(1);
  await dialog.getByRole('button',{name:'Close subtask creator'}).click();
  await page.route('**/worklog/subtask/options?*',route=>route.fulfill({status:400,json:{error:'Sign in with Jira to continue.'}}));
  await page.getByRole('button',{name:'New subtask',exact:true}).click();
  await expect(page.locator('.subtask-dialog [data-status]')).toContainText('Sign in with Jira');
  await expect(page.locator('.subtask-dialog form')).toBeHidden();
});

test('live Jira search finds newly created subtasks after reload before a full fetch',async({page})=>{
  await mockSubtaskJira(page);
  await page.route('**/worklog/subtask/search?*',route=>route.fulfill({json:{issues:[{issue_key:'NEW-3',summary:'Fresh Jira subtask',issue_type:'Sub-task',assignee:'Alice'}],next_cursor:null}}));
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'Log work on 2026-10-05'}).click();
  await expect(page.locator('#worklog-live-search')).toContainText('up to date');
  await page.locator('#worklog-issue').selectOption('NEW-3');
  await expect(page.locator('#worklog-selection')).toContainText('Fresh Jira subtask');
  await page.reload();
  await expect(page.locator('#worklog-live-search')).toContainText('up to date');
  await page.locator('#worklog-issue').selectOption('NEW-3');
  await expect(page.locator('#worklog-selection')).toContainText('Fresh Jira subtask');
});

test('subtask creator explains missing local OAuth setup instead of offering a broken login',async({page})=>{
  await mockSubtaskJira(page);
  await page.route('**/api/jira-auth/status',route=>route.fulfill({json:{configured:false}}));
  await page.route('**/worklog/subtask/options?*',route=>route.fulfill({status:400,json:{error:'Sign in with Jira to continue.'}}));
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'New subtask',exact:true}).click();
  await expect(page.locator('.subtask-dialog [data-status]')).toContainText('needs administrator setup');
  await expect(page.locator('.subtask-dialog a[href="/api/jira-auth/login"]')).toHaveCount(0);
});

test('picker searches live inside its dropdown and Escape returns to the chosen parent',async({page})=>{
  await mockSubtaskJira(page);
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'New subtask',exact:true}).click();
  const dialog=await chooseSubtaskHome(page),space=dialog.getByRole('combobox',{name:'Jira space',exact:true}),search=dialog.getByRole('searchbox',{name:'Search jira space',exact:true});
  await expect(search).toBeHidden();
  await space.click();await expect(search).toBeFocused();
  await expect(dialog.locator('#subtask-projects-menu')).toContainText('Customer delight');
  await search.fill('NEW');
  await expect(dialog.locator('[name="epics"]')).toHaveValue('NEW-1');
  await expect(dialog.locator('[name="stories"]')).toHaveValue('NEW-2');
  await expect(dialog.locator('[data-options="projects"] [role="option"]')).toHaveCount(1);
  await search.press('ArrowDown');
  await expect(dialog.getByRole('option',{name:/^NEW /})).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(search).toBeHidden();await expect(space).toBeFocused();
  await expect(dialog).toBeVisible();
  await dialog.getByRole('combobox',{name:'Epic',exact:true}).click();
  await expect(dialog.getByRole('searchbox',{name:'Search epic',exact:true})).toHaveValue('');
  await dialog.screenshot({path:'tmp/subtask-search-dropdown-desktop.png'});
  await page.setViewportSize({width:390,height:844});
  const menu=dialog.locator('#subtask-epics-menu');
  await expect.poll(()=>menu.evaluate(el=>el.getBoundingClientRect().right)).toBeLessThanOrEqual(390);
  await expect.poll(()=>menu.evaluate(el=>el.getBoundingClientRect().left)).toBeGreaterThanOrEqual(0);
});

test('a space without epics explains the missing hierarchy and never requests user stories',async({page})=>{
  await mockSubtaskJira(page);let storyReads=0;
  await page.route('**/worklog/subtask/options?*',route=>{
    const kind=new URL(route.request().url()).searchParams.get('kind');if(kind==='stories')storyReads++;
    return route.fulfill({json:{options:kind==='projects'?[{key:'EMPTY',summary:'Space without epics'}]:[],next_cursor:null,author:'Alice',today:'2026-11-11',first:'2026-11-01',last:'2026-11-30'}});
  });
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'New subtask',exact:true}).click();
  const dialog=page.locator('.subtask-dialog');await choosePicker(dialog,'Jira space','EMPTY');
  await expect(dialog.locator('[data-hint="epics"]')).toContainText('EMPTY has no visible epics');
  await expect(dialog.getByRole('combobox',{name:'User story',exact:true})).toBeDisabled();
  await expect(dialog.locator('[data-hint="stories"]')).toBeEmpty();expect(storyReads).toBe(0);
});

test('Enter in dropdown search selects its result without submitting the main form',async({page})=>{
  const state=await mockSubtaskJira(page);
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'New subtask',exact:true}).click();
  const dialog=await chooseSubtaskHome(page);
  await dialog.getByRole('combobox',{name:'Jira space',exact:true}).click();
  await dialog.getByRole('searchbox',{name:'Search jira space',exact:true}).press('Enter');
  await expect(dialog.getByRole('searchbox',{name:'Search jira space',exact:true})).toBeHidden();
  await expect(dialog.locator('[name="stories"]')).toHaveValue('NEW-2');
  expect(state.reviews).toBe(0);expect(state.writes).toBe(0);
});

test('late hierarchy results cannot overwrite a different selected space',async({page})=>{
  await mockSubtaskJira(page);let releaseOld;
  const waitForOld=new Promise(resolve=>releaseOld=resolve);
  await page.route('**/worklog/subtask/options?*',async route=>{
    const params=new URL(route.request().url()).searchParams,kind=params.get('kind'),project=params.get('project_key');
    if(kind==='epics'&&project==='NEW')await waitForOld;
    return route.fulfill({json:{options:kind==='projects'?[{key:'NEW',summary:'First space'},{key:'MORE',summary:'Second space'}]:kind==='epics'?[{key:project+'-1',summary:'Epic'}]:[{key:project+'-2',summary:'Story'}],next_cursor:null,author:'Alice',today:'2026-11-11',first:'2026-11-01',last:'2026-11-30'}});
  });
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'New subtask',exact:true}).click();
  const dialog=page.locator('.subtask-dialog');await choosePicker(dialog,'Jira space','NEW');
  await choosePicker(dialog,'Jira space','MORE');await choosePicker(dialog,'Epic','MORE-1');
  await choosePicker(dialog,'User story','MORE-2');releaseOld();
  await expect(dialog.locator('[name="epics"]')).toHaveValue('MORE-1');
  await expect(dialog.locator('[name="stories"]')).toHaveValue('MORE-2');
  await dialog.getByRole('combobox',{name:'Epic',exact:true}).click();
  await expect(dialog.getByRole('option',{name:/^MORE-1 /})).toBeVisible();
  await expect(dialog.getByRole('option',{name:/^NEW-1 /})).toHaveCount(0);
});

test('RLT uses standalone leave trackers and records leave separately from work',async({page})=>{
  await mockSubtaskJira(page);const reads=[];let draft,leaveSaved=false;
  await page.route('**/worklog/subtask/options?*',route=>{
    const params=new URL(route.request().url()).searchParams,kind=params.get('kind');reads.push(Object.fromEntries(params));
    return route.fulfill({json:{options:kind==='projects'?[{key:'RLT',summary:'RnD Leave Tracker'}]:[{key:'RLT-1',summary:'Alice leave tracker'}],next_cursor:null,author:'Alice',today:'2026-11-11',first:'2026-11-01',last:'2026-11-30',leave_types:[{id:'planned',label:'Planned Leave'},{id:'unplanned',label:'Unplanned Leave'}]}});
  });
  await page.route('**/worklog/subtask/review',route=>{draft=route.request().postDataJSON();return route.fulfill({json:{token:'leave-create-token',review:{...draft,flow:'leave',author:'Alice',leave_label:'Planned Leave',story_summary:'Alice leave tracker'}}})});
  await page.route('**/worklog/subtask/submit',route=>route.fulfill({json:{issue_key:'RLT-3',summary:draft.summary,assignee:'Alice',issue_type:'Sub-task',parent_issue_key:'RLT-1',flow:'leave',verified:true}}));
  await page.route('**/worklog/subtask/search?*',route=>route.fulfill({json:{issues:[],next_cursor:null}}));
  await page.route('**/worklog/review',route=>{const body=route.request().postDataJSON();expect(body.entry_kind).toBe('leave');expect(body.issue_key).toBe('RLT-3');return route.fulfill({json:{token:'leave-log-token',review:{...body,author:'Alice',minutes:480,summary:'Annual leave',duration_label:'8h'}}})});
  await page.route('**/worklog/submit',route=>{leaveSaved=true;return route.fulfill({json:{worklog_id:'leave-log',verified:true}})});
  await page.route('**/employee-capacity-utilization/tally?*',route=>route.fulfill({json:{employee:'Alice',month:'2026-11',entries:[],leaves:leaveSaved?[{assignee:'Alice',period_day:'2026-11-11',planned_taken_hours:8,jira_task_ids:'RLT-3'}]:[]}}));
  await page.locator('#month').fill('2026-11');
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'New subtask',exact:true}).click();
  const dialog=page.locator('.subtask-dialog');await dialog.getByRole('combobox',{name:'Jira space',exact:true}).click();
  await dialog.getByRole('searchbox',{name:'Search jira space',exact:true}).fill('RLT');
  await dialog.getByRole('option',{name:/^RLT /}).click();
  await expect(dialog.getByRole('combobox',{name:'Epic',exact:true})).toHaveCount(0);
  await choosePicker(dialog,'Leave tracker','RLT-1');
  expect(reads.some(read=>read.kind==='epics')).toBe(false);
  expect(reads.find(read=>read.kind==='trackers').q).toBe('');
  await dialog.getByLabel('Leave type',{exact:true}).selectOption('planned');
  await dialog.getByLabel('Leave subtask title').fill('Annual leave');
  await dialog.getByRole('button',{name:'Review subtask'}).click();
  await expect(dialog.locator('.subtask-check')).toContainText('Planned Leave');
  expect(draft.epic_key).toBe('');expect(draft.story_key).toBe('RLT-1');
  await dialog.getByRole('button',{name:'Create subtask in Jira'}).click();
  await dialog.getByRole('button',{name:'Record leave on this subtask'}).click();
  await expect(page.locator('#availability-day-details')).toContainText('Record leave');
  await expect(page.locator('#worklog-issue')).toHaveValue('RLT-3');
  await expect(page.locator('#worklog-duration')).toHaveValue('8h');
  await page.locator('#worklog-comment').fill('Annual leave taken');
  await page.getByRole('button',{name:'Review leave',exact:true}).click();
  await expect(page.locator('.worklog-day-total')).toContainText('stays separate from work hours');
  await page.locator('#worklog-confirm').check();
  await page.getByRole('button',{name:'Confirm & record leave 8h'}).click();
  await expect(page.locator('.availability-stat.leave strong')).toHaveText('8h');
  await expect(page.locator('.availability-stat.work strong')).toHaveText('0h');
  await expect(page.locator('[data-leave-date="2026-11-11"]')).toContainText('8h');
});

test('leave recording links restore their RLT issue from live search without submitting',async({page})=>{
  await mockSubtaskJira(page);let writes=0;
  await page.route('**/worklog/subtask/search?*',route=>{expect(new URL(route.request().url()).searchParams.get('flow')).toBe('leave');return route.fulfill({json:{issues:[{issue_key:'RLT-1',issue_type:'Sub-task',summary:'Another leave',assignee:'Alice'},{issue_key:'RLT-3',issue_type:'Sub-task',summary:'Annual leave',assignee:'Alice'}],next_cursor:null}})});
  await page.route('**/worklog/submit',route=>{writes++;return route.fulfill({json:{}})});
  await page.goto('http://calendar.test/report.html#modal=availability&employee=Alice&month=2026-11&panel=leave&day=2026-11-11&issue=RLT-3');
  await expect(page.locator('#worklog-issue')).toHaveValue('RLT-3');
  await expect(page.locator('#availability-day-details')).toContainText('Record leave');
  await expect(page.locator('#worklog-selection')).toContainText('Annual leave');
  await page.reload();await expect(page.locator('#worklog-issue')).toHaveValue('RLT-3');
  expect(writes).toBe(0);
});

test('tally downloads a PDF for the displayed employee/month and reports failures honestly',async({page})=>{
  let requests=0;
  await page.route('**/api/employee-capacity-utilization/tally.pdf?*',route=>{
    const params=new URL(route.request().url()).searchParams;
    expect(params.get('employee')).toBe('Alice');expect(params.get('month')).toBe('2026-10');requests++;
    return route.fulfill({body:Buffer.from('%PDF-1.4\nmock tally'),contentType:'application/pdf'});
  });
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  const downloadPromise=page.waitForEvent('download');
  await page.getByRole('button',{name:'Download Jira tally PDF',exact:true}).click();
  const download=await downloadPromise;
  expect(download.suggestedFilename()).toBe('worklog-tally-Alice-2026-10.pdf');
  await expect(page.locator('.availability-tally-status')).toContainText('Tally downloaded for Alice · 2026-10');
  expect(requests).toBe(1);
  await page.route('**/api/employee-capacity-utilization/tally.pdf?*',route=>route.fulfill({status:502,json:{error:'Jira could not be read completely. No tally was generated.'}}));
  await page.getByRole('button',{name:'Download Jira tally PDF',exact:true}).click();
  await expect(page.locator('.availability-tally-status')).toContainText('No tally was generated.');
  await expect(page.getByRole('button',{name:'Download Jira tally PDF',exact:true})).toBeEnabled();
  await page.setViewportSize({width:390,height:844});
  expect(await page.locator('.availability-toolbar').evaluate(el=>el.scrollWidth-el.clientWidth)).toBeLessThanOrEqual(1);
});
test('personal history auto-loads the month, searches titles, paginates, exports and fits mobile',async({page})=>{
  await page.route('**/worklog/history?*',route=>{
    const more=new URL(route.request().url()).searchParams.has('cursor');
    return route.fulfill({json:{author:'Alice',entries:more?[
      {issue_key:'APP-2',summary:'Review release',worklog_id:'203',started:'2026-10-07T11:00:00+05:00',seconds:3600,comment:'=Review'},
    ]:[
      {issue_key:'APP-1',summary:'Build calendar',worklog_id:'201',started:'2026-10-05T09:00:00+05:00',seconds:1800,comment:{content:[{text:'Design notes <script>'}]}},
      {issue_key:'APP-2',summary:'Review release',worklog_id:'202',started:'2026-10-06T09:00:00+05:00',seconds:7200,comment:'Release checks'},
    ],next_cursor:more?null:'next'}});
  });
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'Manage my logged hours',exact:true}).click();
  const manager=page.locator('.wl-manager');
  await expect(manager.locator('.wl-entry')).toHaveCount(2);
  await expect(manager.locator('[data-stats]')).toContainText('2h 30m');
  await expect(manager.locator('[data-status]')).toContainText('totals cover loaded entries only');
  await manager.getByRole('button',{name:'Load more tasks'}).click();
  await expect(manager.locator('.wl-entry')).toHaveCount(3);
  await manager.locator('[name="search"]').fill('release');
  await expect(manager.locator('.wl-entry')).toHaveCount(2);
  await manager.locator('[name="sort"]').selectOption('longest');
  await expect(manager.locator('.wl-entry').first()).toContainText('2h');
  const downloadPromise=page.waitForEvent('download');
  await manager.getByRole('button',{name:'Export CSV'}).click();
  const download=await downloadPromise;
  const csv=fs.readFileSync(await download.path(),'utf8');
  expect(csv).toContain("'=Review");expect(csv).not.toContain('Design notes');
  await manager.locator('[name="search"]').fill('');
  await manager.locator('[name="task"]').selectOption('APP-1');
  await expect(manager.locator('.wl-entry')).toHaveCount(1);
  await expect(manager.locator('[data-entries]')).toContainText('Design notes <script>');
  await manager.locator('[name="task"]').selectOption('');
  await manager.screenshot({path:'tmp/my-logged-hours-redesign-desktop.png'});
  await page.setViewportSize({width:390,height:844});
  expect(await manager.evaluate(el=>el.scrollWidth-el.clientWidth)).toBeLessThanOrEqual(1);
  await manager.screenshot({path:'tmp/my-logged-hours-redesign-mobile.png'});
  await page.keyboard.press('Escape');
  await expect(manager).toHaveCount(0);
  await expect(page.locator('#ecu-drawer')).toHaveClass(/open/);
});
test('missing personal sign-in gives a useful setup message instead of empty history',async({page})=>{
  await page.route('**/worklog/history?*',route=>route.fulfill({status:400,json:{error:'Sign in with Jira to manage your logged hours.'}}));
  await page.route('**/api/jira-auth/status',route=>route.fulfill({json:{configured:false}}));
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'Manage my logged hours',exact:true}).click();
  await expect(page.locator('.wl-empty')).toContainText('Jira history is unavailable');
  await expect(page.locator('.wl-empty')).toContainText('needs administrator setup');
  await expect(page.locator('.wl-manager [data-export]')).toBeDisabled();
});
test('each metric modal has a reloadable URL with its employee and month',async({page})=>{
  for(const metric of ['logged','capacity','official','leaves','availability','booked','logged','utilization']){
    const index=['logged','capacity','official','leaves','availability','booked','logged','utilization'].indexOf(metric);
    await page.locator('#table tbody tr').first().locator('td').nth(index).click();
    await expect(page).toHaveURL(new RegExp('modal='+metric));
    await page.reload();
    await expect(page.locator('#ecu-drawer')).toHaveClass(/open/);
    await expect(page.locator('#ecu-drawer-subtitle')).toContainText('Alice');
    await expect(page.locator('#month')).toHaveValue('2026-10');
    await expect(page.getByRole('button',{name:'Copy link to details'})).toBeVisible();
    await page.getByRole('button',{name:'Close details'}).click();
  }
});
test('shared day and personal manager URLs restore their context without deleting',async({page})=>{
  let deletes=0;await page.route('**/worklog/delete',route=>{deletes++;return route.fulfill({json:{deleted:true}})});
  await page.route('**/worklog/entries?*',route=>route.fulfill({json:{author:'Alice',issue_key:'APP-1',entries:[{worklog_id:'123',started:'2026-10-05',seconds:1800}]}}));
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.locator('[data-work-date="2026-10-05"]').click();
  const dayUrl=page.url();await page.reload();
  await expect(page.locator('#availability-day-details')).toContainText('5.5h logged');
  await page.getByRole('button',{name:'Delete my entries on APP-1 for 2026-10-05'}).click();
  await expect(page.locator('.wl-manager .wl-entry')).toHaveCount(1);
  expect(page.url()).toContain('panel=hours');
  await page.goBack();
  await expect(page.locator('.wl-manager')).toHaveCount(0);
  await expect(page.locator('#availability-day-details')).toContainText('5.5h logged');
  await page.goForward();
  await expect(page.locator('.wl-manager .wl-entry')).toHaveCount(1);
  await page.reload();
  await expect(page.locator('.wl-manager .wl-entry')).toHaveCount(1);
  await expect(page.locator('.wl-manager [name="date"]')).toHaveValue('2026-10-05');
  expect(deletes).toBe(0);
  await page.locator('.wl-manager [data-close]').click();
  await page.goto(dayUrl);await expect(page.locator('#availability-day-details')).toContainText('5.5h logged');
});
test('worklog composer URL restores the form without sharing private draft text or submitting',async({page})=>{
  let writes=0;
  await page.route('**/worklog/identity',route=>route.fulfill({json:{author:'Alice',signed_in:true,write_enabled:true}}));
  await page.route('**/api/jira-auth/status',route=>route.fulfill({json:{configured:true,signed_in:true}}));
  await page.route('**/worklog/submit',route=>{writes++;return route.fulfill({json:{}})});
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'Log work on 2026-10-05'}).click();
  await page.locator('#worklog-comment').fill('Private draft text must stay out of links');
  expect(page.url()).toContain('panel=log');
  expect(page.url()).not.toContain('Private');
  await page.reload();
  await expect(page.locator('#worklog-comment')).toHaveValue('');
  await expect(page.locator('#availability-day-details')).toContainText('Monday, October 5, 2026');
  expect(writes).toBe(0);
});
test('modal month arrows preserve employee, update daily totals, and keep personal management accessible',async({page})=>{
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  const modal=page.locator('#ecu-drawer');
  await page.locator('[data-work-date="2026-10-05"]').click();
  await modal.getByRole('button',{name:'Previous month',exact:true}).click();
  await expect(page.locator('#month')).toHaveValue('2026-09');
  await expect(page.locator('#ecu-drawer-subtitle')).toContainText('Alice · 2026-09');
  await expect(page.locator('[data-calendar-date="2026-09-05"]')).toContainText('8h logged');
  await expect(page.locator('.availability-day[data-calendar-date]')).toHaveCount(30);
  await expect(page.locator('#availability-day-details')).toContainText('Explore a day');
  await modal.getByRole('button',{name:'Next month',exact:true}).click();
  await expect(page.locator('[data-calendar-date="2026-10-05"]')).toContainText('5.5h logged');
  await modal.getByRole('button',{name:'Next month',exact:true}).click();
  await expect(page.locator('#month')).toHaveValue('2026-11');
  await modal.getByRole('button',{name:'Manage my logged hours',exact:true}).click();
  await expect(page.getByRole('dialog',{name:'Manage my logged hours'})).toBeVisible();
});
test('month arrows cross years and render leap February on mobile',async({page})=>{
  await page.locator('#month').fill('2024-01');
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'Previous month',exact:true}).click();
  await expect(page.locator('#month')).toHaveValue('2023-12');
  await page.getByRole('button',{name:'Next month',exact:true}).click();
  await expect(page.locator('#month')).toHaveValue('2024-01');
  await page.getByRole('button',{name:'Next month',exact:true}).click();
  await expect(page.locator('.availability-day[data-calendar-date]')).toHaveCount(29);
  await page.setViewportSize({width:390,height:844});
  await expect(page.getByRole('button',{name:'Previous month',exact:true})).toBeVisible();
  await expect(page.getByRole('button',{name:'Next month',exact:true})).toBeVisible();
  await expect(page.getByRole('button',{name:'Manage my logged hours',exact:true})).toBeVisible();
  expect(await page.locator('.availability-toolbar').evaluate(el=>el.scrollWidth-el.clientWidth)).toBeLessThanOrEqual(1);
});
test('personal hours manager loads any task and deletes only after confirmation', async ({page}) => {
  let deletes = 0;
  await page.route('**/worklog/entries?*', route => route.fulfill({json:{issue_key:'STORY-1',author:'Alice',entries:[
    {worklog_id:'123',started:'2026-10-05T09:00:00+05:00',seconds:1800,comment:{content:[{text:'Fix entry'}]}},
    {worklog_id:'124',started:'2026-10-06T09:00:00+05:00',seconds:3600},
  ]}}));
  await page.route('**/worklog/delete', route => {
    deletes++; expect(route.request().postDataJSON()).toEqual({issue_key:'STORY-1',worklog_id:'123',confirmed:true});
    return route.fulfill({json:{deleted:true,issue_key:'STORY-1',worklog_id:'123'}});
  });
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'Manage my logged hours',exact:true}).click();
  const dialog=page.getByRole('dialog',{name:'Manage my logged hours'});
  await dialog.getByRole('textbox',{name:'Jira task key'}).fill('STORY-1');
  await dialog.getByRole('button',{name:'Load my entries'}).click();
  await expect(dialog).toContainText('2 entries on STORY-1');
  await dialog.locator('.wl-entry').filter({hasText:'Worklog 123'}).getByRole('button',{name:'Delete entry'}).click();
  await dialog.getByRole('button',{name:'Cancel deletion'}).click();
  expect(deletes).toBe(0);
  await dialog.locator('.wl-entry').filter({hasText:'Worklog 123'}).getByRole('button',{name:'Delete entry'}).click();
  await dialog.getByRole('button',{name:'Confirm delete'}).click();
  await expect(dialog.locator('[data-status]')).toHaveText('Deleted worklog 123 from Jira.');
  await expect(dialog.getByRole('button',{name:'Delete entry'})).toHaveCount(1);
  expect(deletes).toBe(1);
  await dialog.getByRole('button',{name:'Close worklog manager'}).click();
  await expect(dialog).toHaveCount(0);
});
test('day table Delete opens that task and date while retaining other dates after deletion', async ({page}) => {
  let deletes=0;
  await page.route('**/worklog/entries?*',route=>route.fulfill({json:{issue_key:'APP-1',author:'Alice',entries:[
    {worklog_id:'122',started:'2026-10-06T09:00:00+05:00',seconds:3600},
    {worklog_id:'123',started:'2026-10-05T09:00:00+05:00',seconds:1800},
  ]}}));
  await page.route('**/worklog/delete',route=>{
    deletes++;expect(route.request().postDataJSON().worklog_id).toBe('123');
    return route.fulfill({json:{deleted:true}});
  });
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.locator('[data-work-date="2026-10-05"]').click();
  await page.getByRole('button',{name:'Delete my entries on APP-1 for 2026-10-05',exact:true}).click();
  const dialog=page.getByRole('dialog',{name:'Manage my logged hours'});
  await expect(dialog.getByRole('textbox',{name:'Jira task key'})).toHaveValue('APP-1');
  await expect(dialog.locator('input[name="date"]')).toHaveValue('2026-10-05');
  await expect(dialog.getByRole('button',{name:'Delete entry'})).toHaveCount(1);
  await expect(dialog.locator('[data-entries]')).not.toContainText('Worklog 122');
  await dialog.getByRole('button',{name:'Delete entry'}).click();
  await dialog.getByRole('button',{name:'Confirm delete'}).click();
  await expect(dialog.locator('[data-status]')).toHaveText('Deleted worklog 123 from Jira.');
  await dialog.locator('input[name="date"]').fill('');
  await expect(dialog.locator('[data-entries]')).toContainText('Worklog 122');
  expect(deletes).toBe(1);
});
test('direct My logged hours link opens the manager automatically',async({page})=>{
  await page.route('**/my_day.html?manageWorklogs=1',route=>route.fulfill({contentType:'text/html',body:
    '<button data-manage-worklogs="">Manage my logged hours</button><script>'+fs.readFileSync('employee_worklog_manage.js','utf8')+'</script>'}));
  await page.goto('http://calendar.test/my_day.html?manageWorklogs=1');
  await expect(page.getByRole('dialog',{name:'Manage my logged hours'})).toBeVisible();
});
test('uncertain deletion disables blind retry and keeps the entry visible', async ({page}) => {
  let deletes=0;
  await page.route('**/worklog/entries?*',route=>route.fulfill({json:{issue_key:'APP-1',author:'Alice',entries:[{worklog_id:'123',started:'2026-10-05',seconds:1800}]}}));
  await page.route('**/worklog/delete',route=>{deletes++;return route.fulfill({status:502,json:{error:'Jira request timed out.'}})});
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'Manage my logged hours',exact:true}).click();
  const dialog=page.getByRole('dialog',{name:'Manage my logged hours'});
  await dialog.getByRole('textbox',{name:'Jira task key'}).fill('APP-1');
  await dialog.getByRole('button',{name:'Load my entries'}).click();
  await expect(dialog).toContainText('1 entries on APP-1');
  await dialog.getByRole('button',{name:'Delete entry'}).click();
  await dialog.getByRole('button',{name:'Confirm delete'}).click();
  await expect(dialog.locator('[data-status]')).toContainText('Check Jira and load entries again before retrying.');
  await expect(dialog.getByRole('button',{name:'Delete entry'})).toBeDisabled();
  await expect(dialog).toContainText('Worklog 123');
  expect(deletes).toBe(1);
});
test.beforeEach(async ({page}) => {
  await page.route('http://calendar.test/**', route => {
    const url = route.request().url();
    if (url.includes('/worklog/history?')) return route.fulfill({json:{author:'Other',entries:[],next_cursor:null}});
    if (url.endsWith('/api/employee-capacity-utilization/data')) return route.fulfill({json:payload});
    if (url.endsWith('/report.html')) return route.fulfill({body:html, contentType:'text/html'});
    return route.fulfill({body:'', contentType:url.endsWith('.css')?'text/css':'text/javascript'});
  });
  await page.goto('http://calendar.test/report.html');
  await page.locator('#month').fill('2026-10');
  await expect(page.locator('#table tbody tr').first()).toContainText('Alice');
});
test('calendar reconciles complete live month including deletions and preserves last result on errors',async({page})=>{
  let round=0;
  await page.route('**/employee-capacity-utilization/tally?*',route=>{
    if(round===2)return route.fulfill({status:502,json:{error:'Jira unavailable'}});
    if(round===1)return route.fulfill({json:{employee:'Alice',month:'2026-10',leaves:[],entries:[]}});
    return route.fulfill({json:{employee:'Alice',month:'2026-10',leaves:[],entries:[{jira:{issue_key:'APP-1',issue_type:'Sub-task',summary:'Build calendar',worklog_id:'1',date:'2026-10-05',seconds:3600}}]}});
  });
  const mainBefore=await page.locator('#table tbody tr').first().innerText();
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  const refresh=page.getByRole('button',{name:'Refresh Jira worklogs'});
  await expect(refresh.locator('.material-symbols-outlined')).toHaveText('sync');
  await expect(refresh).toHaveCSS('border-radius','9px');
  await expect(page.locator('.availability-live-status')).toContainText('Updated at');
  await expect(page.locator('[data-calendar-date="2026-10-05"]')).toContainText('1h logged');
  await page.locator('[data-work-date="2026-10-05"]').click();
  await expect(page.locator('#availability-day-details')).toContainText('1h logged');
  expect(await page.locator('#table tbody tr').first().innerText()).toBe(mainBefore);
  round=1;await page.getByRole('button',{name:'Refresh Jira worklogs'}).click();
  await expect(page.locator('[data-calendar-date="2026-10-05"]')).toContainText('0h logged');
  round=2;await page.getByRole('button',{name:'Refresh Jira worklogs'}).click();
  await expect(page.locator('.availability-live-status')).toContainText('could not be refreshed');
  await expect(page.locator('[data-calendar-date="2026-10-05"]')).toContainText('0h logged');
});
test('unchanged checks preserve selected day details and show checked status',async({page})=>{
  let reads=0;
  await page.route('**/employee-capacity-utilization/tally?*',route=>{
    reads++;
    return route.fulfill({json:{employee:'Alice',month:'2026-10',leaves:[],entries:[{jira:{issue_key:'APP-1',issue_type:'Sub-task',worklog_id:'1',date:'2026-10-05',seconds:3600}}],sync:{mode:reads===1?'updated':'unchanged',checked_at:'2026-11-11T06:00:00Z'}}});
  });
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await expect(page.locator('.availability-live-status')).toContainText('Updated at');
  await page.locator('[data-work-date="2026-10-05"]').click();
  await page.getByRole('button',{name:'Refresh Jira worklogs',exact:true}).click();
  await expect(page.locator('.availability-live-status')).toContainText('No changes');
  await expect(page.locator('#availability-day-details')).toContainText('1h logged');
});
test('live RLT leave replaces stale zero leave and reduces available capacity',async({page})=>{
  await page.route('**/employee-capacity-utilization/tally?*',route=>route.fulfill({json:{employee:'Alice',month:'2026-10',entries:[],leaves:[19,20,21,22].map(day=>({assignee:'Alice',period_day:'2026-10-'+day,planned_taken_hours:8,unplanned_taken_hours:0,jira_task_ids:'RLT-'+day}))}}));
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await expect(page.locator('.availability-live-status')).toContainText('RLT leave · Updated at');
  await expect(page.locator('.availability-stat.leave strong')).toHaveText('32h');
  await expect(page.locator('.availability-stat:not(.work):not(.leave) strong')).toHaveText('136h');
  for(const day of [19,20,21,22])await expect(page.locator('[data-leave-date="2026-10-'+day+'"]')).toContainText('8h');
  await page.locator('[data-leave-date="2026-10-19"]').click();
  await expect(page.locator('#availability-day-details')).toContainText('RLT-19');
});
test('switching months fetches Jira, shows progress, and ignores late responses',async({page})=>{
  let releaseSeptember;const septemberWait=new Promise(resolve=>releaseSeptember=resolve),months=[];
  await page.route('**/employee-capacity-utilization/tally?*',async route=>{
    const period=new URL(route.request().url()).searchParams.get('month');months.push(period);
    if(period==='2026-09')await septemberWait;
    return route.fulfill({json:{employee:'Alice',month:period,leaves:[],entries:[{jira:{issue_key:'APP-1',issue_type:'Sub-task',summary:'Build',worklog_id:period,date:period+'-05',seconds:period==='2026-09'?7200:3600}}]}}).catch(()=>{});
  });
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await expect(page.locator('.availability-live-status')).toContainText('Updated at');
  await page.getByRole('button',{name:'Previous month',exact:true}).click();
  const progress=page.getByRole('progressbar',{name:'Loading Jira worklogs for 2026-09'});
  await expect(progress).toBeVisible();
  await expect(page.locator('#ecu-drawer-title [role="progressbar"]')).toBeVisible();
  await expect(progress).toHaveCSS('width','15px');
  await expect(page.locator('.availability-live-status')).toContainText('Loading Jira worklogs and RLT leave for 2026-09');
  await expect(page.locator('.availability-calendar-scroll')).toHaveAttribute('aria-busy','true');
  await page.getByRole('button',{name:'Next month',exact:true}).click();
  await expect(page.locator('.availability-live-status')).toContainText('Updated at');
  await expect(page.locator('[data-calendar-date="2026-10-05"]')).toContainText('1h logged');
  await expect(page.getByRole('progressbar')).toHaveCount(0);
  releaseSeptember();
  await expect(page.locator('#month')).toHaveValue('2026-10');
  await expect(page.locator('[data-calendar-date="2026-10-05"]')).toContainText('1h logged');
  expect(months.filter(value=>value==='2026-10').length).toBeGreaterThanOrEqual(2);
  expect(months).toContain('2026-09');
});
test('initial transient Jira failure retries promptly and clears the header spinner',async({page})=>{
  let reads=0;
  await page.route('**/employee-capacity-utilization/tally?*',route=>{
    reads++;
    if(reads===1)return route.fulfill({status:502,json:{error:'Temporary Jira timeout'}});
    return route.fulfill({json:{employee:'Alice',month:'2026-10',leaves:[],entries:[{jira:{issue_key:'APP-1',issue_type:'Sub-task',worklog_id:'1',date:'2026-10-05',seconds:3600}}]}});
  });
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await expect(page.locator('#ecu-drawer-title [role="progressbar"]')).toBeVisible();
  await expect(page.locator('.availability-live-status')).toContainText('Updated at');
  await expect(page.locator('[data-calendar-date="2026-10-05"]')).toContainText('1h logged');
  await expect(page.getByRole('progressbar')).toHaveCount(0);expect(reads).toBe(2);
});
test('another employee calendar reads their own Jira hours without personal worklog authentication',async({page})=>{
  let requestedEmployee='',personalReads=0;
  await page.route('**/worklog/history?*',route=>{personalReads++;return route.fulfill({status:401,json:{error:'No personal Jira session'}})});
  await page.route('**/employee-capacity-utilization/tally?*',route=>{
    const params=new URL(route.request().url()).searchParams;requestedEmployee=params.get('employee');
    return route.fulfill({json:{employee:'Bob',month:params.get('month'),leaves:[],entries:[{jira:{issue_key:'APP-1',issue_type:'Sub-task',summary:'Build',worklog_id:'44',date:'2026-10-05',seconds:10800}}]}});
  });
  await page.locator('#table tbody tr:not(.total)').filter({hasText:'Bob'}).locator('td').nth(4).click();
  await expect(page.locator('.availability-live-status')).toContainText('Updated at');
  await expect(page.locator('[data-calendar-date="2026-10-05"]')).toContainText('3h logged');
  expect(requestedEmployee).toBe('Bob');expect(personalReads).toBe(0);
});
test('employee search filters names, totals, and the Excel request',async({page})=>{
  let filters;
  await page.route('**/api/employee-capacity-utilization/export',route=>{filters=route.request().postDataJSON();return route.fulfill({body:'mock workbook',contentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})});
  const search=page.getByRole('searchbox',{name:'Search employees'});
  await search.fill('  ALI  ');
  await expect(page.locator('#table tbody tr:not(.total)')).toHaveCount(1);
  await expect(page.locator('#table tbody tr:not(.total)')).toContainText('Alice');
  await expect(page.locator('#status')).toContainText('1 employees');
  const download=page.waitForEvent('download');await page.getByRole('button',{name:'Download Excel'}).click();await download;
  expect(filters.employee_search).toBe('  ALI  ');
  await search.fill('nobody');await expect(page.locator('#table')).toContainText('No employees match');
  await search.fill('');await expect(page.locator('#table tbody tr:not(.total)')).toHaveCount(2);
});
test('personal entry edits submit the Jira version and refresh the list',async({page})=>{
  let sent=null;
  await page.route('**/worklog/history?*',route=>route.fulfill({json:{author:'Alice',entries:[{issue_key:'APP-1',issue_type:'Sub-task',summary:'Build',worklog_id:'10',updated:'v1',started:'2026-10-05T09:00:00.000+0500',seconds:1800,comment:'Original'}],next_cursor:null}}));
  await page.route('**/worklog/edit',route=>{sent=route.request().postDataJSON();return route.fulfill({json:{updated:true}})});
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.getByRole('button',{name:'Manage my logged hours'}).click();
  const dialog=page.locator('.wl-manager');await dialog.getByRole('button',{name:'Edit entry'}).click();
  await dialog.locator('[name="duration"]').fill('1h');
  await dialog.getByRole('button',{name:'Save changes to Jira'}).click();
  await expect(dialog.locator('[data-status]')).toContainText('Changes saved');
  expect(sent.updated).toBe('v1');expect(sent.duration).toBe('1h');expect(sent.comment_changed).toBe(false);
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
  await expect(page.locator('#ecu-drawer')).toHaveCSS('visibility','hidden');
  await expect(page.locator('#ecu-drawer')).toHaveCSS('pointer-events','none');
  await expect(page.locator('#table tbody tr').first().locator('td').nth(4)).toBeFocused();
  await expect(page.locator('#table tbody tr').first().locator('td').nth(4)).toHaveCSS('color','rgb(102, 66, 191)');
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
test('hover add opens searchable worklog form; review and explicit confirmation precede a single write', async ({page}) => {
  let writes=0;
  await page.route('**/api/jira-auth/status',route=>route.fulfill({json:{configured:true}}));
  await page.route('**/worklog/identity',route=>route.fulfill({json:{author:'Alice',write_enabled:true,signed_in:true}}));
  await page.route('**/worklog/review',async route=>{
    const body=route.request().postDataJSON();
    expect(body.issue_key).toBe('APP-1');expect(body.minutes).toBe(30);
    return route.fulfill({json:{token:'review-token',review:{...body,author:'Alice',summary:'Build calendar'}}});
  });
  await page.route('**/worklog/submit',route=>{
    writes++;expect(route.request().postDataJSON()).toEqual({token:'review-token',confirmed:true});
    return route.fulfill({json:{worklog_id:'test-1',verified:true}});
  });
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  const add=page.locator('[data-add-date="2026-10-05"]');
  await expect(add).toHaveCSS('opacity','0');
  await page.locator('[data-calendar-date="2026-10-05"]').hover();
  await expect(add).toHaveCSS('opacity','1');
  await add.click();
  await page.locator('#worklog-search').fill('calendar');
  await expect(page.locator('#worklog-issue option')).toHaveCount(1);
  await page.locator('#worklog-issue').selectOption('APP-1');
  await page.locator('#worklog-comment').fill('Calendar implementation');
  await page.getByRole('button',{name:'Review worklog',exact:true}).click();
  await expect(page.locator('.worklog-review')).toContainText('Alice');
  await expect(page.locator('#worklog-submit')).toBeDisabled();expect(writes).toBe(0);
  await page.locator('#worklog-confirm').check();
  await page.locator('#worklog-submit').click();
  await expect(page.locator('#availability-day-details')).toContainText('Worklog test-1 created');
  expect(writes).toBe(1);
});
test('duration input accepts units, rejects over 24h, and preserves edits when returning from review', async ({page}) => {
  let reviews=0;
  await page.route('**/api/jira-auth/status',route=>route.fulfill({json:{configured:true}}));
  await page.route('**/worklog/identity',route=>route.fulfill({json:{author:'Alice',write_enabled:true,signed_in:true}}));
  await page.route('**/worklog/review',route=>{
    reviews++;
    const body=route.request().postDataJSON();
    expect(body.duration).toBe('1h 30m');expect(body.minutes).toBe(90);
    return route.fulfill({json:{token:'token',review:{...body,author:'Alice',summary:'Build calendar',duration_label:'1h 30m',logged_minutes:60,total_minutes:150}}});
  });
  await page.locator('#table tbody tr').first().locator('td').nth(4).click();
  await page.locator('[data-add-date="2026-10-05"]').click();
  await page.locator('#worklog-search').fill('APP-1');
  await page.locator('#worklog-issue').selectOption('APP-1');
  await page.locator('#worklog-comment').fill('Useful work description');
  const duration=page.locator('#worklog-duration');
  for(const [input,preview] of [['1d','24h · 1440 minutes'],['8h','8h · 480 minutes'],['8m','8m · 8 minutes'],['0.5h','30m · 30 minutes']]) {
    await duration.fill(input);await expect(page.locator('#worklog-duration-preview')).toHaveText(preview);
  }
  for(const input of ['2d','25h','1d 1m','1h 2h','garbage']) {
    await duration.fill(input);
    await expect(duration).toHaveAttribute('aria-invalid','true');
    await page.getByRole('button',{name:'Review worklog',exact:true}).click();
    expect(reviews).toBe(0);
  }
  await page.getByRole('button',{name:'4h',exact:true}).click();
  await expect(duration).toHaveValue('4h');
  await duration.fill('1h 30m');
  await page.getByRole('button',{name:'Review worklog',exact:true}).click();
  await expect(page.locator('.worklog-day-total')).toContainText('2.5h / 24h');
  await expect(page.locator('.worklog-review')).toContainText('1h 30m (90 minutes)');
  await page.getByRole('button',{name:'Back',exact:true}).click();
  await expect(page.locator('#worklog-duration')).toHaveValue('1h 30m');
  await expect(page.locator('#worklog-comment')).toHaveValue('Useful work description');
  await expect(page.locator('#worklog-issue')).toHaveValue('APP-1');
  await page.setViewportSize({width:390,height:844});
  expect(await page.locator('#calendar-worklog-form').evaluate(el=>el.scrollWidth-el.clientWidth)).toBeLessThanOrEqual(1);
});
