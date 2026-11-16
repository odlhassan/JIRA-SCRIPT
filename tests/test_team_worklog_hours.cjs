const {test,expect}=require('@playwright/test');
const fs=require('fs');
const html=fs.readFileSync('team_worklog_hours_report.html','utf8');
const payload={resources:{Alice:{},Bob:{}},issues:[],teams:[],profiles:[],support:[],worklogs:[
 {issue_id:'APP-1',item_issue_type:'Sub-task',item_assignee:'Alice',worklog_author:'Alice',worklog_date:'2026-10-05',hours_logged:3},
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
 await page.locator('#scope').selectOption('assigned');await expect(fifth).toContainText('3.00h');
 await page.locator('#include-leaves-logged').check();await expect(fifth).toContainText('4.00h');
 const matrix=await page.evaluate(()=>matrixExportRows);expect(matrix[1][9]).toBe(4);expect(matrix[1][10]).toBe(2);
});
