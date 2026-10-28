const {test,expect}=require('@playwright/test');
const fs=require('node:fs');
async function fixture(page,role='member',approved=true){
  let saved;
  await page.route('http://epr.test/**',route=>{
    const path=new URL(route.request().url()).pathname;
    if(path==='/api/auth/me')return route.fulfill({json:{enabled:true,authenticated:true,user:{display_name:'Hassan Malik',role,approved,active:true,csrf_token:'test-csrf'}}});
    if(path==='/api/jira-auth/status')return route.fulfill({json:{configured:true}});
    if(path==='/api/auth/users/member'&&route.request().method()==='PUT'){saved=route.request().postDataJSON();return route.fulfill({json:{user:saved}})}
    if(path==='/api/auth/users')return route.fulfill({json:{users:[{account_id:'member',display_name:'New Employee',role:'member',approved:false,active:true}]}});
    if(path==='/api/auth/audit')return route.fulfill({json:{events:[]}});
    if(path==='/api/auth/logout')return route.fulfill({json:{ok:true}});
    if(path.startsWith('/api/'))return route.fulfill({json:{}});
    const files={'/login':'epr_login.html','/access-pending':'epr_pending.html','/settings/users':'epr_users.html','/shared-nav.js':'shared-nav.js','/shared-nav.css':'shared-nav.css'};
    const file=files[path];
    return route.fulfill({body:file?fs.readFileSync(file,'utf8'):'<!doctype html><title>Report</title><script src="/shared-nav.js"></script>',contentType:path.endsWith('.js')?'text/javascript':path.endsWith('.css')?'text/css':'text/html'});
  });
  return ()=>saved;
}
test('login is clear on desktop and mobile',async({page})=>{
  await fixture(page);await page.goto('http://epr.test/login');
  await expect(page.getByRole('heading',{name:'Sign in to continue'})).toBeVisible();
  await expect(page.getByRole('link',{name:'Continue with Atlassian'})).toHaveAttribute('href','/api/jira-auth/login');
  await page.screenshot({path:'tmp/epr-login-desktop.png'});
  await page.setViewportSize({width:390,height:844});
  expect(await page.evaluate(()=>document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
  await expect(page.getByRole('link',{name:'Continue with Atlassian'})).toBeVisible();
});
test('member identity shows personal links and hides administration; logout returns to login',async({page})=>{
  await fixture(page);await page.goto('http://epr.test/report.html');
  await expect(page.locator('.epr-nav-identity')).toContainText('Hassan Malik');
  await expect(page.getByRole('link',{name:'My availability'})).toHaveAttribute('href','/my-availability');
  await expect(page.getByRole('button',{name:/Admin Settings/})).toHaveCount(0);
  await page.getByRole('button',{name:'Sign out',exact:true}).click();
  await expect(page).toHaveURL('http://epr.test/login');
});
test('admin can approve a member and access the user management menu',async({page})=>{
  const saved=await fixture(page,'admin');await page.goto('http://epr.test/settings/users');
  await expect(page.getByRole('heading',{name:'Users & access'})).toBeVisible();
  await page.getByRole('button',{name:/Admin Settings/}).click();
  await expect(page.getByRole('link',{name:/Users & access/})).toBeVisible();
  await page.getByRole('checkbox',{name:'Approved for New Employee'}).check();
  await page.getByRole('button',{name:'Save access'}).click();
  await expect(page.getByRole('status')).toHaveText('Access updated.');
  expect(saved()).toEqual({role:'member',approved:true,active:true});
});
test('pending account reaches the approval screen',async({page})=>{
  await fixture(page,'member',false);await page.goto('http://epr.test/report.html');
  await expect(page).toHaveURL('http://epr.test/access-pending');
  await expect(page.getByRole('heading')).toContainText('approval');
});
