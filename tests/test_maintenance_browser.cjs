const {chromium} = require('playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const {spawn, execFileSync} = require('node:child_process');

(async () => {
  const root = path.resolve(__dirname, '..');
  const python = path.join(root, '.venv', process.platform === 'win32' ? 'Scripts/python.exe' : 'bin/python');
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), 'epr-game-'));
  const marker = path.join(temporary, 'maintenance.json');
  const port = process.env.MAINTENANCE_TEST_PORT || '3097';
  const base = `http://127.0.0.1:${port}`;
  const env = {...process.env, EPR_MAINTENANCE_FILE: marker};
  const server = spawn(python, ['tools/preview_maintenance.py', '--port', port], {cwd: root, env, windowsHide: true, stdio: 'ignore'});
  let browser;
  function phase(value) {
    execFileSync(python, ['maintenance_service.py', value, '--file', marker], {cwd: root, windowsHide: true, stdio: 'ignore'});
  }
  try {
    let ready = false;
    for (let n = 0; n < 80; n++) {
      if (server.exitCode !== null) throw new Error('Maintenance preview could not start. Choose another MAINTENANCE_TEST_PORT.');
      try { ready = (await fetch(base + '/healthz')).ok; } catch {}
      if (ready) break;
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    assert.ok(ready, 'preview starts');
    browser = await chromium.launch({headless: true});
    const context1 = await browser.newContext({viewport: {width: 1440, height: 1000}});
    await context1.route(base + '/app-shortcuts.js', route => route.fulfill({path: path.join(root, 'app-shortcuts.js'), contentType: 'application/javascript'}));
    const context2 = await browser.newContext({viewport: {width: 390, height: 844}, isMobile: true, hasTouch: true});
    const police = await context1.newPage(), runner = await context2.newPage();
    for (const page of [police, runner]) await page.addInitScript(() => {
      window.__gameFrames = [];
      const arc = CanvasRenderingContext2D.prototype.arc;
      CanvasRenderingContext2D.prototype.arc = function(x,y,r,...rest) {
        if (this.canvas.id === 'field' && (r === 21 || (r === 15 && this.fillStyle === '#be844a'))) {
          window.__gameFrames.push({t:performance.now(), x,y, local:r===21});
          if (window.__gameFrames.length > 2400) window.__gameFrames.shift();
        }
        return arc.call(this,x,y,r,...rest);
      };
    });
    const errors = [];
    police.on('pageerror', e => errors.push(e.message));
    runner.on('pageerror', e => errors.push(e.message));
    await police.goto(base + '/maintenance');
    await police.locator('#name').fill('Officer');
    await police.getByRole('button', {name: 'Enter playground'}).click();
    await police.waitForFunction(() => document.querySelector('#role').textContent.includes('You’re police'));
    await runner.goto(base + '/maintenance');
    await runner.locator('#name').fill('Samosa');
    await runner.getByRole('button', {name: 'Enter playground'}).click();
    await runner.waitForFunction(() => document.querySelector('#role').textContent.includes('You’re a runner'));
    // Measure actual painted positions, with delayed network responses. Input
    // must respond before a response, and rendering must fill the packet gaps.
    await police.route('**/api/maintenance/game', async route => {
      if (route.request().method() === 'POST' && route.request().postDataJSON().action === 'move') {
        const result = await route.fetch();
        await new Promise(resolve => setTimeout(resolve,120));
        return route.fulfill({response:result});
      }
      return route.continue();
    });
    await police.evaluate(() => { window.__gameFrames=[]; });
    const movementStarted = await police.evaluate(() => performance.now());
    await police.keyboard.down('d');
    await police.waitForTimeout(460);
    await police.keyboard.up('d');
    const painted = await police.evaluate(() => window.__gameFrames.filter(p => p.local));
    const moving = painted.filter((p,n) => n && p.x>painted[n-1].x+.1);
    assert.ok(moving.length >= 10, `movement paints between packets (${moving.length} frames)`);
    assert.ok(moving[0].t-movementStarted < 100, 'movement responds before delayed server acknowledgement');
    assert.ok(Math.max(...painted.slice(1).map((p,n) => Math.abs(p.x-painted[n].x))) < 12, 'no visible network-sized jumps');
    await police.waitForTimeout(300);
    assert.ok(await runner.getByRole('button', {name: 'Move up', exact: true}).isVisible());
    assert.ok(await runner.evaluate(() => document.documentElement.scrollWidth <= innerWidth), 'mobile has no horizontal overflow');
    const before = await runner.evaluate(async () => (await (await fetch('/api/maintenance/game')).json()).players.find(p => p.name === 'Samosa'));
    await runner.getByRole('button', {name: 'Move left', exact: true}).scrollIntoViewIfNeeded();
    const left = await runner.getByRole('button', {name: 'Move left', exact: true}).boundingBox();
    await runner.mouse.move(left.x + left.width / 2, left.y + left.height / 2);
    await runner.mouse.down();
    await police.evaluate(() => { window.__gameFrames=[]; });
    await runner.waitForTimeout(1400);
    await runner.mouse.up();
    const after = await runner.evaluate(async () => (await (await fetch('/api/maintenance/game')).json()).players.find(p => p.name === 'Samosa'));
    assert.ok(after.x < before.x, `touch controls move a runner: ${JSON.stringify({before, after, errors})}`);
    const remote = await police.evaluate(() => window.__gameFrames.filter(p => !p.local));
    const remoteSteps = remote.slice(1).map((p,n) => Math.abs(p.x-remote[n].x));
    assert.ok(remoteSteps.filter(distance => distance>.1).length >= 15, 'other players move between server packets');
    assert.ok(Math.max(...remoteSteps) < 15, 'remote interpolation avoids packet-sized jumps');
    await runner.getByRole('button', {name: 'Ask to join police'}).click();
    await police.getByRole('button', {name: 'Accept', exact: true}).waitFor();
    await police.getByRole('button', {name: 'Accept', exact: true}).click();
    await runner.waitForFunction(() => document.querySelector('#role').textContent.includes('You’re police'));
    assert.ok(await runner.locator('#team').isHidden(), 'two-officer team hides join request');
    await fs.mkdir(path.join(root, 'tmp'), {recursive: true});
    await police.screenshot({path: path.join(root, 'tmp/maintenance-desktop.png'), fullPage: true});
    await runner.screenshot({path: path.join(root, 'tmp/maintenance-mobile.png'), fullPage: true});

    phase('deploying');
    const response = await fetch(base + '/api/private', {method: 'POST', body: '{}'});
    assert.equal(response.status, 503, 'maintenance blocks app mutations');
    assert.equal((await fetch(base + '/protected', {headers: {Accept: 'text/html'}})).status, 503);
    await police.waitForFunction(() => document.querySelector('.availability').dataset.phase === 'deploying', {timeout: 10000});
    phase('available');
    await police.waitForURL(base + '/', {timeout: 12000});
    assert.ok((await police.textContent('body')).includes('App available'), 'availability restores the app automatically');

    const defaults = await browser.newContext();
    const defaultPage = await defaults.newPage();
    await defaultPage.route(base + '/app-shortcuts.js', route => route.fulfill({path: path.join(root, 'app-shortcuts.js'), contentType: 'application/javascript'}));
    await defaultPage.route(base + '/shared-nav.js', route => route.fulfill({path: path.join(root, 'shared-nav.js'), contentType: 'application/javascript'}));
    await defaultPage.route(base + '/api/game-shortcut', route => route.fulfill({json: {shortcut:'Ctrl+G'}}));
    await defaultPage.route(base + '/default.html*', route => route.fulfill({body:'<!doctype html><script src="/shared-nav.js"></script><p>Default shortcut report</p>',contentType:'text/html'}));
    const loadedDefault = defaultPage.waitForResponse(r => r.url().endsWith('/api/game-shortcut'));
    await defaultPage.goto(base + '/default.html?view=2#anchor');
    await loadedDefault;
    await defaultPage.keyboard.press('Control+G');
    await defaultPage.waitForURL(/\/maintenance\?return=/);
    await defaultPage.locator('#name').focus();
    await defaultPage.keyboard.press('Control+G');
    await defaultPage.waitForURL(base + '/default.html?view=2#anchor');
    await defaults.close();

    // Anonymous visitors learn only the non-secret key from the public status,
    // so a custom shortcut also works without an app login or cached preference.
    const anonymous = await browser.newContext();
    const anonymousPage = await anonymous.newPage();
    let protectedPreferenceReads = 0;
    anonymousPage.on('request', r => { if (r.url().endsWith('/api/game-shortcut')) protectedPreferenceReads++; });
    await anonymousPage.route(base + '/api/maintenance/status', route => route.fulfill({json:{available:true,phase:'available',message:'Ready',shortcut:'Alt+7'}}));
    await anonymousPage.route(base + '/signed-out.html*', route => route.fulfill({body:'<!doctype html><p>App login destination</p>',contentType:'text/html'}));
    await anonymousPage.goto(base + '/maintenance?return=' + encodeURIComponent('/signed-out.html?tab=2#top'));
    await anonymousPage.waitForFunction(() => document.querySelector('.availability').dataset.phase === 'available');
    await anonymousPage.keyboard.press('Alt+7');
    await anonymousPage.waitForURL(base + '/signed-out.html?tab=2#top');
    assert.equal(protectedPreferenceReads,0,'anonymous return uses no protected preference API');
    await anonymous.close();

    // The current app loads a configurable shortcut without advertising the game.
    const hidden = await context1.newPage();
    let shortcutLoaded = false;
    await hidden.route(base + '/saved.html', route => route.fulfill({body: '<!doctype html><script src="/shared-nav.js"></script><p>Saved report</p>', contentType: 'text/html'}));
    await hidden.route(base + '/shared-nav.js', route => route.fulfill({path: path.join(root, 'shared-nav.js'), contentType: 'application/javascript'}));
    await hidden.route(base + '/api/game-shortcut', route => { shortcutLoaded = true; return route.fulfill({json: {shortcut: 'Ctrl+Shift+G'}}); });
    await hidden.goto(base + '/saved.html');
    for (let n = 0; n < 30 && !shortcutLoaded; n++) await hidden.waitForTimeout(100);
    await hidden.waitForTimeout(100);
    assert.ok(!(await hidden.textContent('body')).includes('Ctrl+'), 'normal UI never advertises shortcut');
    await hidden.keyboard.press('Control+Shift+G');
    await hidden.waitForURL(/\/maintenance\?return=/);
    await hidden.waitForTimeout(5500);
    assert.ok(hidden.url().includes('/maintenance'), 'healthy app allows uninterrupted play');
    assert.ok((await hidden.locator('#return').getAttribute('href')).includes('/saved.html'));
    if (await hidden.locator('#welcome').isVisible()) {
      await hidden.locator('#name').fill('Private return');
      await hidden.getByRole('button', {name:'Enter playground'}).click();
      await hidden.waitForFunction(() => document.getElementById('welcome').hidden);
    }
    const leaving = hidden.waitForRequest(r => r.url().endsWith('/api/maintenance/game') && r.method()==='POST' && r.postDataJSON().action==='leave');
    await hidden.keyboard.press('Control+Shift+G');
    await hidden.waitForURL(base + '/saved.html');
    await leaving;
    assert.ok((await hidden.textContent('body')).includes('Saved report'), 'the same custom shortcut returns to the original app');
    await hidden.keyboard.press('Control+Shift+G');
    await hidden.waitForURL(/\/maintenance\?return=/);

    // Current status wins even if the availability strip still says healthy.
    phase('deploying');
    const gameUrl = hidden.url();
    await hidden.keyboard.press('Control+Shift+G');
    await hidden.waitForTimeout(400);
    assert.equal(hidden.url(), gameUrl, 'shortcut cannot return to an unavailable app');
    await hidden.route(base + '/api/maintenance/status', route => route.fulfill({status:503, json:{error:'restarting'}}));
    await hidden.keyboard.press('Control+Shift+G');
    await hidden.waitForTimeout(300);
    assert.equal(hidden.url(), gameUrl, 'failed availability check keeps the game open');
    await hidden.unroute(base + '/api/maintenance/status');
    // The private key remains known even while the app's preference API is paused.
    await hidden.route(base + '/api/game-shortcut', route => route.fulfill({status:503,json:{maintenance:true}}));
    await hidden.reload();
    phase('available');
    await hidden.keyboard.press('Control+Shift+G');
    await hidden.waitForURL(base + '/saved.html');

    // The superadmin alone sees and saves the setting in Users & access.
    const settings = await context1.newPage();
    let owner = true, savedShortcut = 'Ctrl+G';
    await settings.route(base + '/settings/users', route => route.fulfill({path: path.join(root, 'epr_users.html'), contentType: 'text/html'}));
    await settings.route(base + '/epr-superadmin-preferences.js', route => route.fulfill({path: path.join(root, 'epr-superadmin-preferences.js'), contentType: 'application/javascript'}));
    await settings.route(base + '/shared-nav.js', route => route.fulfill({path: path.join(root, 'shared-nav.js'), contentType: 'application/javascript'}));
    await settings.route(base + '/api/auth/me', route => route.fulfill({json: {enabled: true, authenticated: true, user: {account_id: 'owner', display_name: 'Owner', role: 'admin', active: true, approved: true, is_superadmin: owner, csrf_token: 'csrf'}}}));
    await settings.route(base + '/api/auth/users', route => route.fulfill({json: {users: []}}));
    await settings.route(base + '/api/auth/audit', route => route.fulfill({json: {events: []}}));
    await settings.route(base + '/api/auth/game-shortcut', async route => {
      if (route.request().method() === 'PUT') savedShortcut = route.request().postDataJSON().shortcut;
      return route.fulfill({json: {shortcut: savedShortcut}});
    });
    await settings.goto(base + '/settings/users');
    await settings.getByLabel('Keyboard shortcut', {exact: true}).fill('Ctrl+Shift+H');
    await settings.getByRole('button', {name: 'Save shortcut'}).click();
    await settings.waitForFunction(() => document.querySelector('.shortcut-status')?.textContent === 'Saved');
    assert.equal(savedShortcut, 'Ctrl+Shift+H');
    owner = false;
    await settings.reload();
    await settings.waitForTimeout(400);
    assert.equal(await settings.getByRole('heading', {name: 'Private playground shortcut'}).count(), 0, 'ordinary admins see no shortcut setting');

    // A saved app page loads the current shared script and moves to maintenance.
    phase('deploying');
    const nav = await context1.newPage();
    await nav.route(base + '/saved.html', route => route.fulfill({body: '<!doctype html><script src="/shared-nav.js"></script><p>Saved report</p>', contentType: 'text/html'}));
    await nav.route(base + '/shared-nav.js', route => route.fulfill({path: path.join(root, 'shared-nav.js'), contentType: 'application/javascript'}));
    await nav.goto(base + '/saved.html');
    await nav.waitForURL(/\/maintenance\?return=/, {timeout: 12000});
    assert.ok(nav.url().includes(encodeURIComponent('/saved.html')), 'navigation preserves original destination');
    assert.deepEqual(errors, []);
    console.log('Passed: desktop/mobile, multiplayer movement, team approval, hidden default/custom shortcut toggles, return URL preservation, maintenance/failure guard, owner preferences and open-page handoff.');
  } finally {
    if (browser) await browser.close();
    server.kill();
    await new Promise(resolve => { if (server.exitCode !== null) resolve(); else server.once('exit', resolve); });
    await fs.rm(temporary, {recursive: true, force: true});
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
