/* No external libraries, reporting data, authentication or persistent game records. */
(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const canvas = $('field'), ctx = canvas.getContext('2d');
  const motion = new ChorrMotion.Motion();
  const background = document.createElement('canvas');
  background.width = 1000; background.height = 600;
  let state = {players: [], obstacles: []}, joined = false, joining = false, posting = false, wantsToPlay = false;
  let initialAvailability, returnOnReady = false, rosterKey = '', stopped = false, pollTimer, pendingInput = false;
  const keys = new Set();
  let pointer = {dx: 0, dy: 0};
  const name = sessionStorage.getItem('epr-play-name') || '';
  $('name').value = name;
  function destination() {
    const value = new URLSearchParams(location.search).get('return');
    if (value && value.startsWith('/') && !value.startsWith('//') && !value.includes('\\') && !/[\x00-\x1f]/.test(value)) {
      const parsed = new URL(value, location.origin);
      if (parsed.origin === location.origin && !parsed.pathname.startsWith('/maintenance') && !parsed.pathname.startsWith('/api/')) return parsed.pathname + parsed.search + parsed.hash;
    }
    return location.pathname === '/maintenance' ? '/' : location.pathname + location.search + location.hash;
  }
  $('return').href = destination();
  async function api(data) {
    const sentAt = performance.now();
    const response = await fetch('/api/maintenance/game', data ? {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(data)} : {cache: 'no-store'});
    const result = await response.json();
    if (!response.ok) {
      if (response.status === 401) { joined = false; $('welcome').hidden = false; }
      throw new Error(result.error || 'Connection interrupted. Reconnecting…');
    }
    if (!motion.receive(result, performance.now(), performance.now()-sentAt, data)) return state;
    state = result;
    joined = !!state.you;
    renderRoster();
    return result;
  }
  function element(tag, text, className) {
    const node = document.createElement(tag); node.textContent = text;
    if (className) node.className = className;
    return node;
  }
  function renderRoster() {
    $('online').textContent = `${state.players.length} online`;
    $('count').textContent = state.players.length;
    $('round').textContent = `ROUND ${String(state.round).padStart(2, '0')}`;
    $('game-message').textContent = state.message;
    const me = state.players.find(p => p.id === state.you);
    $('welcome').hidden = !!me;
    document.querySelector('.touch-controls').hidden = !me;
    $('leave').hidden = !me;
    $('role').textContent = !me ? 'Your next great escape starts here.' : me.jailed ? 'You’re in jail. A runner can free you.' : me.role === 'police' ? 'You’re police. Get close to catch runners.' : 'You’re a runner. Stay free. Rescue your crew.';
    $('team').hidden = !me || me.role === 'police' || me.jailed || state.players.filter(p => p.role === 'police').length >= 2;
    $('team').disabled = !!me?.request;
    $('team').textContent = me?.request ? 'Waiting for police approval…' : 'Ask to join police';
    const key = JSON.stringify([state.you, state.players.map(p => [p.id, p.name, p.role, p.jailed, p.request])]);
    if (key === rosterKey) return;
    rosterKey = key;
    const list = $('players'); list.replaceChildren();
    if (!state.players.length) list.append(element('p', 'Quiet streets, for now. Be the first to enter.', 'empty'));
    state.players.forEach(p => {
      const row = element('div', '', 'player');
      row.append(element('span', p.role === 'police' ? '★' : p.name.slice(0, 1).toUpperCase(), 'avatar ' + p.role), element('span', p.name + (p.id === state.you ? ' (you)' : '')), element('span', p.jailed ? 'IN JAIL' : p.role === 'police' ? 'POLICE' : 'RUNNER', 'badge'));
      list.append(row);
    });
    $('requests').replaceChildren();
    if (me?.role === 'police') state.players.filter(p => p.request && !p.jailed).forEach(p => {
      const row = element('div', '', 'request'); row.append(element('div', `${p.name} wants a badge.`));
      ['accept', 'decline'].forEach(action => {
        const button = element('button', action === 'accept' ? 'Accept' : 'Decline');
        button.addEventListener('click', () => act({action, player_id: p.id})); row.append(button);
      });
      $('requests').append(row);
    });
  }
  async function act(data) {
    try { await api(data); $('error').textContent = ''; }
    catch (error) { $('error').textContent = error.message; }
  }
  async function join() {
    if (joining) return;
    joining = true; $('join').disabled = true;
    try {
      await api({action: 'join', name: $('name').value || 'Guest'});
      wantsToPlay = true; sessionStorage.setItem('epr-play-name', $('name').value); $('error').textContent = '';
      canvas.focus({preventScroll: true}); pendingInput = true; schedulePoll(0);
    }
    catch (error) { $('error').textContent = error.message; }
    finally { joining = false; $('join').disabled = false; }
  }
  $('join-form').addEventListener('submit', event => { event.preventDefault(); join(); });
  $('team').addEventListener('click', () => act({action: 'request'}));
  $('leave').addEventListener('click', async () => { wantsToPlay = false; clearInput(); await act({action: 'leave'}); });
  const allowed = ['w', 'a', 's', 'd', 'arrowup', 'arrowleft', 'arrowdown', 'arrowright'];
  addEventListener('keydown', event => {
    if (event.defaultPrevented || event.ctrlKey || event.altKey || event.metaKey) return;
    if (event.target.matches('input,textarea,select') || event.target.isContentEditable || !joined) return;
    const key = event.key.toLowerCase();
    if (allowed.includes(key)) { event.preventDefault(); if (!keys.has(key)) { keys.add(key); inputChanged(); } }
  });
  addEventListener('keyup', event => { if (keys.delete(event.key.toLowerCase())) inputChanged(); });
  function readInput() {
    return {dx: pointer.dx || Number(keys.has('d') || keys.has('arrowright')) - Number(keys.has('a') || keys.has('arrowleft')),
      dy: pointer.dy || Number(keys.has('s') || keys.has('arrowdown')) - Number(keys.has('w') || keys.has('arrowup'))};
  }
  function inputChanged() {
    motion.setInput(readInput(), performance.now());
    if (joined) { pendingInput = true; schedulePoll(0); }
  }
  function clearInput() { keys.clear(); pointer = {dx: 0, dy: 0}; inputChanged(); }
  addEventListener('blur', clearInput);
  document.addEventListener('visibilitychange', clearInput);
  document.querySelectorAll('.touch-controls button').forEach(button => {
    button.addEventListener('pointerdown', event => { event.preventDefault(); button.setPointerCapture(event.pointerId); pointer = {dx: Number(button.dataset.dx || 0), dy: Number(button.dataset.dy || 0)}; inputChanged(); });
    ['pointerup', 'pointercancel', 'lostpointercapture'].forEach(event => button.addEventListener(event, clearInput));
  });
  function schedulePoll(delay) { clearTimeout(pollTimer); pollTimer = setTimeout(poll, delay); }
  async function poll() {
    if (stopped) return;
    if (posting || joining) { schedulePoll(16); return; }
    if ((pendingInput || !document.hidden) && !stopped) {
      posting = true;
      pendingInput = false;
      try {
        if (joined) {
          await api({action: 'move', ...readInput()});
        } else { await api(); if (wantsToPlay && !joined) await join(); }
      } catch (error) { $('error').textContent = error.message; }
      finally { posting = false; }
    }
    schedulePoll(pendingInput ? 0 : joined ? 200 : 1000);
  }
  function exitGame() {
    wantsToPlay = false; stopped = true; clearTimeout(pollTimer);
    keys.clear(); pointer = {dx: 0, dy: 0};
    motion.setInput(pointer, performance.now());
    if (joined) {
      joined = false;
      fetch('/api/maintenance/game', {method: 'POST', keepalive: true, headers: {'Content-Type': 'application/json'}, body: JSON.stringify({action: 'leave'})}).catch(() => {});
    }
  }
  window.addEventListener('epr-game-return', exitGame);
  $('return').addEventListener('click', exitGame);
  async function availability() {
    try {
      const response = await fetch('/api/maintenance/status', {cache: 'no-store'});
      if (!response.ok) throw new Error();
      const status = await response.json();
      if (typeof status.shortcut === 'string') window.dispatchEvent(new CustomEvent('epr-game-shortcut', {detail: status.shortcut}));
      if (initialAvailability === undefined) { initialAvailability = status.available; returnOnReady = !status.available; }
      if (!status.available) returnOnReady = true;
      document.querySelector('.availability').dataset.phase = status.phase;
      $('deployment').textContent = status.available ? 'All ready. The app is available — thanks for waiting.' : status.message;
      $('return').hidden = !status.available;
      if (status.available && returnOnReady) { stopped = true; location.replace(destination()); }
    } catch { $('deployment').textContent = 'The app is restarting. Reconnecting automatically…'; }
    setTimeout(availability, 5000);
  }
  function tree(x, y, ctx) {
    ctx.fillStyle = '#879e6b'; ctx.beginPath(); ctx.arc(x + 2, y + 4, 13, 0, Math.PI * 2); ctx.fill();
    ctx.fillStyle = '#a9bc87'; ctx.beginPath(); ctx.arc(x, y, 12, 0, Math.PI * 2); ctx.fill();
    ctx.fillStyle = '#b9ca96'; ctx.beginPath(); ctx.arc(x - 3, y - 3, 7, 0, Math.PI * 2); ctx.fill();
  }
  function paintWorld() {
    const ctx = background.getContext('2d');
    ctx.fillStyle = '#dbe5c9'; ctx.fillRect(0, 0, 1000, 600);
    ctx.fillStyle = '#eae8cf'; ctx.fillRect(0, 263, 1000, 74); ctx.fillRect(462, 0, 76, 600);
    ctx.strokeStyle = '#d3d7bb'; ctx.lineWidth = 1; ctx.setLineDash([7, 8]);
    ctx.beginPath(); ctx.moveTo(0, 300); ctx.lineTo(1000, 300); ctx.moveTo(500, 0); ctx.lineTo(500, 600); ctx.stroke(); ctx.setLineDash([]);
    for (let x = 30; x < 1000; x += 47) { tree(x, 30, ctx); tree(x, 568, ctx); }
    for (const [x, y, w, h] of (state.obstacles.length ? state.obstacles : [[180,125,145,70],[675,125,145,70],[180,405,145,70],[675,405,145,70]])) {
      ctx.fillStyle = '#b2bd99'; ctx.fillRect(x + 4, y + 6, w, h);
      ctx.fillStyle = '#d1c5a2'; ctx.fillRect(x, y, w, h);
      ctx.fillStyle = '#b2a485'; ctx.fillRect(x, y, w, 8);
      ctx.fillStyle = '#e9dfbf'; for (let i = 17; i < w - 10; i += 31) ctx.fillRect(x + i, y + 25, 16, 22);
      tree(x - 28, y + 35, ctx); tree(x + w + 28, y + 35, ctx);
    }
    ctx.fillStyle = '#bab99e'; ctx.fillRect(450, 240, 100, 120);
    ctx.strokeStyle = '#87977f'; ctx.lineWidth = 3; ctx.strokeRect(450, 240, 100, 120);
    ctx.lineWidth = 2; for (let x = 461; x < 550; x += 13) { ctx.beginPath(); ctx.moveTo(x, 240); ctx.lineTo(x, 360); ctx.stroke(); }
    ctx.fillStyle = '#64765b'; ctx.font = 'bold 12px Segoe UI'; ctx.textAlign = 'center'; ctx.fillText('JAIL', 500, 226);
    ctx.font = '10px Segoe UI'; ctx.fillText('TOUCH TO RESCUE', 500, 380);
  }
  function resize() {
    const ratio = Math.min(devicePixelRatio || 1, 2);
    canvas.width = Math.round(canvas.clientWidth * ratio);
    canvas.height = Math.round(canvas.clientHeight * ratio);
    ctx.setTransform(canvas.width / 1000, 0, 0, canvas.height / 600, 0, 0);
  }
  function draw(now) {
    requestAnimationFrame(draw);
    if (document.hidden) { motion.lastFrame = now; return; }
    ctx.drawImage(background, 0, 0, 1000, 600);
    ctx.textAlign = 'center';
    for (const p of motion.frame(now)) {
      const me = p.id === state.you;
      if (me) { ctx.strokeStyle = '#fffdf1'; ctx.lineWidth = 3; ctx.beginPath(); ctx.arc(p.x, p.y, 21, 0, Math.PI * 2); ctx.stroke(); }
      ctx.fillStyle = '#283d3230'; ctx.beginPath(); ctx.ellipse(p.x + 2, p.y + 12, 16, 8, 0, 0, Math.PI * 2); ctx.fill();
      ctx.fillStyle = p.jailed ? '#8c8977' : p.role === 'police' ? '#417b9d' : '#be844a';
      ctx.beginPath(); ctx.arc(p.x, p.y, 15, 0, Math.PI * 2); ctx.fill();
      ctx.fillStyle = '#fff5dc'; ctx.font = 'bold 13px Segoe UI'; ctx.fillText(p.role === 'police' ? '★' : '•', p.x, p.y + 5);
      ctx.font = `${me ? 'bold ' : ''}11px Segoe UI`; ctx.fillStyle = '#354b36'; ctx.fillText(p.name + (me ? ' · you' : ''), p.x, p.y - 24);
    }
  }
  paintWorld(); resize(); new ResizeObserver(resize).observe(canvas);
  poll(); availability(); requestAnimationFrame(draw);
})();
