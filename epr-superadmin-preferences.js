/* This panel is created only for the verified bootstrap owner. */
(() => {
  'use strict';
  async function init() {
    try {
      const identity = await fetch('/api/auth/me', {cache: 'no-store'});
      if (!identity.ok) return;
      const user = (await identity.json()).user;
      if (!user?.is_superadmin) return;
      const response = await fetch('/api/auth/game-shortcut', {cache: 'no-store'});
      if (!response.ok) return;
      const settings = await response.json();
      const panel = document.createElement('section');
      panel.className = 'panel';
      panel.style.marginTop = '24px';
      panel.setAttribute('aria-label', 'Superadmin preferences');
      panel.innerHTML = '<div class="panel-head"><div><h2>Private playground shortcut</h2><p class="sub">Superadmin only. The shortcut is never displayed in other users’ menus or help.</p></div></div><form class="toolbar"><label for="private-game-shortcut">Keyboard shortcut</label><input id="private-game-shortcut" maxlength="32" placeholder="Ctrl+G" required autocomplete="off"><button class="action" type="submit">Save shortcut</button><p class="sub" style="flex-basis:100%">Use Ctrl or Alt, optional Shift, and one letter or digit. Example: Ctrl+Shift+G. Changes reach open pages within a minute.</p><span role="status" class="shortcut-status"></span></form>';
      document.querySelector('main').append(panel);
      const input = panel.querySelector('input'), status = panel.querySelector('[role=status]'), button = panel.querySelector('button');
      input.value = settings.shortcut;
      panel.querySelector('form').addEventListener('submit', async event => {
        event.preventDefault(); button.disabled = true; status.textContent = 'Saving…';
        try {
          const saved = await fetch('/api/auth/game-shortcut', {method: 'PUT', headers: {'Content-Type': 'application/json', 'X-EPR-CSRF': user.csrf_token}, body: JSON.stringify({shortcut: input.value})});
          const body = await saved.json();
          if (!saved.ok) throw new Error(body.error || 'Could not save shortcut.');
          input.value = body.shortcut;
          status.textContent = 'Saved';
          window.dispatchEvent(new Event('epr-private-preference-saved'));
        } catch (error) { status.textContent = error.message; }
        finally { button.disabled = false; }
      });
    } catch (_) { /* Existing access controls remain usable if preferences fail to load. */ }
  }
  init();
})();
