(function () {
  if (window.__eprAppShortcutsLoaded) return;
  window.__eprAppShortcutsLoaded = true;
  // Move already-open pages to the playground before the package changes.
  if (window.location.protocol !== 'file:') {
    const inPlayground = document.body?.hasAttribute('data-epr-playground') === true;
    let eprGameShortcut = inPlayground ? 'Ctrl+G' : null;
    let returningToApp = false;
    try { eprGameShortcut = sessionStorage.getItem('epr-private-game-shortcut') || eprGameShortcut; } catch (_) {}
    function rememberShortcut(value) {
      if (typeof value !== 'string') return;
      eprGameShortcut = value;
      try { sessionStorage.setItem('epr-private-game-shortcut', value); } catch (_) {}
    }
    window.addEventListener('epr-game-shortcut', event => rememberShortcut(event.detail));
    async function refreshGameShortcut() {
      if (document.hidden) return;
      try {
        const response = await fetch(inPlayground ? '/api/maintenance/status' : '/api/game-shortcut', { cache: 'no-store' });
        if (response.ok) {
          const value = (await response.json()).shortcut;
          rememberShortcut(value);
        }
      } catch (_) { /* Keep the last saved preference during reconnects. */ }
    }
    // Deliberately no menu item, tooltip, help text or displayed shortcut.
    document.addEventListener('keydown', async function (event) {
      if (!eprGameShortcut || event.isComposing || event.metaKey) return;
      const target = event.target;
      if (!inPlayground && target && (target.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName))) return;
      const chord = [event.ctrlKey ? 'Ctrl' : '', event.altKey ? 'Alt' : '', event.shiftKey ? 'Shift' : '', event.key.toUpperCase()].filter(Boolean).join('+');
      if (chord !== eprGameShortcut) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      if (event.repeat) return;
      if (inPlayground) {
        if (returningToApp) return;
        returningToApp = true;
        try {
          // Check current availability, rather than reopening during a deployment
          // that started since the game's last notice poll.
          const response = await fetch('/api/maintenance/status', {cache: 'no-store'});
          if (!response.ok || (await response.json()).available !== true) return;
          const href = document.getElementById('return')?.getAttribute('href') || '/';
          const destination = new URL(href, location.origin);
          if (destination.origin !== location.origin || destination.pathname.startsWith('/api/') || destination.pathname.startsWith('/maintenance')) return;
          window.dispatchEvent(new Event('epr-game-return'));
          window.location.assign(destination.href);
        } catch (_) { /* Keep the playground open if availability cannot be verified. */ }
        finally { returningToApp = false; }
        return;
      }
      const back = window.location.pathname + window.location.search + window.location.hash;
      window.location.assign('/maintenance?return=' + encodeURIComponent(back));
    }, true);
    refreshGameShortcut();
    window.setInterval(refreshGameShortcut, 60000);
    window.addEventListener('epr-private-preference-saved', refreshGameShortcut);
    document.addEventListener('visibilitychange', function () { if (!document.hidden) refreshGameShortcut(); });
    async function checkAvailability() {
      if (document.hidden || inPlayground) return;
      try {
        const response = await fetch('/api/maintenance/status', { cache: 'no-store' });
        if (!response.ok) return;
        const status = await response.json();
        if (status.available === false) {
          const back = window.location.pathname + window.location.search + window.location.hash;
          window.location.replace('/maintenance?return=' + encodeURIComponent(back));
        }
      } catch (_) { /* The next poll retries a container restart. */ }
    }
    checkAvailability();
    window.setInterval(checkAvailability, 10000);
  }

})();
