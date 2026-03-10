# Expected files and why the app may look broken

Some assets are **expected by the report server but may be missing** after a fresh clone. That can make the report index look plain, the "Dashboard" link return 404, or **icons show as empty boxes** in the sidebar and reports.

## 1. `dashboard.html`

- **Expected:** The nav and root URL expect a report at `report_html/dashboard.html` (or `dashboard.html` in the project root, which sync copies into `report_html`).
- **If missing:** The "Dashboard" link leads to 404, and the server suggests opening `/report_html/` instead of `/dashboard.html`.
- **Fix:** Copy the template into the reports folder:
  ```powershell
  copy "D:\JIRA SCRIPT\dashboard_template.html" "D:\JIRA SCRIPT\report_html\dashboard.html"
  ```
  Restart the server and open http://127.0.0.1:3000/ — you should be redirected to the dashboard.

## 2. Material Symbols icon font (`report_html/fonts/material-symbols-outlined.woff2`)

- **Expected:** `report_html/material-symbols.css` uses a self-hosted font at `report_html/fonts/material-symbols-outlined.woff2` so icons (sidebar, buttons, etc.) render without calling Google.
- **If missing:** The font request returns 404 and **icons appear as empty squares or missing glyphs**; the rest of the page can look unstyled or “wrong.”
- **Fix (option A – recommended):** The repo’s `material-symbols.css` includes a **CDN fallback**. If the local font is missing, the browser loads the font from Google’s CDN and icons should still display. No extra steps unless you want to self-host.
- **Fix (option B – self-host):** Create the fonts folder and add the woff2 file:
  1. Create folder: `report_html\fonts`
  2. Download the [Material Symbols Outlined variable font](https://fonts.google.com/icons?selected=Material+Symbols+Outlined) (e.g. from [Google Fonts](https://fonts.google.com/icons) or [Fontsource](https://fontsource.org/fonts/material-symbols-outlined)) and save as `material-symbols-outlined.woff2` in that folder.

After adding or fixing these, restart the report server and refresh the browser.
