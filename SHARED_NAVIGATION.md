# Shared navigation (sidebar)

`shared-nav.js` and `shared-nav.css` inject the EPR sidebar into every report and settings page. One script builds the menu from the Page Categories configuration, the signed-in identity and the current URL.

## Business Logic

- **One column.** Every section is in a single scrolling column 264px wide. Admin Settings no longer opens a second 290px panel. It is a group in the same column, so the sidebar width and the page's left padding never change when a group opens.
- **Groups (nested menus).** Report categories from `/api/page-categories` (`navigation.reports.categories`) appear as collapsible groups. Admin Settings is a group, and the admin categories (`navigation.admin_settings.categories`) are nested groups inside it. A group header shows an icon, the label, the total page count (including nested groups) and a chevron. Groups without pages are skipped.
- **Which groups start open.**
  1. Any group that contains the current page always opens, and so do its parents.
  2. Otherwise the open/closed state saved in `localStorage["unified-report-nav-groups"]` is used.
  3. If nothing is saved, the first report category opens.
- **Active page.** A link is active when its `file` matches the current file name or its `path` prefixes the current path. Active links get `aria-current="page"`, a filled icon and a left accent bar. A closed group that contains the active page is tinted.
- **Find a page.** Matching is a case-insensitive substring match on the page title. Groups with matches are forced open and empty groups are hidden. If nothing matches, the menu shows "No pages match." Enter opens the first match and Escape clears the search. The search does not change the saved group state.
- **Collapse rail (desktop ≥960px).** The toggle switches between 264px and a 68px icon rail. The choice is stored in `localStorage["unified-report-nav-collapsed"]`. In the rail, clicking a group icon expands the sidebar and opens that group. Labels appear as native tooltips.
- **Mobile (<960px).** The sidebar is a drawer (`min(300px, 86vw)`) opened by the floating menu button. It closes with the × button, a click on the scrim, or Escape. Focus returns to the menu button.
- **Account footer.** When EPR auth is enabled, the bottom of the sidebar shows the avatar, name and role, plus two buttons on one row: **My availability** and **Sign out**. Sign out posts `/api/auth/logout` with `X-EPR-CSRF`. While the request runs, the button reads "Signing out…" and is disabled. On failure it reads "Sign out failed — retry". The footer stays visible in the rail as icon buttons.
- **My Day pin.** A **My Day** link (`page_key: my_day`, icon `wb_sunny`, class `is-pinned`) is always rendered first, above the report items, whatever the Page Categories layout says. Duplicate catalog entries with the same key are skipped. When active it uses the warm yellow `#ffc23d` highlight instead of the usual accent.
- **Usage Analytics.** If an Admin Settings section exists, a **Usage Analytics** item (`/settings/usage-analytics`, icon `insights`) is appended once.
- **Role filtering.** Members do not get the Admin Settings group. Admins also get **Users & access**. The server still enforces access, so the menu is only a convenience.
- If `localStorage` is blocked, the menu still works without saved state.
- Icons have a fixed 1em box, so labels do not wrap while the icon font loads.
- Reduced-motion users get no transitions.

## Business Cases

Admins and leads move between about 40 reports and settings screens every day. The old two-column layout used 540px when Admin Settings was open. On laptops and phones it covered the report. Long menus had no search, and Sign out wrapped onto two lines inside a 34px column. This design keeps the menu narrow, lets users jump to a page by name, and keeps account actions clear.

## Examples

- On `/settings/epics-management`, Admin Settings opens and **Epics Planner** is highlighted. Other groups keep their saved state.
- Typing `planner` shows Epics Planner, Epics Planner Import, IPP Meeting Planner, Team Capacity Planner and Seating Planner. Every other link is hidden.
- With the menu collapsed, clicking the Admin Settings icon expands the sidebar and scrolls the open group into view.
- On a 375px phone, tapping the menu button slides in the drawer. Pressing Escape closes it.

## Explanations

The sidebar reads like a file tree. Top-level reports come first, then categories such as Leadership, Department and Team that open and close in place, then Admin Settings with its own sub-groups. The search box at the top narrows the tree as you type. Your name and the Sign out button are always at the bottom.

## Front-end UI Fields

| Element | Type | Default | Behavior |
|---|---|---|---|
| Collapse toggle (`.unified-nav-toggle`) | button | expanded | Desktop: switches between the full sidebar and the rail and saves the choice. Mobile: closes the drawer. |
| Find a page (`.unified-nav-search input`) | search text | empty | Filters links by title. Escape clears it and Enter opens the first match. |
| Page link (`.unified-nav-main-link`) | link | — | Navigates to the page. The active page has `aria-current="page"`. |
| Group header (`.unified-nav-group-toggle`) | button with `aria-expanded` | see Business Logic | Opens or closes the group. The badge shows its page count. |
| Nested group (`.unified-nav-group.is-nested`) | group | closed | An admin category inside Admin Settings. |
| My Day (`.unified-nav-main-link.is-pinned`) | link | always first | Opens `/my_day.html`, the personal home page. |
| My availability (`.epr-nav-action`) | link | — | Opens `/my-availability`. |
| Sign out (`.epr-nav-signout`) | button | — | Ends the EPR session and goes to `/login`. |
| Mobile menu (`.unified-nav-mobile-btn`) | button | hidden on desktop | Opens the drawer. |

## Script Files

- `shared-nav.js`: the canonical script that builds the menu, groups, search, collapse, mobile drawer and account footer.
- `shared-nav.css`: the canonical styles.
- `report_html/shared-nav.js` and `report_html/shared-nav.css`: promoted copies written by `sync_report_html()` in `report_server.py`.
- `tests/test_admin_navigation.py`: checks the script contract and that pages include the assets.
- `tests/test_epr_auth_browser.cjs`: browser checks for identity, Sign out, the admin group, search, the rail and the mobile drawer.

## Dependent & Impacted Files

- Every report and settings HTML page that loads `/shared-nav.js` and `/shared-nav.css`. These include the root report sources, the generator templates and the inline settings pages in `report_server.py`.
- The Page Categories API and settings screen (`/api/page-categories`, `/settings/page-categories`) supply the categories and titles.
- `my_day.html` (`docs/report-user-guide/screens/17-my-day.md`) and `usage_analytics.html` (`docs/capacity-user-guide/screens/09-usage-analytics.md`) are the targets of the pinned My Day link and the Usage Analytics admin item.
- `epr_auth.py` (`/api/auth/me`, `/api/auth/logout`) supplies the identity, role and CSRF token. See `EPR_AUTHENTICATION.md`.
- Pages that depended on the removed `unified-nav-secondary-open` body class or `--unified-nav-secondary-width` must use `--unified-nav-width-main` instead. No current page does.

## Table Schema

The sidebar reads no SQLite tables directly. It reads the Page Categories tables through `/api/page-categories`. That API is owned by the Page Categories settings module.

## Data Flow

1. The page loads `shared-nav.js`.
2. `resolveNavSections()` builds the static sections, then fetches `/api/page-categories` and applies the catalog titles and the category layout.
3. `/api/auth/me` returns the identity. Anonymous users go to `/login?next=<page>` and pending users to `/access-pending`. Members lose the admin section.
4. The Usage Analytics admin item is ensured and the pinned My Day item is prepared.
5. `initUnifiedNav()` renders the header, search, groups and account footer, then restores the state saved in localStorage.
5. When a user clicks a link, the browser navigates. When a user clicks Sign out, the script posts `/api/auth/logout` and then goes to `/login`.
