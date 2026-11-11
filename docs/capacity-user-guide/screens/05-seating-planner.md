# Seating Planner

## Purpose

The Seating Planner is the office floor layout tool for assigning employees and guests to desks, reviewing team or product distribution, and exporting the current plan.

## Navigation

- Page: `/settings/seating-planner`
- Data source: `assignee_hours_capacity.db`

## Main Areas

| Area | Purpose | Key controls |
| --- | --- | --- |
| Toolbar | Layout editing and zoom. | Add Table, Zoom Out, Zoom In, Save Layout |
| Canvas | Floor plan workspace. | Drag tables, rotate tables, assign seats, select and move occupied seats |
| Employees | Employee and guest roster. | Search, drag employee to seat, WFH flag, add/rename/delete guest |
| Mode bar | Visual overlays and exports. | Team, Product, Legend, PDF, Excel |

## PDF Export

The **PDF** action builds a print-only vector copy of the seating tables and fits the complete plan onto one A4 landscape page. Use the browser print dialog to save the output as PDF. The export excludes editing controls, sidebars, and rotate handles.

## Zoom Behavior

The canvas uses browser layout zoom where available so text and seat labels remain sharp while zooming out. A transform fallback remains for browsers without layout zoom support.

## Dependent & Impacted Files

The page is rendered by `_seating_planner_html()` in `report_server.py`. It loads
the common `app-shortcuts.js` availability integration directly because this screen
does not load the shared sidebar. Operational behavior is described in
`MAINTENANCE_GAME.md`; no new visible planner controls or data fields are added.

## Data Flow

While the page is open, the common availability listener reads the deployment
status. A paused app moves the visitor to the intermission with this page as the
return destination. After deployment, the page reloads under normal access checks.
Existing seating-plan database reads, saves and exports retain their normal flow.
