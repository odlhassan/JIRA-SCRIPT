---
name: localhost-ready-after-code
description: After completing code changes, ensure server, generated HTML, and assets are updated so that restarting the local server shows the latest changes on localhost. Use when changing report logic, server APIs, generators, or shared assets so the user can simply restart the server and see new behavior.
---

# Localhost Ready After Code

Use this skill when a task changes anything that can affect what localhost serves: `report_server.py`, `run_server.py`, generator scripts, root-level report HTML, templates, or shared assets.

## Workflow

1. Identify the canonical source for each changed output.
   - Root HTML such as `introduction.html` or `approved_vs_planned_hours_report.html`
   - Templates such as `dashboard_template.html` or `ipp_meeting_dashboard_template.html`
   - Generator scripts such as `generate_*.py` or `fetch_jira_dashboard.py`
   - Shared assets such as `shared-nav.js`, `shared-nav.css`, `shared-date-filter.js`, and `report_html/material-symbols.css`
2. Decide whether the change needs only a server restart or also a rebuild.
   - Server or API code changed: restart is required.
   - Generator or template changed: rerun the narrowest generator or use `python run_html_only.py --no-server` when multiple generated outputs must be refreshed together.
   - Shared asset changed: edit the canonical source, then rely on startup sync.
3. If a report name or source path changed, update `_resolve_report_html_sources()` in `report_server.py`.
4. Run feasible targeted checks before handoff.
5. Prepare the final localhost instructions with the exact command, URL, and expected visible behavior.

## Exit Criteria

- Restarting `python run_server.py` is enough for the user to see the latest behavior.
- Any required generator output has been refreshed or the reason it could not be refreshed is stated.
- The final response includes:
  - exact commands executed
  - exact additional commands the user can run
  - exact localhost URL or endpoint
  - focused manual steps tied to the latest change

## Repo Defaults

- Normal startup: `python run_server.py`
- Default landing page: `http://127.0.0.1:3000/introduction.html`
- HTML-only rebuild path: `python run_html_only.py --no-server`
- If the user asks to **push to GitHub** for production, use **`main`** per `.cursor/rules/workspace-context.mdc` (`git push origin main` triggers Azure deploy).

## Changelog

- `2026-05-12`: noted default Git push branch for production (`main`).
- `2026-04-30`: converted the skill to a concise trigger-workflow-exit contract with repo-specific commands.
