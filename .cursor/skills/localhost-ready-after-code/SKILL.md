---
name: localhost-ready-after-code
description: After completing code changes, ensure server, generated HTML, and assets are updated so that restarting the local server shows the latest changes on localhost. Use when changing report logic, server APIs, generators, or shared assets so the user can simply restart the server and see new behavior.
---

# Localhost-ready after code completion

The user expects **restart server → see latest changes** on localhost. After any code completion that affects reports or the server, ensure all necessary files are updated so that a naive "restart the server" delivers the new experience. The most common failure is **localhost still serving old files or old code** after core logic updates.

## Why localhost shows old content

- **Server code** (e.g. `report_server.py`): New code runs only after the process is restarted.
- **Generated HTML**: Reports are produced by generator scripts (e.g. `generate_*.py`, `fetch_jira_dashboard.py`). Output lives in **source** paths (project root or env); `sync_report_html()` copies from those sources into `report_html/`, which the server serves. If you change generator logic but never re-run the generator, the **source** file stays stale and sync keeps serving old HTML.
- **Shared assets**: `shared-nav.js`, `shared-nav.css`, `shared-date-filter.js`, `material-symbols.css` are synced from project root or `report_html/` into the report folder. Edits must be in the **source** location (see `_sync_report_html_assets` in `report_server.py`), not only in a copy.
- **Wrong source**: If a generator or config is changed to write somewhere new, `_resolve_report_html_sources` in `report_server.py` must still map that report name to the correct path so sync can copy it.

## Workflow after code completion

1. **Server-side / API changes** (`report_server.py`, `run_server.py`, or any Python the server imports)  
   - No file copy needed. User must **restart** the server for new code to run.  
   - In "How to test locally", state: restart the server (e.g. `python run_server.py`) and verify the change on localhost.

2. **Generator script changes** (e.g. `generate_assignee_hours_report.py`, `generate_rnd_data_story.py`, `fetch_jira_dashboard.py`)  
   - These scripts write HTML (and sometimes data) to **source** paths (project root or paths from env). The server serves from `report_html/` after `sync_report_html()` copies from those sources.  
   - **Re-run the affected generator(s)** so the source file is up to date. Then the next server start (which runs `sync_report_html` by default) will copy the new file into `report_html/`.  
   - If you added a new report or changed where a generator writes, ensure `_resolve_report_html_sources` in `report_server.py` includes the correct source path for that report name.

3. **HTML template or inline HTML in code**  
   - If the canonical source is a file (e.g. `dashboard_template.html`, or a template in `report_html/` that a generator reads), edit that source. Then re-run the generator that produces the served HTML so the output file is updated.  
   - If the server serves a file directly from `report_html/` and that file is also the sync source, edit it and ensure no generator overwrites it (or run the generator so it writes the updated content).

4. **Shared assets** (`shared-nav.js`, `shared-nav.css`, `shared-date-filter.js`, `material-symbols.css`)  
   - Prefer editing the copy at **project root** (or the one under `report_html/` that sync uses as source). On next `sync_report_html`, the updated asset is copied into the report folder.  
   - In "How to test locally", say: restart the server (sync runs on start) and hard-refresh the report page if needed.

5. **New report or new output path**  
   - Add the report name and its **source path** to `_resolve_report_html_sources` in `report_server.py`.  
   - Ensure the generator writes to that path (or set the right env var). Re-run the generator so the file exists; then sync will copy it when the server starts.

## Default server startup

- `python run_server.py` (or `run_all.py` with serve) runs `sync_report_html(base_dir, report_html_dir)` then starts the Flask app. So after you’ve updated **sources** (generator output, root-level assets), a single **restart** is enough for the user to see new content.

## Summary checklist

After code completion that affects reports or the server:

- [ ] **Server logic changed** → Document "restart server" in How to test locally; no extra file updates needed for code path.
- [ ] **Generator logic or template changed** → Re-run the generator so the **source** HTML/file is current; sync on next server start will serve it.
- [ ] **Shared asset changed** → Edit the source (root or report_html as used by sync); sync on start will update the served copy.
- [ ] **New report or new output path** → Register it in `_resolve_report_html_sources` and ensure the generator writes there; run generator once so sync can copy it.

Result: the user can **restart the server** and experience the latest changes on localhost without hunting for stale files or hidden caches.
