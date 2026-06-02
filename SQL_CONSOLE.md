# SQL Console (`/settings/sql-console`)

Admin/developer tool for browsing and querying the read-only SQLite databases
that power the EPR Reporting Tool. The page is served by
`_sql_console_settings_html()` in `report_server.py` and is backed by a small set
of read-only API endpoints under `/api/admin/sql-console/`.

---

## Business Logic

- The console exposes **three** databases, all opened **read-only** (SQLite
  `?mode=ro` URI via `_sqlite_readonly_uri`):
  - `canonical` → `assignee_hours_capacity.db` (authoritative Jira snapshot +
    planner store)
  - `exports` → `jira_exports.db` (flattened export tables)
  - `support_center` → the Support Center DB (`resolve_support_center_db_path()`)
- The active database is chosen in the **Databases** column of the schema
  browser and drives everything else on the page (the query target, the
  IntelliSense schema, and the table list). It is fixed to those three keys
  (`_sql_console_target_path`); any other key raises a `ValueError`.
- **Natural-language → SQL workspace:**
  1. The user types a question in plain English. **Generate SQL**
     (`POST /api/admin/sql-console/generate`) sends the prompt plus the selected
     database's full schema (every `table(column_name column_type, ...)` line),
     a strict rule forbidding invented column names, and curated inter-table join
     hints to an OpenAI model so the generated query is schema-accurate.
  2. When the prompt contains a back-tick quoted phrase (e.g.
     `` `Holidays Management …` ``) the endpoint resolves it against
     `oeh_epics` and appends **"Resolved entity hints from the database"**
     (`epic_key=…, summary=…`) to the model input so the model uses the exact
     key/value.
  3. The model output is stripped of markdown fences and validated through the
     same `_normalize_sql_console_query` guard as `/execute` (single read-only
     statement). Invalid SQL → HTTP 400; an empty model response → HTTP 502; an
     OpenAI request failure → HTTP 502.
  4. The validated SQL lands in an editable **CodeMirror** editor (syntax
     highlighting + schema-aware IntelliSense). The user can edit it and **Run
     Query** (`/execute`) — results render in a table below — or **Download
     Excel** (`/export`).
  - Model selection: `OPENAI_SQL_CONSOLE_MODEL` env var, default
    `gpt-4.1-mini`. Requires `OPENAI_API_KEY` (missing key → HTTP 400, no
    outbound request). Endpoint: `https://api.openai.com/v1/responses`.
- **Editor IntelliSense:** the editor is a CodeMirror 5 instance
  (`text/x-sqlite` mode) loaded from jsDelivr with the `sql-hint` and
  `show-hint` addons. The hint schema is rebuilt from the schema endpoint each
  time the active database changes, so completions include SQL keywords/operators
  **and** the live table and column names. Suggestions pop up while typing
  (`inputRead`) and on `Ctrl-Space`. If CodeMirror fails to load (offline), the
  editor degrades gracefully to a plain `<textarea>`.
- **Browse experience (three-column layout):**
  1. **Databases** column — static list from `SQL_CONSOLE_DATABASE_INFO`
     (label, file name, description).
  2. **Tables** column — populated from `GET /api/admin/sql-console/schema`,
     which lists every non-`sqlite_%` table with its columns and a **brief
     description**. Descriptions come from the curated
     `SQL_CONSOLE_TABLE_DESCRIPTIONS` map; tables not in the map fall back to
     `"{n} columns."`.
  3. **Table preview** column — populated from
     `GET /api/admin/sql-console/table-preview`, which returns the table's
     columns and the **first `SQL_CONSOLE_PREVIEW_ROWS` (10) rows** plus the
     total row count.
- **Safety rules:**
  - The preview endpoint validates the requested table name against
    `sqlite_master` before interpolating it (quoted with doubled `"`), so
    unknown names and injection attempts (`name; DROP ...`) are rejected with
    HTTP 400.
  - The ad-hoc query endpoints (`/execute`, `/export`) only accept a **single**
    read-only statement (`SELECT`, `WITH`, `EXPLAIN`, or a whitelisted schema
    `PRAGMA`) via `_normalize_sql_console_query`; write keywords, multiple
    statements, and `ATTACH`/`DETACH` are rejected.
  - `/execute` truncates at `SQL_CONSOLE_MAX_ROWS` (500); `/export` caps at
    `SQL_CONSOLE_EXPORT_MAX_ROWS` (20000).

## Business Cases

- Operators and developers need to inspect the live report databases without a
  desktop SQLite tool or shell access to the Azure App Service host.
- The browse flow (database → table → first rows) answers the most common
  question — *"what does this table actually contain?"* — in two clicks, with
  no SQL required.
- The query/export endpoints remain available for ad-hoc read-only analysis and
  Excel extraction when the browse preview is not enough.

## Examples

- Selecting **Canonical DB → `canonical_issues`** shows the curated description
  *"One row per Jira issue in the latest sync run…"* and a 10-row preview with
  the column types (e.g. `issue_key TEXT`, `original_estimate_hours REAL`) and a
  pill reading `Total rows: 12,480`.
- Selecting **Exports DB → `work_items`** previews the first 10 work-item rows.
- Preview request/response:
  - `GET /api/admin/sql-console/table-preview?database=canonical&table=canonical_worklogs`
  - → `{ "ok": true, "columns": [...], "rows": [...10 rows...],
        "total_row_count": 95231, "preview_limit": 10 }`

## Explanations

When the page loads it renders the three fixed databases in the left column and
auto-selects the first one. Picking a database asks the server for its table
list and shows each table with a one-line description in the middle column.
Clicking a table asks the server for that table's first ten rows and displays
them as a spreadsheet — column names with their SQLite type in the header, one
row per record, with `NULL` cells shown in muted italics. A header pill shows
how many rows the table holds in total versus how many are previewed. Nothing
the user does can modify any database; every connection is opened read-only.

## Front-end UI Fields

| Field / control | Type | Behavior |
|---|---|---|
| **Target DB** pill | Label | Reflects the active database selected in the Databases column; the query/generate target. |
| **Ask in plain English** | Textarea | Natural-language question. Feeds `Generate SQL`. Back-tick quoted phrases trigger entity-hint resolution against `oeh_epics`. |
| **Generate SQL** | Button | Calls `/generate`; writes the validated SQL into the editor. Disabled while generating. |
| **SQL query (editable)** | CodeMirror editor | Syntax-highlighted, schema-aware IntelliSense (keywords + live table/column names). `Ctrl-Space` or typing shows suggestions. Falls back to a plain textarea if CodeMirror is unavailable. |
| **Run Query** | Button | Calls `/execute`; renders the result table below with `Elapsed`, `Rows`, `Truncated` pills. |
| **Download Excel** | Button | Calls `/export`; downloads the result as `.xlsx`. |
| **Clear** | Button | Empties the editor and results. |
| **Databases** list | Buttons | One per database (`Canonical DB`, `Exports DB`, `Support Center DB`). Shows label, description, file name. Click selects, loads its tables, and refreshes the editor IntelliSense schema + Target DB. First item auto-selected on load. |
| **Tables** list | Buttons | One per table in the selected database. Shows table name, brief description, and column count. Click loads the preview. |
| **Table preview** | Read-only grid | Header row = column name + SQLite type; body = first 10 rows. Pills show `Total rows` and `Columns`. Sub-line shows `Showing first N of M rows.` |
| **Load SELECT * into editor** | Button | Inserts a `SELECT * FROM "<table>" LIMIT 100;` starter query for the previewed table into the editor. |
| Status lines | Text | Inline loading / error messages for generation, query run, tables, and preview. |

## Script Files

- `report_server.py`
  - `_sql_console_settings_html()` — generates the page (query workspace +
    CodeMirror editor + three-column schema browser; HTML/CSS/JS).
  - `SQL_CONSOLE_DATABASE_INFO`, `SQL_CONSOLE_TABLE_DESCRIPTIONS`,
    `SQL_CONSOLE_PREVIEW_ROWS`, `SQL_CONSOLE_MAX_ROWS`,
    `SQL_CONSOLE_EXPORT_MAX_ROWS`, `SQL_CONSOLE_DEFAULT_OPENAI_MODEL`,
    `SQL_CONSOLE_OPENAI_URL` — constants/metadata.
  - `_sql_console_target_path`, `_sqlite_readonly_uri`,
    `_sql_console_open_connection`, `_sql_console_resolve_target_or_error`,
    `_sql_console_run_query`, `_normalize_sql_console_query` — resolution and
    safety helpers.
  - `_sql_console_schema_brief` — now includes column types (`name TYPE`) in
    addition to names so the model sees precise column definitions.
    `_sql_console_extract_backtick_phrases`,
    `_sql_console_resolve_entity_hints`, `_sql_console_build_generation_input`,
    `_sql_console_extract_model_sql` — NL→SQL generation helpers.
  - `_SQL_CONSOLE_JOIN_HINTS` — module-level constant listing key inter-table
    relationships (e.g. `oeh_subtasks.subtask_key = canonical_worklogs.issue_key`)
    injected into every generation prompt so the model knows how to join for
    worklog hours instead of inventing non-existent columns.
  - Routes: `GET /settings/sql-console`,
    `GET /api/admin/sql-console/schema`,
    `GET /api/admin/sql-console/table-preview`,
    `POST /api/admin/sql-console/execute`,
    `POST /api/admin/sql-console/export`,
    `POST /api/admin/sql-console/generate`.
  - Module import: `requests` (used by `/generate` to call OpenAI).
- **CodeMirror 5** (jsDelivr CDN): `lib/codemirror`, `mode/sql`,
  `addon/edit/matchbrackets`, `addon/hint/show-hint`, `addon/hint/sql-hint` —
  editor syntax highlighting + schema-aware IntelliSense.
- `shared-nav.js` — registers the SQL Console settings nav entry.
- `tests/test_admin_sql_console_api.py` — API tests (schema, preview, execute,
  export, generate, validation).
- **Environment:** `OPENAI_API_KEY` (required for `/generate`),
  `OPENAI_SQL_CONSOLE_MODEL` (optional model override, default `gpt-4.1-mini`).

## Dependent & Impacted Files

- `shared-nav.js` / `report_html/shared-nav.js` — surface the SQL Console link
  in the settings navigation.
- `report_server.py` settings nav metadata (`SQL_CONSOLE_SETTINGS_ROUTE`,
  `_settings_top_nav_html`) — the page embeds the shared settings top nav.
- The console is **read-only** and writes to no table, so no downstream report
  module depends on its output; it only *reads* the same databases that the
  capacity/exports/support-center modules produce.

## Table Schema

The SQL Console does **not** own any table. It reads whatever exists in the
three target databases. Notable tables it surfaces (with curated descriptions):

| Database | Table | Meaning |
|---|---|---|
| canonical | `canonical_issues` | One row per Jira issue in the latest sync run. |
| canonical | `canonical_worklogs` | Worklog entries (hours per issue per date). |
| canonical | `canonical_refresh_runs` | Sync run history (status, scope, timestamps). |
| canonical | `canonical_refresh_state` | Pointer to last successful run id. |
| canonical | `oeh_epics` | Epics Planner epics (keys, dates, estimates). |
| exports | `work_items` | Flattened work-item export rows. |
| exports | `subtask_worklogs` | Per-subtask worklog totals. |
| support_center | `support_issues` | Support-tagged issues for the Support Center report. |

Column lists/types are read live via `PRAGMA table_info(...)`.

## Data Flow

1. **Page load** → browser GETs `/settings/sql-console`
   (`_sql_console_settings_html()` renders the static three-column shell and
   embeds `SQL_CONSOLE_DATABASE_INFO` + `SQL_CONSOLE_PREVIEW_ROWS` as JSON).
2. **Select database** → JS calls `GET /api/admin/sql-console/schema?database=<key>`
   → `sql_console_schema()` opens the resolved DB read-only, lists tables via
   `sqlite_master`, reads columns via `PRAGMA table_info`, attaches a
   description, and returns JSON.
3. **Select table** → JS calls
   `GET /api/admin/sql-console/table-preview?database=<key>&table=<name>`
   → `sql_console_table_preview()` validates the table name against
   `sqlite_master`, counts rows, selects the first
   `SQL_CONSOLE_PREVIEW_ROWS` rows, and returns columns + rows + total count.
4. **Generate SQL** → JS POSTs `{{database, prompt}}` to `/generate` →
   `sql_console_generate()` validates prompt + `OPENAI_API_KEY`, builds the
   schema brief (`_sql_console_schema_brief`) and entity hints
   (`_sql_console_resolve_entity_hints`), calls the OpenAI Responses API via
   `requests.post`, extracts the SQL (`_sql_console_extract_model_sql`),
   validates it with `_normalize_sql_console_query`, and returns the SQL. JS
   loads it into the CodeMirror editor.
5. **Run query / export** → `POST /execute` / `POST /export` run a single
   validated read-only statement through `_sql_console_run_query`, returning
   JSON rows (rendered in the results table) or an `.xlsx` download.

The only outbound side effect is the OpenAI request in `/generate`; no database
is ever written — every connection is opened read-only.

---

> **Schema note:** This change added no SQLite schema changes — the console only
> reads existing databases — so the
> `DATABASE_SCHEMA_MIGRATION_PROTOCOL.md` workflow does not apply.
