---
name: update-docs-with-code
description: After completing or changing code, update relevant .md documentation (README, user guides, column specs, runbooks, module docs) so docs stay in sync. Mandatory after every code task. Use when finishing code changes, adding features, changing exports, or when the user asks to keep docs updated.
---

# Update Docs With Code

Use this skill after **every** code change, without exception. A task is not complete
until the primary module doc and all linked/dependent module docs are updated.

## Workflow

### Step 1 — Identify all affected modules

Map the changed files to modules:
- Which module owns the changed file(s)?
- Which other modules **read from**, **write to**, or **share DB tables** with that module?
- Which modules **render UI** that displays the changed module's output?
- Which modules **import** functions or constants from the changed module?

### Step 2 — Update the primary module's `.md` file

Locate (or create) the `.md` file for the primary module. Ensure it contains **all** of
the following sections, updated to reflect the current code:

| Section | What to include |
|---|---|
| **Business Logic** | Rules, calculations, conditions, edge cases. Enough detail to re-implement from the doc alone. |
| **Business Cases** | Why the module exists. What real-world problem it solves. Who uses the output and when. |
| **Examples** | Concrete input → output examples for key logic paths. Use realistic Jira data, dates, numbers. |
| **Explanations** | Plain-English end-to-end walkthrough. Written for a business analyst, not a developer. |
| **Front-end UI Fields** | Every visible field, filter, toggle, column, button, label: name, type, default, valid range, behavior, worked example. |
| **Script Files** | Every `.py`, `.js`, `.html`, `.sql`, config file that belongs to the module and each file's role. |
| **Dependent & Impacted Files** | Files in *other* modules that read from, write to, or are affected by this module. |
| **Table Schema** | For every SQLite table the module reads or writes: table name, columns, types, constraints, meaning. |
| **Data Flow** | Step-by-step: data source → transformation → output. Which functions/routes handle each step. |

### Step 3 — Update linked/dependent module docs

For every linked module identified in Step 1:
- Update its **Dependent & Impacted Files** section.
- Update its **Data Flow** section if the relationship changed.
- If the linked module has no doc file, create a stub with all sections from Step 2
  populated at a minimum level.

### Step 4 — Update root operational docs

Check these root-level docs and update only the sections the change affects:
- `README.md` — if usage, routes, or setup changed
- `EXPECTED_FILES.md` — if files were added or removed
- `DASHBOARD_REFRESH_CLI.md` — if CLI workflow changed
- `GENERATED_EXPORTS_COLUMNS.md` — if export columns changed
- `RLT_LEAVE_REPORT.md`, `ASSIGNEE_HOURS_CAPACITY.md`, `IPP_PHASE_TRANSFORM_LOGIC.md`,
  `NESTED_VIEW_SCORECARD_FORMULAS.md`, `INCREMENTAL_SYNC.md`, `AZURE_APP_SERVICE.md`
  — if their subject module was touched
- `docs/report-user-guide/` and `docs/capacity-user-guide/` — if report UI changed
- `handover/**/*.md` — if handover-relevant behavior changed

### Step 5 — Report updated docs

In the final response, list every `.md` file that was created or updated.

## Doc file locations

| Module type | Doc location |
|---|---|
| Report module | `docs/report-user-guide/screens/<NN>-<module-name>.md` |
| Capacity/settings module | `docs/capacity-user-guide/screens/<NN>-<module-name>.md` |
| Operational script | Root-level `<MODULE-NAME>.md` (e.g. `INCREMENTAL_SYNC.md`) |
| No existing home | Nearest established area; document the new path in `README.md` |

## Exit Criteria

- Primary module `.md` has all 9 sections from Step 2, accurate and current.
- Every linked/dependent module doc has an accurate **Dependent & Impacted Files** and
  **Data Flow** section.
- Root operational docs that cover the changed behavior are updated.
- The final response lists every `.md` file updated.
- No unrelated `.md` rewrites were introduced.

## Changelog

- `2026-05-18`: expanded to full 9-section module doc standard with UI fields, table
  schema, data flow, and mandatory linked-module doc updates.
- `2026-04-30`: added agent-setup doc mapping and explicit exit criteria.
