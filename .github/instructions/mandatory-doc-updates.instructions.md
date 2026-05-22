---
applyTo: "**"
description: "Mandatory rule applied to every code task: after ANY code change, update all relevant .md documentation files before marking the task complete. Use when editing .py, .js, .html, .css, .sql, or any source file."
---

# Mandatory: Update .md Docs After Every Code Change

**This rule is non-negotiable and applies to every code task in this repository.**

## Required Steps Before Task Complete

After making ANY code change, you MUST do all of the following before calling `task_complete`:

1. **Identify** every module/file you changed
2. **Update** the primary module's `.md` doc file (create it if it doesn't exist)
3. **Update** every linked/dependent module's `.md` doc
4. **List** every `.md` file updated in your final response

## Which .md Files to Update

| Changed Area | Update These Docs |
|---|---|
| Report generator (`generate_*.py`) | Root operational `.md` + `docs/report-user-guide/screens/` |
| Capacity / settings module | `ASSIGNEE_HOURS_CAPACITY.md` + `docs/capacity-user-guide/screens/` |
| Jira sync / export | `INCREMENTAL_SYNC.md`, `GENERATED_EXPORTS_COLUMNS.md` |
| HTML / template | Coupled JS/CSS generator + relevant screen doc |
| SQLite schema | `DATABASE_SCHEMA_MIGRATION_PROTOCOL.md` + primary module doc |
| Server routes (`report_server.py`) | `README.md`, `AZURE_APP_SERVICE.md` |
| Any module | `handover/**/*.md` if behavior visible to end user changed |

## Mandatory Sections in Every Module .md

Each module `.md` file must contain:
- **Business Logic** — rules, calculations, edge cases
- **Business Cases** — why it exists, who uses it
- **Examples** — concrete input → output
- **Explanations** — plain-English end-to-end walkthrough
- **Front-end UI Fields** — every visible field, filter, column, button
- **Script Files** — every `.py`, `.js`, `.html`, `.sql` in the module
- **Dependent & Impacted Files** — other modules that read/write this module's output
- **Table Schema** — every SQLite table the module reads or writes
- **Data Flow** — step-by-step from data source to output

## Task Completion Gate

**The task is NOT complete until:**
- ✅ Primary module `.md` updated (or created)
- ✅ All dependent/linked module `.md` files updated
- ✅ Final response explicitly lists every `.md` file that was updated

If you reach `task_complete` without updating docs — **stop, go back, update the docs first**.
