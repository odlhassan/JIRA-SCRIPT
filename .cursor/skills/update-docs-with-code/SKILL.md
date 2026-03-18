---
name: update-docs-with-code
description: After completing or changing code, update relevant .md documentation (README, user guides, column specs, runbooks) so docs stay in sync. Use when finishing code changes, adding features, changing exports, or when the user asks to keep docs updated.
---

# Update docs with code completion

When you complete or change code, **update the relevant .md files** so documentation stays in sync. Do this as part of the same task, not as an afterthought.

## When to apply

- You have just added, removed, or changed behavior in scripts, APIs, or reports.
- You added or changed CLI options, columns, exports, or report outputs.
- You changed logic that is described in existing .md (e.g. formulas, refresh behavior, data contracts).

## Workflow

1. **Identify affected docs**  
   After making code changes, determine which .md files document that code or feature:
   - **Root docs**: `README.md`, `AGENTS.md`, `GENERATED_EXPORTS_COLUMNS.md`, `RLT_LEAVE_REPORT.md`, `EXPECTED_FILES.md`, `NESTED_VIEW_SCORECARD_FORMULAS.md`, `IPP_PHASE_TRANSFORM_LOGIC.md`, `DASHBOARD_REFRESH_CLI.md`, `INCREMENTAL_SYNC.md`, `ASSIGNEE_HOURS_CAPACITY.md`, etc.
   - **User guides**: `docs/report-user-guide/` and `docs/capacity-user-guide/` (overview and screen-specific .md).
   - **Handover/runbooks**: `handover/**/*.md` (e.g. `RUNBOOK.md`, `DATA_CONTRACT.md`).

2. **Update those files**  
   - Add or revise sections that describe the new/changed behavior, columns, options, or UI.
   - Fix examples, field lists, and step-by-step instructions so they match the current code.
   - If you added a new report or script, add or link it in the right overview (e.g. report-overview or README) and create or update a dedicated screen doc if the project uses that pattern.

3. **Keep changes minimal**  
   Only touch .md that are actually affected; don’t rewrite unrelated sections.

## Conventions

- **Column/export changes**: Update `GENERATED_EXPORTS_COLUMNS.md` (or the doc that lists export columns) when you add/remove/rename columns or change export behavior.
- **Report behavior**: Update the corresponding screen under `docs/report-user-guide/screens/` or `docs/capacity-user-guide/screens/` (e.g. filters, new sections, removed features).
- **CLI / scripts**: Update README, `DASHBOARD_REFRESH_CLI.md`, or the relevant runbook so commands and options are accurate.
- **Formulas / logic**: Update docs like `NESTED_VIEW_SCORECARD_FORMULAS.md` or `IPP_PHASE_TRANSFORM_LOGIC.md` when those formulas or steps change in code.

## Summary

After every code completion that affects behavior, exports, or UX: **identify the relevant .md files and update them in the same pass** so the system docs stay accurate.
