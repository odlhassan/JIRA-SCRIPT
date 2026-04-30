---
name: update-docs-with-code
description: After completing or changing code, update relevant .md documentation (README, user guides, column specs, runbooks) so docs stay in sync. Use when finishing code changes, adding features, changing exports, or when the user asks to keep docs updated.
---

# Update Docs With Code

Use this skill whenever a code change affects behavior, outputs, CLI usage, localhost steps, or user-visible fields.

## Workflow

1. Map the changed files to the smallest relevant `.md` set.
   - Root operational docs: `AGENTS.md`, `README.md`, `EXPECTED_FILES.md`, `DASHBOARD_REFRESH_CLI.md`, `GENERATED_EXPORTS_COLUMNS.md`, `RLT_LEAVE_REPORT.md`, `ASSIGNEE_HOURS_CAPACITY.md`, `IPP_PHASE_TRANSFORM_LOGIC.md`, `NESTED_VIEW_SCORECARD_FORMULAS.md`
   - User guides: `docs/report-user-guide/` and `docs/capacity-user-guide/`
   - Handover docs: `handover/**/*.md`
   - Agent-setup docs: `docs/codex-agent-gap-analysis.md`, `docs/codex-task-contract.md`, `docs/codex-agent-validation-report.md`
2. Update only the sections that the code change actually affects.
3. Keep examples, CLI commands, routes, and field names aligned with the codebase.
4. If a new behavior has no doc home yet, add it in the nearest established docs area rather than inventing a new structure.

## Exit Criteria

- Every changed behavior with an existing documentation home is reflected in the relevant `.md` files.
- The final response names which docs were updated.
- No unrelated `.md` rewrites were introduced.

## Changelog

- `2026-04-30`: added agent-setup doc mapping and explicit exit criteria.
