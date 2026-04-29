# Codex Agent Rules/Skills Implementation Plan

## Goal
Create a reusable Codex agent setup that consistently delivers:
- Correct project context usage.
- Safe, complete code changes (including coupled files).
- Verifiable localhost-ready outcomes.
- Synchronized tests and documentation.

This plan is written so Codex can execute it directly with minimal follow-up.

## Scope
- Implement and wire rules/instructions for behavior.
- Add/align reusable skills for testing, localhost readiness, and doc sync.
- Validate the setup through a dry-run task and checklist.

Out of scope:
- Product feature implementation unrelated to agent setup.
- CI/CD architecture changes beyond minimal verification commands.

## Target Outcomes
1. A new Codex run starts with clear workspace defaults and guardrails.
2. Any HTML/UI change forces related JS/CSS/generator/test updates.
3. Every completed task includes exact localhost and test instructions.
4. Docs stay in sync with code changes for modified modules.

## Required Artifacts
Create or update the following files (adapt paths if your environment differs):

1. `.cursor/rules/workspace-context.mdc`
2. `.cursor/rules/html-and-associated-scripts.mdc`
3. `.cursor/rules/experience-latest-changes.mdc`
4. `AGENTS.md`
5. `.cursor/skills/localhost-ready-after-code/SKILL.md`
6. `.cursor/skills/update-docs-with-code/SKILL.md`
7. (Optional but recommended) `~/.codex/skills/completion-test-fix/SKILL.md`
8. (Optional but recommended) `~/.codex/skills/localhost-ready-check/SKILL.md`
9. (Optional but recommended) `~/.codex/skills/module-doc-sync/SKILL.md`
10. (Optional) `~/.codex/skills/regression-audit/SKILL.md`

## Implementation Phases

## Phase 1: Baseline Inventory
### Objective
Capture current agent behavior and identify missing enforcement points.

### Steps
1. Collect existing rule files and skill files from project and global locations.
2. Map overlaps and conflicts (same behavior defined in multiple places).
3. Mark each required capability as `Present`, `Partial`, or `Missing`:
   - Workspace defaults.
   - HTML + script coupling.
   - Localhost verification section in final response.
   - Test-first/targeted test execution.
   - Documentation sync after code changes.

### Deliverable
- `docs/codex-agent-gap-analysis.md` with a matrix of capabilities vs current state.

## Phase 2: Normalize Rule Set
### Objective
Define strict, non-ambiguous rules Codex follows in every task.

### Steps
1. Update `workspace-context.mdc` to include:
   - Project root path conventions.
   - Primary server script and default localhost URL.
   - Known databases/caches and route defaults.
2. Update `html-and-associated-scripts.mdc` to enforce:
   - HTML changes require JS/CSS/template/generator/test/server alignment.
   - No partial HTML-only edits when selectors/handlers are affected.
3. Update `experience-latest-changes.mdc` to require:
   - Exact CLI commands.
   - Ordered manual validation steps.
   - URL(s) and expected visible behavior.
4. Update `AGENTS.md` with:
   - Required `How to test locally` section.
   - Rule to run narrow, relevant checks when feasible.
   - Requirement to report executed commands and any non-executed checks.

### Deliverable
- Finalized rules with no contradictory instructions.

## Phase 3: Skill Hardening
### Objective
Ensure Codex has reusable, task-triggered procedures.

### Steps
1. `localhost-ready-after-code` skill:
   - Verify generated artifacts are refreshed if applicable.
   - Verify server startup command and target URL.
2. `update-docs-with-code` skill:
   - Detect impacted docs automatically.
   - Update only relevant sections (avoid broad unrelated rewrites).
3. `completion-test-fix` skill:
   - Add/adjust focused tests for changed behavior.
   - Run targeted test command(s).
   - Fix straightforward regressions before final response.
4. `localhost-ready-check` skill:
   - Produce concise manual QA path for the changed behavior.
5. `module-doc-sync` skill:
   - Keep module-level business logic docs aligned.
6. `regression-audit` skill (optional):
   - For broad changes, run workflow-level checks and summarize risks.

### Deliverable
- Skills with clear trigger conditions and output format requirements.

## Phase 4: Prompt Contract for Codex
### Objective
Provide a standard “task contract” so Codex executes predictably.

### Add this reusable prompt block to new tasks
```md
Use project rules and skills strictly.

Completion requirements:
1) Implement requested change end-to-end.
2) Run focused verification commands where feasible.
3) Fix any issues introduced by your edits.
4) Update relevant documentation when behavior/output changes.
5) End your response with:
   - How to experience latest changes on live localhost
   - How to test locally
6) Include exact commands you executed, and separate commands I can run.
```

### Deliverable
- Saved snippet in `docs/codex-task-contract.md`.

## Phase 5: Validation Dry Run
### Objective
Confirm the setup behaves correctly on a realistic small task.

### Test Scenario
Pick a contained UI/report tweak that touches:
- At least one HTML output/template.
- One related JS/CSS or generator file.
- One focused test.
- One docs page.

### Pass Criteria
1. Codex updates coupled files, not HTML alone.
2. Codex runs at least one targeted check and reports result.
3. Final response includes both required sections:
   - `How to experience latest changes on live localhost`
   - `How to test locally`
4. Docs are updated only where relevant.

### Deliverable
- `docs/codex-agent-validation-report.md` with:
  - Task run summary.
  - Commands executed.
  - Failures encountered and fixes applied.
  - Remaining gaps.

## Governance and Maintenance
1. Review rules monthly or after major workflow changes.
2. Add a changelog section to each rule/skill file for traceability.
3. When a regression occurs, update the nearest rule/skill that should have prevented it.
4. Keep rule text concise and imperative; avoid vague guidance.

## Risks and Mitigations
- Risk: Overlapping rules create conflicting behavior.
  - Mitigation: Maintain one canonical rule per behavior domain.
- Risk: Skills become verbose and ignored.
  - Mitigation: Keep each skill procedural with explicit trigger and exit criteria.
- Risk: “Test locally” becomes boilerplate.
  - Mitigation: Enforce actual command execution reporting.

## Suggested Execution Order (1-2 Days)
Day 1:
1. Phase 1 inventory.
2. Phase 2 rule normalization.
3. Phase 3 core skills (`completion-test-fix`, localhost skills, docs sync).

Day 2:
1. Phase 4 task contract snippet.
2. Phase 5 dry run and validation report.
3. Tighten wording based on observed misses.

## Definition of Done
This implementation is complete when:
1. All required artifacts exist and are updated.
2. Dry-run task passes all criteria.
3. Codex final responses consistently include localhost + local-test sections.
4. At least one completed task demonstrates test execution and doc sync.

## Handoff Notes (for Codex)
- Prioritize deterministic instructions over “best effort” wording.
- If a rule cannot be followed due to environment constraints, state why and provide fallback verification steps.
- Never silently skip testing/doc updates when the change affects behavior or outputs.
