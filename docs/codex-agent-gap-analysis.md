# Codex Agent Gap Analysis

Snapshot date: `2026-04-30`

## Inventory Reviewed

- Repo rules: `.cursor/rules/workspace-context.mdc`, `.cursor/rules/html-and-associated-scripts.mdc`, `.cursor/rules/experience-latest-changes.mdc`
- Repo skills: `.cursor/skills/localhost-ready-after-code/SKILL.md`, `.cursor/skills/update-docs-with-code/SKILL.md`
- Repo instruction file: `AGENTS.md`
- Global skills already available in this environment: `completion-test-fix`, `localhost-ready-check`, `module-doc-sync`

## Capability Matrix

| Capability | Status before this implementation | Evidence | Gap |
| --- | --- | --- | --- |
| Workspace defaults | Partial | `workspace-context.mdc` already named the repo root, localhost host, key routes, and DBs. | It mixed in unrelated GitHub push instructions and did not capture the actual `run_server.py` and `run_html_only.py` usage contract clearly enough. |
| HTML + script coupling | Partial | `html-and-associated-scripts.mdc` already required JS, CSS, generator, test, and server alignment. | It did not explicitly require canonical-source edits, `report_html/` sync awareness, or doc updates when visible behavior changed. |
| Localhost verification section in final response | Partial | `experience-latest-changes.mdc` required a localhost section, and `AGENTS.md` required `How to test locally`. | The two files were not aligned on exact section names, executed-command reporting, ordered manual steps, or expected visible behavior. |
| Test-first / targeted test execution | Partial | The global `completion-test-fix` skill already existed and instructed targeted testing. | The repo had no setup-specific regression test that guarded the required contract artifacts. |
| Documentation sync after code changes | Partial | The repo already had `update-docs-with-code`, and the environment already had `module-doc-sync`. | The project-local doc skill did not call out agent-setup docs, minimal scope rules, or exit criteria. |

## Overlaps And Conflicts

- `AGENTS.md` required `How to test locally`, while `.cursor/rules/experience-latest-changes.mdc` separately required a localhost section. The behavior domain was correct, but the output contract was split across two places and only partially aligned.
- `workspace-context.mdc` contained a stability-and-push rule that was unrelated to workspace defaults and could conflict with execution-first task handling.
- The repo-local skills existed, but both were descriptive rather than procedural. They lacked concise trigger, workflow, and exit criteria blocks.

## Missing Or Under-Specified Artifacts Before The Patch

- `docs/codex-agent-gap-analysis.md`
- `docs/codex-task-contract.md`
- `docs/codex-agent-validation-report.md`
- A focused automated test that proves the required agent-contract files exist and contain the critical constraints

## Upgrade Direction

1. Keep one canonical rule per behavior domain.
2. Make localhost and local-test output requirements exact instead of suggestive.
3. Keep repo-local skills procedural and short enough that they are easy to follow.
4. Add a regression test so future edits do not silently remove the setup contract.
