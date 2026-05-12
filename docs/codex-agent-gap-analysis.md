# Codex Agent Gap Analysis

## Verified current state (`2026-05-12`)

Automated check: `tests/test_codex_agent_setup_contract.py` (plus `tests/test_run_server.py` for the documented `python run_server.py` path).

| Capability | Status | Evidence |
| --- | --- | --- |
| Workspace defaults | Present | `.cursor/rules/workspace-context.mdc` — root path, `run_server.py`, `run_html_only.py`, DBs, routes |
| HTML + script coupling | Present | `.cursor/rules/html-and-associated-scripts.mdc` — canonical source, JS/CSS, generators, tests, `report_server.py` sync |
| Localhost section in final response | Present | `.cursor/rules/experience-latest-changes.mdc` — exact title, commands, URLs, steps |
| Local test section + executed commands | Present | `AGENTS.md` — `How to test locally`, narrow checks, report executed commands |
| Targeted test execution | Present | Global `completion-test-fix` skill + repo regression `tests/test_codex_agent_setup_contract.py` |
| Documentation sync after code changes | Present | `.cursor/skills/update-docs-with-code/SKILL.md` + global `module-doc-sync` |
| Task contract snippet | Present | `docs/codex-task-contract.md` |
| Phase 5 validation narrative | Present | `docs/codex-agent-validation-report.md` |

### Inventory (repo + global skills)

- Repo rules: `.cursor/rules/workspace-context.mdc`, `.cursor/rules/html-and-associated-scripts.mdc`, `.cursor/rules/experience-latest-changes.mdc`
- Repo skills: `.cursor/skills/localhost-ready-after-code/SKILL.md`, `.cursor/skills/update-docs-with-code/SKILL.md`
- Repo instruction file: `AGENTS.md`
- Global skills (recommended in plan): `%USERPROFILE%\.codex\skills\completion-test-fix`, `localhost-ready-check`, `module-doc-sync`, optional `regression-audit`

### Overlaps to keep intentional

- `AGENTS.md` and `.cursor/rules/experience-latest-changes.mdc` both govern the final response: **localhost** steps vs **local test** steps. Domains stay split; wording is aligned on exact section titles.

---

## Historical baseline (`2026-04-30`)

The following captured the state **before** the upgrade described in `docs/CODEX_AGENT_SETUP_IMPLEMENTATION_PLAN.md` (retained for audit trail).

### Capability matrix (before upgrade)

| Capability | Status before this implementation | Evidence | Gap |
| --- | --- | --- | --- |
| Workspace defaults | Partial | `workspace-context.mdc` already named the repo root, localhost host, key routes, and DBs. | It mixed in unrelated GitHub push instructions and did not capture the actual `run_server.py` and `run_html_only.py` usage contract clearly enough. |
| HTML + script coupling | Partial | `html-and-associated-scripts.mdc` already required JS, CSS, generator, test, and server alignment. | It did not explicitly require canonical-source edits, `report_html/` sync awareness, or doc updates when visible behavior changed. |
| Localhost verification section in final response | Partial | `experience-latest-changes.mdc` required a localhost section, and `AGENTS.md` required `How to test locally`. | The two files were not aligned on exact section names, executed-command reporting, ordered manual steps, or expected visible behavior. |
| Test-first / targeted test execution | Partial | The global `completion-test-fix` skill already existed and instructed targeted testing. | The repo had no setup-specific regression test that guarded the required contract artifacts. |
| Documentation sync after code changes | Partial | The repo already had `update-docs-with-code`, and the environment already had `module-doc-sync`. | The project-local doc skill did not call out agent-setup docs, minimal scope rules, or exit criteria. |

### Overlaps and conflicts (before upgrade)

- `AGENTS.md` required `How to test locally`, while `.cursor/rules/experience-latest-changes.mdc` separately required a localhost section. The behavior domain was correct, but the output contract was split across two places and only partially aligned.
- `workspace-context.mdc` contained a stability-and-push rule that was unrelated to workspace defaults and could conflict with execution-first task handling.
- The repo-local skills existed, but both were descriptive rather than procedural. They lacked concise trigger, workflow, and exit criteria blocks.

### Missing or under-specified artifacts (before upgrade)

- `docs/codex-agent-gap-analysis.md`
- `docs/codex-task-contract.md`
- `docs/codex-agent-validation-report.md`
- A focused automated test that proves the required agent-contract files exist and contain the critical constraints
