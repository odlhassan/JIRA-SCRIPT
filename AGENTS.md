# AGENTS.md

## Final response requirement
- End every completed task with two short sections in this order:
  - `How to experience latest changes on live localhost`
  - `How to test locally`
- Before writing those sections, inspect project-native sources such as `AGENTS.md`, `README.md`, `package.json`, test files, helper scripts, and existing CLI/docs in the repo instead of guessing commands, ports, or URLs.
- When verification is feasible in the current environment, execute the narrowest relevant local CLI commands yourself before completing the task. Prefer focused checks over broad expensive runs unless the change requires broader coverage.
- If the change touches HTML, UI, templates, generated reports, or served assets, update coupled JavaScript, CSS, generator, server, test, and relevant `.md` documentation files instead of changing HTML alone.
- In the final response, include the exact CLI commands you executed and separately list any additional local commands the user can re-run on the same machine.
- If a local server run is applicable, include the exact command to start the server, the exact localhost URL or route to inspect, ordered manual steps, and the expected visible behavior.
- If a command could not be executed, say so briefly and explain why.
- If no localhost run is applicable, explicitly state `How to experience latest changes on live localhost: Not applicable for this change`.
- If no local server run is applicable, explicitly state `How to test locally: Not applicable for this change`.
