# AGENTS.md

## Final response requirement
- At the end of every completed task, include a short section named `How to test locally`.
- Before writing that section, explore the project and identify the exact local verification commands that apply to the latest change. Prefer project-native sources such as `AGENTS.md`, `README.md`, `package.json`, test files, helper scripts, and existing CLI/docs in the repo instead of generic guesses.
- When verification is feasible in the current environment, execute the relevant local CLI commands yourself before completing the task. Use the narrowest command set that proves the change, and prefer targeted checks over broad expensive runs unless the change requires broader coverage.
- In the final response, include the exact CLI commands you executed, and separately list any additional local commands the user can re-run to verify the change on their own machine.
- If local server run is applicable, include the exact command to start the local server and the relevant local URL or endpoint to check.
- If a command could not be executed, say so briefly and explain why.
- If no local server run is applicable, explicitly state `How to test locally: Not applicable for this change`.
