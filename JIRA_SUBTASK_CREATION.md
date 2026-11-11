# Create a Jira subtask from My Availability

## Business Cases

Employees can turn planned work into a real Jira subtask without leaving Availability Details or losing an unfinished worklog description. The chooser reads the personal Jira account's visible spaces and hierarchy directly; it is independent of the full-fetch project list.

## Business Logic

- The creation path is **Jira space → existing epic → existing user story → new subtask**. The epic must be in the chosen space; the story must be in that space and belong to that epic. Tasks, orphan stories, epics and other subtasks cannot be substituted for the story. Modern Parent and legacy Epic Link (resolved by Jira field schema) are supported. Story issue types are named `Story` or `User Story`; Epic uses hierarchy level 1 or the name `Epic`.
- Creation always uses personal OAuth and the verified `/myself` account ID. Shared server credentials cannot create a subtask. Opening another employee's calendar still creates work assigned to the signed-in user, as the dialog explains.
- Start and due dates are required ISO calendar dates. Start must be on or before due. Due must fall between the first and last day of the **actual current month**, independently of the calendar month being browsed. The current date uses the personal Jira profile's timezone, with Asia/Karachi as fallback. `tzdata` supplies IANA zones on Windows. The server repeats validation at submission, including across month rollover; changing the HTML date limits cannot bypass it. Start may be in an earlier month.
- Original estimate defaults to `1d`. Shortcuts are `30m`, `4h`, `1d`, `1w`. Custom positive durations accept w/d/h/m combinations with each unit once (for example `1w 2d`, `1h 30m`). Values retain Jira's working-day/week semantics; the separate worklog duration parser's 24-hour day and maximum do not apply. Jira performs final estimate validation.
- The server discovers the space's subtask issue type and paginated create fields. It prefers a type called Sub-task/Subtask, otherwise the first subtask type. One unambiguous date field called Start date, a Due date field, Assignee, Parent, Summary and Time tracking must be available on the create screen. Due date uses the standard `duedate` field when available, otherwise a single date field named Due Date (such as ODL’s `customfield_10163`). Unsupported mandatory fields without defaults produce a specific message. A supplied description also needs its create field enabled. Jira still enforces validators and permission rules.
- Browse, Create, Assign and Schedule Issues permissions are checked, together with assignability of the personal account. Review and submit both re-read the hierarchy, fields and permissions.
- A ten-minute signed review token identifies one creation attempt and its account. Submit requires confirmation from the review button. Before sending the POST, the server claims a durable exclusive UUID receipt and serializes writes. Replaying that token returns the known key or an uncertain-result message without another POST. An ambiguous failure must be checked in Jira before starting a different request. Optional verification failure retains the returned key; verification checks assignment, parent, both dates and a positive saved original estimate.
- Receipts reuse the existing worklog directory and account/time marker, so existing privacy reporting and account erasure cover stored profile data. There is no SQLite schema change or canonical snapshot write.

## Examples

On November 11, 2026, Alice chooses NEW → NEW-1 (epic) → NEW-2 (story), title “Build customer search”, start November 11, due November 30 and `1d`. Jira creates NEW-3 beneath NEW-2, assigned to Alice, using Jira's interpretation of one working day. A December 1 due date fails even if the user is browsing December. November 10 due fails when start is November 11. A review prepared in November cannot create a November-due task after the current month becomes December.

## Explanations

Open **My availability**, select **New subtask** in the toolbar or the Log work panel, and choose its three parent levels. Search or load more to find items beyond the first results. Add a title, optional description, dates and planned effort. **Review subtask** checks Jira without creating anything; **Edit details** preserves the form. **Create subtask in Jira** performs the real write. Success shows its key, Jira link and **Log work on this subtask**. The last action opens your current-month calendar and preselects the new subtask with today's log date. Closing the creator instead preserves the existing worklog draft and adds the new task to its chooser.

## Front-end UI Fields

| Field/action | Default and validation | Behavior |
| --- | --- | --- |
| New subtask | Toolbar and Log work panel | Opens an accessible native modal over Availability. |
| Jira space | Empty, required | Personal Jira live search by key/name, with Load more. |
| Epic | Disabled until space selected, required | Live title/key search within that space. |
| User story | Disabled until epic selected, required | Live child search; only Story/User Story choices. Changing a parent clears descendants and ignores stale responses. |
| Subtask title | Empty, 1–255 characters | Trimmed by server. |
| Description | Empty, optional, maximum 2000 characters | Stored as Jira rich-text document containing plain text. |
| Start date | Selected logging day when in current month, otherwise today | Required; may be before current month; cannot exceed due. |
| Due date | Same initial date as start | Required, min/max actual current month in Jira timezone. |
| Original estimate | `1d`, required | Editable w/d/h/m; shortcut selection is visibly highlighted. |
| Review subtask / Edit details | No write | Validates live Jira state, previews assignment, hierarchy, dates and estimate. |
| Create subtask in Jira | Review screen only | Single attempt; disables conflicting actions while saving. |
| Open in Jira | Success only | Opens new issue in a separate tab. |
| Log work on this subtask | Success only | Preselects issue without logging any time. |
| Close / Escape | Available except during review/write request | Restores focus to original trigger and preserves underlying worklog draft. |
| Log work: live Jira search / Load more Jira subtasks | Latest 50 assigned subtasks on opening | Searching a title/key searches eligible subtasks across spaces, including tasks assigned to others. Merges results with saved and in-session tasks; paginates. Missing OAuth/network failure retains saved choices with a status message. |

## Script Files

- `employee_subtask.py`: live options/search, hierarchy and create-field checks, date/estimate validation, signed review, durable submit and direct verification.
- `employee_worklog.py`: registers routes with existing personal identity, origin guard, signer, errors and write serialization.
- `employee_subtask_form.js` / `.css`: guided responsive dialog, live cascading search, review, success and in-session task chooser data.
- `employee_worklog_form.js`: entry point, live subtask search and immediate selection while preserving drafts.
- `employee_availability_calendar.js` / `.css`: toolbar entry point, live refresh coordination and calendar metadata.
- `generate_employee_capacity_utilization_report.py`: inlines all UI sources into the root report; server startup promotes it to `report_html`.
- `requirements.txt`: `tzdata` dependency for profile timezone boundaries.
- `tests/test_employee_subtask.py`: mocked Jira API tests; `tests/test_employee_availability_calendar.cjs`: full browser flow and regression coverage.

## Dependent & Impacted Files

`report_server.py` already registers the worklog routes, serves the report and live tally. `jira_oauth.py` supplies personal tokens using the existing read/write Jira scopes. `jira_calendar_cache.py` invalidates the created issue and parent. `jira_privacy.py` reads/erases account markers in the existing UUID receipt format. [Calendar worklogs](JIRA_WORKLOG_SIGNIN.md), [capacity report](docs/report-user-guide/screens/16-employee-capacity-utilization.md), and [monthly tally](WORKLOG_TALLY.md) retain their source rules. Main reports, exports and My Day planned tasks continue to use the canonical snapshot and therefore require a full fetch for the new issue.

## Table Schema

No tables, columns, constraints or indexes are added or changed. This workflow does not read or write SQLite directly. The surrounding report still reads existing `canonical_issues`/`canonical_worklogs`, capacity profiles and RLT leave snapshots. Durable JSON receipts live under the existing persistent `worklog_requests` directory; they contain the created issue key, returned details, verification state and the account/timestamp marker used by existing privacy handling. Pending receipts hold an error/request ID only.

## Data Flow

1. `GET /api/employee-capacity-utilization/worklog/subtask/options?kind=projects|epics|stories` reads personal Jira projects/search and returns paginated choices, account label and current-month bounds. `project_key`, `epic_key`, `q`, `cursor` control descendants/search/pages.
2. `POST .../subtask/review` validates the complete draft, reads parent/epic, Jira field catalogue, paginated create metadata, permissions and assignability; returns a signed review without writes.
3. `POST .../subtask/submit` verifies personal account/token/confirmation, revalidates, claims the receipt and sends one `POST /rest/api/3/issue` with parent story, personal assignee, discovered start field, due date and original estimate. It persists the key before direct verification.
4. The UI merges the returned issue into an in-session chooser map and dispatches `epr-subtask-created`. Existing calendar refresh is paused while the creator is open. The new issue can be selected immediately, without waiting for Jira search indexing. `GET .../subtask/search?q=...&cursor=...` supplies live loggable subtasks after reload; external search indexing may briefly delay discovery.
5. Creation does **not** alter logged hours, leave or available capacity. After actual work is logged, existing issue invalidation and live calendar reads update the affected daily/monthly work figures. Saved summary rows, booked estimates, exports and other modules reflect it after the next full fetch.

Jira API contracts: [issue creation and metadata](https://developer.atlassian.com/cloud/jira/platform/rest/v3/api-group-issues/) and [working-time configuration](https://developer.atlassian.com/cloud/jira/platform/rest/v3/api-group-jira-settings/).

## Local verification

From `E:\JIRA SCRIPT`:

```powershell
.venv/Scripts/python.exe -m pip install -r requirements.txt
.venv/Scripts/python.exe generate_employee_capacity_utilization_report.py
.venv/Scripts/python.exe -m unittest tests.test_employee_subtask tests.test_employee_worklog tests.test_employee_capacity_utilization_report tests.test_jira_privacy
node node_modules/@playwright/test/cli.js test --config tests/employee-availability.playwright.config.cjs --workers=1 --reporter=line
.venv/Scripts/python.exe run_server.py --no-sync --keep-cache --port 3119
```

Open `http://127.0.0.1:3119/employee_capacity_utilization_report.html`, sign in through configured local personal OAuth, open My availability and try New subtask. Tests mock Jira and create no live issue. Actual creation requires the space's Jira permissions and create-screen fields above. Standard `.venv/Scripts/python.exe run_server.py --keep-cache --port 3119` also rebuilds/promotes the latest UI at startup.

## Read-only Jira configuration check — November 11, 2026

The local server’s personal OAuth status is currently unconfigured; its creator explains that administrator setup is required instead of offering an unusable login. Live create/edit metadata was read for DIGITALLOG and ODL; no issue was created or edited. Both spaces expose Start date (`customfield_10015`) and Assignee. DIGITALLOG exposes standard Due date; ODL exposes custom Due Date (`customfield_10163`), which this workflow supports. Jira reports time tracking globally enabled, but neither inspected subtask create/edit screen exposes Time tracking. Their Jira administrator must add **Time tracking** to the relevant subtask create screen before this feature can set the mandatory original estimate. The app blocks creation with a specific field message until that external configuration is ready. The app does not change Jira administration settings. Automated writes remain mocked.

Automated verification: the final focused Python command above passes 54 tests; the Availability browser suite covers desktop/mobile layout, default/shortcut estimates, month bounds, parent reset/search/pagination, review/one-submit behavior, failure messages, preserved worklog drafts, live search after reload and the existing calendar/worklog flows. The actual localhost report returns HTTP 200 with the new inlined UI, and anonymous creation reads correctly require personal Jira sign-in.
