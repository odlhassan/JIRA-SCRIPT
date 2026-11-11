# JIRA-SCRIPT

My availability now includes **New subtask**: choose a live Jira space → epic → user story, enter start/due dates and an original estimate (default `1d`; shortcuts `30m`, `4h`, `1d`, `1w`), then review and create it assigned to your personal Jira account. Due dates must remain in the actual current month. Live subtask search makes it loggable before a full fetch; other reports retain their saved snapshot. Search lives inside each dropdown and preserves selections until a different parent is chosen. RLT has a separate leave path through existing tracker Tasks/Stories, with Planned/Unplanned Leave classification and explicit personal leave recording. See [Jira subtask creation](JIRA_SUBTASK_CREATION.md).

## Production startup

Azure App Service starts the Flask app through `startup.txt`, which runs `wsgi:app` with
Gunicorn. Keep production on one Gunicorn worker because the app uses SQLite-backed
settings and refresh state during WSGI startup; multi-worker cold starts can race on
schema initialization before the warmup probe succeeds.

The deploy workflow vendors `requirements.txt` into `.python_packages/lib/site-packages`
inside the ZIP and marks `startup.txt` executable before packaging. The script prepends
`.python_packages` to `PYTHONPATH`, which keeps the site bootable even when Azure starts
from read-only ZIP contents without an Oryx-created virtual environment.
Before uploading the ZIP, the workflow also verifies the WSGI startup files and direct
startup helpers such as `db_journal_mode.py` are present in the staged package.

Local development uses:

Signed-in users can delete their own individual Jira worklogs from **My availability → Manage my logged hours** in the personal account area or the red **Delete** button in **Availability → day details**, using any task key. See [personal worklog management](JIRA_WORKLOG_SIGNIN.md#delete-your-own-logged-hours) for permissions, confirmation and snapshot-refresh behavior. The personal manager browses monthly history with search, filters, totals and CSV export. Report detail modals have Copy link URLs that restore the view without submitting an action.

```powershell
python run_server.py
```

## Product release readiness design

An interactive release-readiness design is available at:

http://127.0.0.1:3000/settings/product-releases/readiness-design

It presents products as vertical tabs, then a focused release list and inline release
detail board. Release number, release date, epic assignments, and release completion use
the existing Product Releases APIs and remain synchronized with that module. Release,
epic, checklist, and content readiness can be changed inline; checklist dates are
intentionally excluded and Planned can instead be manually marked delayed. Checklist
readiness and the prototype archive remain in browser local storage. Product tabs come
only from managed projects or live release records; demonstration releases cannot create
an unconfigured product. Responsible/owner selectors load active people through the
existing Performance assignee API; the final dropdown option can add a name that is saved
only inside the selected release's browser-local board. See
PRODUCT_RELEASE_READINESS_DESIGN.md for field behavior and implementation notes.

## Reduce Colossal Refresh storage

On `/settings/canonical-refresh`, choose a smaller year/start-month scope and run
**Full Refresh** with **Replace previous fetched Jira data and reclaim disk space**
checked. After reports succeed, older canonical Jira snapshots are removed and the
database is compacted, preserving settings and manual plans. The result shows reclaimed
space or a pending-maintenance warning if disk headroom/locks prevent compaction.
Uncheck to retain snapshots. Smart Refresh and standalone Fetch keep their existing behavior.
See [storage maintenance details](CANONICAL_COLOSSAL_FETCH_IMPLEMENTATION_PLAN.md#full-refresh-snapshot-replacement-and-disk-reclamation).
