# JIRA-SCRIPT

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
readiness and the prototype archive remain in browser local storage. See
PRODUCT_RELEASE_READINESS_DESIGN.md for field behavior and implementation notes.
