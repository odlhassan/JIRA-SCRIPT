# JIRA-SCRIPT

## Production startup

Azure App Service starts the Flask app through `startup.txt`, which runs `wsgi:app` with
Gunicorn. Keep production on one Gunicorn worker because the app uses SQLite-backed
settings and refresh state during WSGI startup; multi-worker cold starts can race on
schema initialization before the warmup probe succeeds.

The deploy workflow vendors `requirements.txt` into `.python_packages/lib/site-packages`
inside the ZIP, and `startup.txt` prepends that folder to `PYTHONPATH`. This keeps the
site bootable even when Azure starts from the ZIP contents without an Oryx-created
virtual environment.

Local development uses:

```powershell
python run_server.py
```
