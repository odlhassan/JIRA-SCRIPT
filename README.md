# JIRA-SCRIPT

## Production startup

Azure App Service starts the Flask app through `startup.txt`, which runs `wsgi:app` with
Gunicorn. Keep production on one Gunicorn worker because the app uses SQLite-backed
settings and refresh state during WSGI startup; multi-worker cold starts can race on
schema initialization before the warmup probe succeeds.

Local development uses:

```powershell
python run_server.py
```
