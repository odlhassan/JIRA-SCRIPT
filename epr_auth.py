"""EPR access control reusing the existing Atlassian OAuth session store."""
from collections import deque
from datetime import datetime, timezone
from functools import wraps
import json
import os
from pathlib import Path
import re
import secrets
import time
from urllib.parse import quote, unquote

from flask import g, jsonify, redirect, request, send_file
import requests
from jira_oauth import COOKIE, config, oauth_session, save_private, same_origin
from jira_privacy import ACCESS_LOCK

OWNER_ACCOUNT_ID = '641adec89796ea0a87185666'
_lock = ACCESS_LOCK
MEMBER_POSTS = {
    '/api/employee-capacity-utilization/worklog/review', '/api/employee-capacity-utilization/worklog/submit',
    '/api/employee-capacity-utilization/export', '/api/planned-actual-table-view/export', '/api/capacity/calculate',
}
MEMBER_SETTINGS = {'/settings/product-releases/calendar', '/settings/rnd-muscle-utilization/view'}
ADMIN_PARTS = ('/admin/', '/sql-console', '/db-migration', '/canonical-', '/report-refresh', '/prepare-offline-html',
               '/populate-jira', '/import', '/diagnostics', '/debug-', '/publish', '/refresh', '/cancel', '/resource-records')


def safe_next(value):
    value = str(value or '/')
    decoded = unquote(value)
    if not value.startswith('/') or decoded.startswith('//') or '\\' in decoded or any(ord(c)<32 for c in decoded):
        return '/'
    if decoded.split('?')[0] in ('/login', '/access-pending'):
        return '/'
    return value


def member_allowed(path, method, report_files):
    if path in ('/api/auth/me', '/api/auth/logout', '/my-availability'):
        return True
    if path.startswith('/api/auth/') or path.startswith('/settings/'):
        return method in ('GET', 'HEAD') and path in MEMBER_SETTINGS
    if any(part in path for part in ADMIN_PARTS):
        return False
    if method not in ('GET', 'HEAD'):
        return method == 'POST' and path in MEMBER_POSTS
    if path.startswith('/api/') or path.startswith('/project-images/'):
        return True
    return path == '/' or path.lstrip('/') in report_files or path in (
        '/shared-nav.js', '/shared-nav.css', '/shared-date-filter.js', '/product-release-readiness-design.js',
        '/material-symbols.css') or path.endswith(('.woff', '.woff2'))


def register_epr_auth(app, oauth_directory: Path, report_files):
    directory = oauth_directory.parent / 'epr_access'
    users_file, audit_file = directory / 'users.json', directory / 'audit.jsonl'
    enabled = bool(os.getenv('WEBSITE_SITE_NAME') or all(config().values()) or os.getenv('EPR_AUTH_ENABLED') == '1')
    app.config['EPR_AUTH_ENABLED'] = enabled
    app.config['EPR_AUTH_DIRECTORY'] = directory
    app.config.setdefault('EPR_SESSION_TTL', 8*3600)
    assets = Path(__file__).parent

    def read_users():
        return json.loads(users_file.read_text(encoding='utf-8')) if users_file.exists() else {}

    def audit(event, actor=None, **fields):
        directory.mkdir(parents=True, exist_ok=True)
        row = dict(at=datetime.now(timezone.utc).isoformat(), event=event, actor=actor, **fields)
        with _lock:
            fd = os.open(audit_file, os.O_WRONLY|os.O_CREAT|os.O_APPEND, 0o600)
            with os.fdopen(fd, 'a', encoding='utf-8') as stream:
                stream.write(json.dumps(row, ensure_ascii=True)+'\n')

    app.extensions['epr_audit'] = audit

    def current_user():
        if hasattr(g, 'epr_identity'):
            return g.epr_identity
        g.epr_identity = None
        sid = request.cookies.get(COOKIE, '')
        if not re.fullmatch('[a-f0-9]{64}', sid):
            return None
        path = oauth_directory / (sid+'.json')
        if not path.exists():
            return None
        try:
            saved = json.loads(path.read_text(encoding='utf-8'))
            if saved.get('session_expires_at', 0) <= time.time():
                return None
            signed_at = saved.get('epr_signed_at')
            if signed_at and time.time()-signed_at > app.config['EPR_SESSION_TTL']:
                return None
            identity = saved.get('epr_identity')
            if not identity:
                session = oauth_session(oauth_directory)
                if session is None:
                    return None
                response = session.get(session.epr_base_url+'/rest/api/3/myself', timeout=(10, 25))
                response.raise_for_status()
                profile = response.json()
                identity = {'account_id': profile['accountId'], 'display_name': profile.get('displayName', ''),
                            'avatar_url': (profile.get('avatarUrls') or {}).get('48x48', '')}
                # Re-read after token refresh so rotating tokens are never overwritten with old ones.
                saved = json.loads(path.read_text(encoding='utf-8'))
                saved.update(epr_identity=identity, epr_signed_at=time.time(), epr_csrf=secrets.token_hex(32))
                save_private(path, saved)
            account = identity['account_id']
            with _lock:
                users = read_users()
                if account not in users:
                    owners = {x.strip() for x in os.getenv('EPR_ADMIN_ACCOUNT_IDS', OWNER_ACCOUNT_ID).split(',') if x.strip()}
                    owner = account in owners
                    users[account] = dict(account_id=account, display_name=identity['display_name'],
                        role='admin' if owner else 'member', approved=owner, active=True, personal_data_retrieved_at=time.time())
                    directory.mkdir(parents=True, exist_ok=True)
                    save_private(users_file, users)
                    audit('bootstrap_admin' if owner else 'access_requested', account)
                elif users[account].get('display_name') != identity['display_name']:
                    users[account]['display_name'] = identity['display_name']
                    users[account]['personal_data_retrieved_at'] = time.time()
                    save_private(users_file, users)
                user = {**identity, **users[account], 'csrf_token': saved.get('epr_csrf', ''),
                        'expires_at': saved.get('epr_signed_at', 0)+app.config['EPR_SESSION_TTL']}
            g.epr_identity = user
            return user
        except (ValueError, KeyError, OSError, requests.RequestException):
            return None

    def denial(message, code, user=None):
        audit('access_denied', (user or {}).get('account_id'), path=request.path, method=request.method, status=code)
        if request.path.startswith('/api/'):
            return jsonify(error=message, login_url='/login'), code
        if code == 401:
            return redirect('/login?next='+quote(safe_next(request.full_path.rstrip('?')), safe=''))
        return '<!doctype html><title>Access restricted</title><h1>Access restricted</h1><p>Your account does not have access to this page.</p><a href="/">Return to EPR</a>', code

    public = {'/login', '/epr-login.css', '/epr-login.js', '/access-pending', '/healthz', '/api/auth/me', '/api/auth/logout',
              '/api/jira-auth/login', '/api/jira-auth/callback', '/api/jira-auth/status', '/api/jira-auth/logout'}

    @app.before_request
    def protect_epr():
        if not enabled or request.path in public:
            return None
        user = current_user()
        if not user:
            return denial('Sign in to EPR with Atlassian.', 401)
        if not user['active'] or not user['approved']:
            if request.path.startswith('/api/'):
                return denial('Your EPR access is pending approval or disabled.', 403, user)
            return redirect('/access-pending')
        g.epr_user = user
        if user['role'] != 'admin' and not member_allowed(request.path, request.method, report_files):
            return denial('Administrator access is required.', 403, user)
        if request.method not in ('GET', 'HEAD', 'OPTIONS'):
            csrf = request.headers.get('X-EPR-CSRF', '')
            valid_csrf = bool(csrf and user['csrf_token'] and secrets.compare_digest(csrf, user['csrf_token']))
            if not (request.headers.get('Origin') and same_origin()) and not valid_csrf:
                return denial('Send this change from EPR.', 403, user)

    @app.after_request
    def protect_responses(response):
        if enabled:
            response.headers['Cache-Control'] = 'no-store, private'
            response.headers['X-Content-Type-Options'] = 'nosniff'
            response.headers['X-Frame-Options'] = 'SAMEORIGIN'
            user = getattr(g, 'epr_user', None)
            if user and request.method not in ('GET', 'HEAD', 'OPTIONS'):
                audit('request', user['account_id'], path=request.path, method=request.method, status=response.status_code)
        return response

    @app.get('/healthz')
    def epr_health():
        return jsonify(ok=True)

    @app.get('/epr-login.css')
    @app.get('/epr-login.js')
    def epr_login_asset():
        return send_file(assets / request.path.lstrip('/'))

    @app.get('/login')
    def epr_login():
        user = current_user() if enabled else None
        if user and user['approved'] and user['active']:
            return redirect(safe_next(request.args.get('next')))
        response = send_file(assets / 'epr_login.html')
        response.set_cookie('epr_next', safe_next(request.args.get('next')), httponly=True,
                            secure=config()['REDIRECT_URI'].startswith('https://'), samesite='Lax', max_age=600)
        return response

    @app.get('/access-pending')
    def epr_pending():
        user = current_user()
        if not user:
            return redirect('/login')
        if user['active'] and user['approved']:
            return redirect('/')
        return send_file(assets / 'epr_pending.html')

    @app.get('/api/auth/me')
    def epr_me():
        if not enabled:
            return jsonify(enabled=False, authenticated=False)
        user = current_user()
        if not user:
            return jsonify(enabled=True, authenticated=False, login_url='/login'), 401
        return jsonify(enabled=True, authenticated=True, user=user)

    @app.post('/api/auth/logout')
    def epr_logout():
        user = current_user()
        if not same_origin() or not request.is_json:
            return jsonify(error='Sign out from EPR.'), 403
        audit('logout', (user or {}).get('account_id'))
        return app.view_functions['jira_auth_logout']()

    @app.get('/my-availability')
    def epr_my_availability():
        user = current_user()
        if not user:
            return redirect('/login')
        return redirect('/employee_capacity_utilization_report.html?availability=1&employee='+quote(user['display_name']))

    @app.get('/settings/users')
    def epr_users_page():
        return send_file(assets / 'epr_users.html')

    def admin_api(function):
        @wraps(function)
        def wrapped(*args, **kwargs):
            user = current_user()
            if not user or not user['active'] or not user['approved'] or user['role'] != 'admin':
                return jsonify(error='Administrator access required.'), 403
            return function(*args, **kwargs)
        return wrapped

    @app.get('/api/auth/users')
    @admin_api
    def epr_users():
        return jsonify(users=list(read_users().values()))

    @app.put('/api/auth/users/<account_id>')
    @admin_api
    def epr_update_user(account_id):
        body = request.get_json(silent=True) or {}
        if body.get('role') not in ('admin', 'member') or not all(isinstance(body.get(x), bool) for x in ('approved', 'active')):
            return jsonify(error='Choose Admin or Member and valid approval/active states.'), 400
        with _lock:
            users = read_users()
            if account_id not in users:
                return jsonify(error='User must sign in once before access can be approved.'), 404
            proposed = {**users[account_id], **{key: body[key] for key in ('role', 'approved', 'active')}}
            updated = {**users, account_id: proposed}
            if not any(x['role']=='admin' and x['approved'] and x['active'] for x in updated.values()):
                return jsonify(error='Keep at least one approved, active Admin.'), 400
            if account_id == current_user()['account_id'] and not (proposed['approved'] and proposed['active']):
                return jsonify(error='You cannot disable your own access.'), 400
            save_private(users_file, updated)
            audit('user_access_changed', current_user()['account_id'], target=account_id,
                  previous={key: users[account_id][key] for key in ('role','approved','active')},
                  new={key: proposed[key] for key in ('role','approved','active')})
        return jsonify(user=proposed)

    @app.get('/api/auth/audit')
    @admin_api
    def epr_audit_log():
        with _lock:
            with audit_file.open(encoding='utf-8') if audit_file.exists() else open(os.devnull) as stream:
                rows = [json.loads(line) for line in deque(stream, maxlen=200)]
        return jsonify(events=list(reversed(rows)))
