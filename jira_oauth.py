"""Atlassian 3LO sign-in; browser holds only an opaque session ID."""
import json
import os
from pathlib import Path
import re
import secrets
import time
from urllib.parse import urlencode, urlsplit
from flask import jsonify, redirect, request
import requests
from jira_client import BASE_URL

COOKIE = 'epr_jira_session'
CALLBACK = '/api/jira-auth/callback'


def config():
    return {key: os.getenv('JIRA_OAUTH_' + key, '').strip() for key in ('CLIENT_ID', 'CLIENT_SECRET', 'REDIRECT_URI')}


def same_origin():
    origin = request.headers.get('Origin')
    if not origin:
        return True
    callback = urlsplit(config()['REDIRECT_URI'])
    expected = (callback.scheme + '://' + callback.netloc) if callback.scheme and callback.netloc else request.host_url.rstrip('/')
    return origin.rstrip('/') == expected


def save_private(path, data):
    temp = path.with_name(path.name + '.' + secrets.token_hex(8) + '.tmp')
    fd = os.open(temp, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
    with os.fdopen(fd, 'w', encoding='utf-8') as stream:
        json.dump(data, stream)
    os.replace(temp, path)


def oauth_session(directory):
    sid = request.cookies.get(COOKIE, '')
    if not sid:
        return None
    if not re.fullmatch('[a-f0-9]{64}', sid):
        raise ValueError('Your Jira sign-in is invalid. Sign in again.')
    path = directory / (sid + '.json')
    if not path.exists():
        raise ValueError('Your Jira sign-in expired. Sign in again.')
    saved = refresh_oauth_file(path)
    session = requests.Session()
    session.headers.update({'Authorization': 'Bearer ' + saved['access_token'], 'Accept': 'application/json', 'Content-Type': 'application/json'})
    session.epr_base_url = 'https://api.atlassian.com/ex/jira/' + saved['cloud_id']
    return session


def refresh_oauth_file(path):
    saved = json.loads(path.read_text(encoding='utf-8'))
    if saved.get('session_expires_at', float('inf')) <= time.time():
        from jira_privacy import erase_oauth_account
        erase_oauth_account(path.parent, saved.get('account_id'))
        raise ValueError('Your Jira sign-in expired. Sign in again.')
    if saved.get('expires_at', 0) <= time.time() + 60:
        cfg = config()
        response = requests.post('https://auth.atlassian.com/oauth/token', json={
            'grant_type': 'refresh_token', 'client_id': cfg['CLIENT_ID'], 'client_secret': cfg['CLIENT_SECRET'],
            'refresh_token': saved.get('refresh_token', '')}, timeout=(10, 25))
        if response.status_code in (400, 401):
            from jira_privacy import erase_oauth_account
            erase_oauth_account(path.parent, saved.get('account_id'))
        response.raise_for_status()
        tokens = response.json()
        saved.update(tokens)
        saved['expires_at'] = time.time() + tokens['expires_in']
        save_private(path, saved)
    return saved


def register_jira_oauth(app, directory: Path):
    if all(config().values()) and not app.testing:
        from jira_privacy import start_privacy_worker
        start_privacy_worker(directory)
    @app.get('/api/jira-auth/status')
    def jira_auth_status():
        return jsonify(configured=all(config().values()), signed_in=bool(request.cookies.get(COOKIE)), login_url='/api/jira-auth/login')

    @app.get('/api/jira-auth/login')
    def jira_auth_login():
        cfg = config()
        if not all(cfg.values()):
            return 'Jira sign-in needs an Atlassian OAuth app. Configure JIRA_OAUTH_CLIENT_ID, JIRA_OAUTH_CLIENT_SECRET and JIRA_OAUTH_REDIRECT_URI in Azure. No account was connected.', 503
        state = secrets.token_hex(32)
        url = 'https://auth.atlassian.com/authorize?' + urlencode({
            'audience': 'api.atlassian.com', 'client_id': cfg['CLIENT_ID'],
            'scope': 'read:jira-user read:jira-work write:jira-work report:personal-data offline_access',
            'redirect_uri': cfg['REDIRECT_URI'], 'state': state, 'response_type': 'code', 'prompt': 'consent'})
        response = redirect(url)
        response.set_cookie('epr_jira_state', state, max_age=600, httponly=True, secure=cfg['REDIRECT_URI'].startswith('https://'), samesite='Lax')
        return response

    @app.get(CALLBACK)
    def jira_auth_callback():
        state, expected = request.args.get('state', ''), request.cookies.get('epr_jira_state', '')
        if not state or not expected or not secrets.compare_digest(state, expected):
            return 'Jira sign-in state is invalid or expired. Start sign-in again.', 400
        if not request.args.get('code'):
            return 'Jira sign-in was cancelled. No account was connected.', 400
        cfg = config()
        try:
            response = requests.post('https://auth.atlassian.com/oauth/token', json={
                'grant_type': 'authorization_code', 'client_id': cfg['CLIENT_ID'], 'client_secret': cfg['CLIENT_SECRET'],
                'code': request.args['code'], 'redirect_uri': cfg['REDIRECT_URI']}, timeout=(10, 25))
            response.raise_for_status()
            tokens = response.json()
            resources = requests.get('https://api.atlassian.com/oauth/token/accessible-resources', headers={
                'Authorization': 'Bearer ' + tokens['access_token']}, timeout=(10, 25))
            resources.raise_for_status()
            resource = next((item for item in resources.json() if item.get('url', '').rstrip('/') == BASE_URL.rstrip('/')), None)
            if not resource or not {'read:jira-user', 'read:jira-work', 'write:jira-work'}.issubset(set(resource.get('scopes', []))):
                return 'Grant access to the report’s Jira site with read and worklog permissions, then try again.', 403
            sid = secrets.token_hex(32)
            directory.mkdir(parents=True, exist_ok=True)
            identity_response = requests.get('https://api.atlassian.com/ex/jira/' + resource['id'] + '/rest/api/3/myself',
                headers={'Authorization': 'Bearer ' + tokens['access_token']}, timeout=(10, 25))
            identity_response.raise_for_status()
            tokens.update(cloud_id=resource['id'], expires_at=time.time() + tokens['expires_in'],
                          account_id=identity_response.json()['accountId'], personal_data_retrieved_at=time.time(), session_expires_at=time.time()+30*86400)
            save_private(directory / (sid + '.json'), tokens)
            response = redirect('/employee_capacity_utilization_report.html')
            response.set_cookie(COOKIE, sid, httponly=True, secure=cfg['REDIRECT_URI'].startswith('https://'), samesite='Lax', max_age=30*86400)
            response.delete_cookie('epr_jira_state')
            return response
        except (KeyError, ValueError, requests.RequestException):
            return 'Could not connect Jira. Sign in again; no worklog was created.', 502

    @app.post('/api/jira-auth/logout')
    def jira_auth_logout():
        if not same_origin():
            return jsonify(error='Sign out from this site.'), 403
        sid = request.cookies.get(COOKIE, '')
        if re.fullmatch('[a-f0-9]{64}', sid):
            path = directory / (sid + '.json')
            if path.exists():
                from jira_privacy import erase_oauth_account
                saved = json.loads(path.read_text(encoding='utf-8'))
                if saved.get('account_id'):
                    erase_oauth_account(directory, saved['account_id'])
                path.unlink(missing_ok=True)
        response = jsonify(ok=True)
        response.delete_cookie(COOKIE)
        return response
