"""Report and erase the OAuth integration's locally retained account data."""
from datetime import datetime, timezone
import json
import logging
from pathlib import Path
import re
import threading
import time

import requests

LOGGER = logging.getLogger(__name__)
_workers = set()
_lock = threading.Lock()
ACCESS_LOCK = threading.RLock()


def iso_time(timestamp):
    return datetime.fromtimestamp(timestamp, timezone.utc).isoformat().replace('+00:00', 'Z')


def session_files(directory):
    return [path for path in directory.glob('*.json') if re.fullmatch(r'[a-f0-9]{64}\.json', path.name)]


def erase_oauth_account(directory, account_id, erase_access=False):
    """Erase app-owned profile references while keeping non-personal deduplication receipts."""
    from jira_oauth import save_private
    if erase_access:
        with ACCESS_LOCK:
            _erase_access_account(directory, account_id)
    state_path = directory / 'privacy-report-state.json'
    if state_path.exists():
        state = json.loads(state_path.read_text(encoding='utf-8'))
        state.get('accounts', {}).pop(account_id, None)
        save_private(state_path, state)
    for path in session_files(directory):
        saved = json.loads(path.read_text(encoding='utf-8'))
        if saved.get('account_id') == account_id:
            path.unlink(missing_ok=True)
    for path in directory.parent.glob('*.json'):
        if not re.fullmatch(r'[a-f0-9]{32}\.json', path.name):
            continue
        saved = json.loads(path.read_text(encoding='utf-8'))
        if (saved.get('review') or {}).get('account_id') == account_id:
            save_private(path, {key: saved[key] for key in ('worklog_id', 'issue_key', 'verified', 'request_id', 'error') if key in saved})


def _erase_access_account(directory, account_id):
    from jira_oauth import save_private
    access_file = directory.parent / 'epr_access' / 'users.json'
    if access_file.exists():
        access = json.loads(access_file.read_text(encoding='utf-8'))
        access.pop(account_id, None)
        save_private(access_file, access)
        audit_file = access_file.parent / 'audit.jsonl'
        if audit_file.exists():
            rows = [json.loads(line) for line in audit_file.read_text(encoding='utf-8').splitlines() if line]
            for row in rows:
                for key in ('actor', 'target'):
                    if row.get(key) == account_id:
                        row[key] = None
            temp = audit_file.with_suffix('.tmp')
            temp.write_text(''.join(json.dumps(row)+'\n' for row in rows), encoding='utf-8')
            temp.chmod(0o600)
            temp.replace(audit_file)


def cycle_seconds(header):
    value = str(header or '').strip()
    if value.isdigit():
        return max(1, int(value))
    match = re.fullmatch(r'P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?', value)
    if match and any(match.groups()):
        return max(1, sum(int(number or 0)*scale for number, scale in zip(match.groups(), (86400, 3600, 60, 1))))
    return 7*86400


def report_accounts(directory: Path, now=None):
    """One resumable reporting pass. No canonical reporting DB or Jira worklogs are changed."""
    from jira_oauth import refresh_oauth_file, save_private
    now = time.time() if now is None else now
    state_path = directory / 'privacy-report-state.json'
    state = json.loads(state_path.read_text(encoding='utf-8')) if state_path.exists() else {'accounts': {}}
    if now < state.get('retry_at', 0):
        return {'deferred': True}
    accounts, token = {}, None
    for path in session_files(directory):
        if not path.exists():
            continue
        saved = json.loads(path.read_text(encoding='utf-8'))
        account = saved.get('account_id')
        if not account or account == 'unknown':
            continue
        if saved.get('session_expires_at', float('inf')) <= now:
            erase_oauth_account(directory, account)
            continue
        accounts[account] = min(accounts.get(account, now), saved.get('personal_data_retrieved_at', now))
        if token is None:
            try:
                token = refresh_oauth_file(path)['access_token']
            except (requests.RequestException, ValueError, KeyError):
                if not path.exists():
                    accounts.pop(account, None)
                continue
    # Worklog receipt profile fields are also covered, using their original retrieval time.
    for path in directory.parent.glob('*.json'):
        if not re.fullmatch(r'[a-f0-9]{32}\.json', path.name):
            continue
        saved = json.loads(path.read_text(encoding='utf-8'))
        review = saved.get('review') or {}
        account = review.get('account_id')
        if account and account != 'unknown':
            retrieved = review.get('personal_data_retrieved_at', path.stat().st_mtime)
            accounts[account] = min(accounts.get(account, now), retrieved)
    access_file = directory.parent / 'epr_access' / 'users.json'
    if access_file.exists():
        for account, user in json.loads(access_file.read_text(encoding='utf-8')).items():
            accounts[account] = min(accounts.get(account, now), user.get('personal_data_retrieved_at', now))
    state['accounts'] = {account: value for account, value in state['accounts'].items() if account in accounts}
    pending = [account for account in accounts if now >= state['accounts'].get(account, {}).get('next_report_at', 0)]
    if not token or not pending:
        return {'reported': 0}
    reported = 0
    for start in range(0, len(pending), 90):
        batch = pending[start:start+90]
        response = requests.post('https://api.atlassian.com/app/report-accounts/',
            headers={'Authorization': 'Bearer ' + token, 'Accept': 'application/json'},
            json={'accounts': [{'accountId': account, 'updatedAt': iso_time(accounts[account])} for account in batch]}, timeout=(10, 30))
        if response.status_code == 429:
            state['retry_at'] = now + max(1, int(response.headers.get('Retry-After', '3600')))
            save_private(state_path, state)
            return {'reported': reported, 'rate_limited': True}
        response.raise_for_status()
        instructions = response.json().get('accounts', []) if response.status_code == 200 else []
        erase = {item['accountId'] for item in instructions if item.get('status') in ('closed', 'updated')}
        for account in batch:
            if account in erase:
                erase_oauth_account(directory, account, erase_access=True)
                state['accounts'].pop(account, None)
            else:
                state['accounts'][account] = {'next_report_at': now + cycle_seconds(response.headers.get('Cycle-Period'))}
        state.pop('retry_at', None)
        save_private(state_path, state)
        reported += len(batch)
    return {'reported': reported}


def start_privacy_worker(directory):
    """The app has one production worker; an hourly wake honors each account's API cycle."""
    key = str(directory.resolve())
    with _lock:
        if key in _workers:
            return
        _workers.add(key)
    def run():
        while True:
            try:
                report_accounts(directory)
            except Exception:
                # Do not log response bodies, account IDs, or OAuth tokens.
                LOGGER.warning('Jira personal-data reporting needs attention; retrying on the next hourly pass.')
            time.sleep(3600)
    threading.Thread(target=run, name='jira-personal-data-reporting', daemon=True).start()
