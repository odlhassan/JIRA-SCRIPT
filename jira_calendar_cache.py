"""Transactional, bounded, process-local Jira read cache for the single-worker server."""
from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json
import re
import threading
import time
from urllib.parse import urlsplit

from jira_client import BASE_URL, get_session
from worklog_tally import _get, month_bounds

LOCK = threading.RLock()
STORES = {}
CHECK_SECONDS = 30
PAGE_SECONDS = 900


def _signature(result):
    return hashlib.sha256(json.dumps({'entries': result['entries'], 'leaves': result['leaves']}, sort_keys=True).encode()).hexdigest()


class CachedReply:
    status_code = 200
    headers = {}
    def __init__(self, data): self.data = data
    def json(self): return deepcopy(self.data)
    def raise_for_status(self): pass


class ReadSession:
    def __init__(self, session, state, now):
        self.session, self.state, self.now = session, state, now
        self.network_reads = 0

    def get(self, url, **kwargs):
        path = urlsplit(url).path.split('/rest/api/3/', 1)[-1]
        params = kwargs.get('params') or {}
        key = json.dumps([path, params], sort_keys=True)
        match = re.fullmatch(r'issue/([A-Z][A-Z0-9_]*-\d+)/worklog', path)
        ttl = PAGE_SECONDS if match else 3600 if path in ('myself', 'user/search') else 0
        cached = self.state['pages'].get(key)
        if ttl and cached and self.now - cached['at'] < ttl:
            return CachedReply(cached['data'])
        self.network_reads += 1
        response = self.session.get(url, **kwargs)
        if response.status_code >= 400:
            return response  # Let the existing GET retry/permission handling run.
        response.raise_for_status()
        data = response.json()
        if path == 'search/jql' and isinstance(data, dict):
            for issue in data.get('issues', []):
                issue_key, identifier = issue.get('key'), str(issue.get('id', ''))
                if identifier: self.state['issues'][identifier] = issue_key
                # All callers re-read discovery metadata. Fingerprints also cover
                # leave assignment, dates, type and estimates, not just time spent.
                fields = issue.get('fields') or {}
                fingerprint_key = (issue_key, params.get('fields', ''))
                previous = self.state['metadata'].get(fingerprint_key)
                if previous is not None and previous != fields:
                    self.evict(issue_key)
                self.state['metadata'][fingerprint_key] = deepcopy(fields)
        elif re.fullmatch(r'issue/[A-Z][A-Z0-9_]*-\d+', path) and isinstance(data, dict) and data.get('id'):
            self.state['issues'][str(data['id'])] = path.split('/')[1]
        if match:
            for row in data.get('worklogs', []):
                self.state['logs'][str(row.get('id'))] = match[1]
        if ttl:
            self.state['pages'][key] = {'at': self.now, 'data': deepcopy(data), 'issue': match[1] if match else None}
        return CachedReply(data)

    def evict(self, issue):
        self.state['pages'] = {key: row for key, row in self.state['pages'].items() if row['issue'] != issue}

    def changes(self):
        if self.now - self.state['feed_at'] < CHECK_SECONDS: return
        since = self.state['cursor']
        updates, deleted, untils = [], [], []
        for kind, target in [('updated', updates), ('deleted', deleted)]:
            position, seen = since, set()
            for _ in range(200):
                page = _get(self, 'worklog/' + kind, since=position)
                if not isinstance(page, dict) or not isinstance(page.get('values'), list) or type(page.get('until')) is not int:
                    raise RuntimeError('Jira change feed was incomplete; cached calendar was retained.')
                for row in page['values']:
                    identifier, stamp = str(row.get('worklogId', '')), row.get('updatedTime')
                    if not identifier.isdecimal() or type(stamp) is not int: raise RuntimeError('Invalid Jira change record.')
                    ledger_key = (kind, identifier)
                    if stamp > self.state['seen'].get(ledger_key, -1):
                        target.append(identifier)
                        self.state['seen'][ledger_key] = stamp
                if page.get('lastPage') is True:
                    untils.append(page['until']); break
                if page['until'] <= position or page['until'] in seen: raise RuntimeError('Jira change pagination repeated.')
                position = page['until']; seen.add(position)
            else: raise RuntimeError('Too many Jira change pages.')
        for identifier in set(deleted + updates):
            if identifier in self.state['logs']: self.evict(self.state['logs'][identifier])
        unknown = list(set(updates) - self.state['logs'].keys())
        for start in range(0, len(unknown), 1000):
            # Jira's bulk worklog POST is a read operation; it does not create logs.
            self.network_reads += 1
            response = self.session.post(BASE_URL + '/rest/api/3/worklog/list', json={'ids': [int(identifier) for identifier in unknown[start:start+1000]]}, timeout=(10, 30))
            response.raise_for_status(); rows = response.json()
            if not isinstance(rows, list): raise RuntimeError('Jira changed worklog lookup failed.')
            for row in rows:
                issue = self.state['issues'].get(str(row.get('issueId', '')))
                if issue: self.evict(issue)
        # Overlap compensates for Jira's one-minute feed delay; the ledger avoids
        # repeatedly invalidating records returned in that overlap.
        self.state['cursor'] = max(0, min(untils) - 120000)
        self.state['feed_at'] = self.now
        self.state['seen'] = {key: stamp for key, stamp in self.state['seen'].items() if stamp >= self.state['cursor']}


def calendar_read(db, employee, month, builder, session=None, now=None, force=False):
    employee = str(employee or '').strip()
    month_bounds(month)
    if not employee or len(employee) > 200 or employee.casefold() in ('grand total', 'unassigned'):
        raise ValueError('Open one named employee calendar.')
    real_clock = now is None
    now = time.time() if real_clock else now
    session = session or get_session()
    auth = str(session.headers.get('Authorization', ''))
    namespace = (str(db), BASE_URL, hashlib.sha256(auth.encode()).hexdigest())
    with LOCK:
        old = STORES.get(namespace) or {'pages': {}, 'issues': {}, 'logs': {}, 'metadata': {}, 'results': {}, 'seen': {}, 'cursor': int((now-120)*1000), 'feed_at': now}
        result_key = (employee.casefold(), month)
        cached = old['results'].get(result_key)
        if cached and not force and now - cached['checked'] < CHECK_SECONDS:
            result = deepcopy(cached['result'])
            result['sync'] = {**result['sync'], 'mode': 'cached', 'network_reads': 0}
            return result
        state = deepcopy(old)
        reader = ReadSession(session, state, now)
        reader.changes()
        result = builder(reader)
        if real_clock:
            now = time.time()
        signature = _signature(result)
        same = cached is not None and signature == cached['signature']
        checked_at = datetime.fromtimestamp(now, timezone.utc).isoformat()
        result['sync'] = {'mode': 'unchanged' if same else 'updated', 'checked_at': checked_at,
                          'changed_at': cached['result']['sync']['changed_at'] if same else checked_at,
                          'network_reads': reader.network_reads}
        state['results'][result_key] = {'result': deepcopy(result), 'checked': now, 'signature': signature}
        # Bound memory and discard old data. A restart simply does a fresh first read.
        state['pages'] = {key: row for key, row in state['pages'].items() if now-row['at'] < 3600}
        if len(state['pages']) > 4000 or len(state['results']) > 100 or len(state['metadata']) > 8000 or len(state['logs']) > 100000 or len(state['issues']) > 20000 or len(state['seen']) > 100000:
            state = {'pages': {}, 'issues': {}, 'logs': {}, 'metadata': {}, 'results': {result_key:state['results'][result_key]}, 'seen': {}, 'cursor':int((now-120)*1000), 'feed_at':now}
        if len(STORES) >= 4 and namespace not in STORES: STORES.clear()
        STORES[namespace] = state
        return result


def invalidate_issue(issue):
    """After an EPR mutation, every employee/month rechecks and the task re-reads."""
    with LOCK:
        for state in STORES.values():
            state['pages'] = {key: row for key, row in state['pages'].items() if row['issue'] != issue}
            for result in state['results'].values(): result['checked'] = 0
