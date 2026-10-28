"""Reviewed Jira writes with durable one-attempt receipts; no database schema changes."""
from datetime import datetime
from decimal import Decimal
from functools import wraps
import json
import os
from pathlib import Path
import re
import secrets
import time
import threading
import uuid

from flask import jsonify, request, has_request_context, current_app
from jira_oauth import oauth_session, register_jira_oauth, same_origin
from itsdangerous import URLSafeTimedSerializer, BadData
import requests
from jira_client import BASE_URL, get_session, get_write_session

_submit_lock = threading.Lock()


def serialize_submit(function):
    @wraps(function)
    def serialized(*args, **kwargs):
        with _submit_lock:
            return function(*args, **kwargs)
    return serialized


def parse_duration(value):
    """Calendar durations: one day is 24 hours, with whole-minute precision."""
    if not isinstance(value, str) or len(value) > 100:
        raise ValueError('Enter time such as 30m, 8h, 1d, or 1h 30m.')
    text = value.strip().lower()
    number = r'\d+(?:\.\d+)?'
    if re.fullmatch(number, text):
        minutes = Decimal(text)
    elif re.fullmatch(r'(?:' + number + r'\s*[dhm]\s*)+', text):
        minutes, seen = Decimal(0), set()
        for amount, unit in re.findall(r'(' + number + r')\s*([dhm])', text):
            if unit in seen:
                raise ValueError('Use each unit once, for example 1h 30m.')
            seen.add(unit)
            minutes += Decimal(amount) * {'d': 1440, 'h': 60, 'm': 1}[unit]
    else:
        raise ValueError('Enter time such as 30m, 8h, 1d, or 1h 30m.')
    if minutes <= 0 or minutes > 1440:
        raise ValueError('Time must be greater than zero and no more than 24h (1d).')
    if minutes != minutes.to_integral_value():
        raise ValueError('Time must resolve to whole minutes, for example 0.5h or 30m.')
    return int(minutes)


def duration_label(minutes):
    hours, remainder = divmod(minutes, 60)
    return ' '.join(part for part in (f'{hours}h' if hours else '', f'{remainder}m' if remainder else '') if part) or '0m'


def daily_logged_seconds(session, issue_key, account_id, started):
    """Read all Jira pages; compare dates in the selected calendar's UTC offset."""
    offset, seconds, seen = 0, 0, set()
    while True:
        response = session.get(jira_base(session) + '/rest/api/3/issue/' + issue_key + '/worklog',
                               params={'startAt': offset, 'maxResults': 100}, timeout=(10, 25))
        response.raise_for_status()
        payload = response.json()
        if not isinstance(payload, dict):
            raise ValueError('Could not verify the day’s existing Jira hours. No worklog was created.')
        rows, total = payload.get('worklogs'), payload.get('total')
        if not isinstance(rows, list) or not isinstance(total, int) or total < 0:
            raise ValueError('Could not verify the day’s existing Jira hours. No worklog was created.')
        for row in rows:
            if (row.get('author') or {}).get('accountId') != account_id:
                continue
            identifier = row.get('id')
            if identifier and identifier in seen:
                continue
            if identifier:
                seen.add(identifier)
            logged_at = datetime.fromisoformat(str(row.get('started', '')))
            if not logged_at.tzinfo:
                raise ValueError('Jira returned a worklog without a timezone; daily hours could not be verified.')
            if logged_at.astimezone(started.tzinfo).date() == started.date():
                amount = row.get('timeSpentSeconds')
                if isinstance(amount, bool) or not isinstance(amount, int) or amount < 0:
                    raise ValueError('Jira returned invalid worklog hours; daily hours could not be verified.')
                seconds += amount
        offset += len(rows)
        if offset >= total:
            return seconds
        if not rows:
            raise ValueError('Could not read every Jira worklog page. No worklog was created.')


def check_daily_limit(session, issue_key, account_id, started, minutes):
    existing = daily_logged_seconds(session, issue_key, account_id, started)
    if existing + minutes * 60 > 86400:
        remaining = max(0, (86400-existing)//60)
        raise ValueError(f'This would exceed 24h for you on {issue_key} on {started.date()}. '
                         f'Already logged: {existing/3600:g}h. You can add at most {remaining} minutes.')
    return existing


def jira_base(session):
    return getattr(session, 'epr_base_url', BASE_URL)


def jira_identity():
    if has_request_context() and current_app.config.get('EPR_JIRA_OAUTH_DIRECTORY') is not None:
        session = oauth_session(current_app.config['EPR_JIRA_OAUTH_DIRECTORY'])
        if session is not None:
            response = session.get(jira_base(session) + '/rest/api/3/myself', timeout=(10, 25))
            response.raise_for_status()
            return session, response.json()
    for factory in (get_write_session, get_session):
        try:
            session = factory()
            response = session.get(jira_base(session) + '/rest/api/3/myself', timeout=(10, 25))
            response.raise_for_status()
            identity = response.json()
            if identity.get('accountId'):
                return session, identity
        except (ValueError, requests.RequestException):
            continue
    raise ValueError('No valid Jira account is configured for work logging.')


def checked_issue(session, issue_key):
    if not re.fullmatch(r'[A-Z][A-Z0-9_]*-\d+', issue_key):
        raise ValueError('Choose a valid Jira subtask.')
    response = session.get(jira_base(session) + '/rest/api/3/issue/' + issue_key,
                           params={'fields': 'summary,issuetype,assignee,status'}, timeout=(10, 25))
    response.raise_for_status()
    issue = response.json()
    fields = issue.get('fields') or {}
    if not (fields.get('issuetype') or {}).get('subtask') or issue_key.startswith('RLT-'):
        raise ValueError('Only non-leave Jira subtasks can receive work here.')
    response = session.get(jira_base(session) + '/rest/api/3/mypermissions', params={
        'issueKey': issue_key, 'permissions': 'BROWSE_PROJECTS,WORK_ON_ISSUES'}, timeout=(10, 25))
    response.raise_for_status()
    permissions = response.json().get('permissions') or {}
    if not all((permissions.get(key) or {}).get('havePermission') for key in ('BROWSE_PROJECTS', 'WORK_ON_ISSUES')):
        raise ValueError('The connected Jira account cannot log work on this subtask.')
    return fields


def register_employee_worklog_routes(app, receipt_dir: Path):
    app.config['EPR_JIRA_OAUTH_DIRECTORY'] = receipt_dir / 'oauth'
    register_jira_oauth(app, receipt_dir / 'oauth')
    prefix = '/api/employee-capacity-utilization/worklog'

    def signer():
        receipt_dir.mkdir(parents=True, exist_ok=True)
        key_file = receipt_dir / 'review-secret'
        try:
            fd = os.open(key_file, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
            with os.fdopen(fd, 'w', encoding='utf-8') as stream:
                stream.write(secrets.token_hex(32))
        except FileExistsError:
            pass
        return URLSafeTimedSerializer(key_file.read_text(encoding='utf-8'), salt='employee-worklog-v1')

    def guard():
        if request.headers.get('X-EPR-Worklog') != '1' or not request.is_json:
            raise ValueError('Worklog requests must come from the review form.')
        if not same_origin():
            raise ValueError('Worklog requests must come from this site.')
        if not request.cookies.get('epr_jira_session') and os.getenv('EPR_WORKLOG_ALLOW_SHARED_ACCOUNT') != '1':
            raise ValueError('Sign in with Jira before reviewing or submitting work. Personal sign-in requires administrator OAuth setup.')

    def failure(exc):
        if isinstance(exc, requests.RequestException):
            status = exc.response.status_code if exc.response is not None else None
            return jsonify(error='Jira could not complete this request' + (f' (HTTP {status})' if status else '') + '. Verify Jira before retrying any submission.', jira_status=status), 502
        return jsonify(error=str(exc)), 400

    @app.after_request
    def private_jira_responses(response):
        if request.path.startswith((prefix, '/api/jira-auth/')):
            response.headers['Cache-Control'] = 'no-store'
        return response

    @app.get(prefix + '/identity')
    def worklog_identity():
        try:
            _, identity = jira_identity()
            return jsonify(author=identity['displayName'], account_id=identity['accountId'], signed_in=bool(request.cookies.get('epr_jira_session')),
                           write_enabled=bool(request.cookies.get('epr_jira_session')) or os.getenv('EPR_WORKLOG_ALLOW_SHARED_ACCOUNT') == '1')
        except (ValueError, requests.RequestException) as exc:
            return failure(exc)

    @app.post(prefix + '/review')
    def worklog_review():
        try:
            guard()
            body = request.get_json()
            session, identity = jira_identity()
            if str(body.get('employee', '')).strip().casefold() != identity['displayName'].strip().casefold():
                raise ValueError('Jira logs work as ' + identity['displayName'] + '. Open that employee’s calendar to continue.')
            minutes = parse_duration(body['duration']) if 'duration' in body else body.get('minutes')
            if isinstance(minutes, bool) or not isinstance(minutes, int) or not 1 <= minutes <= 1440:
                raise ValueError('Enter whole minutes between 1 and 1440.')
            started = datetime.fromisoformat(str(body.get('started', '')))
            if not started.tzinfo or started.utcoffset() is None:
                raise ValueError('A start time with a timezone is required.')
            if started.date() > datetime.now(started.tzinfo).date():
                raise ValueError('Work cannot be logged on a future date.')
            comment = str(body.get('comment', '')).strip()
            if not 1 <= len(comment) <= 2000:
                raise ValueError('Enter a work description of 1–2000 characters.')
            issue_key = str(body.get('issue_key', '')).strip().upper()
            fields = checked_issue(session, issue_key)
            existing_seconds = check_daily_limit(session, issue_key, identity['accountId'], started, minutes)
            review = dict(request_id=uuid.uuid4().hex, issue_key=issue_key, summary=fields.get('summary', ''),
                          account_id=identity['accountId'], author=identity['displayName'], minutes=minutes,
                          duration_label=duration_label(minutes), logged_minutes=existing_seconds/60,
                          total_minutes=(existing_seconds/60)+minutes,
                          started=started.strftime('%Y-%m-%dT%H:%M:%S.000%z'), comment=comment, personal_data_retrieved_at=time.time())
            return jsonify(review=review, token=signer().dumps(review))
        except (ValueError, TypeError, requests.RequestException) as exc:
            return failure(exc)

    @app.post(prefix + '/submit')
    @serialize_submit
    def worklog_submit():
        try:
            guard()
            body = request.get_json()
            if body.get('confirmed') is not True:
                raise ValueError('Confirm the reviewed employee, subtask, date, and duration first.')
            try:
                review = signer().loads(str(body.get('token', '')), max_age=600)
            except BadData:
                raise ValueError('The review expired or changed. Review the entry again.')
            receipt = receipt_dir / (review['request_id'] + '.json')
            if receipt.exists():
                result = json.loads(receipt.read_text(encoding='utf-8'))
                return jsonify(result), 200 if result.get('worklog_id') else 409
            session, identity = jira_identity()
            if identity['accountId'] != review['account_id']:
                raise ValueError('The connected Jira author changed. Review again.')
            fields = checked_issue(session, review['issue_key'])
            check_daily_limit(session, review['issue_key'], identity['accountId'], datetime.fromisoformat(review['started']), review['minutes'])
            # An exclusive file is the durable claim: a timed-out/parallel request never writes twice.
            try:
                fd = os.open(receipt, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
            except FileExistsError:
                return jsonify(error='Submission already in progress. Verify Jira before trying again.'), 409
            pending = {'error': 'Submission outcome needs verification in Jira; this request will not be resent.', 'request_id': review['request_id']}
            with os.fdopen(fd, 'w', encoding='utf-8') as stream:
                json.dump(pending, stream)
                stream.flush()
                os.fsync(stream.fileno())
            response = session.post(jira_base(session) + '/rest/api/3/issue/' + review['issue_key'] + '/worklog',
                params={'adjustEstimate': 'leave', 'notifyUsers': 'false'}, timeout=(10, 40), json={
                    'started': review['started'], 'timeSpentSeconds': review['minutes'] * 60,
                    'comment': {'type': 'doc', 'version': 1, 'content': [{'type': 'paragraph', 'content': [
                        {'type': 'text', 'text': review['comment']}]}]},
                    'properties': [{'key': 'epr.request', 'value': {'request_id': review['request_id']}}],
                })
            response.raise_for_status()
            worklog = response.json()
            if not worklog.get('id'):
                raise ValueError(pending['error'])
            result = dict(worklog_id=str(worklog['id']), issue_key=review['issue_key'], review=review,
                          author=(worklog.get('author') or {}).get('displayName', ''), verified=False)
            # Store the created ID before verifying so a failed verification cannot cause another POST.
            receipt.write_text(json.dumps(result), encoding='utf-8')
            try:
                check = session.get(jira_base(session) + '/rest/api/3/issue/' + review['issue_key'] + '/worklog/' + str(worklog['id']), timeout=(10, 25))
                check.raise_for_status()
                saved = check.json()
                result['verified'] = (saved.get('timeSpentSeconds') == review['minutes'] * 60 and
                    (saved.get('author') or {}).get('accountId') == review['account_id'])
            except (ValueError, requests.RequestException):
                pass
            result['worklog'] = dict(issue_id=review['issue_key'], worklog_author=review['author'],
                worklog_date=review['started'][:10], hours_logged=review['minutes']/60,
                item_issue_type=(fields.get('issuetype') or {}).get('name', 'Sub-task'),
                item_summary=fields.get('summary', ''), item_assignee=(fields.get('assignee') or {}).get('displayName', ''))
            receipt.write_text(json.dumps(result), encoding='utf-8')
            return jsonify(result)
        except (ValueError, TypeError, requests.RequestException) as exc:
            return failure(exc)
