"""Live RLT leave for the Availability calendar, using existing RLT business rules."""
from datetime import datetime
import re

from jira_client import get_session
from worklog_tally import _get, month_bounds
from generate_rlt_leave_report import (
    DEFAULT_START_DATE_FIELD, LEAVE_TYPE_FIELD, _normalize_subtasks, _normalize_worklogs,
    _redistribute_continuous_leave_worklogs, _compute_aggregates, _day_hours_profile_from_env,
)


def build_live_rlt_leaves(employee, month, account_id, session=None):
    first, last = month_bounds(month)
    if not re.fullmatch(r'[A-Za-z0-9:_-]+', str(account_id)):
        raise ValueError('Invalid Jira employee account.')
    session = session or get_session()
    fields = 'summary,issuetype,assignee,parent,status,created,updated,duedate,timeoriginalestimate,timespent,customfield_10015,' + DEFAULT_START_DATE_FIELD + ',' + LEAVE_TYPE_FIELD
    issues, cursor, seen = {}, None, set()
    for _ in range(200):
        params = dict(jql=f'project = RLT AND assignee = "{account_id}" ORDER BY key ASC', fields=fields, maxResults=100)
        if cursor:
            params['nextPageToken'] = cursor
        page = _get(session, 'search/jql', **params)
        if not isinstance(page, dict) or not isinstance(page.get('issues'), list):
            raise RuntimeError('RLT leave issue discovery was incomplete.')
        for issue in page['issues']:
            key = issue.get('key', '')
            meta = issue.get('fields') or {}
            if not re.fullmatch(r'RLT-\d+', key):
                raise RuntimeError('Invalid RLT issue returned by Jira.')
            if (meta.get('assignee') or {}).get('accountId') != account_id:
                continue
            if (meta.get('issuetype') or {}).get('subtask') is True:
                issues[key] = issue
        if page.get('isLast') is True:
            break
        cursor = page.get('nextPageToken')
        if not isinstance(cursor, str) or not cursor or cursor in seen:
            raise RuntimeError('RLT leave issue pagination was incomplete.')
        seen.add(cursor)
    else:
        raise RuntimeError('Too many RLT leave issue pages.')
    logs_by_issue = {}
    for key in issues:
        logs, offset, seen_ids = [], 0, set()
        for _ in range(1000):
            page = _get(session, 'issue/' + key + '/worklog', startAt=offset, maxResults=100)
            rows, total = page.get('worklogs'), page.get('total')
            if not isinstance(rows, list) or type(total) is not int or total < 0:
                raise RuntimeError('RLT leave worklog pagination was incomplete.')
            for log in rows:
                identifier = str(log.get('id', ''))
                if not identifier or identifier in seen_ids:
                    raise RuntimeError('RLT leave worklog page repeated or invalid.')
                seen_ids.add(identifier)
                started = datetime.fromisoformat(str(log.get('started', '')).replace('Z', '+00:00'))
                seconds = log.get('timeSpentSeconds')
                if started.utcoffset() is None or type(seconds) is not int or seconds < 0:
                    raise RuntimeError('Invalid RLT leave worklog date or hours.')
                logs.append(log)
            offset += len(rows)
            if offset >= total:
                break
            if not rows:
                raise RuntimeError('RLT leave worklog pagination ended early.')
        else:
            raise RuntimeError('Too many RLT worklog pages.')
        logs_by_issue[key] = _normalize_worklogs(key, logs)
    subtasks = _normalize_subtasks(list(issues.values()), {}, logs_by_issue, DEFAULT_START_DATE_FIELD)
    # Leave belongs to the RLT subtask's assignee, irrespective of the worklog author.
    for subtask in subtasks:
        subtask.assignee = employee
    logs = [row for rows in logs_by_issue.values() for row in rows]
    profile = _day_hours_profile_from_env()
    distributed = _redistribute_continuous_leave_worklogs(subtasks, logs, profile)
    return _compute_aggregates(subtasks, distributed, first.isoformat(), last.isoformat(), profile)['daily']
