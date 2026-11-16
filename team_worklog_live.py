"""Live, read-only employee worklog and leave data for the team matrix."""
from datetime import datetime, timezone
import re
import threading

from jira_client import get_session
from live_rlt_leave import build_live_rlt_leaves
from worklog_tally import _account, _live_entries, _snapshot, month_bounds

# The production server has one worker: bound Jira pressure across all viewers.
LIVE_READ_SLOTS = threading.BoundedSemaphore(2)


def read_team_worklog_live(db_path, employee, month, kind, session=None):
    employee = str(employee or '').strip()
    if not employee or len(employee) > 200 or employee.casefold() in ('grand total', 'unassigned'):
        raise ValueError('Choose one named employee.')
    first, last = month_bounds(month)
    if kind not in ('worklogs', 'leave'):
        raise ValueError('Invalid live data type.')
    session = session or get_session()
    account = _account(session, employee)
    account_id = account.get('accountId', '')
    if not re.fullmatch(r'[A-Za-z0-9:_-]+', account_id):
        raise RuntimeError('Jira returned an invalid employee account.')
    if kind == 'worklogs':
        snapshot = _snapshot(db_path, employee, month)
        values = _live_entries(session, account_id, first, last, snapshot['entries'], include_assignee=True)
        # Moved entries may be needed for reconciliation but belong to another month.
        values = [row for row in values if first.isoformat() <= row['date'] <= last.isoformat()]
    else:
        values = build_live_rlt_leaves(employee, month, account_id, session=session)
    return {'employee': employee, 'month': month, 'kind': kind, 'values': values,
            'checked_at': datetime.now(timezone.utc).isoformat()}
