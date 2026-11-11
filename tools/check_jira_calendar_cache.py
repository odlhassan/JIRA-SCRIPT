"""Read-only real Jira smoke test for Hassan's calendar and unchanged-page reuse."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from jira_calendar_cache import calendar_read
from jira_client import get_session
from worklog_tally import build_worklog_tally
from live_rlt_leave import build_live_rlt_leaves

session = get_session()
calls = []
original_get = session.get
def counted_get(url, **kwargs):
    calls.append(url.split('/rest/api/3/')[-1])
    return original_get(url, **kwargs)
session.get = counted_get
db = Path(__file__).resolve().parents[1] / 'assignee_hours_capacity.db'
def build(reader):
    result = build_worklog_tally(db, 'Hassan Malik', '2026-10', reader)
    result['leaves'] = build_live_rlt_leaves(result['employee'], result['month'], result['account_id'], reader)
    return result
for label, force in [('first read', False), ('shared result', False), ('explicit check', True)]:
    before = len(calls)
    result = calendar_read(db, 'Hassan Malik', '2026-10', build, session, force=force)
    pages = sum(path.endswith('/worklog') for path in calls[before:])
    print(label, result['sync'], 'worklog_pages:', pages, flush=True)
    if label == 'shared result': assert result['sync']['network_reads'] == 0
    if label == 'explicit check': assert pages == 0
