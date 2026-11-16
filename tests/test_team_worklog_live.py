from datetime import date
from unittest import TestCase
from unittest.mock import Mock, patch

from team_worklog_live import read_team_worklog_live
from worklog_tally import _live_entries


class TeamWorklogLiveTests(TestCase):
    @patch('team_worklog_live._live_entries')
    @patch('team_worklog_live._snapshot')
    @patch('team_worklog_live._account')
    def test_live_worklogs_include_assignee_and_exclude_moved_month(self, account, snapshot, read):
        account.return_value = {'accountId': 'alice'}
        snapshot.return_value = {'entries': [{'issue_key': 'APP-1'}]}
        read.return_value = [{'date': '2026-10-05'}, {'date': '2026-11-01'}]
        session = Mock()
        result = read_team_worklog_live('unused.db', 'Alice', '2026-10', 'worklogs', session=session)
        self.assertEqual(result['values'], [{'date': '2026-10-05'}])
        self.assertTrue(read.call_args.kwargs['include_assignee'])

    @patch('team_worklog_live.build_live_rlt_leaves')
    @patch('team_worklog_live._account')
    def test_leave_uses_existing_live_leave_calculation(self, account, leave):
        account.return_value = {'accountId': 'alice'}
        leave.return_value = [{'assignee': 'Alice', 'period_day': '2026-10-05'}]
        session = Mock()
        result = read_team_worklog_live('unused.db', 'Alice', '2026-10', 'leave', session=session)
        self.assertEqual(result['values'], leave.return_value)
        leave.assert_called_once_with('Alice', '2026-10', 'alice', session=session)

    def test_live_reader_requires_assignee_metadata(self):
        session = Mock()
        def fake_get(path, **kwargs):
            response = Mock()
            response.json.return_value = {'issues': [{'key': 'APP-1', 'fields': {'issuetype': {'name': 'Sub-task'}}}], 'isLast': True}
            return response
        session.get.side_effect = fake_get
        with self.assertRaisesRegex(RuntimeError, 'assignee metadata'):
            _live_entries(session, 'alice', date(2026,10,1), date(2026,10,31), [], include_assignee=True)
