import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch
from jira_privacy import report_accounts, erase_oauth_account, cycle_seconds


class JiraPrivacyTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name) / 'oauth'
        self.directory.mkdir()
        self.session = self.directory / ('a'*64+'.json')
        self.session.write_text(json.dumps({'account_id': 'person', 'personal_data_retrieved_at': 100, 'access_token': 'secret', 'expires_at': 1000000000000}), encoding='utf-8')
        self.receipt = self.directory.parent / ('b'*32+'.json')
        self.receipt.write_text(json.dumps({'worklog_id': '1', 'review': {'account_id': 'person', 'author': 'Name', 'comment': 'Private', 'personal_data_retrieved_at': 90}, 'worklog': {'worklog_author': 'Name'}}), encoding='utf-8')

    def test_reporting_honors_cycle_and_oldest_data_age(self):
        response = Mock(status_code=204, headers={'Cycle-Period': 'P7D'})
        with patch('jira_privacy.requests.post', return_value=response) as post:
            self.assertEqual(report_accounts(self.directory, now=1000)['reported'], 1)
            body = post.call_args.kwargs['json']
            self.assertEqual(body['accounts'][0], {'accountId': 'person', 'updatedAt': '1970-01-01T00:01:30Z'})
            self.assertEqual(report_accounts(self.directory, now=2000)['reported'], 0)
            post.assert_called_once()

    def test_closed_and_updated_erase_profiles_but_keep_worklog_id(self):
        for status in ('closed', 'updated'):
            with self.subTest(status=status):
                self.setUp()
                response = Mock(status_code=200, headers={})
                response.json.return_value = {'accounts': [{'accountId': 'person', 'status': status}]}
                with patch('jira_privacy.requests.post', return_value=response):
                    report_accounts(self.directory, now=1000)
                self.assertFalse(self.session.exists())
                self.assertEqual(json.loads(self.receipt.read_text()), {'worklog_id': '1'})
                self.assertNotIn('person', (self.directory/'privacy-report-state.json').read_text())

    def test_rate_limit_and_expired_session(self):
        response = Mock(status_code=429, headers={'Retry-After': '180'})
        with patch('jira_privacy.requests.post', return_value=response) as post:
            self.assertTrue(report_accounts(self.directory, now=1000)['rate_limited'])
            self.assertTrue(report_accounts(self.directory, now=1100)['deferred'])
            post.assert_called_once()
        saved = json.loads(self.session.read_text())
        saved['session_expires_at'] = 1200
        self.session.write_text(json.dumps(saved), encoding='utf-8')
        with patch('jira_privacy.requests.post') as post:
            report_accounts(self.directory, now=1300)
            self.assertFalse(self.session.exists())
            post.assert_not_called()

    def test_cycle_header(self):
        self.assertEqual(cycle_seconds('PT12H'), 43200)
        self.assertEqual(cycle_seconds('3600'), 3600)
        self.assertEqual(cycle_seconds(None), 604800)

