import json
from pathlib import Path
import tempfile
import unittest
import time
from concurrent.futures import ThreadPoolExecutor
from unittest.mock import Mock, patch
from jira_privacy import report_accounts, erase_oauth_account, cycle_seconds
from jira_oauth import refresh_oauth_file


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

    def test_expired_session_does_not_remove_another_session(self):
        other = self.directory / ('c'*64+'.json')
        data = json.loads(self.session.read_text())
        other.write_text(json.dumps({**data, 'session_expires_at': time.time()+5000}), encoding='utf-8')
        self.session.write_text(json.dumps({**data, 'session_expires_at': 1200}), encoding='utf-8')
        with patch('jira_privacy.requests.post', return_value=Mock(status_code=204, headers={})):
            report_accounts(self.directory, now=1300)
        self.assertFalse(self.session.exists())
        self.assertTrue(other.exists())

    def test_rotating_token_is_refreshed_only_once_across_threads(self):
        data=json.loads(self.session.read_text())
        data.update(expires_at=0, refresh_token='old', session_expires_at=time.time()+3600)
        self.session.write_text(json.dumps(data),encoding='utf-8')
        reply=Mock(status_code=200)
        reply.json.return_value={'access_token':'new','refresh_token':'rotated','expires_in':3600}
        with patch('jira_oauth.requests.post',return_value=reply) as post:
            with ThreadPoolExecutor(max_workers=4) as pool:
                results=list(pool.map(lambda _: refresh_oauth_file(self.session),range(4)))
        self.assertEqual(post.call_count,1)
        self.assertTrue(all(result['refresh_token']=='rotated' for result in results))

    def test_invalid_grant_only_removes_the_failed_session(self):
        other=self.directory/('c'*64+'.json')
        data=json.loads(self.session.read_text())
        data.update(expires_at=0, refresh_token='old', session_expires_at=time.time()+3600)
        self.session.write_text(json.dumps(data),encoding='utf-8')
        other.write_text(json.dumps(data),encoding='utf-8')
        reply=Mock(status_code=400)
        reply.json.return_value={'error':'invalid_grant'}
        reply.raise_for_status.side_effect=__import__('requests').HTTPError(response=reply)
        with patch('jira_oauth.requests.post',return_value=reply):
            with self.assertRaises(__import__('requests').HTTPError):
                refresh_oauth_file(self.session)
        self.assertFalse(self.session.exists())
        self.assertTrue(other.exists())

    def test_other_refresh_errors_do_not_delete_sessions(self):
        data=json.loads(self.session.read_text())
        data.update(expires_at=0, refresh_token='old', session_expires_at=time.time()+3600)
        self.session.write_text(json.dumps(data),encoding='utf-8')
        reply=Mock(status_code=400)
        reply.json.return_value={'error':'invalid_client'}
        reply.raise_for_status.side_effect=__import__('requests').HTTPError(response=reply)
        with patch('jira_oauth.requests.post',return_value=reply):
            with self.assertRaises(__import__('requests').HTTPError):
                refresh_oauth_file(self.session)
        self.assertTrue(self.session.exists())

    def test_approval_survives_logout_but_privacy_erasure_removes_it(self):
        access = self.directory.parent / 'epr_access'
        access.mkdir()
        (access/'users.json').write_text(json.dumps({'person': {'account_id':'person','display_name':'Name'}}), encoding='utf-8')
        (access/'audit.jsonl').write_text(json.dumps({'actor':'person','target':'person','event':'user_access_changed'})+'\n', encoding='utf-8')
        erase_oauth_account(self.directory, 'person')
        self.assertIn('person', json.loads((access/'users.json').read_text()))
        erase_oauth_account(self.directory, 'person', erase_access=True)
        self.assertEqual(json.loads((access/'users.json').read_text()), {})
        self.assertNotIn('person', (access/'audit.jsonl').read_text())

    def test_updated_profile_preserves_approval_but_closed_removes_it(self):
        access=self.directory.parent/'epr_access';access.mkdir()
        users=access/'users.json'
        users.write_text(json.dumps({'person':{'account_id':'person','display_name':'Old Name','role':'member','approved':True,'active':True}}))
        response=Mock(status_code=200,headers={})
        response.json.return_value={'accounts':[{'accountId':'person','status':'updated'}]}
        with patch('jira_privacy.requests.post',return_value=response):
            report_accounts(self.directory,now=1000)
        user=json.loads(users.read_text())['person']
        self.assertTrue(user['approved']);self.assertEqual(user['role'],'member')
        self.assertNotIn('display_name',user)
        self.assertFalse(self.session.exists())
        self.assertNotIn('author',json.loads(self.receipt.read_text()).get('review',{}))
        erase_oauth_account(self.directory,'person',erase_access=True)
        self.assertNotIn('person',json.loads(users.read_text()))

