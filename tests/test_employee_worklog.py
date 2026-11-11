import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch
from flask import Flask
import requests
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import urlparse, parse_qs
from employee_worklog import register_employee_worklog_routes, parse_duration, daily_logged_seconds
from datetime import datetime
from jira_client import BASE_URL

PREFIX = '/api/employee-capacity-utilization/worklog'
HEADERS = {'X-EPR-Worklog': '1'}


class EmployeeWorklogTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.app = Flask(__name__)
        register_employee_worklog_routes(self.app, Path(self.tmp.name))
        self.client = self.app.test_client()
        self.session = Mock()
        self.session.epr_base_url = 'https://jira.example'
        self.identity = {'displayName': 'Hassan Malik', 'accountId': 'hassan'}
        self.daily_logs = []
        def jira_get(url, **kwargs):
            response = Mock()
            response.json.return_value = ({'worklogs': self.daily_logs, 'total': len(self.daily_logs)} if url.endswith('/worklog')
                else {'timeSpentSeconds': 1800, 'author': self.identity})
            return response
        self.session.get.side_effect = jira_get
        self.fields = {'summary': 'Build calendar', 'issuetype': {'subtask': True, 'name': 'Sub-task'}, 'assignee': {'displayName': 'Hassan Malik'}}
        for patcher in [patch.dict(os.environ, {'EPR_WORKLOG_ALLOW_SHARED_ACCOUNT': '1'}),
                        patch('employee_worklog.jira_identity', return_value=(self.session, self.identity)),
                        patch('employee_worklog.checked_issue', return_value=self.fields)]:
            patcher.start()
            self.addCleanup(patcher.stop)
        self.body = {'employee': 'Hassan Malik', 'issue_key': 'APP-1', 'minutes': 30,
                     'started': '2026-10-01T09:00:00+05:00', 'comment': 'Test work'}

    def review(self):
        return self.client.post(PREFIX + '/review', json=self.body, headers=HEADERS)

    def personal_jira(self, rows=None, owner=None):
        rows = rows if rows is not None else [{'id': '123', 'author': self.identity, 'started': self.body['started'], 'timeSpentSeconds': 1800}]
        def get(url, **kwargs):
            response = Mock()
            if url.endswith('/myself'):
                response.json.return_value = self.identity
            elif url.endswith('/worklog'):
                offset = kwargs['params']['startAt']
                response.json.return_value = {'worklogs': rows[offset:offset+1], 'total': len(rows)}
            else:
                response.json.return_value = {'id': '123', 'author': owner or self.identity}
            return response
        self.session.get.side_effect = get
        self.session.delete.return_value.status_code = 204
        patcher = patch('employee_worklog.oauth_session', return_value=self.session)
        patcher.start(); self.addCleanup(patcher.stop)
        self.client.set_cookie('epr_jira_session', 'personal')

    def test_list_only_own_entries_across_pages_on_any_issue(self):
        self.personal_jira(rows=[
            {'id': '1', 'author': {'accountId': 'other'}, 'timeSpentSeconds': 3600},
            {'id': '2', 'author': self.identity, 'timeSpentSeconds': 1800, 'started': self.body['started']},
        ])
        result = self.client.get(PREFIX + '/entries?issue_key=story-1')
        self.assertEqual(result.status_code, 200)
        self.assertEqual([row['worklog_id'] for row in result.json['entries']], ['2'])
        self.assertEqual(self.session.get.call_args.kwargs['params']['startAt'], 1)
        self.session.delete.assert_not_called()

    def test_month_history_discovers_tasks_and_filters_exact_author_and_month(self):
        self.personal_jira(rows=[
            {'id':'1','author':self.identity,'started':'2026-10-05T09:00:00+05:00','timeSpentSeconds':1800},
            {'id':'2','author':{'accountId':'other','displayName':'Hassan Malik'},'started':'2026-10-05','timeSpentSeconds':3600},
            {'id':'3','author':self.identity,'started':'2026-09-30','timeSpentSeconds':3600},
        ])
        get = self.session.get.side_effect
        def search(url, **kwargs):
            if url.endswith('/search/jql'):
                self.assertIn('currentUser()',kwargs['params']['jql'])
                self.assertIn('2026-10-31',kwargs['params']['jql'])
                response=Mock();response.json.return_value={'issues':[{'key':'APP-1','fields':{'summary':'Build calendar'}}], 'nextPageToken':'next'}
                return response
            return get(url,**kwargs)
        self.session.get.side_effect=search
        response=self.client.get(PREFIX+'/history?month=2026-10')
        self.assertEqual(response.status_code,200)
        self.assertEqual(response.json['next_cursor'],'next')
        self.assertEqual([row['worklog_id'] for row in response.json['entries']],['1'])
        self.assertEqual(response.json['entries'][0]['summary'],'Build calendar')
        self.session.delete.assert_not_called()

    def test_history_cursor_month_validation_and_personal_session(self):
        self.personal_jira()
        get=self.session.get.side_effect
        def search(url,**kwargs):
            if url.endswith('/search/jql'):
                self.assertEqual(kwargs['params']['nextPageToken'],'next')
                self.assertIn('2024-02-29',kwargs['params']['jql'])
                response=Mock();response.json.return_value={'issues':[], 'isLast':True}
                return response
            return get(url,**kwargs)
        self.session.get.side_effect=search
        self.assertEqual(self.client.get(PREFIX+'/history?month=2024-02&cursor=next').json['entries'],[])
        for value in ['2026-13','2026-00','2026-2','bad','0000-01']:
            self.assertEqual(self.client.get(PREFIX+'/history?month='+value).status_code,400)
        with patch('employee_worklog.oauth_session',return_value=None):
            self.assertEqual(self.client.get(PREFIX+'/history?month=2026-10').status_code,400)

    def test_confirmed_own_delete_uses_personal_session_and_leaves_estimate(self):
        self.personal_jira()
        result = self.client.post(PREFIX+'/delete', json={'issue_key':'STORY-1', 'worklog_id':'123', 'confirmed':True}, headers=HEADERS)
        self.assertEqual(result.status_code, 200)
        self.assertTrue(result.json['deleted'])
        self.session.delete.assert_called_once_with('https://jira.example/rest/api/3/issue/STORY-1/worklog/123', params={'adjustEstimate':'leave'}, timeout=(10,40))

    def test_edit_checks_ownership_conflict_and_excludes_current_entry_from_limit(self):
        self.personal_jira(rows=[{'id':'123','author':self.identity,'started':'2026-10-01T09:00:00+05:00','timeSpentSeconds':86400}])
        original_get = self.session.get.side_effect
        current = {'id':'123','author':self.identity,'updated':'v1','comment':{'type':'doc','content':[]}}
        def get(url, **kwargs):
            if url.endswith('/worklog/123'):
                response=Mock(); response.json.return_value=current.copy(); return response
            return original_get(url, **kwargs)
        self.session.get.side_effect=get
        body={'issue_key':'APP-1','worklog_id':'123','updated':'v1','duration':'24h',
              'started':'2026-10-01T09:00:00+05:00','comment':'Work','comment_changed':False,'confirmed':True}
        current['author']={'accountId':'other'}
        self.assertEqual(self.client.post(PREFIX+'/edit',json=body,headers=HEADERS).status_code,403)
        current['author']=self.identity;current['updated']='v2'
        self.assertEqual(self.client.post(PREFIX+'/edit',json=body,headers=HEADERS).status_code,409)
        self.session.put.assert_not_called()
        current['updated']='v1'
        response=self.client.post(PREFIX+'/edit',json=body,headers=HEADERS)
        self.assertEqual(response.status_code,200)
        self.assertEqual(self.session.put.call_args.kwargs['json']['comment'], current['comment'])
        self.assertEqual(self.session.put.call_args.kwargs['params']['adjustEstimate'],'leave')

    def test_delete_blocks_other_author_even_with_same_display_name(self):
        self.personal_jira(owner={'accountId':'other', 'displayName':self.identity['displayName']})
        result = self.client.post(PREFIX+'/delete', json={'issue_key':'APP-1','worklog_id':'123','confirmed':True}, headers=HEADERS)
        self.assertEqual(result.status_code, 403)
        self.session.delete.assert_not_called()

    def test_delete_requires_confirmation_valid_ids_and_same_origin(self):
        self.personal_jira()
        body = {'issue_key':'APP-1','worklog_id':'123','confirmed':True}
        for patch_body in [{'confirmed':False}, {'issue_key':'../APP-1'}, {'worklog_id':'../123'}]:
            self.assertEqual(self.client.post(PREFIX+'/delete',json={**body,**patch_body},headers=HEADERS).status_code,400)
        for headers in [{}, {**HEADERS,'Origin':'https://evil.test'}]:
            self.assertEqual(self.client.post(PREFIX+'/delete',json=body,headers=headers).status_code,400)
        self.session.delete.assert_not_called()

    def test_delete_and_list_never_fall_back_to_shared_credentials(self):
        self.client.set_cookie('epr_jira_session', 'expired')
        with patch('employee_worklog.oauth_session', return_value=None):
            self.assertEqual(self.client.get(PREFIX+'/entries?issue_key=APP-1').status_code,400)
            self.assertEqual(self.client.post(PREFIX+'/delete',json={'issue_key':'APP-1','worklog_id':'123','confirmed':True},headers=HEADERS).status_code,400)
        self.session.delete.assert_not_called()

    def test_jira_permission_failure_and_timeout_do_not_report_deleted(self):
        self.personal_jira()
        denied = Mock(status_code=403)
        for error, status in [(requests.HTTPError(response=denied),403),(requests.Timeout(),502)]:
            self.session.delete.side_effect = error
            result = self.client.post(PREFIX+'/delete',json={'issue_key':'APP-1','worklog_id':'123','confirmed':True},headers=HEADERS)
            self.assertEqual(result.status_code,status)
            self.assertNotIn('deleted',result.json)
        self.assertEqual(self.session.delete.call_count,2)

    def test_review_never_writes_and_checks_author_duration_and_origin(self):
        self.assertEqual(self.review().status_code, 200)
        self.session.post.assert_not_called()
        for field, value in [('employee', 'Bob'), ('minutes', 0), ('minutes', True), ('minutes', 0.5), ('minutes', 1441), ('started', '2026-10-01T09:00:00')]:
            with self.subTest(field=field, value=value):
                body = {**self.body, field: value}
                self.assertEqual(self.client.post(PREFIX + '/review', json=body, headers=HEADERS).status_code, 400)
        self.assertEqual(self.client.post(PREFIX + '/review', json=self.body).status_code, 400)
        self.assertEqual(self.client.post(PREFIX + '/review', json=self.body, headers={**HEADERS, 'Origin': 'https://other.test'}).status_code, 400)
        self.session.post.assert_not_called()

    def test_submit_is_confirmed_verified_and_idempotent(self):
        token = self.review().json['token']
        self.session.post.return_value.json.return_value = {'id': '123', 'author': self.identity}
        self.session.get.return_value.json.return_value = {'timeSpentSeconds': 1800, 'author': self.identity}
        self.assertEqual(self.client.post(PREFIX + '/submit', json={'token': token}, headers=HEADERS).status_code, 400)
        for _ in range(2):
            response = self.client.post(PREFIX + '/submit', json={'token': token, 'confirmed': True}, headers=HEADERS)
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.json['worklog_id'], '123')
            self.assertTrue(response.json['verified'])
        self.session.post.assert_called_once()
        call = self.session.post.call_args.kwargs
        self.assertEqual(call['json']['timeSpentSeconds'], 1800)
        self.assertEqual(call['params']['adjustEstimate'], 'leave')
        self.assertIsInstance(call['json']['properties'][0]['value'], dict)

    def test_timeout_is_never_retried_and_tampered_review_cannot_write(self):
        token = self.review().json['token']
        self.session.post.side_effect = requests.Timeout()
        response = self.client.post(PREFIX + '/submit', json={'token': token, 'confirmed': True}, headers=HEADERS)
        self.assertEqual(response.status_code, 502)
        self.assertEqual(self.client.post(PREFIX + '/submit', json={'token': token, 'confirmed': True}, headers=HEADERS).status_code, 409)
        self.assertEqual(self.client.post(PREFIX + '/submit', json={'token': token+'bad', 'confirmed': True}, headers=HEADERS).status_code, 400)
        self.session.post.assert_called_once()

    def test_public_shared_account_cannot_write_without_signin(self):
        with patch.dict(os.environ, {'EPR_WORKLOG_ALLOW_SHARED_ACCOUNT': ''}):
            self.assertEqual(self.review().status_code, 400)
            self.session.post.assert_not_called()

    def test_flexible_duration_and_review_daily_total(self):
        for value, expected in [('1d', 1440), ('8h', 480), ('8m', 8), ('1h 30m', 90), ('0.5h', 30), ('90', 90), (' 2H 5M ', 125)]:
            with self.subTest(value=value):
                self.assertEqual(parse_duration(value), expected)
        for value in ['2d', '25h', '1d 1m', '0m', '-1h', 'nonsense', '1h 2h', '1 2h', '1.5m', '', True]:
            with self.subTest(value=value), self.assertRaises(ValueError):
                parse_duration(value)
        self.body['duration'] = '8h'
        response = self.review()
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json['review']['minutes'], 480)
        self.assertEqual(response.json['review']['duration_label'], '8h')
        self.assertEqual(response.json['review']['total_minutes'], 480)

    def test_daily_limit_checked_at_review_and_rechecked_at_submit(self):
        self.daily_logs = [{'id': 'old', 'author': self.identity, 'started': self.body['started'], 'timeSpentSeconds': 23*3600}]
        self.body['duration'] = '1h'
        response = self.review()
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json['review']['total_minutes'], 1440)
        token = response.json['token']
        self.body['duration'] = '1h 1m'
        self.assertEqual(self.review().status_code, 400)
        # Another entry has arrived since the review; the older review must not bypass the cap.
        self.daily_logs[0]['timeSpentSeconds'] = 24*3600
        blocked = self.client.post(PREFIX+'/submit', json={'token': token, 'confirmed': True}, headers=HEADERS)
        self.assertEqual(blocked.status_code, 400)
        self.assertIn('exceed 24h', blocked.json['error'])
        self.session.post.assert_not_called()

    def test_daily_reader_pages_authors_and_timezone_boundaries(self):
        started = datetime.fromisoformat('2026-10-01T09:00:00+05:00')
        rows = [
            {'id': 'a', 'author': self.identity, 'started': '2026-09-30T20:00:00.000+0000', 'timeSpentSeconds': 3600},
            {'id': 'b', 'author': {'accountId': 'other'}, 'started': self.body['started'], 'timeSpentSeconds': 86400},
            {'id': 'c', 'author': self.identity, 'started': '2026-10-01T20:00:00.000+0000', 'timeSpentSeconds': 86400},
            {'id': 'd', 'author': self.identity, 'started': '2026-10-01T18:00:00.000+0000', 'timeSpentSeconds': 1800},
        ]
        first, second = Mock(), Mock()
        first.json.return_value = {'worklogs': rows[:2], 'total': 4}
        second.json.return_value = {'worklogs': rows[2:], 'total': 4}
        self.session.get.side_effect = [first, second]
        self.assertEqual(daily_logged_seconds(self.session, 'APP-1', 'hassan', started), 5400)
        self.assertEqual(self.session.get.call_args.kwargs['params']['startAt'], 2)

    def test_two_reviewed_submissions_cannot_race_past_day_limit(self):
        self.body['duration'] = '15h'
        tokens = [self.review().json['token'] for _ in range(2)]
        def post_worklog(url, **kwargs):
            row = {'id': 'created', 'author': self.identity, 'started': kwargs['json']['started'], 'timeSpentSeconds': kwargs['json']['timeSpentSeconds']}
            self.daily_logs.append(row)
            response = Mock()
            response.json.return_value = row
            return response
        self.session.post.side_effect = post_worklog
        def submit(token):
            with self.app.test_client() as client:
                return client.post(PREFIX+'/submit', json={'token': token, 'confirmed': True}, headers=HEADERS).status_code
        with ThreadPoolExecutor(max_workers=2) as pool:
            statuses = list(pool.map(submit, tokens))
        self.assertEqual(sorted(statuses), [200, 400])
        self.session.post.assert_called_once()

    def test_unreadable_jira_history_blocks_writing(self):
        self.session.get.side_effect = requests.Timeout()
        self.assertEqual(self.review().status_code, 502)
        self.session.post.assert_not_called()

    def test_oauth_setup_and_state_checks(self):
        with patch.dict(os.environ, {'JIRA_OAUTH_CLIENT_ID': '', 'JIRA_OAUTH_CLIENT_SECRET': '', 'JIRA_OAUTH_REDIRECT_URI': ''}):
            self.assertFalse(self.client.get('/api/jira-auth/status').json['configured'])
            self.assertEqual(self.client.get('/api/jira-auth/login').status_code, 503)
        self.assertEqual(self.client.get('/api/jira-auth/callback?state=bad&code=bad').status_code, 400)
        with patch.dict(os.environ, {'JIRA_OAUTH_CLIENT_ID': 'client', 'JIRA_OAUTH_CLIENT_SECRET': 'secret', 'JIRA_OAUTH_REDIRECT_URI': 'https://report.test/api/jira-auth/callback'}):
            response = self.client.get('/api/jira-auth/login')
            self.assertEqual(response.status_code, 302)
            self.assertIn('auth.atlassian.com/authorize', response.location)
            self.assertIn('Secure;', response.headers['Set-Cookie'])
            self.assertIn('HttpOnly;', response.headers['Set-Cookie'])

    def test_oauth_callback_stores_tokens_only_on_server(self):
        cfg = {'JIRA_OAUTH_CLIENT_ID': 'client', 'JIRA_OAUTH_CLIENT_SECRET': 'secret', 'JIRA_OAUTH_REDIRECT_URI': 'https://report.test/api/jira-auth/callback'}
        with patch.dict(os.environ, cfg):
            response = self.client.get('/api/jira-auth/login')
            state = parse_qs(urlparse(response.location).query)['state'][0]
            tokens, resources = Mock(), Mock()
            tokens.json.return_value = {'access_token': 'access-secret', 'refresh_token': 'refresh-secret', 'expires_in': 3600}
            resources.json.return_value = [{'url': BASE_URL, 'id': 'cloud123', 'scopes': ['read:jira-user', 'read:jira-work', 'write:jira-work']}]
            identity_response = Mock()
            identity_response.json.return_value = self.identity
            with patch('jira_oauth.requests.post', return_value=tokens), patch('jira_oauth.requests.get', side_effect=[resources, identity_response]):
                response = self.client.get('/api/jira-auth/callback?code=code&state=' + state)
            self.assertEqual(response.status_code, 302)
            self.assertNotIn('access-secret', response.headers.get('Set-Cookie', ''))
            files = list((Path(self.tmp.name) / 'oauth').glob('*.json'))
            self.assertEqual(len(files), 1)
            self.assertIn('refresh-secret', files[0].read_text())
            self.client.post('/api/jira-auth/logout')
            self.assertFalse(files[0].exists())


if __name__ == '__main__':
    unittest.main()
