import json
import gc
import os
from pathlib import Path
import tempfile
import time
import unittest
from unittest.mock import patch
from flask import Flask, jsonify
from epr_auth import register_epr_auth, OWNER_ACCOUNT_ID, safe_next
from jira_oauth import register_jira_oauth


class EprAuthTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.addCleanup(gc.collect)
        self.directory = Path(self.tmp.name)/'oauth'
        self.directory.mkdir()
        self.env = patch.dict(os.environ, {'EPR_AUTH_ENABLED':'1', 'JIRA_OAUTH_CLIENT_ID':'', 'JIRA_OAUTH_CLIENT_SECRET':'', 'JIRA_OAUTH_REDIRECT_URI':''})
        self.env.start(); self.addCleanup(self.env.stop)
        self.app = Flask(__name__)
        self.app.testing=True
        register_jira_oauth(self.app, self.directory)
        register_epr_auth(self.app, self.directory, {'report.html'})
        for route in ['/', '/report.html', '/settings/capacity', '/api/report-data', '/api/admin/sql-console/schema']:
            self.app.add_url_rule(route, route, lambda: jsonify(ok=True))
        for route in ['/api/capacity', '/api/employee-capacity-utilization/export', '/api/employee-capacity-utilization/worklog/review']:
            self.app.add_url_rule(route, route+'post', lambda: jsonify(ok=True), methods=['POST'])
        self.client = self.app.test_client()
        self.headers = {'Origin':'http://localhost', 'Content-Type':'application/json'}

    def login(self, account=OWNER_ACCOUNT_ID, name='Hassan Malik', sid=None, age=0):
        sid = sid or ('a'*64 if account==OWNER_ACCOUNT_ID else 'b'*64)
        (self.directory/(sid+'.json')).write_text(json.dumps({'account_id':account, 'session_expires_at':time.time()+86400,
            'epr_signed_at':time.time()-age, 'epr_csrf':'csrf-value',
            'epr_identity':{'account_id':account,'display_name':name,'avatar_url':''}}),encoding='utf-8')
        self.client.set_cookie('epr_jira_session',sid)
        return sid

    def test_anonymous_api_and_direct_files_require_login(self):
        self.assertEqual(self.client.get('/report.html').status_code,302)
        self.assertIn('/login?next=',self.client.get('/report.html').location)
        self.assertEqual(self.client.get('/api/report-data').status_code,401)
        self.assertEqual(self.client.post('/api/capacity',json={}).status_code,401)
        with self.client.get('/login') as response:
            self.assertEqual(response.status_code,200)
        for path, content in [('/epr-login.css','text/css'),('/epr-login.js','javascript')]:
            with self.client.get(path) as response:
                self.assertEqual(response.status_code,200)
                self.assertIn(content,response.content_type)
        self.assertEqual(self.client.get('/epr-login.secret').status_code,302)
        self.assertEqual(self.client.get('/healthz').json,{'ok':True})
        self.client.set_cookie('epr_jira_session','f'*64)
        self.assertEqual(self.client.get('/api/report-data').status_code,401)

    def test_bootstrap_owner_and_other_first_user_pending(self):
        self.login('new-user','New User')
        self.assertEqual(self.client.get('/report.html').location,'/access-pending')
        user=self.client.get('/api/auth/me').json['user']
        self.assertEqual(user['role'],'member');self.assertFalse(user['approved'])
        self.assertEqual(self.client.get('/api/report-data').status_code,403)
        self.login()
        self.assertEqual(self.client.get('/api/auth/me').json['user']['role'],'admin')
        self.assertEqual(self.client.get('/settings/capacity').status_code,200)

    def test_approval_member_permissions_and_immediate_revocation(self):
        member_sid=self.login('member','Member Name')
        self.client.get('/api/auth/me')
        self.login();self.client.get('/api/auth/me')
        payload={'role':'member','active':True,'approved':True}
        self.assertEqual(self.client.put('/api/auth/users/member',json=payload,headers=self.headers).status_code,200)
        self.client.set_cookie('epr_jira_session',member_sid)
        self.assertEqual(self.client.get('/report.html').status_code,200)
        self.assertEqual(self.client.get('/settings/capacity').status_code,403)
        self.assertEqual(self.client.get('/api/admin/sql-console/schema').status_code,403)
        self.assertEqual(self.client.get('/api/auth/users').status_code,403)
        self.assertEqual(self.client.post('/api/capacity',json={},headers=self.headers).status_code,403)
        self.assertEqual(self.client.post('/api/employee-capacity-utilization/export',json={},headers=self.headers).status_code,200)
        self.assertEqual(self.client.post('/api/employee-capacity-utilization/worklog/review',json={},headers=self.headers).status_code,200)
        self.login();self.client.get('/api/auth/me')
        self.client.put('/api/auth/users/member',json={**payload,'active':False},headers=self.headers)
        self.client.set_cookie('epr_jira_session',member_sid)
        self.assertEqual(self.client.get('/api/report-data').status_code,403)

    def test_admin_csrf_last_admin_and_audit_redaction(self):
        self.login();self.client.get('/api/auth/me')
        self.assertEqual(self.client.post('/api/capacity',json={'secret':'never-in-audit'},headers={'Origin':'https://evil.test'}).status_code,403)
        self.assertEqual(self.client.post('/api/capacity',json={'secret':'never-in-audit'},headers={'X-EPR-CSRF':'csrf-value'}).status_code,200)
        self.assertEqual(self.client.put('/api/auth/users/'+OWNER_ACCOUNT_ID,json={'role':'member','approved':True,'active':True},headers=self.headers).status_code,400)
        self.assertEqual(self.client.put('/api/auth/users/'+OWNER_ACCOUNT_ID,json={'role':'admin','approved':True,'active':False},headers=self.headers).status_code,400)
        response=self.client.get('/api/auth/audit')
        self.assertEqual(response.status_code,200)
        self.assertNotIn('never-in-audit',response.get_data(as_text=True))
        self.assertTrue(any(x['event']=='bootstrap_admin' for x in response.json['events']))

    def test_session_expiry_logout_and_safe_return(self):
        self.login(age=8*3600+1)
        self.assertEqual(self.client.get('/api/auth/me').status_code,401)
        self.login();self.client.get('/api/auth/me')
        self.assertEqual(self.client.post('/api/auth/logout',json={},headers=self.headers).status_code,200)
        self.assertEqual(self.client.get('/api/auth/me').status_code,401)
        for value in ['https://evil.test','//evil.test','/%2fevil.test','/\\evil.test','/%0d%0a']:
            self.assertEqual(safe_next(value),'/')
        self.assertEqual(safe_next('/report.html?month=2026-10'),'/report.html?month=2026-10')

    def test_real_server_registration_protects_existing_routes(self):
        from report_server import create_report_server_app
        root=Path(self.tmp.name)/'server'
        root.mkdir();(root/'report_html').mkdir()
        app=create_report_server_app(root,'report_html')
        client=app.test_client()
        for path in ['/api/employee-capacity-utilization/data','/api/admin/sql-console/schema','/api/canonical-refresh/current']:
            response=client.get(path)
            self.assertEqual(response.status_code,401,path)
            self.assertNotIn('Access-Control-Allow-Origin',response.headers)
        self.assertEqual(client.get('/employee_capacity_utilization_report.html').status_code,302)
        with client.get('/login') as response:
            self.assertEqual(response.status_code,200)


if __name__=='__main__':unittest.main()
