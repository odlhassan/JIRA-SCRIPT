import json
from datetime import datetime
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

from flask import Flask
import requests

from employee_worklog import register_employee_worklog_routes
from employee_subtask import estimate_text, month_bounds


API = '/api/employee-capacity-utilization/worklog/subtask'
HEADERS = {'X-EPR-Worklog': '1'}


class EmployeeSubtaskTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.app = Flask(__name__)
        register_employee_worklog_routes(self.app, Path(self.tmp.name))
        self.client = self.app.test_client()
        self.client.set_cookie('epr_jira_session', 'personal')
        self.session = Mock()
        self.session.epr_base_url = 'https://jira.example'
        self.identity = dict(accountId='alice', displayName='Alice', timeZone='Asia/Karachi')
        self.epic = dict(summary='Customer experience', issuetype={'name':'Epic','hierarchyLevel':1}, project={'key':'NEW'})
        self.story = dict(summary='Search customers', issuetype={'name':'Story'}, project={'key':'NEW'}, parent={'key':'NEW-1'})
        self.fields = [dict(fieldId=f, name=f) for f in ('summary','project','issuetype','parent','assignee','duedate','timetracking','description')]
        self.fields += [dict(fieldId='customfield_55',name='Start date',schema={'type':'date'})]
        self.body = dict(project_key='NEW',epic_key='NEW-1',story_key='NEW-2',summary='Build customer search',description='Search by name',start_date='2026-11-11',due_date='2026-11-30',estimate='1d')
        def get(url, **kwargs):
            if url.endswith('/myself'):
                payload = self.identity
            elif url.endswith('/field'):
                payload = [dict(id='customfield_88',schema={'custom':'com.pyxis.greenhopper.jira:gh-epic-link'})]
            elif url.endswith('/issue/NEW-1'):
                payload = {'fields':self.epic}
            elif url.endswith('/issue/NEW-2'):
                payload = {'fields':self.story}
            elif url.endswith('/issuetypes'):
                payload = {'issueTypes':[dict(id='7',name='Sub-task',subtask=True)],'total':1}
            elif '/issuetypes/7' in url:
                payload = {'fields':self.fields,'total':len(self.fields)}
            elif url.endswith('/mypermissions'):
                payload = {'permissions':{key:{'havePermission':True} for key in ('BROWSE_PROJECTS','CREATE_ISSUES','ASSIGN_ISSUES','SCHEDULE_ISSUES')}}
            elif url.endswith('/user/assignable/search'):
                payload = [self.identity]
            elif url.endswith('/project/search'):
                payload = {'values':[dict(key='NEW',name='A space absent from the saved snapshot')],'isLast':False}
            elif url.endswith('/search/jql'):
                payload = {'issues':[{'key':'NEW-2','fields':self.story}],'isLast':True}
            elif url.endswith('/issue/NEW-3'):
                payload = {'fields':dict(assignee=self.identity,parent={'key':'NEW-2'},issuetype={'name':'Sub-task'},duedate='2026-11-30',customfield_55='2026-11-11',timetracking={'originalEstimateSeconds':28800})}
            else:
                raise AssertionError(url)
            response = Mock(); response.json.return_value = payload
            return response
        self.session.get.side_effect = get
        self.session.post.return_value.json.return_value = {'key':'NEW-3','id':'123'}
        for patcher in (patch('employee_worklog.oauth_session', return_value=self.session),
                        patch('employee_subtask.month_bounds',return_value=('2026-11-11','2026-11-01','2026-11-30')),
                        patch('employee_subtask.invalidate_issue')):
            patcher.start(); self.addCleanup(patcher.stop)

    def review(self, **changes):
        return self.client.post(API+'/review',json={**self.body,**changes},headers=HEADERS)

    def submit(self, token):
        return self.client.post(API+'/submit',json={'token':token,'confirmed':True},headers=HEADERS)

    def test_create_assigns_personal_account_preserves_native_estimate_and_replay(self):
        result = self.review(account_id='attacker',assignee='attacker')
        self.assertEqual(result.status_code,200,result.json)
        self.session.post.assert_not_called()
        created = self.submit(result.json['token'])
        self.assertEqual(created.status_code,200,created.json)
        self.assertTrue(created.json['verified'])
        self.assertEqual(created.json['original_estimate_hours'],8)
        fields = self.session.post.call_args.kwargs['json']['fields']
        self.assertEqual(fields['assignee'],{'accountId':'alice'})
        self.assertEqual(fields['parent'],{'key':'NEW-2'})
        self.assertEqual(fields['project'],{'key':'NEW'})
        self.assertEqual(fields['customfield_55'],'2026-11-11')
        self.assertEqual(fields['duedate'],'2026-11-30')
        self.assertEqual(fields['timetracking'],{'originalEstimate':'1d'})
        self.assertEqual(self.submit(result.json['token']).json['issue_key'],'NEW-3')
        self.session.post.assert_called_once()

    def test_dates_current_month_boundaries_order_and_invalid_calendar_dates(self):
        for start, due in [('2026-10-01','2026-10-31'),('2026-11-11','2026-12-01'),('2026-11-21','2026-11-20'),('2026-11-11','2026-11-31'),('2026-11-11','2026-1-01')]:
            with self.subTest(start=start,due=due):
                self.assertEqual(self.review(start_date=start,due_date=due).status_code,400)
        # Only the due date has a month restriction; past starts and both month edges are valid.
        for due in ('2026-11-01','2026-11-30'):
            self.assertEqual(self.review(start_date='2026-10-01',due_date=due).status_code,200)
        self.session.post.assert_not_called()

    def test_month_rollover_and_hierarchy_are_rechecked_before_write(self):
        token = self.review().json['token']
        with patch('employee_subtask.month_bounds',return_value=('2026-12-01','2026-12-01','2026-12-31')):
            self.assertEqual(self.submit(token).status_code,400)
        self.story['parent'] = {'key':'NEW-9'}
        self.assertEqual(self.submit(token).status_code,400)
        self.session.post.assert_not_called()

    def test_rejects_non_story_or_orphan_wrong_space_and_non_epic(self):
        for change in ({'issuetype':{'name':'Task'}},{'parent':{}},{'parent':{'key':'NEW-9'}},{'project':{'key':'OTHER'}}):
            original = self.story.copy();self.story.update(change)
            self.assertEqual(self.review().status_code,400)
            self.story = original
        self.epic['issuetype'] = {'name':'Story','hierarchyLevel':0}
        self.assertEqual(self.review().status_code,400)
        self.session.post.assert_not_called()

    def test_legacy_epic_link_is_resolved_from_jira_catalogue(self):
        self.story.pop('parent'); self.story['customfield_88'] = 'NEW-1'
        self.assertEqual(self.review().status_code,200)

    def test_search_pagination_live_spaces_and_story_filter(self):
        result = self.client.get(API+'/options?kind=projects&cursor=50&q=New')
        self.assertEqual(result.status_code,200)
        self.assertEqual(result.json['next_cursor'],'51')
        self.assertEqual(result.json['author'],'Alice')
        self.assertEqual(result.json['last'],'2026-11-30')
        self.assertEqual(self.session.get.call_args.kwargs['params']['startAt'],50)
        result = self.client.get(API+'/options?kind=stories&project_key=NEW&epic_key=NEW-1&q=NEW-2')
        self.assertEqual(result.json['options'][0]['key'],'NEW-2')
        self.assertIn('parent = "NEW-1"',self.session.get.call_args.kwargs['params']['jql'])
        self.story['issuetype'] = {'name':'Task'}
        self.assertEqual(self.client.get(API+'/options?kind=stories&project_key=NEW&epic_key=NEW-1').json['options'],[])

    def test_metadata_fields_across_pages_and_required_field_diagnostics(self):
        get = self.session.get.side_effect
        def paginated(url, **kwargs):
            if '/issuetypes/7' in url:
                offset = kwargs['params']['startAt'];response = Mock()
                response.json.return_value = {'fields':self.fields[offset:offset+3],'total':len(self.fields)}
                return response
            return get(url,**kwargs)
        self.session.get.side_effect = paginated
        self.assertEqual(self.review().status_code,200)
        self.fields.append(dict(fieldId='customfield_99',name='Cost centre',required=True,hasDefaultValue=False))
        self.assertIn('Cost centre',self.review().json['error'])
        self.fields.pop();self.fields = [f for f in self.fields if f['fieldId'] != 'customfield_55']
        self.assertIn('create screen',self.review().json['error'])

    def test_requires_personal_auth_same_origin_and_matching_submit_account(self):
        token = self.review().json['token']
        self.identity['accountId'] = 'other'
        self.assertEqual(self.submit(token).status_code,400)
        for headers in ({},{**HEADERS,'Origin':'https://evil.test'}):
            self.assertEqual(self.client.post(API+'/review',json=self.body,headers=headers).status_code,400)
        with patch('employee_worklog.oauth_session',return_value=None):
            self.assertEqual(self.review().status_code,400)
            self.assertEqual(self.client.get(API+'/options').status_code,400)
        self.session.post.assert_not_called()

    def test_ambiguous_timeout_is_not_resent_and_known_key_survives_read_failure(self):
        token = self.review().json['token']
        self.session.post.side_effect = requests.Timeout()
        self.assertEqual(self.submit(token).status_code,502)
        self.assertEqual(self.submit(token).status_code,409)
        self.session.post.assert_called_once()
        self.session.post.side_effect = None
        token = self.review().json['token']
        get = self.session.get.side_effect
        def failed_verify(url,**kwargs):
            if url.endswith('/issue/NEW-3'):
                raise requests.Timeout()
            return get(url,**kwargs)
        self.session.get.side_effect = failed_verify
        result = self.submit(token)
        self.assertEqual(result.status_code,200)
        self.assertFalse(result.json['verified'])
        self.assertEqual(self.submit(token).json['issue_key'],'NEW-3')
        self.assertEqual(self.session.post.call_count,2)

    def test_permissions_assignment_and_jira_required_field_failure(self):
        get = self.session.get.side_effect
        def denied(url,**kwargs):
            if url.endswith('/user/assignable/search'):
                response=Mock();response.json.return_value=[];return response
            return get(url,**kwargs)
        self.session.get.side_effect=denied
        self.assertIn('assigning',self.review().json['error'])
        self.session.get.side_effect=get
        token=self.review().json['token']
        response=Mock(status_code=400);response.json.return_value={'errors':{'customfield_4':'Cost centre is required'}}
        self.session.post.side_effect=requests.HTTPError(response=response)
        result=self.submit(token)
        self.assertEqual(result.status_code,502)
        self.assertIn('Cost centre is required',result.json['error'])
        self.assertEqual(self.submit(token).status_code,409)

    def test_estimate_shortcuts_and_no_worklog_24h_limit(self):
        for value in ('30m','4h','1d','1w','1w 2d 3h 30m','0.5h'):
            self.assertEqual(estimate_text(value),value)
        for value in ('0m','-1d','1d 2d','10','1x','abc'):
            with self.assertRaises(ValueError):
                estimate_text(value)

    def test_month_bounds_uses_jira_timezone_and_leap_month(self):
        # This calls the real helper (the route helper is otherwise frozen by setUp).
        with patch('employee_subtask.datetime') as clock:
            clock.now.return_value = datetime(2024,2,29,23,59)
            self.assertEqual(month_bounds({'timeZone':'America/New_York'}),('2024-02-29','2024-02-01','2024-02-29'))
            self.assertEqual(str(clock.now.call_args.args[0]),'America/New_York')
            clock.now.return_value = datetime(2026,12,31,23,59)
            self.assertEqual(month_bounds({'timeZone':'Invalid/Zone'}),('2026-12-31','2026-12-01','2026-12-31'))

    def test_live_subtask_search_is_personal_paginated_and_accepts_any_space(self):
        self.story['issuetype'] = {'name':'Sub-task','subtask':True}
        result = self.client.get(API+'/search?q=NEW-2&cursor=next')
        self.assertEqual(result.status_code,200)
        self.assertEqual(result.json['issues'][0]['issue_key'],'NEW-2')
        params = self.session.get.call_args.kwargs['params']
        self.assertEqual(params['nextPageToken'],'next')
        self.assertIn('key = "NEW-2"',params['jql'])
        self.client.get(API+'/search')
        self.assertIn('assignee = currentUser()',self.session.get.call_args.kwargs['params']['jql'])
        with patch('employee_worklog.oauth_session',return_value=None):
            self.assertEqual(self.client.get(API+'/search').status_code,400)

    def test_missing_permissions_blocks_creation_and_tokens_require_confirmation(self):
        get=self.session.get.side_effect
        def denied(url,**kwargs):
            if url.endswith('/mypermissions'):
                response=Mock();response.json.return_value={'permissions':{}};return response
            return get(url,**kwargs)
        self.session.get.side_effect=denied
        self.assertIn('permissions',self.review().json['error'])
        self.session.get.side_effect=get
        token=self.review().json['token']
        self.assertEqual(self.client.post(API+'/submit',json={'token':token,'confirmed':False},headers=HEADERS).status_code,400)
        self.assertEqual(self.submit(token+'tampered').status_code,400)
        self.session.post.assert_not_called()

    def test_custom_due_date_field_and_missing_time_tracking(self):
        self.fields = [f for f in self.fields if f['fieldId'] != 'duedate']
        self.fields.append(dict(fieldId='customfield_10163',name='Due Date',schema={'type':'date'}))
        token=self.review().json['token']
        result=self.submit(token)
        self.assertEqual(result.status_code,200)
        fields=self.session.post.call_args.kwargs['json']['fields']
        self.assertEqual(fields['customfield_10163'],'2026-11-30')
        self.assertNotIn('duedate',fields)
        self.fields = [f for f in self.fields if f['fieldId'] != 'timetracking']
        self.assertIn('needs Time tracking',self.review().json['error'])

    def test_subtask_receipt_reuses_privacy_erasure_and_keeps_duplicate_guard(self):
        from jira_privacy import erase_oauth_account
        token=self.review().json['token']
        self.assertEqual(self.submit(token).status_code,200)
        erase_oauth_account(Path(self.tmp.name)/'oauth','alice')
        replay=self.submit(token)
        self.assertEqual(replay.status_code,200)
        self.assertEqual(replay.json,{'issue_key':'NEW-3','verified':True})
        self.session.post.assert_called_once()

    def rlt_jira(self):
        from generate_rlt_leave_report import LEAVE_TYPE_FIELD
        self.fields.append(dict(fieldId=LEAVE_TYPE_FIELD,name='Leave Type',allowedValues=[
            {'id':'planned','value':'Planned Leave'},{'id':'unplanned','value':'Unplanned Leave'}]))
        self.session.post.return_value.json.return_value={'key':'RLT-3'}
        original=self.session.get.side_effect
        def get(url,**kwargs):
            response=Mock()
            if url.endswith('/issue/RLT-1'):
                response.json.return_value={'fields':{'summary':'My leave tracker','project':{'key':'RLT'},'issuetype':{'name':'Task','subtask':False}}}
            elif url.endswith('/issue/RLT-3'):
                response.json.return_value={'fields':dict(assignee=self.identity,parent={'key':'RLT-1'},issuetype={'name':'Sub-task'},duedate='2026-11-30',customfield_55='2026-11-11',timetracking={'originalEstimateSeconds':28800},**{LEAVE_TYPE_FIELD:{'id':'planned'}})}
            else:
                return original(url,**kwargs)
            return response
        self.session.get.side_effect=get

    def test_rlt_leave_creation_uses_standalone_tracker_and_jira_leave_type(self):
        from generate_rlt_leave_report import LEAVE_TYPE_FIELD
        self.rlt_jira()
        result=self.review(project_key='RLT',epic_key='',story_key='RLT-1',leave_type='planned')
        self.assertEqual(result.status_code,200,result.json)
        self.assertEqual(result.json['review']['flow'],'leave')
        self.assertEqual(result.json['review']['leave_label'],'Planned Leave')
        saved=self.submit(result.json['token'])
        self.assertEqual(saved.status_code,200,saved.json)
        self.assertTrue(saved.json['verified'])
        fields=self.session.post.call_args.kwargs['json']['fields']
        self.assertEqual(fields['parent'],{'key':'RLT-1'})
        self.assertEqual(fields[LEAVE_TYPE_FIELD],{'id':'planned'})
        self.assertEqual(fields['assignee'],{'accountId':'alice'})
        self.assertEqual(saved.json['flow'],'leave')

    def test_rlt_does_not_relax_work_hierarchy_or_accept_forged_leave_types(self):
        self.rlt_jira()
        self.assertEqual(self.review(project_key='RLT',story_key='NEW-2',leave_type='planned').status_code,400)
        self.assertEqual(self.review(project_key='RLT',story_key='RLT-1',leave_type='forged').status_code,400)
        self.assertEqual(self.review(project_key='NEW',epic_key='',flow='leave',leave_type='planned').status_code,400)
        self.session.post.assert_not_called()

    def test_rlt_trackers_do_not_request_epics_and_return_available_leave_types(self):
        self.rlt_jira()
        self.story['issuetype']={'name':'Task'}
        result=self.client.get(API+'/options?kind=trackers&project_key=RLT&q=Leave')
        self.assertEqual(result.status_code,200,result.json)
        self.assertEqual(result.json['leave_types'],[{'id':'planned','label':'Planned Leave'},{'id':'unplanned','label':'Unplanned Leave'}])
        jql=self.session.get.call_args.kwargs['params']['jql']
        self.assertIn('project = RLT',jql)
        self.assertNotIn('parent',jql)
        self.assertNotIn('Epic',jql)
        self.assertEqual(self.client.get(API+'/options?kind=trackers&project_key=NEW').status_code,400)

    def test_leave_search_stays_in_personal_rlt_scope_even_with_query(self):
        self.client.get(API+'/search?flow=leave&q=Leave')
        jql=self.session.get.call_args.kwargs['params']['jql']
        self.assertIn('project = RLT AND assignee = currentUser()',jql)
        self.assertIn('summary ~',jql)

    def test_empty_epics_has_specific_reason_and_never_loads_stories_implicitly(self):
        original=self.session.get.side_effect
        def get(url,**kwargs):
            if url.endswith('/search/jql'):
                response=Mock();response.json.return_value={'issues':[],'isLast':True};return response
            return original(url,**kwargs)
        self.session.get.side_effect=get
        result=self.client.get(API+'/options?kind=epics&project_key=NEW')
        self.assertIn('NEW has no visible epics',result.json['empty_reason'])
        result=self.client.get(API+'/options?kind=epics&project_key=NEW&q=other')
        self.assertIn('Clear its search',result.json['empty_reason'])


if __name__ == '__main__':
    unittest.main()
