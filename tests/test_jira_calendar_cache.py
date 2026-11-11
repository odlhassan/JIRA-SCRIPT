from concurrent.futures import ThreadPoolExecutor
from copy import deepcopy
import unittest
from unittest.mock import Mock
import requests

from jira_calendar_cache import calendar_read, STORES, invalidate_issue
from worklog_tally import _get


class CalendarCacheTests(unittest.TestCase):
    def setUp(self):
        STORES.clear()
        self.session = Mock(headers={'Authorization':'reporter'})
        self.logs = [{'id':'1','timeSpentSeconds':3600}]
        self.events = {'updated':[], 'deleted':[]}
        self.fields = {'summary':'Task','issuetype':{'name':'Sub-task'}}
        self.fail = False
        self.calls = []
        def get(url, **kwargs):
            path=url.split('/rest/api/3/')[1];self.calls.append(path)
            response=Mock(status_code=200)
            if path.startswith('worklog/'):
                if self.fail and path=='worklog/deleted': raise requests.Timeout()
                response.json.return_value={'values':deepcopy(self.events[path.split('/')[1]]),'until':9999999,'lastPage':True}
            elif path=='search/jql':
                response.json.return_value={'issues':[{'key':'APP-1','id':'100','fields':deepcopy(self.fields)}],'isLast':True}
            else:
                response.json.return_value={'worklogs':deepcopy(self.logs),'total':len(self.logs)}
            return response
        self.session.get.side_effect=get
        self.session.post.return_value.status_code=200
        self.session.post.return_value.json.return_value=[{'id':'2','issueId':'100'}]

    def builder(self, reader):
        _get(reader,'search/jql',fields='summary,issuetype')
        rows=_get(reader,'issue/APP-1/worklog',startAt=0,maxResults=100)['worklogs']
        return {'employee':'Alice','month':'2026-10','entries':rows,'leaves':[{'field':deepcopy(self.fields)}]}

    def read(self, now, builder=None):
        return calendar_read('db','Alice','2026-10',builder or self.builder,self.session,now)

    def test_no_change_reuses_pages_and_shared_result(self):
        self.read(1000);self.calls.clear()
        cached=self.read(1010)
        self.assertEqual(cached['sync']['network_reads'],0);self.assertEqual(self.calls,[])
        unchanged=self.read(1031)
        self.assertEqual(unchanged['sync']['mode'],'unchanged')
        self.assertNotIn('issue/APP-1/worklog',self.calls)
        self.assertIn('worklog/deleted',self.calls)

    def test_edit_delete_and_new_entry_invalidate_only_changed_issue(self):
        self.read(1000)
        self.logs[0]['timeSpentSeconds']=7200
        self.events['updated']=[{'worklogId':1,'updatedTime':1001000}]
        self.assertEqual(self.read(1031)['entries'][0]['timeSpentSeconds'],7200)
        self.logs=[];self.events['deleted']=[{'worklogId':1,'updatedTime':1035000}]
        self.assertEqual(self.read(1062)['entries'],[])
        self.logs=[{'id':'2','timeSpentSeconds':1800}]
        self.events['updated']=[{'worklogId':2,'updatedTime':1065000}]
        self.assertEqual(self.read(1093)['entries'][0]['id'],'2')
        self.session.post.assert_called_once()

    def test_rlt_metadata_change_and_write_invalidation(self):
        self.read(1000)
        self.fields['leave_type']='Unplanned'
        self.assertEqual(self.read(1031)['sync']['mode'],'updated')
        self.logs=[];invalidate_issue('APP-1')
        self.assertEqual(self.read(1032)['entries'],[])

    def test_failure_retains_cursor_and_result(self):
        self.read(1000);before=deepcopy(STORES)
        self.events['updated']=[{'worklogId':1,'updatedTime':1001000}]
        self.fail=True
        with self.assertRaises(requests.Timeout):self.read(1031)
        self.assertEqual(STORES,before)

    def test_concurrent_viewers_coalesce_and_months_share_worklog_pages(self):
        with ThreadPoolExecutor(max_workers=4) as pool:
            results=list(pool.map(lambda _:self.read(1000),range(4)))
        self.assertEqual(self.calls.count('issue/APP-1/worklog'),1)
        self.assertEqual(sum(r['sync']['network_reads']==0 for r in results),3)
        calendar_read('db','Alice','2026-09',self.builder,self.session,1010)
        self.assertEqual(self.calls.count('issue/APP-1/worklog'),1)

    def test_periodic_reconciliation_and_incomplete_feed(self):
        self.read(1000);self.calls.clear()
        self.read(1901)
        self.assertEqual(self.calls.count('issue/APP-1/worklog'),1)

    def test_overlapping_feed_event_is_not_reapplied(self):
        self.read(1000)
        self.events['updated']=[{'worklogId':1,'updatedTime':1001000}]
        self.read(1031);self.calls.clear()
        # Use a realistic until so the event remains in the overlap ledger.
        state=next(iter(STORES.values()));state['seen'][('updated','1')]=1001000
        self.read(1062)
        self.assertNotIn('issue/APP-1/worklog',self.calls)
