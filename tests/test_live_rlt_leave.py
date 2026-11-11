import unittest
from unittest.mock import Mock
from live_rlt_leave import build_live_rlt_leaves


class LiveRltLeaveTests(unittest.TestCase):
    def test_leave_uses_assignee_not_worklog_author_and_actual_dates(self):
        session=Mock()
        def get(url,**kwargs):
            response=Mock(status_code=200)
            if url.endswith('search/jql'):
                self.assertIn('assignee = "namra"',kwargs['params']['jql'])
                response.json.return_value={'isLast':True,'issues':[
                    {'key':f'RLT-{day}','id':str(day),'fields':{'summary':f'{day}-October-26','issuetype':{'subtask':True},
                     'assignee':{'accountId':'namra','displayName':'Namra Zahid'},'customfield_10584':{'value':'Planned Leave'},'timeoriginalestimate':28800}}
                    for day in range(19,23)]}
            else:
                day=url.split('RLT-')[1].split('/')[0]
                response.json.return_value={'total':1,'worklogs':[{'id':day,'started':f'2026-10-{day}T09:00:00+05:00','timeSpentSeconds':28800,'author':{'accountId':'hr','displayName':'HR'}}]}
            return response
        session.get.side_effect=get
        rows=build_live_rlt_leaves('Namra Zahid','2026-10','namra',session)
        self.assertEqual([r['period_day'] for r in rows],['2026-10-19','2026-10-20','2026-10-21','2026-10-22'])
        self.assertEqual(sum(r['planned_taken_hours'] for r in rows),32)
        self.assertTrue(all(r['assignee']=='Namra Zahid' for r in rows))
        session.post.assert_not_called();session.delete.assert_not_called()

    def test_incomplete_leave_read_fails_instead_of_returning_zero(self):
        session=Mock();session.get.return_value.status_code=200
        session.get.return_value.json.return_value={'issues':[],'isLast':False}
        with self.assertRaises(RuntimeError):build_live_rlt_leaves('Namra Zahid','2026-10','namra',session)

