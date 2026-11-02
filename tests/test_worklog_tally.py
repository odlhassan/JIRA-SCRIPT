from datetime import date
from contextlib import closing
from io import BytesIO
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import Mock, patch

from pypdf import PdfReader
import requests

from worklog_tally import _live_entries, _snapshot, build_worklog_tally, month_bounds, reconcile, tally_pdf


def entry(identifier, day='2026-10-26', seconds=3600, task='APP-1', kind='Sub-task'):
    return dict(worklog_id=identifier, issue_key=task, date=day, seconds=seconds,
                issue_type=kind, summary='Work <script> & title', planned_start='2026-10-19', planned_due='2026-10-19')


def snapshot(entries):
    return dict(issue_run='issues', worklog_run='logs', fetched_at='2026-11-02T07:24:00Z', entries=entries)


class WorklogTallyTests(unittest.TestCase):
    def test_reconcile_ids_dates_durations_and_every_day_not_just_equal_totals(self):
        old = [entry('match'), entry('moved'), entry('stale'), entry('changed', seconds=1800)]
        live = [entry('match'), entry('moved', day='2026-11-02'), entry('new'), entry('changed', seconds=1801)]
        result = reconcile('Alice', '2026-10', snapshot(old), live)
        by_id = {r['worklog_id']: r for r in result['entries']}
        self.assertEqual(by_id['match']['status'], 'Match')
        self.assertEqual(by_id['moved']['status'], 'Different')
        self.assertEqual(by_id['changed']['status'], 'Different')
        self.assertEqual(by_id['new']['status'], 'Only in Jira')
        self.assertEqual(by_id['stale']['status'], 'Not returned by Jira')
        self.assertEqual(len(result['daily']), 31)
        self.assertEqual(result['daily'][27]['status'], 'No entries in either source')
        self.assertEqual(result['jira_seconds'], 9001)  # The moved November row is listed but excluded.
        equal = reconcile('Alice','2026-10',snapshot([entry('a')]),[entry('b')])
        self.assertEqual(equal['jira_seconds'],equal['epr_seconds'])
        self.assertEqual(equal['counts'],{'Only in Jira':1,'Not returned by Jira':1})

    def test_calendar_scope_keeps_non_subtasks_and_leaves_visible_as_exclusions(self):
        rows = [entry('sub'), entry('bug', kind='Bug Subtask'), entry('story', kind='Story'),
                entry('leave', task='RLT-1'), entry('missing-meta', kind='')]
        result = reconcile('Alice', '2026-10', snapshot(rows), rows)
        self.assertEqual(result['jira_seconds'], 5*3600)
        self.assertEqual(result['calendar_seconds'], 2*3600)
        self.assertEqual(result['counts'], {'Match': 5})
        reasons={r['worklog_id']:r['calendar'] for r in result['entries']}
        self.assertIn('Leave',reasons['leave'])
        self.assertEqual(reasons['story'],'Not a subtask')
        self.assertEqual(reasons['missing-meta'],'Issue metadata unavailable')

    def test_snapshot_uses_same_two_runs_as_report_and_preserves_ids(self):
        with tempfile.TemporaryDirectory() as td:
            db=Path(td)/'data.db'
            with closing(sqlite3.connect(db)) as c:
                c.executescript('''
                    CREATE TABLE canonical_refresh_state(id INTEGER, last_success_run_id TEXT);
                    INSERT INTO canonical_refresh_state VALUES(1,'issue-only');
                    CREATE TABLE canonical_refresh_runs(run_id TEXT, ended_at_utc TEXT);
                    INSERT INTO canonical_refresh_runs VALUES('latest-logs','2026-11-02T07:24:00Z');
                    CREATE TABLE canonical_issues(run_id TEXT,issue_key TEXT,summary TEXT,issue_type TEXT,start_date TEXT,due_date TEXT);
                    INSERT INTO canonical_issues VALUES('issue-only','APP-1','Build','Sub-task','2026-10-19','2026-10-19');
                    CREATE TABLE canonical_worklogs(run_id TEXT,worklog_id TEXT,issue_key TEXT,worklog_author TEXT,started_date TEXT,hours_logged REAL);
                    INSERT INTO canonical_worklogs VALUES('old','old','APP-1','Alice','2026-10-26',2);
                    INSERT INTO canonical_worklogs VALUES('latest-logs','new','APP-1',' Alice ','2026-10-26',3);
                    INSERT INTO canonical_worklogs VALUES('latest-logs','other','APP-1','Bob','2026-10-26',8);
                ''')
            result=_snapshot(db,'alice','2026-10')
            self.assertEqual(result['issue_run'],'issue-only')
            self.assertEqual(result['worklog_run'],'latest-logs')
            self.assertEqual([r['worklog_id'] for r in result['entries']],['new'])
            self.assertEqual(result['entries'][0]['seconds'],10800)
            self.assertEqual(result['entries'][0]['planned_start'],'2026-10-19')

    def test_live_reads_issue_and_worklog_pages_checks_account_and_moved_logs(self):
        s=Mock(); search_count=0
        def get(url,**kwargs):
            nonlocal search_count
            response=Mock();params=kwargs['params']
            if url.endswith('search/jql'):
                search_count+=1
                response.json.return_value={'issues':[{'key':'APP-1','fields':{'summary':'Build','issuetype':{'name':'Sub-task'}}}],
                    'nextPageToken':'page2','isLast':False} if search_count==1 else {'issues':[],'isLast':True}
            elif url.endswith('issue/APP-2'):
                response.json.return_value={'fields':{'summary':'Moved','issuetype':{'name':'Sub-task'}}}
            else:
                task='APP-2' if '/APP-2/' in url else 'APP-1'
                offset=params['startAt']
                rows=[{'id':'moved' if task=='APP-2' else 'own','author':{'accountId':'alice'},
                       'started':'2026-11-02T09:00:00+05:00' if task=='APP-2' else '2026-10-26T09:00:00+05:00','timeSpentSeconds':180},
                      {'id':'other','author':{'accountId':'bob'},'started':'2026-10-26T09:00:00+05:00','timeSpentSeconds':86400}]
                response.json.return_value={'worklogs':rows[offset:offset+1],'total':2}
            return response
        s.get.side_effect=get
        rows=_live_entries(s,'alice',date(2026,10,1),date(2026,10,31),[entry('moved',task='APP-2')])
        self.assertEqual({r['worklog_id'] for r in rows},{'own','moved'})
        self.assertEqual(search_count,2)
        s.post.assert_not_called();s.delete.assert_not_called()

    def test_incomplete_or_repeated_pages_fail_instead_of_returning_zero(self):
        for payload in [{'issues':[],'isLast':False}, {'issues':[],'nextPageToken':'again','isLast':False}]:
            s=Mock();s.get.return_value.json.return_value=payload
            with self.subTest(payload=payload),self.assertRaises(RuntimeError):
                _live_entries(s,'alice',date(2026,10,1),date(2026,10,31),[])
        s=Mock();s.get.return_value.json.side_effect=[{'issues':[{'key':'APP-1','fields':{}}],'isLast':True},{'worklogs':[],'total':1}]
        with self.assertRaises(RuntimeError):
            _live_entries(s,'alice',date(2026,10,1),date(2026,10,31),[])

    def test_validation_and_ambiguous_employee_never_guess(self):
        for month in ['2026-13','2026-1',None,'0001-01','9999-12']:
            with self.subTest(month=month),self.assertRaises(ValueError):month_bounds(month)
        self.assertEqual(month_bounds('2024-02')[1].day,29)
        s=Mock();s.get.return_value.json.side_effect=[{'displayName':'Reader'},[
            {'accountId':'a','displayName':'Alice'},{'accountId':'b','displayName':'Alice'}]]
        with patch('worklog_tally._snapshot',return_value=snapshot([])),self.assertRaises(ValueError):
            build_worklog_tally(Path('unused'),'Alice','2026-10',s)
        for name in ['', 'Grand Total', 'Unassigned']:
            with self.assertRaises(ValueError):build_worklog_tally(Path('unused'),name,'2026-10',s)

    def test_pdf_is_readable_includes_zero_days_ids_and_limitations(self):
        result=reconcile('Faiza Nasir','2026-10',snapshot([entry('123')]),[entry('123')])
        result.update(jira_reader='Reporting account',account_id='faiza')
        output=tally_pdf(result)
        self.assertTrue(output.getvalue().startswith(b'%PDF'))
        pages=PdfReader(output).pages
        text='\n'.join(p.extract_text() for p in pages)
        for expected in ['2026-10-26','2026-10-28','2026-10-29','2026-10-30','123','No entries in either source',
                         'does not prove','Work <script> & title','2026-10-19']:
            self.assertIn(expected,text)
        self.assertGreaterEqual(len(pages),3)

    def test_routes_stream_pdf_and_never_disguise_jira_failure(self):
        from report_server import create_report_server_app
        result=reconcile('Alice','2026-10',snapshot([]),[])
        result.update(jira_reader='Reader',account_id='alice')
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            root=Path(td);(root/'report_html').mkdir()
            client=create_report_server_app(root,'report_html').test_client()
            with patch('report_server.build_worklog_tally',return_value=result) as build:
                response=client.get('/api/employee-capacity-utilization/tally.pdf?employee=Alice&month=2026-10')
                self.assertEqual(response.status_code,200)
                self.assertEqual(response.mimetype,'application/pdf')
                self.assertIn('no-store',response.headers['Cache-Control'])
                self.assertIn('worklog-tally-Alice-2026-10.pdf',response.headers['Content-Disposition'])
                build.assert_called_once_with(root/'assignee_hours_capacity.db','Alice','2026-10')
            with patch('report_server.build_worklog_tally',side_effect=requests.Timeout()):
                response=client.get('/api/employee-capacity-utilization/tally.pdf?employee=Alice&month=2026-10')
                self.assertEqual(response.status_code,502)
                self.assertIn('No tally was generated',response.json['error'])


if __name__=='__main__':
    unittest.main()
