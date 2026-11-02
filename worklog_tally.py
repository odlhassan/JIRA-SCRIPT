"""Read-only, per-worklog reconciliation of live Jira and the calendar snapshot."""
from collections import Counter
from contextlib import closing
from datetime import date, datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
import calendar
import math
import re
import sqlite3
from xml.sax.saxutils import escape

from jira_client import BASE_URL, get_session


def month_bounds(month):
    if not re.fullmatch(r'\d{4}-\d{2}', str(month)):
        raise ValueError('Choose a valid month.')
    try:
        first = date.fromisoformat(month + '-01')
        last = first.replace(day=calendar.monthrange(first.year, first.month)[1])
        # A day either side also covers Jira search/profile timezone differences.
        first - timedelta(days=1)
        last + timedelta(days=1)
    except (ValueError, OverflowError):
        raise ValueError('Choose a month between January 0002 and December 9998.') from None
    return first, last


def _snapshot(db_path, employee, month):
    """Use exactly the issue/worklog runs selected by the report's existing readers."""
    if not Path(db_path).exists():
        raise RuntimeError('The EPR reporting snapshot is unavailable. Run Colossal Refresh first.')
    with closing(sqlite3.connect(Path(db_path).resolve().as_uri() + '?mode=ro', uri=True)) as conn:
        conn.row_factory = sqlite3.Row
        conn.execute('BEGIN')
        state = conn.execute('SELECT last_success_run_id FROM canonical_refresh_state WHERE id=1').fetchone()
        issue_run = state[0] if state else ''
        if not issue_run:
            raise RuntimeError('No successful EPR reporting snapshot is available.')
        latest = conn.execute('SELECT run_id FROM canonical_worklogs ORDER BY rowid DESC LIMIT 1').fetchone()
        worklog_run = latest[0] if latest else issue_run
        fetched = conn.execute('SELECT ended_at_utc FROM canonical_refresh_runs WHERE run_id=?', (worklog_run,)).fetchone()
        issues = {r['issue_key']: dict(r) for r in conn.execute(
            'SELECT issue_key,summary,issue_type,start_date,due_date FROM canonical_issues WHERE run_id=?', (issue_run,))}
        rows = [dict(r) for r in conn.execute(
            'SELECT worklog_id,issue_key,worklog_author,started_date,hours_logged '
            'FROM canonical_worklogs WHERE run_id=? AND started_date>=? AND started_date<=?',
            (worklog_run, month + '-01', month + '-31'))
            if str(r['worklog_author']).strip().casefold() == employee.casefold()]
    for row in rows:
        meta = issues.get(row['issue_key'], {})
        hours = float(row.pop('hours_logged'))
        if not math.isfinite(hours) or hours < 0:
            raise RuntimeError('EPR contains an invalid worklog duration; tally could not be completed.')
        row.update(date=row.pop('started_date'), seconds=round(hours * 3600),
                   summary=meta.get('summary', ''), issue_type=meta.get('issue_type', ''),
                   planned_start=meta.get('start_date', ''), planned_due=meta.get('due_date', ''))
    return dict(issue_run=issue_run, worklog_run=worklog_run,
                fetched_at=fetched[0] if fetched else '', entries=rows)


def _get(session, path, **params):
    response = session.get(BASE_URL + '/rest/api/3/' + path, params=params, timeout=(10, 30))
    response.raise_for_status()
    return response.json()


def _account(session, employee):
    matches = {}
    for page in range(20):
        users = _get(session, 'user/search', query=employee, startAt=page*100, maxResults=100)
        if not isinstance(users, list):
            raise RuntimeError('Jira did not return a valid employee search.')
        for user in users:
            if str(user.get('displayName', '')).strip().casefold() == employee.casefold() and user.get('accountId'):
                matches[user['accountId']] = user
        if len(users) < 100:
            break
    else:
        raise RuntimeError('Employee lookup was incomplete; no tally was generated.')
    if len(matches) != 1:
        raise ValueError('Jira must identify exactly one account with this employee name. Ask an administrator to check the name.')
    return next(iter(matches.values()))


def _live_entries(session, account_id, first, last, snapshot):
    jql = (f'worklogAuthor = "{account_id}" AND worklogDate >= "{first-timedelta(days=1)}" '
           f'AND worklogDate <= "{last+timedelta(days=1)}" ORDER BY key ASC')
    # Read snapshot tasks too: a log may have moved to another date/month in Jira.
    tasks = {r['issue_key']: None for r in snapshot}
    token, seen_tokens = None, set()
    for _ in range(200):
        params = dict(jql=jql, fields='summary,issuetype', maxResults=100)
        if token:
            params['nextPageToken'] = token
        page = _get(session, 'search/jql', **params)
        if not isinstance(page, dict) or not isinstance(page.get('issues'), list):
            raise RuntimeError('Jira issue discovery was incomplete; no tally was generated.')
        for task in page['issues']:
            task_key = task.get('key', '')
            if not re.fullmatch(r'[A-Z][A-Z0-9_]*-\d+', task_key):
                raise RuntimeError('Jira returned an invalid issue key.')
            tasks[task_key] = task.get('fields') or {}
        token = page.get('nextPageToken')
        if page.get('isLast') is True:
            break
        if not isinstance(token, str) or not token or token in seen_tokens:
            raise RuntimeError('Jira issue pagination was incomplete; no tally was generated.')
        seen_tokens.add(token)
    else:
        raise RuntimeError('Jira returned too many issue pages; no partial tally was generated.')
    rows, baseline_ids = {}, {r['worklog_id'] for r in snapshot}
    for task_key, fields in sorted(tasks.items()):
        if fields is None:
            fields = _get(session, 'issue/' + task_key, fields='summary,issuetype').get('fields') or {}
        offset, seen_pages = 0, set()
        for _ in range(1000):
            page = _get(session, 'issue/' + task_key + '/worklog', startAt=offset, maxResults=100)
            logs, total = page.get('worklogs'), page.get('total')
            if not isinstance(logs, list) or type(total) is not int or total < 0:
                raise RuntimeError('Jira worklog pagination was incomplete; no tally was generated.')
            ids = tuple(str(row.get('id', '')) for row in logs)
            if logs and ids in seen_pages:
                raise RuntimeError('Jira repeated a worklog page; no partial tally was generated.')
            seen_pages.add(ids)
            for row in logs:
                if (row.get('author') or {}).get('accountId') != account_id:
                    continue
                identifier, started, seconds = str(row.get('id', '')), str(row.get('started', '')), row.get('timeSpentSeconds')
                try:
                    parsed = datetime.fromisoformat(started.replace('Z', '+00:00'))
                except ValueError:
                    raise RuntimeError('Jira returned an invalid worklog date.') from None
                if not identifier or parsed.utcoffset() is None or type(seconds) is not int or seconds < 0:
                    raise RuntimeError('Jira returned an invalid worklog.')
                if first <= parsed.date() <= last or identifier in baseline_ids:
                    rows[identifier] = dict(worklog_id=identifier, issue_key=task_key, date=started[:10],
                        seconds=seconds, summary=fields.get('summary', ''),
                        issue_type=(fields.get('issuetype') or {}).get('name', ''))
            offset += len(logs)
            if offset >= total:
                break
            if not logs:
                raise RuntimeError('Jira ended its worklog pages early; no partial tally was generated.')
        else:
            raise RuntimeError('Jira returned too many worklog pages; no partial tally was generated.')
    return list(rows.values())


def calendar_reason(row):
    if row['issue_key'].upper().startswith('RLT-'):
        return 'Leave (separate from calendar work)'
    if not re.search(r'sub[ -]?task', row.get('issue_type', ''), re.I):
        return 'Not a subtask' if row.get('issue_type') else 'Issue metadata unavailable'
    return 'Included'


def reconcile(employee, month, snapshot, live):
    first, last = month_bounds(month)
    local = {r['worklog_id']: r for r in snapshot['entries']}
    remote = {r['worklog_id']: r for r in live}
    details = []
    for identifier in local.keys() | remote.keys():
        jira, epr = remote.get(identifier), local.get(identifier)
        if not epr:
            status = 'Only in Jira'
        elif not jira:
            status = 'Not returned by Jira'
        else:
            status = 'Match' if all(jira[k] == epr[k] for k in ('issue_key', 'date', 'seconds')) else 'Different'
        row = jira or epr
        details.append(dict(worklog_id=identifier, issue_key=row['issue_key'], summary=row.get('summary', ''),
                            jira=jira, epr=epr, status=status,
                            calendar=calendar_reason(epr) if epr else 'Not in EPR snapshot'))
    details.sort(key=lambda r: ((r['jira'] or r['epr'])['date'], r['issue_key'], r['worklog_id']))
    daily = []
    for day in range(1, last.day+1):
        iso = first.replace(day=day).isoformat()
        jira_seconds = sum(r['seconds'] for r in live if r['date'] == iso)
        epr_seconds = sum(r['seconds'] for r in local.values() if r['date'] == iso)
        shown_seconds = sum(r['seconds'] for r in local.values() if r['date'] == iso and calendar_reason(r) == 'Included')
        different = any(r['status'] != 'Match' and any(x and x['date'] == iso for x in (r['jira'], r['epr'])) for r in details)
        daily.append(dict(date=iso, jira_seconds=jira_seconds, epr_seconds=epr_seconds, calendar_seconds=shown_seconds,
                          status='Check entries' if different else 'No entries in either source' if not jira_seconds and not epr_seconds else 'Match'))
    return dict(employee=employee, month=month, generated_at=datetime.now(timezone.utc).isoformat(),
                issue_run=snapshot['issue_run'], worklog_run=snapshot['worklog_run'], fetched_at=snapshot['fetched_at'],
                counts=dict(Counter(r['status'] for r in details)), entries=details, daily=daily,
                jira_seconds=sum(r['jira_seconds'] for r in daily), epr_seconds=sum(r['epr_seconds'] for r in daily),
                calendar_seconds=sum(r['calendar_seconds'] for r in daily))


def build_worklog_tally(db_path, employee, month, session=None):
    employee = str(employee or '').strip()
    if not employee or len(employee) > 200 or employee.casefold() in ('grand total', 'unassigned'):
        raise ValueError('Open one named employee calendar to generate a tally.')
    first, last = month_bounds(month)
    snapshot = _snapshot(db_path, employee, month)
    session = session or get_session()
    reader = _get(session, 'myself')
    account = _account(session, employee)
    # Account IDs come from Jira, but still validate before embedding into JQL.
    if not re.fullmatch(r'[A-Za-z0-9:_-]+', account['accountId']):
        raise RuntimeError('Jira returned an invalid account identifier.')
    result = reconcile(employee, month, snapshot, _live_entries(session, account['accountId'], first, last, snapshot['entries']))
    result.update(jira_reader=reader.get('displayName', 'Configured reporting account'), account_id=account['accountId'])
    return result


def tally_pdf(result):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), leftMargin=32, rightMargin=32, topMargin=30, bottomMargin=30,
                            title=f"Worklog tally - {result['employee']} - {result['month']}", author='EPR Reporting')
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Cell', fontName='Helvetica', fontSize=8, leading=10, spaceAfter=0))
    styles.add(ParagraphStyle(name='CellHeader', parent=styles['Cell'], textColor=colors.white, fontName='Helvetica-Bold'))
    styles['BodyText'].fontSize, styles['BodyText'].leading = 10, 14
    def p(text, style='BodyText'):
        return Paragraph(escape(str(text)), styles[style])
    def hours(seconds):
        return f'{seconds/3600:.2f}h'
    def table(headers, rows, widths, compact=False):
        data = [[p(x, 'CellHeader') for x in headers]] + [[p(x, 'Cell') for x in row] for row in rows]
        t = Table(data, colWidths=widths, repeatRows=1, hAlign='LEFT')
        t.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#153d6c')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f0f5fb')]),
            ('VALIGN', (0,0), (-1,-1), 'TOP'), ('TOPPADDING', (0,0), (-1,-1), 2 if compact else 4), ('BOTTOMPADDING', (0,0), (-1,-1), 2 if compact else 4),
            ('LINEBELOW', (0,0), (-1,0), 1, colors.HexColor('#153d6c'))]))
        return t
    story = [p('EPR / JIRA WORKLOG TALLY', 'Title'), p(f"{result['employee']} | {result['month']}", 'Heading1'),
        p('A read-only comparison of saved Jira worklogs and the EPR reporting snapshot.'), Spacer(1,12)]
    story.append(table(['Live Jira (all types)', 'EPR snapshot (all types)', 'EPR calendar work', 'Entries requiring review'],
        [[hours(result['jira_seconds']), hours(result['epr_seconds']), hours(result['calendar_seconds']),
          sum(v for k,v in result['counts'].items() if k != 'Match')]], [190,190,190,200]))
    for line in [f"Comparison generated (UTC): {result['generated_at']}",
                 f"Jira read access: {result['jira_reader']} | Employee account: {result['account_id']}",
                 f"EPR worklog snapshot: {result['worklog_run']} | Fetch completed (UTC): {result['fetched_at'] or 'Unknown'}",
                 f"EPR issue snapshot: {result['issue_run']}"]:
        story.extend([Spacer(1,8), p(line)])
    story.extend([Spacer(1,12), p('How to read this tally', 'Heading2')])
    for line in [
        'Dates are the saved worklog started dates, not task planned start/due dates. Every day of the month is listed, including days with no entries.',
        'All work-item types and leave worklogs are compared. Calendar work includes only non-RLT subtasks, including work assigned to other people. Main-table scope/team filters do not narrow this employee tally.',
        'Match compares worklog ID, task, date and duration. Equal totals alone do not establish a match. Different includes entries moved outside the selected month; those rows are listed but do not add to the live month total.',
        'Jira results cover the reporting account\'s visible records at generation time. Restricted or not-yet-indexed worklogs may not be returned. Not returned by Jira does not prove deletion, failure to save, or user error. A failed or incomplete page read prevents PDF generation.',
        'This report reads saved data only. It does not add, edit, delete, refresh or repair worklogs. If an entry is absent from both sources, check the Jira task, intended date and final submission confirmation.'
    ]:
        story.extend([p(line), Spacer(1,8)])
    story.extend([PageBreak(), p('Daily reconciliation', 'Heading1'), p(f"{result['employee']} | {result['month']}"), Spacer(1,8),
        table(['Worklog date', 'Jira all hours', 'EPR all hours', 'EPR calendar work', 'Jira minus EPR', 'Result'],
              [[r['date'], hours(r['jira_seconds']), hours(r['epr_seconds']), hours(r['calendar_seconds']),
                hours(r['jira_seconds']-r['epr_seconds']), r['status']] for r in result['daily']], [100,115,115,130,110,200], compact=True),
        PageBreak(), p('Individual saved worklogs', 'Heading1'),
        p('A dash means that source did not return the entry. EPR planned dates are shown to distinguish them from saved worklog dates.'), Spacer(1,8)])
    rows = []
    for r in result['entries']:
        jira, epr = r['jira'], r['epr']
        rows.append([r['issue_key'] + '\n' + r['summary'], r['worklog_id'],
            (jira['date']+' / '+hours(jira['seconds'])) if jira else '-',
            (epr['date']+' / '+hours(epr['seconds'])) if epr else '-',
            ((epr.get('planned_start') or '-')+' to '+(epr.get('planned_due') or '-')) if epr else '-', r['status'], r['calendar']])
    story.append(table(['Task / title', 'Log ID', 'Jira date / hours', 'EPR date / hours', 'EPR task planned dates', 'Result', 'Calendar work'],
                       rows or [['No saved entries found in either source.', '', '', '', '', '', '']], [170,50,105,105,115,95,130]))
    def footer(canvas, doc):
        canvas.setFont('Helvetica', 8)
        canvas.setFillColor(colors.HexColor('#586b82'))
        canvas.drawString(32, 16, f"EPR Worklog Tally | {result['month']} | Saved records visible to the reporting account")
        canvas.drawRightString(810, 16, f'Page {doc.page}')
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    output.seek(0)
    return output
