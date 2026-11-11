"""Personal Jira subtask creation. Reuses worklog auth and durable write receipts."""
import calendar
from datetime import datetime, date
from decimal import Decimal
import json
import os
import re
import time
import uuid
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from flask import jsonify, request
from itsdangerous import BadData
import requests

from jira_calendar_cache import invalidate_issue


def month_bounds(identity):
    try:
        zone = ZoneInfo(identity.get('timeZone') or 'Asia/Karachi')
    except ZoneInfoNotFoundError:
        zone = ZoneInfo('Asia/Karachi')
    today = datetime.now(zone).date()
    return today.isoformat(), today.replace(day=1).isoformat(), today.replace(day=calendar.monthrange(today.year, today.month)[1]).isoformat()


def estimate_text(value):
    """Keep Jira's native d/w semantics; worklog durations deliberately differ."""
    value = str(value or '').strip().lower()
    if len(value) > 100 or not re.fullmatch(r'(?:\d+(?:\.\d+)?\s*[wdhm]\s*)+', value):
        raise ValueError('Enter an estimate such as 30m, 4h, 1d, 1w or 1h 30m.')
    parts = re.findall(r'(\d+(?:\.\d+)?)\s*([wdhm])', value)
    if len({unit for _, unit in parts}) != len(parts) or not any(Decimal(amount) > 0 for amount, _ in parts):
        raise ValueError('Use each estimate unit once and enter more than zero.')
    return value


def register_subtask_routes(app, receipt_dir, *, personal_identity, guard, signer, failure, issue_key, jira_base, serialize_submit):
    prefix = '/api/employee-capacity-utilization/worklog/subtask'

    def read(session, path, **params):
        response = session.get(jira_base(session) + '/rest/api/3/' + path, params=params, timeout=(10, 25))
        response.raise_for_status()
        return response.json()

    def project_key(value):
        value = str(value or '').strip().upper()
        if not re.fullmatch(r'[A-Z][A-Z0-9_]*', value):
            raise ValueError('Choose a Jira space.')
        return value

    def metadata_pages(session, path):
        rows, offset = [], 0
        while True:
            page = read(session, path, startAt=offset, maxResults=100)
            batch = page.get('values', page.get('issueTypes', page.get('fields')))
            if not isinstance(batch, list):
                raise ValueError('Jira create fields could not be read. Try again.')
            rows.extend(batch)
            offset += len(batch)
            if page.get('isLast') is True or (isinstance(page.get('total'), int) and offset >= page['total']):
                return rows
            if not batch:
                raise ValueError('Jira returned incomplete create metadata.')

    def hierarchy(session, project, epic, story=None):
        ef = read(session, 'issue/' + epic, fields='summary,issuetype,project')['fields']
        et = ef.get('issuetype') or {}
        if not (et.get('hierarchyLevel') == 1 or et.get('name', '').casefold() == 'epic'):
            raise ValueError('Choose an existing Jira epic.')
        if (ef.get('project') or {}).get('key') != project:
            raise ValueError('The epic must belong to the selected space.')
        if story:
            # The field catalogue resolves legacy Epic Link installations without hard-coded IDs.
            links = [f['id'] for f in read(session, 'field')
                     if (f.get('schema') or {}).get('custom') == 'com.pyxis.greenhopper.jira:gh-epic-link']
            sf = read(session, 'issue/' + story, fields=','.join(['summary', 'issuetype', 'project', 'parent'] + links))['fields']
            st = sf.get('issuetype') or {}
            parent = (sf.get('parent') or {}).get('key')
            if not parent:
                for field in links:
                    link = sf.get(field)
                    parent = link.get('key') if isinstance(link, dict) else link
                    if parent:
                        break
            if st.get('subtask') or st.get('name', '').casefold() not in ('story', 'user story'):
                raise ValueError('The subtask must be inside a user story.')
            if (sf.get('project') or {}).get('key') != project or parent != epic:
                raise ValueError('This user story no longer belongs to the selected epic and space. Choose it again.')
            return ef, sf
        return ef, None

    def create_fields(session, project):
        types = metadata_pages(session, 'issue/createmeta/' + project + '/issuetypes')
        subtasks = [t for t in types if t.get('subtask') is True or t.get('hierarchyLevel') == -1]
        if not subtasks:
            raise ValueError('This space has no subtask type you can create. Ask its Jira administrator.')
        chosen = next((t for t in subtasks if str(t.get('name', '')).casefold() in ('sub-task', 'subtask')), subtasks[0])
        fields = metadata_pages(session, 'issue/createmeta/' + project + '/issuetypes/' + str(chosen['id']))
        by_id = {f.get('fieldId', f.get('key')): f for f in fields}
        starts = [fid for fid, f in by_id.items() if str(f.get('name', '')).strip().casefold() == 'start date'
                  and (f.get('schema') or {}).get('type') == 'date']
        dues = [fid for fid, f in by_id.items() if str(f.get('name', '')).strip().casefold() == 'due date'
                and (f.get('schema') or {}).get('type') == 'date']
        due_id = 'duedate' if 'duedate' in by_id else dues[0] if len(dues) == 1 else None
        missing_fields = [label for present, label in ((len(starts) == 1, 'Start date'), (bool(due_id), 'Due date'),
                          ('assignee' in by_id, 'Assignee'), ('timetracking' in by_id, 'Time tracking')) if not present]
        if missing_fields or not {'summary', 'parent'}.issubset(by_id):
            raise ValueError('This space needs ' + ', '.join(missing_fields or ['Summary and Parent']) + ' on its subtask create screen. Ask its Jira administrator.')
        supported = {'project', 'issuetype', 'summary', 'parent', 'assignee', due_id, 'timetracking', 'description', starts[0]}
        missing = [f.get('name', fid) for fid, f in by_id.items() if f.get('required') and fid not in supported and not f.get('hasDefaultValue')]
        if missing:
            raise ValueError('This space requires additional Jira fields: ' + ', '.join(missing) + '. Create it in Jira or ask its administrator to supply defaults.')
        return str(chosen['id']), starts[0], due_id, by_id

    def validate_dates(body, identity):
        start, due = str(body.get('start_date', '')), str(body.get('due_date', ''))
        for value in (start, due):
            if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', value):
                raise ValueError('Choose a start date and due date.')
            date.fromisoformat(value)
        _, first, last = month_bounds(identity)
        if not first <= due <= last:
            raise ValueError(f'Due date must be in the current month ({first} to {last}).')
        if start > due:
            raise ValueError('Start date cannot be after the due date.')
        return start, due

    def validated(session, identity, body):
        project, epic, story = project_key(body.get('project_key')), issue_key(body.get('epic_key')), issue_key(body.get('story_key'))
        summary = str(body.get('summary', '')).strip()
        description = str(body.get('description', '')).strip()
        if not 1 <= len(summary) <= 255:
            raise ValueError('Give your subtask a title of 1–255 characters.')
        if len(description) > 2000:
            raise ValueError('Keep the description within 2000 characters.')
        start, due = validate_dates(body, identity)
        estimate = estimate_text(body.get('estimate', '1d'))
        ef, sf = hierarchy(session, project, epic, story)
        type_id, start_id, due_id, fields = create_fields(session, project)
        if description and 'description' not in fields:
            raise ValueError('Description is not enabled on this space’s subtask create screen.')
        permissions = read(session, 'mypermissions', projectKey=project, permissions='BROWSE_PROJECTS,CREATE_ISSUES,ASSIGN_ISSUES,SCHEDULE_ISSUES').get('permissions') or {}
        if not all((permissions.get(p) or {}).get('havePermission') for p in ('BROWSE_PROJECTS', 'CREATE_ISSUES', 'ASSIGN_ISSUES', 'SCHEDULE_ISSUES')):
            raise ValueError('Jira needs Browse, Create, Assign and Schedule issues permissions in this space.')
        assignable = read(session, 'user/assignable/search', project=project, accountId=identity['accountId'], maxResults=50)
        if not any(user.get('accountId') == identity['accountId'] for user in assignable):
            raise ValueError('Jira does not allow assigning subtasks to your account in this space.')
        return dict(project_key=project, epic_key=epic, story_key=story, summary=summary, description=description,
                    start_date=start, due_date=due, estimate=estimate, type_id=type_id, start_id=start_id, due_id=due_id,
                    epic_summary=ef.get('summary', ''), story_summary=sf.get('summary', ''))

    @app.get(prefix + '/search')
    def subtask_search():
        """Find loggable subtasks without waiting for the canonical full fetch."""
        try:
            session, _ = personal_identity()
            query, cursor = request.args.get('q', '').strip(), request.args.get('cursor', '')
            if len(query) > 100 or len(cursor) > 10000:
                raise ValueError('Search text or page cursor is too long.')
            jql = 'issuetype in subTaskIssueTypes() AND project != RLT'
            if not query:
                jql += ' AND assignee = currentUser()'
            elif re.fullmatch(r'[A-Za-z][A-Za-z0-9_]*-\d+', query):
                jql += f' AND key = "{query.upper()}"'
            else:
                safe = re.sub(r'[^\w\s-]', ' ', query).strip()
                if not safe:
                    raise ValueError('Search by a Jira key or words in the title.')
                jql += ' AND summary ~ "' + '\\"' + safe + '\\"' + '"'
            params = dict(jql=jql + ' ORDER BY updated DESC', fields='summary,issuetype,assignee,parent,duedate,timetracking', maxResults=50)
            if cursor:
                params['nextPageToken'] = cursor
            page = read(session, 'search/jql', **params)
            rows = page.get('issues')
            if not isinstance(rows, list):
                raise ValueError('Jira subtasks could not be read.')
            items = []
            for item in rows:
                fields = item.get('fields') or {}
                if not (fields.get('issuetype') or {}).get('subtask') or item['key'].startswith('RLT-'):
                    continue
                items.append(dict(issue_key=item['key'], summary=fields.get('summary', ''),
                                  issue_type=(fields.get('issuetype') or {}).get('name', 'Sub-task'),
                                  assignee=(fields.get('assignee') or {}).get('displayName', ''),
                                  parent_issue_key=(fields.get('parent') or {}).get('key', ''), due_date=fields.get('duedate', ''),
                                  original_estimate_hours=(fields.get('timetracking') or {}).get('originalEstimateSeconds', 0) / 3600))
            next_cursor = page.get('nextPageToken') if page.get('isLast') is not True else None
            if next_cursor and next_cursor == cursor:
                raise ValueError('Jira returned an invalid page cursor.')
            return jsonify(issues=items, next_cursor=next_cursor)
        except (ValueError, TypeError, KeyError, requests.RequestException) as exc:
            return failure(exc)

    @app.get(prefix + '/options')
    def subtask_options():
        try:
            session, identity = personal_identity()
            today, first, last = month_bounds(identity)
            kind, query, cursor = request.args.get('kind', 'projects'), request.args.get('q', '').strip(), request.args.get('cursor', '')
            if len(query) > 100 or len(cursor) > 10000:
                raise ValueError('Search text or page cursor is too long.')
            if kind == 'projects':
                if cursor and (not cursor.isdigit() or len(cursor) > 9):
                    raise ValueError('Invalid space page.')
                page = read(session, 'project/search', startAt=int(cursor or 0), maxResults=50, orderBy='name', query=query)
                rows = page.get('values')
                if not isinstance(rows, list):
                    raise ValueError('Jira spaces could not be read.')
                next_cursor = str(int(cursor or 0) + len(rows)) if page.get('isLast') is False and rows else None
                options = [dict(key=p['key'], summary=p['name']) for p in rows]
            else:
                project = project_key(request.args.get('project_key'))
                if kind == 'epics':
                    jql = f'project = "{project}" AND issuetype = Epic'
                elif kind == 'stories':
                    epic = issue_key(request.args.get('epic_key'))
                    hierarchy(session, project, epic)
                    jql = f'project = "{project}" AND parent = "{epic}" AND issuetype in standardIssueTypes()'
                else:
                    raise ValueError('Choose spaces, epics or user stories.')
                if query:
                    # Quoted phrase matching; user text cannot introduce JQL clauses.
                    safe = re.sub(r'[^\w\s-]', ' ', query).strip()
                    if not safe:
                        raise ValueError('Search by a Jira key or words in the title.')
                    if re.fullmatch(r'[A-Za-z][A-Za-z0-9_]*-\d+', query):
                        jql += f' AND key = "{query.upper()}"'
                    else:
                        jql += ' AND summary ~ "' + '\\"' + safe + '\\"' + '"'
                params = dict(jql=jql + ' ORDER BY key ASC', fields='summary,issuetype', maxResults=50)
                if cursor:
                    params['nextPageToken'] = cursor
                page = read(session, 'search/jql', **params)
                rows = page.get('issues')
                if not isinstance(rows, list):
                    raise ValueError('Jira work items could not be read.')
                next_cursor = page.get('nextPageToken') if page.get('isLast') is not True else None
                if next_cursor and next_cursor == cursor:
                    raise ValueError('Jira returned an invalid page cursor.')
                options = [dict(key=p['key'], summary=p['fields'].get('summary', '')) for p in rows
                           if kind != 'stories' or p['fields'].get('issuetype', {}).get('name', '').casefold() in ('story', 'user story')]
            return jsonify(options=options, next_cursor=next_cursor, author=identity.get('displayName', ''), today=today, first=first, last=last)
        except (ValueError, TypeError, KeyError, requests.RequestException) as exc:
            return failure(exc)

    @app.post(prefix + '/review')
    def subtask_review():
        try:
            guard()
            body = request.get_json()
            if not isinstance(body, dict):
                raise ValueError('Complete the subtask form.')
            session, identity = personal_identity()
            review = validated(session, identity, body)
            review.update(account_id=identity['accountId'], author=identity.get('displayName', ''), request_id=uuid.uuid4().hex)
            return jsonify(review=review, token=signer().dumps({'kind': 'subtask', **review}))
        except (ValueError, TypeError, KeyError, requests.RequestException) as exc:
            return failure(exc)

    @app.post(prefix + '/submit')
    @serialize_submit
    def subtask_submit():
        try:
            guard()
            body = request.get_json()
            if not isinstance(body, dict) or body.get('confirmed') is not True:
                raise ValueError('Confirm the subtask details first.')
            try:
                review = signer().loads(str(body.get('token', '')), max_age=600)
            except BadData:
                raise ValueError('The review expired or changed. Review the subtask again.')
            if review.get('kind') != 'subtask':
                raise ValueError('Review the subtask again.')
            session, identity = personal_identity()
            if identity['accountId'] != review['account_id']:
                raise ValueError('Your Jira account changed. Review the subtask again.')
            receipt = receipt_dir / (review['request_id'] + '.json')
            if receipt.exists():
                result = json.loads(receipt.read_text(encoding='utf-8'))
                return jsonify(result), 200 if result.get('issue_key') else 409
            details = validated(session, identity, review)
            fields = dict(project={'key': details['project_key']}, parent={'key': details['story_key']},
                          issuetype={'id': details['type_id']}, summary=details['summary'],
                          assignee={'accountId': identity['accountId']},
                          timetracking={'originalEstimate': details['estimate']})
            fields[details['start_id']] = details['start_date']
            fields[details['due_id']] = details['due_date']
            if details['description']:
                fields['description'] = {'type': 'doc', 'version': 1, 'content': [{'type': 'paragraph', 'content': [{'type': 'text', 'text': details['description']}]}]}
            pending = dict(error='The creation outcome needs verification in Jira. This request will not be resent.', request_id=review['request_id'])
            try:
                fd = os.open(receipt, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
            except FileExistsError:
                return jsonify(pending), 409
            with os.fdopen(fd, 'w', encoding='utf-8') as stream:
                json.dump(pending, stream)
                stream.flush()
                os.fsync(stream.fileno())
            response = session.post(jira_base(session) + '/rest/api/3/issue', json={'fields': fields}, timeout=(10, 40))
            response.raise_for_status()
            created = response.json()
            new_key = issue_key(created.get('key'))
            result = dict(issue_key=new_key, summary=details['summary'], assignee=review['author'],
                          issue_type='Sub-task', parent_issue_key=details['story_key'], start_date=details['start_date'],
                          due_date=details['due_date'], estimate=details['estimate'], verified=False,
                          review=dict(account_id=identity['accountId'], personal_data_retrieved_at=time.time()))
            # Persist the returned key before optional verification; never resend an ambiguous POST.
            receipt.write_text(json.dumps(result), encoding='utf-8')
            invalidate_issue(new_key)
            invalidate_issue(details['story_key'])
            try:
                saved = read(session, 'issue/' + new_key, fields='summary,assignee,parent,issuetype,timetracking,' + details['due_id'] + ',' + details['start_id'])['fields']
                result['verified'] = ((saved.get('assignee') or {}).get('accountId') == identity['accountId'] and
                                      (saved.get('parent') or {}).get('key') == details['story_key'] and
                                      saved.get(details['due_id']) == details['due_date'] and saved.get(details['start_id']) == details['start_date'] and
                                      (saved.get('timetracking') or {}).get('originalEstimateSeconds', 0) > 0)
                result['issue_type'] = (saved.get('issuetype') or {}).get('name', 'Sub-task')
                result['original_estimate_hours'] = (saved.get('timetracking') or {}).get('originalEstimateSeconds', 0) / 3600
            except (ValueError, TypeError, KeyError, requests.RequestException):
                pass
            receipt.write_text(json.dumps(result), encoding='utf-8')
            return jsonify(result)
        except (ValueError, TypeError, KeyError, requests.RequestException) as exc:
            if isinstance(exc, requests.HTTPError) and exc.response is not None:
                try:
                    errors = exc.response.json()
                    message = '; '.join([str(v) for v in (errors.get('errors') or {}).values()] + [str(v) for v in errors.get('errorMessages', [])])
                    if message:
                        return jsonify(error='Jira could not create the subtask: ' + message + '. Check Jira before starting another request.'), 502
                except (ValueError, TypeError):
                    pass
            return failure(exc)
