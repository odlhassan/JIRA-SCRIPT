"""
Jira API client for fetching board, issues, stories, and subtasks.
Used by the Jira Deadline Dashboard to get data from Jira.
"""
import base64
import os
import re
from collections import defaultdict

import requests
from dotenv import load_dotenv

load_dotenv()

BASE_URL = f"https://{os.getenv('JIRA_SITE', 'octopusdtlsupport')}.atlassian.net"
EMAIL = os.getenv("JIRA_EMAIL", "hassan.malik@octopusdtl.com")
API_TOKEN = os.getenv("JIRA_API_TOKEN")
BOARD_NAME = os.getenv("JIRA_BOARD", "O2")


def extract_jira_key_from_url(url):
    """Extract Jira key (e.g. O2-941) from browse URL."""
    if not url:
        return None
    match = re.search(r"/browse/([A-Za-z0-9]+-\d+)", str(url))
    return match.group(1) if match else None


def get_auth_header():
    """Build Basic auth header from email and API token."""
    if not API_TOKEN:
        raise ValueError(
            "JIRA_API_TOKEN not set. Create a .env file (copy from env.example) with your token."
        )
    credentials = f"{EMAIL}:{API_TOKEN}"
    encoded = base64.b64encode(credentials.encode()).decode()
    return {"Authorization": f"Basic {encoded}"}


def get_session():
    """Return a requests Session with Jira auth headers."""
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        **get_auth_header(),
    }
    session = requests.Session()
    session.headers.update(headers)
    return session


def get_board_id(session):
    """Resolve board name to numeric ID. Matches exact or board containing name."""
    url = f"{BASE_URL}/rest/agile/1.0/board"
    r = session.get(url)
    r.raise_for_status()
    data = r.json()
    boards = data.get("values", [])
    for b in boards:
        if b.get("name") == BOARD_NAME or BOARD_NAME in (b.get("name") or ""):
            return b["id"]
    raise ValueError(f'Board "{BOARD_NAME}" not found. Available: {[b["name"] for b in boards]}')


def get_board_config(session, board_id):
    """Get board configuration to determine JQL context."""
    url = f"{BASE_URL}/rest/agile/1.0/board/{board_id}/configuration"
    r = session.get(url)
    r.raise_for_status()
    return r.json()


def get_stories(session, board_id, excel_jira_keys=None):
    """Fetch open stories due in next 14 days from board. If excel_jira_keys provided, only fetch epics/stories mentioned in Excel."""
    jql = (
        'issuetype = Story AND status NOT IN (Done, Closed) '
        'AND duedate >= startOfDay() AND duedate <= endOfDay("+2w")'
    )
    if excel_jira_keys:
        keys_str = ", ".join(f'"{k}"' for k in excel_jira_keys)
        jql += f" AND (parent in ({keys_str}) OR key in ({keys_str}))"
    url = f"{BASE_URL}/rest/agile/1.0/board/{board_id}/issue"
    all_issues = []
    start_at = 0
    while True:
        params = {
            "jql": jql,
            "startAt": start_at,
            "maxResults": 50,
            "fields": "summary,status,duedate,parent,subtasks,issuetype,customfield_10014,assignee",
        }
        r = session.get(url, params=params)
        r.raise_for_status()
        data = r.json()
        issues = data.get("issues", [])
        all_issues.extend(issues)
        if start_at + len(issues) >= data.get("total", 0):
            break
        start_at += len(issues)
    return all_issues


def get_issue(session, key, fields=None):
    """Fetch single issue with optional fields."""
    url = f"{BASE_URL}/rest/api/3/issue/{key}"
    params = {"fields": fields or "summary,status,assignee"}
    r = session.get(url, params=params)
    r.raise_for_status()
    return r.json()


def get_subtasks_for_story(session, story):
    """Fetch sub-tasks for a story. Jira may include them in fields.subtasks."""
    subtasks_data = story.get("fields", {}).get("subtasks", [])
    subtasks = []
    for st in subtasks_data:
        # May only have key; fetch full details if needed
        key = st.get("key")
        if not key:
            continue
        try:
            issue = get_issue(
                session, key,
                fields="summary,status,assignee"
            )
            fields = issue.get("fields", {})
            assignee = fields.get("assignee")
            assignee_name = assignee.get("displayName", "Unassigned") if assignee else "Unassigned"
            status = fields.get("status", {}).get("name", "Unknown")
            subtasks.append({
                "key": key,
                "summary": fields.get("summary", ""),
                "status": status,
                "assignee": assignee_name,
            })
        except Exception:
            subtasks.append({
                "key": key,
                "summary": "(fetch failed)",
                "status": "Unknown",
                "assignee": "Unassigned",
            })
    return subtasks


def get_epic_key(story):
    """Extract epic key from story. Supports parent and Epic Link."""
    fields = story.get("fields", {})
    # Parent (when story is child of epic)
    parent = fields.get("parent")
    if parent:
        return parent.get("key")
    # Epic Link custom field (can be string key or object with key)
    epic_link = fields.get("customfield_10014")
    if epic_link:
        if isinstance(epic_link, str):
            return epic_link
        return epic_link.get("key") if isinstance(epic_link, dict) else None
    return None


def build_dashboard_data(session, stories):
    """Group stories by epic, fetch subtasks and epic summaries."""
    epic_keys = set()
    story_list = []

    for s in stories:
        fields = s.get("fields", {})
        epic_key = get_epic_key(s)
        if epic_key:
            epic_keys.add(epic_key)
        assignee = fields.get("assignee")
        assignee_name = assignee.get("displayName", "Unassigned") if assignee else "Unassigned"

        story_list.append({
            "key": s["key"],
            "summary": fields.get("summary", ""),
            "status": fields.get("status", {}).get("name", "Unknown"),
            "duedate": fields.get("duedate", ""),
            "assignee": assignee_name,
            "epic_key": epic_key,
            "raw": s,
        })

    # Fetch epic summaries
    epics = {}
    for key in epic_keys:
        try:
            issue = get_issue(session, key, fields="summary")
            epics[key] = {
                "key": key,
                "summary": issue.get("fields", {}).get("summary", key),
            }
        except Exception:
            epics[key] = {"key": key, "summary": key}

    # Add "No Epic" for stories without epic
    epics["_NO_EPIC_"] = {"key": "No Epic", "summary": "No Epic"}

    # Fetch subtasks and attach to stories
    for st in story_list:
        st["subtasks"] = get_subtasks_for_story(session, st["raw"])
        del st["raw"]

    # Group by epic
    grouped = defaultdict(list)
    for st in story_list:
        ek = st["epic_key"] or "_NO_EPIC_"
        grouped[ek].append({
            "key": st["key"],
            "summary": st["summary"],
            "status": st["status"],
            "duedate": st["duedate"],
            "assignee": st.get("assignee", "Unassigned"),
            "subtasks": st["subtasks"],
        })

    result = {"epics": {}, "generated_at": ""}
    for ek, stories_list in grouped.items():
        epic = epics.get(ek, {"key": ek, "summary": ek})
        result["epics"][ek] = {
            "key": epic["key"],
            "summary": epic["summary"],
            "stories": stories_list,
        }
    return result


def fetch_stories_for_epics(session, board_id, epic_keys_from_excel, excel_jira_keys):
    """
    Fetch all stories for the given epic keys (from Excel) plus board stories.
    Returns combined list of raw story dicts from Jira API.
    """
    all_stories = []
    if epic_keys_from_excel:
        print(f"Fetching stories for {len(epic_keys_from_excel)} epics: {', '.join(epic_keys_from_excel)}")
        epic_keys_str = ", ".join(f'"{k}"' for k in epic_keys_from_excel)
        jql = f'issuetype = Story AND (parent in ({epic_keys_str}) OR customfield_10014 in ({epic_keys_str}))'
        url = f"{BASE_URL}/rest/api/3/search/jql"
        next_page_token = None
        while True:
            payload = {
                "jql": jql,
                "maxResults": 50,
                "fields": ["summary", "status", "duedate", "parent", "subtasks", "issuetype", "customfield_10014", "assignee"],
            }
            if next_page_token:
                payload["nextPageToken"] = next_page_token
            try:
                r = session.post(url, json=payload)
                r.raise_for_status()
                data = r.json()
                issues = data.get("issues", [])
                all_stories.extend(issues)
                next_page_token = data.get("nextPageToken")
                if not next_page_token:
                    break
            except Exception as e:
                print(f"Warning: Could not fetch stories for epics: {e}")
                print("Trying alternative approach: fetching stories per epic...")
                for epic_key in epic_keys_from_excel:
                    try:
                        alt_jql = f'issuetype = Story AND (parent = {epic_key} OR customfield_10014 = {epic_key})'
                        alt_next_page_token = None
                        while True:
                            alt_payload = {
                                "jql": alt_jql,
                                "maxResults": 50,
                                "fields": ["summary", "status", "duedate", "parent", "subtasks", "issuetype", "customfield_10014", "assignee"],
                            }
                            if alt_next_page_token:
                                alt_payload["nextPageToken"] = alt_next_page_token
                            alt_r = session.post(url, json=alt_payload)
                            alt_r.raise_for_status()
                            alt_data = alt_r.json()
                            all_stories.extend(alt_data.get("issues", []))
                            alt_next_page_token = alt_data.get("nextPageToken")
                            if not alt_next_page_token:
                                break
                    except Exception as alt_e:
                        print(f"  Warning: Could not fetch stories for epic {epic_key}: {alt_e}")
                break

    # Also get stories from the board query (due in 14 days)
    board_stories = get_stories(session, board_id, excel_jira_keys=excel_jira_keys)

    # Combine and deduplicate
    story_keys_seen = set()
    combined = []
    for story in all_stories + board_stories:
        key = story.get("key")
        if key and key not in story_keys_seen:
            story_keys_seen.add(key)
            combined.append(story)
    return combined
