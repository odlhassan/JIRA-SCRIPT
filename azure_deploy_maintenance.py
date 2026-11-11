"""GitHub deployment handshake. Credentials stay in SCM requests, never page APIs."""
from __future__ import annotations

import argparse
import base64
import json
import os
import time
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET


def scm_marker(phase, deployment, message):
    profiles = ET.fromstring(os.environ['AZURE_PUBLISH_PROFILE']).findall('publishProfile')
    profile = next((p for p in profiles if p.get('publishMethod') == 'ZipDeploy'), profiles[0])
    username, password = profile.get('userName', ''), profile.get('userPWD', '')
    if not username or not password:
        raise RuntimeError('Missing SCM credentials.')
    payload = base64.b64encode(json.dumps({'phase': phase, 'deployment': deployment, 'message': message, 'updated_at': time.time()}).encode()).decode()
    # The SCM sidecar can update /home/data even while the app container restarts.
    script = f"""import base64,json,os,pathlib,secrets
p=pathlib.Path('/home/data/deployment-maintenance.json')
state=json.loads(base64.b64decode('{payload}'))
if state['phase'] != 'draining':
    old=json.loads(p.read_text())
    if old.get('deployment') != state['deployment']:
        raise RuntimeError('Deployment owner changed; refusing to change availability')
p.parent.mkdir(parents=True,exist_ok=True)
tmp=p.with_name(p.name+'.'+secrets.token_hex(6)+'.tmp')
tmp.write_text(json.dumps(state))
os.replace(tmp,p)
print('Maintenance status updated')
"""
    encoded = base64.b64encode(script.encode()).decode()
    command = 'python3 -c "import base64;exec(base64.b64decode(\'' + encoded + '\'))"'
    token = base64.b64encode(f'{username}:{password}'.encode()).decode()
    req = urllib.request.Request(
        'https://' + os.environ['AZURE_WEBAPP_NAME'] + '.scm.azurewebsites.net/api/command',
        data=json.dumps({'command': command, 'dir': '/home'}).encode(),
        headers={'Authorization': 'Basic ' + token, 'Content-Type': 'application/json'},
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as response:
            result = json.load(response)
    except Exception:
        raise RuntimeError('SCM maintenance update failed; inspect Azure deployment diagnostics.') from None
    if result.get('ExitCode') != 0:
        raise RuntimeError('SCM rejected the maintenance update; check marker ownership and storage.')


def get_status():
    req = urllib.request.Request('https://' + os.environ['AZURE_WEBAPP_NAME'] + '.azurewebsites.net/api/maintenance/status', headers={'Cache-Control': 'no-cache'})
    with urllib.request.urlopen(req, timeout=15) as response:
        state = json.load(response)
    if not isinstance(state, dict) or not isinstance(state.get('ready'), bool) or not isinstance(state.get('available'), bool):
        raise ValueError('The running release does not support maintenance.')
    return state


def wait_for(deployment, *, revision=None, timeout=1200, bootstrap=False):
    deadline = time.monotonic() + timeout
    quiet = 0
    while time.monotonic() < deadline:
        try:
            state = get_status()
        except (urllib.error.URLError, ValueError):
            if bootstrap and not revision:
                print('First installation: old release has no maintenance gate. New release will start paused.')
                return
            time.sleep(5)
            continue
        owned = state.get('deployment') == deployment
        if revision:
            if owned and state['ready'] and state.get('revision') == revision:
                return
        else:
            drained = owned and state['ready'] and state.get('phase') == 'draining' and not state['available'] and state.get('active_requests') == 0 and state.get('background_jobs') == 0
            quiet = quiet + 1 if drained else 0
            if quiet >= 2:
                return
        time.sleep(5)
    raise RuntimeError('Maintenance handshake timed out. App remains paused; do not force deployment over active work.')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('action', choices=['enable', 'drain', 'deploying', 'reopen', 'failed'])
    parser.add_argument('--bootstrap', action='store_true', help='First rollout only: allow a release without the gate.')
    args = parser.parse_args()
    deployment = os.environ['GITHUB_RUN_ID'] + '-' + os.environ.get('GITHUB_RUN_ATTEMPT', '1')
    if args.action == 'enable':
        scm_marker('draining', deployment, 'Update queued. Finishing active work before deployment…')
    elif args.action == 'drain':
        wait_for(deployment, bootstrap=args.bootstrap)
    elif args.action == 'deploying':
        scm_marker('deploying', deployment, 'Deployment in progress. Play a round while we get ready.')
    elif args.action == 'reopen':
        scm_marker('checking', deployment, 'Update installed. Checking that the app is ready…')
        wait_for(deployment, revision=os.environ['GITHUB_SHA'])
        scm_marker('available', deployment, 'The update is complete. The app is ready.')
    elif args.action == 'failed':
        scm_marker('failed', deployment, 'The update needs attention. The app is safely paused. Please check back shortly.')
    print('Deployment maintenance action complete: ' + args.action)


if __name__ == '__main__':
    main()
