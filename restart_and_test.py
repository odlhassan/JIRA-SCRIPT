#!/usr/bin/env python
import os
import subprocess
import time
import sys
import glob
import requests

os.chdir('E:\\JIRA SCRIPT')

# Step 1: Kill process 937236
print('Step 1: Killing process 937236...')
try:
    os.system('taskkill /PID 937236 /F 2>nul')
    print('Process killed')
except Exception as e:
    print(f'No process to kill: {e}')

# Step 2: Delete marker files
print('\nStep 2: Deleting marker files...')
for f in glob.glob('.codex_run_server_*_pid'):
    try:
        os.remove(f)
        print(f'Deleted {f}')
    except Exception as e:
        pass

# Step 3: Wait 2 seconds
print('\nStep 3: Waiting 2 seconds...')
time.sleep(2)

# Step 4: Start server in background
print('Step 4: Starting Python server...')
startupinfo = None
if sys.platform == 'win32':
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE

proc = subprocess.Popen([sys.executable, 'run_server.py'], 
                       cwd='E:\\JIRA SCRIPT',
                       startupinfo=startupinfo,
                       stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL)
print(f'Server started with PID: {proc.pid}')

# Step 5: Wait 8 seconds for it to start
print('\nStep 5: Waiting 8 seconds for server to start...')
time.sleep(8)

print('\nServer startup complete, testing endpoints...')

# Step 6 & 7: Test the endpoints
print('\n' + '='*70)
print('Testing Endpoint 1: /api/team-capacity-planner/assignments')
print('='*70)
try:
    response = subprocess.run(
        ['curl', '-s', 'http://127.0.0.1:3000/api/team-capacity-planner/assignments'],
        capture_output=True,
        text=True
    )
    print(response.stdout)
    if response.stderr:
        print(f'Error: {response.stderr}')
except Exception as e:
    print(f'curl command failed: {e}')
    # Try with urllib if curl is not available
    try:
        response = requests.get('http://127.0.0.1:3000/api/team-capacity-planner/assignments', timeout=5)
        print(response.text)
    except Exception as e2:
        print(f'Failed to reach endpoint: {e2}')

print('\n' + '='*70)
print('Testing Endpoint 2: /api/team-capacity-planner/epic-children?epic_key=TEST')
print('='*70)
try:
    response = subprocess.run(
        ['curl', '-s', 'http://127.0.0.1:3000/api/team-capacity-planner/epic-children?epic_key=TEST'],
        capture_output=True,
        text=True
    )
    print(response.stdout)
    if response.stderr:
        print(f'Error: {response.stderr}')
except Exception as e:
    print(f'curl command failed: {e}')
    # Try with urllib if curl is not available
    try:
        response = requests.get('http://127.0.0.1:3000/api/team-capacity-planner/epic-children?epic_key=TEST', timeout=5)
        print(response.text)
    except Exception as e2:
        print(f'Failed to reach endpoint: {e2}')

print('\n' + '='*70)
print('Server restart complete')
print('='*70)
