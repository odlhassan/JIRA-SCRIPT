#!/usr/bin/env python3
import os
import signal
import glob
import time
import subprocess
import urllib.request
import json

os.chdir('E:\\JIRA SCRIPT')

# Kill processes from PID files
print('Step 1: Killing processes from PID files...')
killed_count = 0
for pidfile in glob.glob('.codex_run_server_*_pid'):
    try:
        with open(pidfile, 'r') as f:
            pid = int(f.read().strip())
        try:
            os.kill(pid, signal.SIGTERM)
            print(f'  Killed PID {pid} from {pidfile}')
            killed_count += 1
        except OSError as e:
            print(f'  PID {pid} not found (already dead): {e}')
    except Exception as e:
        print(f'  Error reading {pidfile}: {e}')

if killed_count == 0:
    print('  No running processes found')

# Start new server
print('\nStep 2: Starting server...')
proc = subprocess.Popen(['python', 'run_server.py'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
print(f'  Server started with PID {proc.pid}')

# Wait for startup
print('\nStep 3: Waiting 8 seconds for server initialization...')
time.sleep(8)

# Test endpoint
print('\nStep 4: Testing work-items endpoint...')
try:
    url = 'http://127.0.0.1:3000/api/team-capacity-planner/work-items?project=MN'
    with urllib.request.urlopen(url, timeout=5) as response:
        data = response.read().decode('utf-8')
        response_json = json.loads(data)
        
        print('\n=== SERVER STATUS ===')
        if 'MN-137' in data:
            print('✓ Server is running and responding correctly')
            print('✓ MN-137 FOUND in the work-items response')
        else:
            print('✓ Server is running and responding')
            print('✓ MN-137 does NOT appear in the work-items response')
            # Show a sample of what was returned
            if 'workItems' in response_json and response_json['workItems']:
                print(f'  Sample items returned: {len(response_json.get("workItems", []))} items')
                if response_json['workItems']:
                    print(f'    First item: {response_json["workItems"][0]}')
except urllib.error.URLError as e:
    print('✗ Server failed to respond')
    print(f'  Error: {e}')
except json.JSONDecodeError as e:
    print('✗ Invalid JSON response from server')
    print(f'  Error: {e}')
except Exception as e:
    print(f'✗ Unexpected error: {e}')
