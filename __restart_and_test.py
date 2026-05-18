#!/usr/bin/env python3
"""
Test server restart sequence as requested
"""
import os
import sys
import time
import signal
import glob
import subprocess
import urllib.request
import urllib.error
import json
from pathlib import Path

os.chdir("E:\\JIRA SCRIPT")

print("=" * 70)
print("REPORT SERVER RESTART TEST")
print("=" * 70)

# STEP 1: Kill existing processes running run_server.py
print("\nSTEP 1: Killing existing python processes running run_server.py")
print("-" * 70)
try:
    # Use tasklist to find python processes
    result = subprocess.run(["tasklist", "/FI", "ImageName eq python.exe"], 
                          capture_output=True, text=True)
    print(f"Active python processes:\n{result.stdout}")
except Exception as e:
    print(f"  (Could not list processes: {e})")

# STEP 2: Kill by PID files
print("\nSTEP 2: Killing processes from PID marker files")
print("-" * 70)
pids_killed = 0
for pidfile in glob.glob(".codex_run_server_*_pid"):
    try:
        with open(pidfile, 'r') as f:
            pid_str = f.read().strip()
        if pid_str.isdigit():
            pid = int(pid_str)
            try:
                # On Windows, os.kill with signal.SIGTERM is a force kill
                os.kill(pid, signal.SIGTERM)
                print(f"  ✓ Killed PID {pid} from {pidfile}")
                pids_killed += 1
            except ProcessLookupError:
                print(f"  - PID {pid} not found (already dead)")
            except Exception as e:
                print(f"  ! Error killing {pid}: {e}")
    except Exception as e:
        print(f"  ! Error reading {pidfile}: {e}")

if pids_killed == 0:
    print("  (No existing processes found)")

time.sleep(1)

# STEP 3: Clean up PID files
print("\nSTEP 3: Cleaning up PID marker files")
print("-" * 70)
for pidfile in glob.glob(".codex_run_server_*_pid"):
    try:
        os.remove(pidfile)
        print(f"  ✓ Deleted {pidfile}")
    except Exception as e:
        print(f"  ! Error deleting {pidfile}: {e}")

# STEP 4: Start fresh server
print("\nSTEP 4: Starting fresh report server")
print("-" * 70)
cmd = [sys.executable, "run_server.py"]
print(f"  Command: {' '.join(cmd)}")
try:
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    print(f"  ✓ Server process started (PID: {proc.pid})")
except Exception as e:
    print(f"  ✗ Failed to start server: {e}")
    sys.exit(1)

# STEP 5: Wait for server initialization
print("\nSTEP 5: Waiting 8 seconds for server to initialize...")
print("-" * 70)
time.sleep(8)
print("  ✓ Wait complete")

# STEP 6: Test the work-items endpoint
print("\nSTEP 6: Testing /api/team-capacity-planner/work-items endpoint")
print("-" * 70)
test_url = "http://127.0.0.1:3000/api/team-capacity-planner/work-items?project=MN"
print(f"  URL: {test_url}")
print()

try:
    with urllib.request.urlopen(test_url, timeout=10) as response:
        data = response.read().decode('utf-8')
        
        print(f"  HTTP Status: {response.status}")
        print(f"  Response size: {len(data)} bytes")
        
        # Check for MN-137
        if "MN-137" in data:
            print("\n  " + "=" * 66)
            print("  ✓ SUCCESS: MN-137 FOUND in the work-items response")
            print("  " + "=" * 66)
        else:
            print("\n  " + "=" * 66)
            print("  ✓ Server responding but MN-137 NOT found in response")
            print("  " + "=" * 66)
            
            # Try to parse and show sample
            try:
                response_json = json.loads(data)
                if 'workItems' in response_json:
                    items = response_json['workItems']
                    print(f"\n  Response contains {len(items)} work items")
                    if items:
                        print(f"  First few items:")
                        for item in items[:3]:
                            print(f"    - {item.get('key', 'unknown')}: {item.get('summary', '')[:50]}")
            except json.JSONDecodeError:
                print("\n  Response preview (first 300 chars):")
                print(f"  {data[:300]}...")
                
except urllib.error.URLError as e:
    print(f"  ✗ Connection Error: {e}")
    print("\n  Server may still be starting. Check after a moment.")
except urllib.error.HTTPError as e:
    print(f"  ✗ HTTP Error {e.code}: {e.reason}")
except socket.timeout:
    print(f"  ✗ Request timeout after 10 seconds")
except Exception as e:
    print(f"  ✗ Unexpected error: {e}")

print("\n" + "=" * 70)
print("RESTART SEQUENCE COMPLETE")
print("=" * 70)
