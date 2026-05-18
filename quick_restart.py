import subprocess
import sys

result = subprocess.run([
    sys.executable,
    '-c',
    '''
import os, subprocess, time, glob, sys
os.chdir("E:\\\\JIRA SCRIPT")
print("Killing process 937236...")
os.system("taskkill /PID 937236 /F 2>nul")
print("Deleting marker files...")
for f in glob.glob(".codex_run_server_*_pid"):
    try: os.remove(f)
    except: pass
print("Waiting 2 seconds...")
time.sleep(2)
print("Starting server...")
subprocess.Popen([sys.executable, "run_server.py"], creationflags=subprocess.CREATE_NO_WINDOW)
print("Waiting 8 seconds...")
time.sleep(8)
print("Testing endpoints...")
import subprocess as sp
print("\\n" + "="*70)
print("Endpoint 1: /api/team-capacity-planner/assignments")
print("="*70)
r1 = sp.run(["curl", "-s", "http://127.0.0.1:3000/api/team-capacity-planner/assignments"], capture_output=True, text=True)
print(r1.stdout if r1.stdout else r1.stderr)
print("\\n" + "="*70)
print("Endpoint 2: /api/team-capacity-planner/epic-children?epic_key=TEST")
print("="*70)
r2 = sp.run(["curl", "-s", "http://127.0.0.1:3000/api/team-capacity-planner/epic-children?epic_key=TEST"], capture_output=True, text=True)
print(r2.stdout if r2.stdout else r2.stderr)
print("\\nDone.")
    '''
], capture_output=False, text=True)
sys.exit(result.returncode)
