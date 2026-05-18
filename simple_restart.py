import subprocess
import sys
import time
import os

# List of PIDs to kill based on the marker files
pids_to_kill = [937236, 990352]

print("[restart] Killing old Python processes...")
for pid in pids_to_kill:
    try:
        # Use os.kill on Windows (which uses SIGTERM equivalent)
        os.kill(pid, 9)  # 9 is SIGKILL
        print(f"[restart] Killed PID {pid}")
    except Exception as e:
        print(f"[restart] Could not kill PID {pid}: {e}")

time.sleep(2)

print("[restart] Starting fresh server on port 3000...")
os.chdir("E:\\JIRA SCRIPT")

# Start the server
cmd = [sys.executable, "run_server.py", "--no-sync"]
print(f"[restart] Running: {' '.join(cmd)}")
subprocess.run(cmd)
