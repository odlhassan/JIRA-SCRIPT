#!/usr/bin/env python3
"""
Direct server restart and test script
"""
import os
import sys
import subprocess
import time
import socket
import json
from pathlib import Path

def kill_pids(pids):
    """Kill processes by PID"""
    for pid in pids:
        try:
            os.kill(pid, 9)  # SIGKILL on Unix, force kill on Windows
            print(f"[kill] Killed PID {pid}")
        except ProcessLookupError:
            print(f"[kill] PID {pid} already dead")
        except Exception as e:
            print(f"[kill] Error killing PID {pid}: {e}")

def read_pid_files():
    """Read all PID marker files and extract PIDs"""
    base_dir = Path("E:\\JIRA SCRIPT")
    pids = []
    for pid_file in base_dir.glob(".codex_run_server_*_pid"):
        try:
            with open(pid_file) as f:
                pid_text = f.read().strip()
                if pid_text.isdigit():
                    pids.append(int(pid_text))
                    print(f"[read_pids] Found PID {pid_text} in {pid_file.name}")
        except Exception as e:
            print(f"[read_pids] Error reading {pid_file}: {e}")
    return pids

def clean_pid_files():
    """Delete all PID marker files"""
    base_dir = Path("E:\\JIRA SCRIPT")
    for pid_file in base_dir.glob(".codex_run_server_*_pid"):
        try:
            pid_file.unlink()
            print(f"[clean] Deleted {pid_file.name}")
        except Exception as e:
            print(f"[clean] Error deleting {pid_file}: {e}")

def port_available(port):
    """Check if a port is available"""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.bind(("127.0.0.1", port))
        return True
    except OSError:
        return False

def main():
    print("=" * 60)
    print("[restart] Report Server Restart Sequence")
    print("=" * 60)
    
    # Step 1: Kill old processes
    print("\n[step 1] Killing old processes...")
    pids = read_pid_files()
    if pids:
        kill_pids(pids)
        time.sleep(2)
    else:
        print("[step 1] No PIDs found in marker files")
    
    # Step 2: Clean up marker files
    print("\n[step 2] Cleaning up marker files...")
    clean_pid_files()
    
    # Step 3: Find available port
    print("\n[step 3] Finding available port...")
    port = None
    for p in range(3000, 3010):
        if port_available(p):
            port = p
            print(f"[step 3] Found available port: {port}")
            break
    
    if port is None:
        print("[ERROR] No available ports 3000-3009")
        sys.exit(1)
    
    # Step 4: Start the server
    print(f"\n[step 4] Starting server on port {port}...")
    os.chdir("E:\\JIRA SCRIPT")
    cmd = [sys.executable, "run_server.py", "--no-sync", "--port", str(port)]
    print(f"[step 4] Command: {' '.join(cmd)}")
    
    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        print(f"[step 4] Server started with PID: {proc.pid}")
        
        # Wait for output indicating server is running
        print(f"\n[step 5] Waiting for server to start...")
        for i in range(30):
            try:
                # Try to read from stderr (where Flask logs go)
                line = proc.stderr.readline()
                if line:
                    print(f"[server] {line.rstrip()}")
                    if "Running on" in line or "Serving" in line:
                        print(f"[step 5] Server is ready!")
                        break
            except Exception as e:
                print(f"[step 5] Read error: {e}")
                break
            time.sleep(0.5)
        
        # Give it a moment
        time.sleep(2)
        
        # Step 6: Test endpoints
        print(f"\n[step 6] Testing endpoints...")
        try:
            import urllib.request
            import urllib.error
            
            # Test 1
            url1 = f"http://127.0.0.1:{port}/api/team-capacity-planner/epic-children?epic_key=TEST"
            print(f"[test] GET {url1}")
            try:
                with urllib.request.urlopen(url1, timeout=5) as resp:
                    data = resp.read().decode()
                    print(f"[test] Status: {resp.status}")
                    print(f"[test] Response: {data[:300]}...")
            except urllib.error.HTTPError as e:
                print(f"[test] HTTP Error {e.code}: {e.reason}")
            except Exception as e:
                print(f"[test] Error: {e}")
            
            # Test 2
            url2 = f"http://127.0.0.1:{port}/api/team-capacity-planner/assignments"
            print(f"\n[test] GET {url2}")
            try:
                with urllib.request.urlopen(url2, timeout=5) as resp:
                    data = resp.read().decode()
                    print(f"[test] Status: {resp.status}")
                    print(f"[test] Response: {data[:300]}...")
            except urllib.error.HTTPError as e:
                print(f"[test] HTTP Error {e.code}: {e.reason}")
            except Exception as e:
                print(f"[test] Error: {e}")
        
        except Exception as e:
            print(f"[step 6] Testing failed: {e}")
        
        print("\n" + "=" * 60)
        print(f"[SUCCESS] Server restarted on port {port}")
        print(f"[INFO] Server PID: {proc.pid}")
        print(f"[INFO] Access at: http://127.0.0.1:{port}")
        print("=" * 60)
        
    except Exception as e:
        print(f"[ERROR] Failed to start server: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
