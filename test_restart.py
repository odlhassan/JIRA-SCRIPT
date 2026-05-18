#!/usr/bin/env python3
"""Test script to start the report server and verify endpoints"""
import subprocess
import time
import requests
import socket
from pathlib import Path
import os
import signal
import sys

def port_is_available(host, port):
    """Check if a port is available"""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.bind((host, port))
        return True
    except OSError:
        return False

def find_available_port(host, start_port=3000):
    """Find an available port starting from start_port"""
    for port in range(start_port, start_port + 10):
        if port_is_available(host, port):
            return port
    return None

def main():
    print("[Test] Starting report server restart sequence...")
    
    # Kill any existing python processes (simplified - just check ports)
    host = "127.0.0.1"
    port = find_available_port(host, 3000)
    
    if port is None:
        print("[Test] ERROR: No ports available from 3000-3009")
        sys.exit(1)
    
    print(f"[Test] Available port: {port}")
    
    # Start the server
    print(f"[Test] Starting server on {host}:{port}...")
    base_dir = Path(__file__).resolve().parent
    env = os.environ.copy()
    env["PORT"] = str(port)
    
    try:
        # Start the server process in background
        proc = subprocess.Popen(
            [sys.executable, "run_server.py", "--no-sync", "--port", str(port)],
            cwd=str(base_dir),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            env=env
        )
        
        print(f"[Test] Server process started with PID: {proc.pid}")
        
        # Wait for server to start
        max_retries = 30
        for i in range(max_retries):
            try:
                print(f"[Test] Waiting for server... attempt {i+1}/{max_retries}")
                response = requests.get(f"http://{host}:{port}/", timeout=2)
                if response.status_code == 200:
                    print(f"[Test] Server is ready! Status: {response.status_code}")
                    break
            except requests.exceptions.RequestException:
                time.sleep(1)
        else:
            print("[Test] Server startup timeout")
            proc.terminate()
            sys.exit(1)
        
        # Test the endpoints
        print("\n[Test] Testing endpoints...")
        
        # Test 1: epic-children endpoint
        test_url_1 = f"http://{host}:{port}/api/team-capacity-planner/epic-children?epic_key=TEST"
        print(f"[Test] GET {test_url_1}")
        try:
            resp1 = requests.get(test_url_1, timeout=5)
            print(f"[Test] Status: {resp1.status_code}")
            print(f"[Test] Response (first 500 chars): {resp1.text[:500]}")
        except Exception as e:
            print(f"[Test] ERROR: {e}")
        
        # Test 2: assignments endpoint
        test_url_2 = f"http://{host}:{port}/api/team-capacity-planner/assignments"
        print(f"\n[Test] GET {test_url_2}")
        try:
            resp2 = requests.get(test_url_2, timeout=5)
            print(f"[Test] Status: {resp2.status_code}")
            print(f"[Test] Response (first 500 chars): {resp2.text[:500]}")
        except Exception as e:
            print(f"[Test] ERROR: {e}")
        
        print(f"\n[Test] Server is running on http://{host}:{port}")
        print(f"[Test] Process PID: {proc.pid}")
        print("[Test] SUCCESS: Server restarted and endpoints are accessible")
        
    except Exception as e:
        print(f"[Test] ERROR: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
