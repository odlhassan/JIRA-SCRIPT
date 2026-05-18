#!/usr/bin/env python3
"""
Simple HTTP test to check if server is running and inspect page structure
"""
import sys
import socket
from urllib.request import urlopen
from urllib.error import URLError

def test_port(port):
    """Test if a port is open"""
    try:
        response = urlopen(f"http://127.0.0.1:{port}/team_capacity_planner.html", timeout=3)
        return True, response.status
    except URLError as e:
        return False, str(e)
    except socket.timeout:
        return False, "Timeout"
    except Exception as e:
        return False, str(e)

ports = [3000, 3001, 4173, 5000, 8000, 8080]
print("Testing ports for running server...")
for port in ports:
    success, result = test_port(port)
    if success:
        print(f"✓ Port {port} is running (status: {result})")
        sys.exit(0)
    else:
        print(f"✗ Port {port}: {result}")

print("\nNo server found on any port")
