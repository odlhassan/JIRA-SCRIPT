#!/usr/bin/env python3
import base64
import json
import subprocess
import os
import sys

# Step 1: Read and encode the file
file_path = r'E:\JIRA SCRIPT\monthly_epic_plan_progress_report.html'
work_dir = os.path.dirname(file_path)

print("Step 1: Reading and encoding file...")
with open(file_path, 'rb') as f:
    content = f.read()

b64_content = base64.b64encode(content).decode('utf-8')
print(f"✓ File encoded: {len(content)} bytes -> {len(b64_content)} chars")

# Step 2: Find GitHub token
print("\nStep 2: Looking for GitHub token...")
token = None
for env_var in ['GITHUB_TOKEN', 'GH_TOKEN']:
    token = os.environ.get(env_var)
    if token:
        print(f"✓ Found token in {env_var}")
        break

if not token:
    print("✗ GitHub token not found in environment variables")
    sys.exit(1)

# Step 3: Create JSON payload
print("\nStep 3: Creating JSON payload...")
message = """feat: redesign carry-forward epic detection with typed reasons

HTML changes: colour-coded reason chips banner in expanded rows, count badge in Carry Forward column. Reasons include: start slipped, end slipped, due date passed, no worklog in 7 days, over budget. carried_forward and brought_forward are mutually exclusive.

Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"""

payload = {
    "message": message,
    "content": b64_content,
    "sha": "596a595dbddd098ee8126dd9b6a69f24d6e7266c",
    "branch": "main"
}

payload_file = os.path.join(work_dir, '_push_payload.json')
with open(payload_file, 'w') as f:
    json.dump(payload, f)

print(f"✓ Payload saved to: {payload_file}")

# Step 4: Push to GitHub using curl
print("\nStep 4: Pushing to GitHub...")
url = "https://api.github.com/repos/odlhassan/JIRA-SCRIPT/contents/monthly_epic_plan_progress_report.html"
headers = [
    f"Authorization: Bearer {token}",
    "Content-Type: application/json"
]

curl_cmd = ['curl', '-s', '-X', 'PUT', url]
for header in headers:
    curl_cmd.extend(['-H', header])
curl_cmd.extend(['--data', f'@{payload_file}'])

result = subprocess.run(curl_cmd, capture_output=True, text=True)

print("\n=== HTTP Response ===")
print(result.stdout)
if result.stderr:
    print("Stderr:", result.stderr)

# Step 5: Parse response to check status
try:
    response = json.loads(result.stdout)
    if 'message' in response and 'Content API' not in response['message']:
        print("\n✓ Push successful!")
    else:
        print("\n✗ Push failed or returned error")
except:
    print("\n? Could not parse response as JSON")

# Step 6: Cleanup
print("\nStep 5: Cleaning up temporary files...")
for temp_file in [payload_file]:
    if os.path.exists(temp_file):
        os.remove(temp_file)
        print(f"✓ Removed: {temp_file}")
