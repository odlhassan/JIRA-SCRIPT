"""Read-only Jira probe through Azure SCM; never print credentials or response bodies."""
import base64
import json
import os
import shlex
import urllib.request
import xml.etree.ElementTree as ET

profiles = ET.fromstring(os.environ['AZURE_PUBLISH_PROFILE']).findall('publishProfile')
profile = next((p for p in profiles if p.get('publishMethod') == 'ZipDeploy'), profiles[0])
auth = base64.b64encode((profile.get('userName') + ':' + profile.get('userPWD')).encode()).decode()
probe = '''
import sys, os, urllib.parse
sys.path[:0] = ['/home/site/wwwroot', '/home/site/wwwroot/.python_packages/lib/site-packages']
for key, value in list(os.environ.items()):
 if key.startswith('APPSETTING_'): os.environ.setdefault(key[11:], value)
try:
 import worklog_tally
 original=worklog_tally._get
 def traced(session,path,**params):
  try:
   result=original(session,path,**params)
   print('Jira read OK:',path,flush=True)
   return result
  except Exception as exc:
   response=getattr(exc,'response',None)
   print('Jira read FAILED:',path,'HTTP',getattr(response,'status_code',None),'exception',type(exc).__name__,flush=True)
   raise
 worklog_tally._get=traced
 result=worklog_tally.build_worklog_tally('/home/data/assignee_hours_capacity.db','Namra Zahid','2026-10')
 print('Completed read; live seconds:',result['jira_seconds'],flush=True)
except Exception as exc:
 print('Probe outcome:',type(exc).__name__,flush=True)
'''
payload = json.dumps({'command': 'python3 -c ' + shlex.quote(probe), 'dir': '/home/site/wwwroot'}).encode()
request = urllib.request.Request('https://epreporting.scm.azurewebsites.net/api/command', data=payload,
    headers={'Authorization': 'Basic ' + auth, 'Content-Type': 'application/json'})
with urllib.request.urlopen(request, timeout=180) as response:
    result = json.load(response)
print(result.get('Output', ''))
# Error output may contain runtime internals; report exit status only.
print('SCM exit code:', result.get('ExitCode'))
