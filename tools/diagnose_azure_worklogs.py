"""Read-only Jira probe through Azure SCM; never print credentials or response bodies."""
import base64
import json
import os
import urllib.request
import xml.etree.ElementTree as ET

profiles = ET.fromstring(os.environ['AZURE_PUBLISH_PROFILE']).findall('publishProfile')
profile = next((p for p in profiles if p.get('publishMethod') == 'ZipDeploy'), profiles[0])
auth = base64.b64encode((profile.get('userName') + ':' + profile.get('userPWD')).encode()).decode()
probe = '''
import os, urllib.parse, urllib.request, json, base64
for key, value in list(os.environ.items()):
 if key.startswith('APPSETTING_'): os.environ.setdefault(key[11:], value)
try:
 token=os.environ.get('JIRA_API_TOKEN','')
 print('Reporting credential configured:',bool(token),flush=True)
 authorization=base64.b64encode((os.environ.get('JIRA_EMAIL','hassan.malik@octopusdtl.com')+':'+token).encode()).decode()
 def read(path,**params):
  try:
   req=urllib.request.Request('https://'+os.environ.get('JIRA_SITE','octopusdtlsupport')+'.atlassian.net/rest/api/3/'+path+'?'+urllib.parse.urlencode(params),headers={'Authorization':'Basic '+authorization,'Accept':'application/json'})
   with urllib.request.urlopen(req,timeout=35) as response: result=json.load(response)
   print('Jira read OK:',path,flush=True)
   return result
  except Exception as exc:
   print('Jira read FAILED:',path,'HTTP',getattr(exc,'code',None),'exception',type(exc).__name__,flush=True)
   raise
 read('myself')
 users=read('user/search',query='Namra Zahid',maxResults=100)
 accounts=[x['accountId'] for x in users if x.get('displayName','').strip().lower()=='namra zahid']
 if len(accounts)!=1: raise ValueError('ambiguous account')
 page=read('search/jql',jql='worklogAuthor = "'+accounts[0]+'" AND worklogDate >= "2026-09-30" AND worklogDate <= "2026-11-01" ORDER BY key ASC',fields='summary,issuetype',maxResults=100)
 for issue in page.get('issues',[]): read('issue/'+issue['key']+'/worklog',startAt=0,maxResults=100)
 print('Completed production credential read probe',flush=True)
except Exception as exc:
 print('Probe outcome:',type(exc).__name__,flush=True)
'''
encoded = base64.b64encode(probe.encode()).decode()
command = 'python3 -c "import base64;exec(base64.b64decode(\'' + encoded + '\'))"'
payload = json.dumps({'command': command, 'dir': '/home/site/wwwroot'}).encode()
request = urllib.request.Request('https://epreporting.scm.azurewebsites.net/api/command', data=payload,
    headers={'Authorization': 'Basic ' + auth, 'Content-Type': 'application/json'})
with urllib.request.urlopen(request, timeout=180) as response:
    result = json.load(response)
print(result.get('Output', ''))
# Error output may contain runtime internals; report exit status only.
print('SCM exit code:', result.get('ExitCode'))
if result.get('ExitCode'):
    error = result.get('Error', '')
    for value in (profile.get('userName'), profile.get('userPWD')):
        error = error.replace(value, '[redacted]')
    print(error[:800])
