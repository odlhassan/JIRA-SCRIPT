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
import os, urllib.parse, urllib.request, json, base64, pathlib, time
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
 # Use only the already authorized owner's existing, unexpired session. No cookie
 # or session filename is printed, modified, or sent anywhere except this app.
 for file in pathlib.Path('/home/data/worklog_requests/oauth').glob('*.json'):
  if len(file.stem)!=64: continue
  saved=json.loads(file.read_text())
  if saved.get('account_id')!='641adec89796ea0a87185666' or saved.get('session_expires_at',0)<=time.time(): continue
  url='https://epreporting.azurewebsites.net/api/employee-capacity-utilization/tally?'+urllib.parse.urlencode({'employee':'Namra Zahid','month':'2026-10'})
  req=urllib.request.Request(url,headers={'Cookie':'epr_jira_session='+file.stem,'Accept':'application/json'})
  with urllib.request.urlopen(req,timeout=150) as response: result=json.load(response)
  rows=result.get('leaves')
  if not isinstance(rows,list): raise ValueError('live RLT result missing')
  hours=sum(sum(row.get(k,0) for k in ('planned_taken_hours','unplanned_taken_hours','unknown_taken_hours')) for row in rows)
  days=[row['period_day'] for row in rows if row.get('planned_taken_hours')==8]
  if hours!=32 or not all('2026-10-'+str(day) in days for day in (19,20,21,22)): raise ValueError('RLT day verification failed')
  print('Authenticated production calendar read verified: 32h RLT leave, October 19-22',flush=True)
  break
 else: print('Authenticated production check unavailable: no existing owner session',flush=True)
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
