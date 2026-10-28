const WORKLOG_API='/api/employee-capacity-utilization/worklog';
function parseWorklogDuration(value){
 const text=String(value).trim().toLowerCase(),number='\\d+(?:\\.\\d+)?';let minutes=0;
 if(new RegExp('^'+number+'$').test(text))minutes=Number(text);
 else if(new RegExp('^(?:'+number+'\\s*[dhm]\\s*)+$').test(text)){
  const seen=new Set();for(const match of text.matchAll(/(\d+(?:\.\d+)?)\s*([dhm])/g)){
   if(seen.has(match[2]))throw new Error('Use each unit once, for example 1h 30m.');
   seen.add(match[2]);minutes+=Number(match[1])*({d:1440,h:60,m:1}[match[2]]);
  }
 }else throw new Error('Use 30m, 8h, 1d, or a combination such as 1h 30m.');
 if(!Number.isFinite(minutes)||minutes<=0||minutes>1440)throw new Error('Enter more than zero and no more than 24h (1d).');
 if(Math.abs(minutes-Math.round(minutes))>1e-8)throw new Error('Use a duration that resolves to whole minutes, such as 0.5h.');
 return Math.round(minutes);
}
function formatWorklogDuration(minutes){return [Math.floor(minutes/60)?Math.floor(minutes/60)+'h':'',minutes%60?(minutes%60)+'m':''].filter(Boolean).join(' ')||'0m'}
async function worklogRequest(path,body){
 const response=await fetch(WORKLOG_API+path,{method:body?'POST':'GET',headers:body?{'Content-Type':'application/json','X-EPR-Worklog':'1'}:{},body:body?JSON.stringify(body):undefined});
 const result=await response.json();if(!response.ok||result.error)throw new Error(result.error||'Jira request failed');return result;
}
async function openCalendarWorklog(label,date,draft={}){
 const target=document.querySelector('#availability-day-details');
 target.innerHTML=`<h3>＋ Log work · ${esc(date)}</h3><p>Checking Jira account…</p>`;
 target.scrollIntoView({behavior:'smooth',block:'nearest'});
 try{
  const [identity,status]=await Promise.all([worklogRequest('/identity'),fetch('/api/jira-auth/status').then(r=>r.json())]);
  const canWrite=identity.write_enabled&&key(label)===key(identity.author),subtasks=issues.filter(x=>/sub[ -]?task/i.test(txt(x.issue_type))&&!txt(x.issue_key).startsWith('RLT-'));
  const authText=identity.signed_in?'Signed in to Jira':'Configured Jira account';
  target.innerHTML=`<div class="worklog-heading"><div class="worklog-heading-title">${availabilityIcon('add_task')}<div><h3>Log work</h3><p>${esc(new Date(date+'T12:00:00').toLocaleDateString(undefined,{weekday:'long',day:'numeric',month:'long',year:'numeric'}))}</p></div></div><span class="worklog-date-chip">${availabilityIcon('calendar_today')}${esc(date)}</span></div><div class="worklog-account">${availabilityIcon('verified_user')}<span>${authText}: <strong>${esc(identity.author)}</strong></span>${status.configured?'<a href="/api/jira-auth/login">Sign in with Jira</a>':'<span>Personal Jira sign-in awaits administrator setup.</span>'}${identity.signed_in?' <button type="button" id="worklog-signout">Sign out</button>':''}</div>${canWrite?'':`<p class="worklog-error">Sign in with your personal Jira account. The configured account is ${esc(identity.author)}. Sign in as ${esc(label)} before logging from this calendar.</p>`}<form id="calendar-worklog-form" class="calendar-worklog-form"><div class="worklog-compose"><section class="worklog-panel"><h4><span>1</span> Choose a subtask</h4><label class="worklog-search">Find a subtask<input id="worklog-search" type="search" placeholder="Search Jira key or title" autocomplete="off" ${canWrite?'':'disabled'}><select id="worklog-issue" size="4" required aria-label="Choose a subtask" ${canWrite?'':'disabled'}></select></label><div id="worklog-selection" class="worklog-selection" aria-live="polite">Select the subtask you worked on.</div></section><section class="worklog-panel"><h4><span>2</span> Add your time</h4><div class="worklog-inputs"><label>Time spent<input id="worklog-duration" type="text" maxlength="100" value="${esc(draft.duration||'30m')}" placeholder="e.g. 1h 30m" required aria-describedby="worklog-duration-help worklog-duration-preview" autocomplete="off"></label><label>Start time<input id="worklog-start" type="time" value="${esc(draft.start||'09:00')}" required></label></div><p id="worklog-duration-help" class="worklog-help">Use d, h, m or mix them. <strong>1d = 24h</strong>; numbers alone mean minutes.</p><div class="worklog-presets" aria-label="Quick durations">${['30m','1h','4h','8h'].map(value=>`<button type="button" data-duration="${value}">${value}</button>`).join('')}</div><p id="worklog-duration-preview" class="worklog-duration-preview" aria-live="polite"></p><label class="worklog-description">Work description<textarea id="worklog-comment" maxlength="2000" rows="3" required placeholder="What did you work on?">${esc(draft.comment||'')}</textarea></label></section></div><div class="worklog-compose-footer"><p>${availabilityIcon('shield')} Max 24h per person, subtask & day. Existing Jira hours are checked again before saving.</p><button type="submit" class="worklog-primary" ${canWrite?'':'disabled'}>Review worklog</button></div><p class="worklog-help">Start time uses your browser’s timezone. Estimates stay unchanged. You’ll confirm everything on the next step.</p><p id="worklog-form-status" role="status"></p></form>`;
  const form=target.querySelector('form'),select=target.querySelector('#worklog-issue'),search=target.querySelector('#worklog-search'),message=target.querySelector('#worklog-form-status');
  const duration=target.querySelector('#worklog-duration'),preview=target.querySelector('#worklog-duration-preview');
  function selection(){const item=subtasks.find(x=>x.issue_key===select.value);target.querySelector('#worklog-selection').innerHTML=item?`<strong>${esc(item.issue_key)}</strong><span>${esc(item.summary)}</span><small>Assigned to ${esc(item.assignee||'Unassigned')}</small>`:'Select the subtask you worked on.'}
  function options(){const previous=select.value,q=key(search.value);select.innerHTML=subtasks.filter(x=>key(x.issue_key+' '+x.summary).includes(q)).sort((a,b)=>(key(b.assignee)===key(label))-(key(a.assignee)===key(label))||txt(a.issue_key).localeCompare(txt(b.issue_key))).slice(0,100).map(x=>`<option value="${esc(x.issue_key)}">${esc(x.issue_key)} · ${esc(x.summary)} (${esc(x.assignee||'Unassigned')})</option>`).join('');if([...select.options].some(x=>x.value===previous))select.value=previous;selection()}
  function validateDuration(){try{const minutes=parseWorklogDuration(duration.value);duration.setCustomValidity('');duration.setAttribute('aria-invalid','false');preview.classList.remove('worklog-error');preview.textContent=`${formatWorklogDuration(minutes)} · ${minutes} minutes`;return minutes}catch(error){duration.setCustomValidity(error.message);duration.setAttribute('aria-invalid','true');preview.classList.add('worklog-error');preview.textContent=error.message;return null}}
  search.value=draft.search||'';options();if(draft.issue){select.value=draft.issue;selection()}search.addEventListener('input',options);select.addEventListener('change',selection);duration.addEventListener('input',validateDuration);validateDuration();
  for(const preset of target.querySelectorAll('[data-duration]'))preset.addEventListener('click',()=>{duration.value=preset.dataset.duration;validateDuration();duration.focus()});
  if(canWrite)search.focus();
  target.querySelector('#worklog-signout')?.addEventListener('click',async()=>{await fetch('/api/jira-auth/logout',{method:'POST'});openCalendarWorklog(label,date)});
  form.addEventListener('submit',async event=>{
   event.preventDefault();const button=form.querySelector('[type=submit]');button.disabled=true;message.textContent='Checking subtask and Jira permissions…';
   try{
    const started=new Date(date+'T'+target.querySelector('#worklog-start').value+':00');
    const offset=-started.getTimezoneOffset(),sign=offset>=0?'+':'-',zone=sign+String(Math.floor(Math.abs(offset)/60)).padStart(2,'0')+':'+String(Math.abs(offset)%60).padStart(2,'0');
    const editDraft={search:search.value,issue:select.value,duration:duration.value,start:target.querySelector('#worklog-start').value,comment:target.querySelector('#worklog-comment').value};
    const result=await worklogRequest('/review',{employee:label,issue_key:select.value,duration:duration.value,minutes:parseWorklogDuration(duration.value),started:date+'T'+target.querySelector('#worklog-start').value+':00'+zone,comment:editDraft.comment}),r=result.review;
    target.innerHTML=`<h3>Review before logging to Jira</h3><div class="worklog-review"><p><strong>Author:</strong> ${esc(r.author)}</p><p><strong>Subtask:</strong> ${detailCell(issueLink(r.issue_key))} · ${esc(r.summary)}</p><p><strong>Start:</strong> ${esc(r.started)}</p><p><strong>Duration:</strong> ${esc(r.duration_label||formatWorklogDuration(r.minutes))} (${r.minutes} minutes)</p><p><strong>Description:</strong> ${esc(r.comment)}</p>${typeof r.total_minutes==='number'?`<div class="worklog-day-total">${availabilityIcon('timelapse')} Day total after this entry: <strong>${Number((r.total_minutes/60).toFixed(2))}h / 24h</strong> · already logged ${Number((r.logged_minutes/60).toFixed(2))}h</div>`:''}<p>Original and remaining estimates will stay unchanged.</p></div><label class="worklog-confirm"><input id="worklog-confirm" type="checkbox"> I checked the author, subtask, date, and duration. Create this Jira worklog.</label><div class="worklog-actions"><button id="worklog-back" type="button">Back</button><button id="worklog-submit" class="worklog-primary" type="button" disabled>Confirm & log ${esc(r.duration_label||formatWorklogDuration(r.minutes))}</button></div><p id="worklog-save-status" role="status"></p>`;
    const confirm=target.querySelector('#worklog-confirm'),submit=target.querySelector('#worklog-submit'),saveStatus=target.querySelector('#worklog-save-status');
    confirm.addEventListener('change',()=>submit.disabled=!confirm.checked);
    target.querySelector('#worklog-back').addEventListener('click',()=>openCalendarWorklog(label,date,editDraft));
    let sending=false;
    submit.addEventListener('click',async()=>{
     if(sending||!confirm.checked)return;sending=true;submit.disabled=true;confirm.disabled=true;target.querySelector('#worklog-back').disabled=true;saveStatus.textContent='Saving once to Jira…';
     try{
      const saved=await worklogRequest('/submit',{token:result.token,confirmed:true});
      if(saved.worklog&&!allWorklogs.some(x=>x.worklog_id===saved.worklog_id))allWorklogs.push({...saved.worklog,worklog_id:saved.worklog_id});
      logs=includeLeaves.checked?allWorklogs:allWorklogs.filter(x=>!isLeaveWorklog(x));data.worklogs=logs;render();
      window.openEcuDrawer(label,'availability',drawerPreviousFocus);
      const detail=document.querySelector('#availability-day-details');detail.innerHTML=`<h3>✓ Worklog ${esc(saved.worklog_id)} created</h3><p>${r.minutes} minutes logged by ${esc(r.author)} on ${detailCell(issueLink(r.issue_key))}. ${saved.verified?'Verified in Jira.':'Created; verify the entry in Jira before doing anything else.'}</p><p class="color-help">This calendar is updated. Other report sessions receive the entry after the next canonical refresh.</p>`;
     }catch(error){saveStatus.textContent=error.message+' Do not create a fresh request until you have checked the subtask’s Jira worklog.';saveStatus.className='worklog-error'}
    });
   }catch(error){message.textContent=error.message;message.className='worklog-error';button.disabled=false}
  });
 }catch(error){target.innerHTML=`<h3>Log work</h3><p class="worklog-error">${esc(error.message)}</p><a href="/api/jira-auth/login">Sign in with Jira</a>`}
}
