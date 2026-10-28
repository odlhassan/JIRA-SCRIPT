const WORKLOG_API='/api/employee-capacity-utilization/worklog';
async function worklogRequest(path,body){
 const response=await fetch(WORKLOG_API+path,{method:body?'POST':'GET',headers:body?{'Content-Type':'application/json','X-EPR-Worklog':'1'}:{},body:body?JSON.stringify(body):undefined});
 const result=await response.json();if(!response.ok||result.error)throw new Error(result.error||'Jira request failed');return result;
}
async function openCalendarWorklog(label,date){
 const target=document.querySelector('#availability-day-details');
 target.innerHTML=`<h3>＋ Log work · ${esc(date)}</h3><p>Checking Jira account…</p>`;
 target.scrollIntoView({behavior:'smooth',block:'nearest'});
 try{
  const [identity,status]=await Promise.all([worklogRequest('/identity'),fetch('/api/jira-auth/status').then(r=>r.json())]);
  const canWrite=identity.write_enabled&&key(label)===key(identity.author),subtasks=issues.filter(x=>/sub[ -]?task/i.test(txt(x.issue_type))&&!txt(x.issue_key).startsWith('RLT-'));
  const authText=identity.signed_in?'Signed in to Jira':'Configured Jira account';
  target.innerHTML=`<h3>＋ Log work · ${esc(date)}</h3><p class="worklog-account">${authText}: <strong>${esc(identity.author)}</strong>. ${status.configured?'<a href="/api/jira-auth/login">Sign in with Jira</a>':'<span>Personal Jira sign-in awaits administrator setup.</span>'}${identity.signed_in?' <button type="button" id="worklog-signout">Sign out</button>':''}</p>${canWrite?'':`<p class="worklog-error">Sign in with your personal Jira account. The configured account is ${esc(identity.author)}. Sign in as ${esc(label)} before logging from this calendar.</p>`}<form id="calendar-worklog-form" class="calendar-worklog-form"><label class="worklog-search">Find a subtask<input id="worklog-search" type="search" placeholder="Search Jira key or title" autocomplete="off" ${canWrite?'':'disabled'}><select id="worklog-issue" size="5" required aria-label="Choose a subtask" ${canWrite?'':'disabled'}></select></label><div class="worklog-inputs"><label>Start time<input id="worklog-start" type="time" value="09:00" required></label><label>Time spent (minutes)<input id="worklog-minutes" type="number" min="1" max="1440" step="1" value="30" required></label><label>Work description<textarea id="worklog-comment" maxlength="2000" rows="2" required placeholder="Describe the work completed"></textarea></label></div><p class="color-help">The date is ${esc(date)} in your browser’s timezone. Jira estimates stay unchanged. Nothing is written until you review and confirm.</p><button type="submit" class="worklog-primary" ${canWrite?'':'disabled'}>Review worklog</button><p id="worklog-form-status" role="status"></p></form>`;
  const form=target.querySelector('form'),select=target.querySelector('#worklog-issue'),search=target.querySelector('#worklog-search'),message=target.querySelector('#worklog-form-status');
  function options(){const q=key(search.value);select.innerHTML=subtasks.filter(x=>key(x.issue_key+' '+x.summary).includes(q)).sort((a,b)=>(key(b.assignee)===key(label))-(key(a.assignee)===key(label))||txt(a.issue_key).localeCompare(txt(b.issue_key))).slice(0,100).map(x=>`<option value="${esc(x.issue_key)}">${esc(x.issue_key)} · ${esc(x.summary)} (${esc(x.assignee||'Unassigned')})</option>`).join('')}
  options();search.addEventListener('input',options);if(canWrite)search.focus();
  target.querySelector('#worklog-signout')?.addEventListener('click',async()=>{await fetch('/api/jira-auth/logout',{method:'POST'});openCalendarWorklog(label,date)});
  form.addEventListener('submit',async event=>{
   event.preventDefault();const button=form.querySelector('[type=submit]');button.disabled=true;message.textContent='Checking subtask and Jira permissions…';
   try{
    const started=new Date(date+'T'+target.querySelector('#worklog-start').value+':00');
    const offset=-started.getTimezoneOffset(),sign=offset>=0?'+':'-',zone=sign+String(Math.floor(Math.abs(offset)/60)).padStart(2,'0')+':'+String(Math.abs(offset)%60).padStart(2,'0');
    const result=await worklogRequest('/review',{employee:label,issue_key:select.value,minutes:Number(target.querySelector('#worklog-minutes').value),started:date+'T'+target.querySelector('#worklog-start').value+':00'+zone,comment:target.querySelector('#worklog-comment').value}),r=result.review;
    target.innerHTML=`<h3>Review before logging to Jira</h3><div class="worklog-review"><p><strong>Author:</strong> ${esc(r.author)}</p><p><strong>Subtask:</strong> ${detailCell(issueLink(r.issue_key))} · ${esc(r.summary)}</p><p><strong>Start:</strong> ${esc(r.started)}</p><p><strong>Duration:</strong> ${r.minutes} minutes</p><p><strong>Description:</strong> ${esc(r.comment)}</p><p>Original and remaining estimates will stay unchanged.</p></div><label class="worklog-confirm"><input id="worklog-confirm" type="checkbox"> I checked the author, subtask, date, and duration. Create this Jira worklog.</label><div class="worklog-actions"><button id="worklog-back" type="button">Back</button><button id="worklog-submit" class="worklog-primary" type="button" disabled>Confirm & log ${r.minutes} minutes</button></div><p id="worklog-save-status" role="status"></p>`;
    const confirm=target.querySelector('#worklog-confirm'),submit=target.querySelector('#worklog-submit'),saveStatus=target.querySelector('#worklog-save-status');
    confirm.addEventListener('change',()=>submit.disabled=!confirm.checked);
    target.querySelector('#worklog-back').addEventListener('click',()=>openCalendarWorklog(label,date));
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
