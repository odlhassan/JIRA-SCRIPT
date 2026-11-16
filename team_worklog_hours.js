/* The matrix shares the capacity report's filters and canonical data. */
let matrixExportRows=[];
let liveMonth='',liveStates=new Map(),liveQueue=[],liveRunning=false,liveEpoch=0,liveAbort=null,liveVisible=new Set();
const liveRefreshButton=document.createElement('button');
liveRefreshButton.type='button';liveRefreshButton.className='matrix-refresh';liveRefreshButton.textContent='Refresh live hours and leave';
document.querySelector('.meta').append(liveRefreshButton);
function resetLiveUpdates(){
 liveEpoch++;liveAbort?.abort();liveAbort=null;liveQueue=[];liveRunning=false;liveStates=new Map();liveMonth=month.value;
}
function liveState(name){
 if(!liveStates.has(name))liveStates.set(name,{worklogs:{status:'new'},leave:{status:'new'}});
 return liveStates.get(name);
}
function queueLiveUpdates(rows){
 liveVisible=new Set(rows.map(row=>key(row.label)));
 for(const row of rows){
  const nk=key(row.label),state=liveState(nk);
  for(const kind of ['worklogs','leave']){
   if(state[kind].status!=='new')continue;
   if(nk==='unassigned'||nk==='grand total'){state[kind].status='error';state[kind].error='No Jira employee account';continue}
   state[kind].status='queued';liveQueue.push({name:row.label,nk,kind,epoch:liveEpoch});
  }
 }
 if(!liveRunning&&liveQueue.length){liveRunning=true;setTimeout(runLiveQueue,0)}
}
async function runLiveQueue(){
 const epoch=liveEpoch;
 while(epoch===liveEpoch&&liveQueue.length){
  const task=liveQueue.shift(),state=liveStates.get(task.nk)?.[task.kind];
  if(!state)continue;
  if(!liveVisible.has(task.nk)){state.status='new';continue}
  state.status='loading';render();
  const controller=new AbortController();liveAbort=controller;
  try{
   const params=new URLSearchParams({employee:task.name,month:liveMonth,kind:task.kind});
   const response=await fetch('/api/team-worklog-hours/live?'+params,{cache:'no-store',signal:controller.signal});
   const result=await response.json();
   if(!response.ok)throw new Error(result.error||'Update failed');
   if(!Array.isArray(result.values)||key(result.employee)!==task.nk||result.month!==liveMonth||result.kind!==task.kind||!result.checked_at)throw new Error('Incomplete Jira response');
   if(task.kind==='worklogs'&&result.values.some(x=>!x.issue_key||!x.issue_type||typeof x.assigned_to_author!=='boolean'||!Number.isFinite(x.seconds)||x.seconds<0||!/^\d{4}-\d{2}-\d{2}$/.test(x.date)))throw new Error('Incomplete Jira worklog data');
   if(epoch!==liveEpoch)return;
   state.values=result.values;state.checkedAt=result.checked_at;state.status='live';state.error='';
  }catch(error){if(epoch!==liveEpoch)return;state.status='error';state.error=error.message||'Update failed'}
  finally{if(epoch===liveEpoch)liveAbort=null}
  if(epoch===liveEpoch)render();
 }
 if(epoch===liveEpoch){liveRunning=false;render()}
}
function liveStatus(state){
 const parts=['worklogs','leave'].map(kind=>{const item=state[kind],label=kind==='worklogs'?'Work':'Leave';
  if(item.status==='error')return `${label}: update failed · ${item.values?'last live '+new Date(item.checkedAt).toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'}):'snapshot'}`;
  if(item.status==='live')return `${label}: live ${new Date(item.checkedAt).toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'})}`;
  return `${label}: updating${item.values?' · previous live '+new Date(item.checkedAt).toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'}):' · snapshot'}`;
 });return parts.join(' | ');
}
liveRefreshButton.addEventListener('click',()=>{
 liveEpoch++;liveAbort?.abort();liveAbort=null;liveQueue=[];liveRunning=false;
 for(const state of liveStates.values())for(const kind of ['worklogs','leave'])state[kind].status='new';
 render();
});
setInterval(()=>{if(!document.hidden&&liveStates.size)liveRefreshButton.click()},300000);
function render(){
 if(!month.value)return;
 if(liveMonth!==month.value)resetLiveUpdates();
 updateScopeExplanation();updateFilterSummary();
 const range=days(), calendar=capacityCalendar(activeProfile(),range), todayDate=new Date();
 const today=`${todayDate.getFullYear()}-${String(todayDate.getMonth()+1).padStart(2,'0')}-${String(todayDate.getDate()).padStart(2,'0')}`;
 const names=new Set([...Object.keys(resources),...issues.map(x=>x.assignee),...allWorklogs.map(x=>x.worklog_author||x.issue_assignee),...leaves.map(x=>x.assignee)].map(key));
 const rows=[];
 for(const nk of names){
  if(!nk||!teamAllows(nk)||!nk.includes(key(employeeSearch.value)))continue;
  const resourceKey=Object.keys(resources).find(x=>key(x)===nk);
  if(!resigned.checked&&resources[resourceKey]?.resigned)continue;
  const label=resourceKey||issues.find(x=>key(x.assignee)===nk)?.assignee||allWorklogs.find(x=>key(x.worklog_author||x.issue_assignee)===nk)?.worklog_author||allWorklogs.find(x=>key(x.issue_assignee)===nk)?.issue_assignee||leaves.find(x=>key(x.assignee)===nk)?.assignee||'';
  const live=liveState(nk);
  const sourceLogs=live.worklogs.values?live.worklogs.values.map(x=>({issue_id:x.issue_key,worklog_author:label,worklog_date:x.date,hours_logged:x.seconds/3600,item_issue_type:x.issue_type,item_assignee:x.assigned_to_author?label:'',project_key:x.issue_key.split('-')[0]})).filter(x=>includeLeaves.checked||!isLeaveWorklog(x)):logs;
  const sourceLeaves=live.leave.values||leaves;
  const dayValues=calendar.map(day=>{
   const work=sourceLogs.filter(x=>key(x.worklog_author||x.issue_assignee)===nk&&txt(x.worklog_date)===day.date&&(scope.value==='any'||(key(x.item_assignee)===nk&&/sub-?task/i.test(txt(x.item_issue_type))))) .reduce((sum,x)=>sum+n(x.hours_logged),0);
   const leave=sourceLeaves.filter(x=>key(x.assignee)===nk&&txt(x.period_day)===day.date).reduce((sum,x)=>sum+n(x.planned_taken_hours)+n(x.unplanned_taken_hours),0);
   return {work,leave};
  });
  rows.push({label,live,dayValues,total:dayValues.reduce((s,v)=>s+v.work,0)});
 }
 rows.sort((a,b)=>a.label.localeCompare(b.label));
 const fmt=value=>value.toFixed(2);
 matrixExportRows=[['Employee',...calendar.flatMap(d=>[d.date+' logged hours',d.date+' RLT leave hours']),'Total logged hours','Live update status'],...rows.map(row=>[row.label,...row.dayValues.flatMap(v=>[v.work,v.leave]),row.total,liveStatus(row.live)])];
 const header=calendar.map(d=>`<th class="${d.type==='Weekend'?'weekend':''} ${d.date===today?'today':''}" title="${esc(d.type)}${d.date===today?' · Today':''}">${esc(d.date.slice(8))}<small>${d.date===today?'Today · ':''}${esc(d.type)}</small></th>`).join('');
 const htmlRows=rows.map(row=>`<tr><td class="employee">${esc(row.label)}<small class="matrix-live-status" aria-label="Live update status">${esc(liveStatus(row.live))}</small>${supportToggle.checked&&support.has(key(row.label))?'<span class="support">Support</span>':''}</td>${row.dayValues.map((v,i)=>`<td data-date="${esc(calendar[i].date)}" class="matrix-day ${v.work?'has-hours':'no-hours'} ${calendar[i].type==='Weekend'?'weekend':''} ${calendar[i].date===today?'today':''}" title="${esc(calendar[i].date)}: ${fmt(v.work)} logged hours; ${fmt(v.leave)} RLT leave hours"><span>${fmt(v.work)}h</span>${v.leave?`<small>Leave ${fmt(v.leave)}h</small>`:''}</td>`).join('')}<td class="num">${fmt(row.total)}h</td></tr>`).join('');
 table.innerHTML=rows.length?`<table class="matrix"><thead><tr><th>Employee</th>${header}<th>Total</th></tr></thead><tbody>${htmlRows}</tbody></table>`:'<div class="empty">No employees match the selected filters.</div>';
 // Every cell opens the same employee availability drawer as the capacity report.
 table.querySelectorAll('tbody tr').forEach((row,index)=>{
  const label=rows[index].label;
  row.querySelectorAll('td').forEach(cell=>{
   cell.tabIndex=0;cell.setAttribute('role','button');
   const date=cell.dataset.date;
   cell.setAttribute('aria-label',date?`Open log work for ${label} on ${date}`:`Open availability for ${label}`);
   cell.title=(cell.title?cell.title+' · ':'')+(date?'Click to log work for '+label+' on '+date:'Click to open availability for '+label);
   const open=()=>{window.openEcuDrawer(label,'availability',cell);if(date)openCalendarWorklog(label,date)};
   cell.addEventListener('click',open);
   cell.addEventListener('keydown',event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();open()}});
  });
 });
 queueLiveUpdates(rows);
 statusEl.textContent=`${rows.length} employees · ${loggedHoursScope().label.toLowerCase()} · filters apply automatically`;
 statusEl.className='status';snapshot.textContent=data.generated_at?`Canonical snapshot ${data.canonical_run_id||''} · ${new Date(data.generated_at).toLocaleString()}`:'';
}
async function downloadMatrix(){
 try{const response=await fetch('/api/team-worklog-hours/export',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({rows:matrixExportRows})});if(!response.ok)throw new Error('Could not export matrix');const url=URL.createObjectURL(await response.blob()),a=document.createElement('a');a.href=url;a.download=`team_worklog_hours_${month.value}.xlsx`;a.click();URL.revokeObjectURL(url)}catch(error){alert(error.message)}
}
// Keep report filters shareable without adding a browser-history entry per keystroke.
queueMicrotask(()=>{
 const restore=window.eprRestoreModalRoute;
 window.eprRestoreModalRoute=()=>{
  const params=new URLSearchParams(location.hash.slice(1));
  employeeSearch.value=params.get('employee_search')||'';
  if(!params.has('profile'))profile.value='';
  return restore?.();
 };
 function syncFilters(){
  window.eprUpdateModalUrl?.({month:month.value,profile:profile.value,scope:scope.value,
   leaves:includeLeaves.checked?'1':'0',resigned:resigned.checked?'1':'0',support:supportToggle.checked?'1':'0',
   teams:JSON.stringify([...team.querySelectorAll('input[type=checkbox]:checked')].map(input=>input.value)),
   employee_search:employeeSearch.value},true);
 }
 for(const input of [month,profile,scope,includeLeaves,resigned,supportToggle])input.addEventListener('change',syncFilters);
 team.addEventListener('change',syncFilters);
 team.addEventListener('click',event=>{if(event.target?.dataset?.teamAction)syncFilters()});
 employeeSearch.addEventListener('input',syncFilters);
});
