// Pure daily model: subtask worklogs authored by the selected employee(s),
// independent of the summary's logged-hours scope and leave inclusion switch.
let availabilityLive=null, availabilityTimer=null, availabilityGeneration=0, availabilityEvents=null;
function availabilityDays(label){
 const range=days(),keys=detailNameKeys(label),itemMap=new Map(issues.map(x=>[txt(x.issue_key).toUpperCase(),x]));
 const calendar=capacityCalendar(activeProfile(),range).map(day=>({...day,work:0,leave:0,leaveEntries:[],subtasks:new Map()}));
 const byDate=new Map(calendar.map(day=>[day.date,day]));
 const source=availabilityLive && availabilityLive.label===label && availabilityLive.month===month.value ? availabilityLive.rows : allWorklogs;
 for(const row of source){
  const day=byDate.get(txt(row.worklog_date).slice(0,10)),issueKey=txt(row.issue_id||row.issue_key).toUpperCase(),item=itemMap.get(issueKey)||{};
  if(!day||!keys.has(key(row.worklog_author||row.issue_assignee))||isLeaveWorklog(row)||!/sub[ -]?task/i.test(txt(row.item_issue_type||item.issue_type)))continue;
  const hours=n(row.hours_logged);day.work+=hours;
  const groupKey=key(row.worklog_author||row.issue_assignee)+'|'+issueKey;
  if(!day.subtasks.has(groupKey))day.subtasks.set(groupKey,{issueKey,title:row.item_summary||item.summary||'',author:row.worklog_author||row.issue_assignee,start:row.item_start_date||item.start_date||'—',due:row.item_due_date||item.due_date||'—',estimate:n(row.original_estimate_hours??item.original_estimate_hours),hours:0});
  day.subtasks.get(groupKey).hours+=hours;
 }
 for(const row of leaves){const day=byDate.get(txt(row.period_day).slice(0,10));if(day&&keys.has(key(row.assignee))){day.leave+=n(row.planned_taken_hours)+n(row.unplanned_taken_hours);if(n(row.planned_taken_hours)+n(row.unplanned_taken_hours)>0)day.leaveEntries.push(row)}}
 return calendar;
}
function availabilityHours(value){return `${Number(value.toFixed(2))}h`}
function availabilityIcon(name){return `<span class="material-symbols-outlined" aria-hidden="true">${name}</span>`}
function availabilityDetails(label){
 const current=new Date(),today=[current.getFullYear(),String(current.getMonth()+1).padStart(2,'0'),String(current.getDate()).padStart(2,'0')].join('-');
 const calendar=availabilityDays(label),first=new Date(calendar[0].date+'T12:00:00'),offset=(first.getDay()+6)%7,slots=Math.ceil((offset+calendar.length)/7)*7;
 const totals=calendar.reduce((sum,day)=>({work:sum.work+day.work,leave:sum.leave+day.leave,capacity:sum.capacity+day.hours*detailNames(label).length}),{work:0,leave:0,capacity:0});
 const cards=Array.from({length:slots},(_,index)=>{
  const day=calendar[index-offset];if(!day)return '<div class="availability-day outside" aria-hidden="true"></div>';
  const isToday=day.date===today,dayLabel=[isToday?'Today':'',day.type==='Workday'?'':day.type].filter(Boolean).join(' · ');
  return `<div class="availability-day ${day.type==='Weekend'?'weekend':''} ${isToday?'today':''}" ${isToday?'aria-current="date"':''} data-calendar-date="${day.date}"><div class="availability-date"><span>${Number(day.date.slice(-2))}</span><small>${esc(dayLabel)}</small><button type="button" class="availability-add" data-add-date="${day.date}" aria-label="Log work on ${day.date}">+</button></div>${day.work>0?`<button type="button" class="availability-hours" data-work-date="${day.date}" aria-label="${day.date}: ${availabilityHours(day.work)} logged; show subtasks">${availabilityIcon('task_alt')}${availabilityHours(day.work)} logged</button>`:`<span class="availability-hours">${availabilityIcon('schedule')}0h logged</span>`}${day.leave>0?`<button type="button" class="availability-hours leave" data-leave-date="${day.date}" aria-label="${day.date}: ${availabilityHours(day.leave)} leave; show leave entries">${availabilityIcon('beach_access')}${availabilityHours(day.leave)} leave</button>`:''}</div>`;
 }).join('');
 return {title:'Availability Details',subtitle:`${label} · ${month.value} · Daily subtask work & leave`,html:`<div class="availability-toolbar"><div class="availability-month-controls"><button type="button" class="availability-month-step" data-month-step="-1" aria-label="Previous month" title="Previous month">${availabilityIcon('chevron_left')}</button><h3 aria-live="polite">${esc(first.toLocaleDateString(undefined,{month:'long',year:'numeric'}))}</h3><button type="button" class="availability-month-step" data-month-step="1" aria-label="Next month" title="Next month">${availabilityIcon('chevron_right')}</button></div><div class="availability-actions">${label==='Grand Total'?'':'<button type="button" class="availability-tally" data-worklog-tally>Download Jira tally PDF</button>'}<button type="button" class="worklog-primary availability-manage" data-manage-worklogs="">Manage my logged hours</button></div></div><p class="availability-tally-status" role="status" aria-live="polite"></p><div class="availability-intro"><div><p>${esc(label)} · A little clarity for every day</p></div>${availabilityIcon('calendar_month')}</div><div class="availability-stats"><div class="availability-stat">Available capacity<strong>${availabilityHours(Math.max(0,totals.capacity-totals.leave))}</strong></div><div class="availability-stat work">Subtask hours logged<strong>${availabilityHours(totals.work)}</strong></div><div class="availability-stat leave">Leave taken<strong>${availabilityHours(totals.leave)}</strong></div></div><div class="availability-calendar-scroll"><div class="availability-weekdays">${['Mon','Tue','Wed','Thu','Fri','Sat','Sun'].map(name=>`<span>${name}</span>`).join('')}</div><div class="availability-grid" role="group" aria-label="Monthly work and leave calendar">${cards}</div></div><div class="availability-legend"><span><i></i>No subtask time logged</span><span><i class="work"></i>Subtask work · click hours for details</span><span><i class="leave"></i>Leave taken · click hours for details</span></div><p class="color-help">All subtask time authored by this employee is shown, including work on subtasks assigned to someone else. Leave stays separate, even when work and leave occur on the same day.</p><section id="availability-day-details" class="availability-details" aria-live="polite"><h3>${availabilityIcon('touch_app')} Explore a day</h3><p class="color-help">Choose green hours for subtask details or red hours for leave entries.</p></section>`};
}
function wireAvailabilityCalendar(label){
 const body=document.querySelector('#ecu-drawer-body');let calendar=availabilityDays(label);
 clearInterval(availabilityTimer);
 availabilityEvents?.abort();availabilityEvents=new AbortController();
 const generation=++availabilityGeneration, period=month.value;
 const status=document.createElement('p');status.className='availability-live-status';status.setAttribute('role','status');
 const refresh=document.createElement('button');refresh.type='button';refresh.className='availability-refresh';refresh.innerHTML=availabilityIcon('sync')+'<span>Refresh Jira worklogs</span>';
 body.querySelector('.availability-toolbar').after(status);body.querySelector('.availability-actions').prepend(refresh);
 let running=false;
 async function refreshLive(){
  if(running||!status.isConnected||document.hidden||!document.querySelector('#ecu-drawer')?.classList.contains('open')||body.querySelector('#calendar-worklog-form')||document.querySelector('.wl-manager'))return;
  running=true;refresh.disabled=true;status.textContent='Reading Jira worklogs…';
  try{
   let cursor='',rows=[],author='',seen=new Set();
   do{
    const params=new URLSearchParams({month:period});if(cursor)params.set('cursor',cursor);
    const response=await fetch('/api/employee-capacity-utilization/worklog/history?'+params,{cache:'no-store'});
    const result=await response.json();if(!response.ok)throw new Error(result.error||'Jira unavailable');
    if(!Array.isArray(result.entries)||!result.author)throw new Error('Incomplete Jira response');
    if(key(result.author)!==key(label)){status.textContent='Showing saved report data for this employee.';return}
    if(author && author!==result.author)throw new Error('Jira account changed');author=result.author;
    for(const entry of result.entries){
     if(!entry.issue_type)throw new Error('Jira issue type is unavailable');
     rows.push({issue_id:entry.issue_key,item_issue_type:entry.issue_type,item_summary:entry.summary,item_assignee:entry.assignee,worklog_author:author,worklog_date:entry.started.slice(0,10),hours_logged:Number(entry.seconds)/3600,worklog_id:entry.worklog_id});
    }
    cursor=result.next_cursor||'';if(cursor && seen.has(cursor))throw new Error('Jira pagination failed');if(cursor)seen.add(cursor);
   }while(cursor);
   if(generation!==availabilityGeneration||!status.isConnected||period!==month.value||!document.querySelector('#ecu-drawer')?.classList.contains('open'))return;
   const selected=body.querySelector('.availability-day.selected')?.dataset.calendarDate;
   availabilityLive={label,month:period,rows};
   calendar=availabilityDays(label);
   const temp=document.createElement('div');temp.innerHTML=availabilityDetails(label).html;
   body.querySelector('.availability-grid').innerHTML=temp.querySelector('.availability-grid').innerHTML;
   body.querySelector('.availability-stat.work strong').textContent=temp.querySelector('.availability-stat.work strong').textContent;
   wireDays();
   if(selected){const chosen=body.querySelector(`[data-work-date="${selected}"]`);if(chosen)chosen.click();else{const detail=body.querySelector('#availability-day-details');if(detail)detail.textContent='Jira worklogs updated. Select a day to see its details.'}}
   status.textContent='Jira worklogs · Updated at '+new Date().toLocaleTimeString()+'. Capacity and leave use the saved report.';
  }catch(error){if(generation===availabilityGeneration)status.textContent='Jira worklogs could not be refreshed. Showing last available data. '+error.message}
  finally{running=false;if(refresh.isConnected)refresh.disabled=false}
 }
 refresh.addEventListener('click',refreshLive);
 availabilityTimer=setInterval(()=>{if(!status.isConnected||!document.querySelector('#ecu-drawer')?.classList.contains('open')){clearInterval(availabilityTimer);availabilityEvents?.abort();availabilityGeneration++;return}refreshLive()},60000);
 document.addEventListener('visibilitychange',()=>{if(generation===availabilityGeneration&&!document.hidden)refreshLive()},{signal:availabilityEvents.signal});
 document.addEventListener('epr-worklog-changed',()=>refreshLive(),{signal:availabilityEvents.signal});
 queueMicrotask(refreshLive);
 body.querySelector('[data-worklog-tally]')?.addEventListener('click',async event=>{
  const button=event.currentTarget,period=month.value,status=body.querySelector('.availability-tally-status');
  button.disabled=true;status.textContent='Reading live Jira and the saved EPR snapshot for '+label+' · '+period+'…';
  try{
   const params=new URLSearchParams({employee:label,month:period});
   const response=await fetch('/api/employee-capacity-utilization/tally.pdf?'+params,{cache:'no-store'});
   if(!response.ok){let error;try{error=await response.json()}catch(_error){}throw new Error(error?.error||'The tally could not be completed. No PDF was generated.')}
   if(!response.headers.get('Content-Type')?.includes('application/pdf'))throw new Error('The server did not return a PDF. Sign in again and retry.');
   const url=URL.createObjectURL(await response.blob()),link=document.createElement('a');
   link.href=url;link.download='worklog-tally-'+label.replace(/[^A-Za-z0-9_-]+/g,'-')+'-'+period+'.pdf';link.click();setTimeout(()=>URL.revokeObjectURL(url),60000);
   status.textContent='Tally downloaded for '+label+' · '+period+'. It compares saved Jira entries with the current EPR snapshot; no hours were changed.';
  }catch(error){status.textContent=error.message}
  finally{button.disabled=false}
 });
 for(const button of body.querySelectorAll('[data-month-step]'))button.addEventListener('click',()=>{
  const direction=button.dataset.monthStep,next=new Date(month.value+'-01T12:00:00');
  next.setMonth(next.getMonth()+Number(direction));
  if(next.getFullYear()<1||next.getFullYear()>9999)return;
  month.value=String(next.getFullYear()).padStart(4,'0')+'-'+String(next.getMonth()+1).padStart(2,'0');
  render();drawerPreviousFocus=month;
  window.eprUpdateModalUrl?.({month:month.value,day:null,kind:null,panel:null,issue:null,date:null});
  const detail=availabilityDetails(label);
  document.querySelector('#ecu-drawer-subtitle').textContent=detail.subtitle;
  body.innerHTML=detail.html;wireAvailabilityCalendar(label);body.scrollTop=0;
  body.querySelector('[data-month-step="'+direction+'"]').focus();
 });
 function wireDays(){
 for(const button of body.querySelectorAll('[data-add-date]'))button.addEventListener('click',()=>openCalendarWorklog(label,button.dataset.addDate));
 for(const button of body.querySelectorAll('[data-work-date],[data-leave-date]'))button.addEventListener('click',()=>{
  const isLeave=Boolean(button.dataset.leaveDate),day=calendar.find(x=>x.date===(button.dataset.leaveDate||button.dataset.workDate)),target=body.querySelector('#availability-day-details');
  window.eprUpdateModalUrl?.({day:day.date,kind:isLeave?'leave':'work',panel:null,issue:null,date:null});
  for(const cell of body.querySelectorAll('.availability-day'))cell.classList.toggle('selected',cell.dataset.calendarDate===day.date);
  for(const badge of body.querySelectorAll('[data-work-date],[data-leave-date]'))badge.setAttribute('aria-pressed',String(badge===button));
  if(isLeave){
   const linkedItems=row=>{
    const ids=txt(row.jira_task_ids).split(',').map(x=>x.trim()).filter(Boolean),urls=txt(row.jira_task_links).split(',').map(x=>x.trim()).filter(Boolean);
    return htmlValue((ids.length?ids:urls).map((id,index)=>detailCell(urls[index]?externalLink(urls[index],ids[index]||'Open leave'):issueLink(id))).join('<br>'));
   };
   target.innerHTML=`<h3>${availabilityIcon('beach_access')} ${esc(day.date)} · ${availabilityHours(day.leave)} leave taken</h3>`+detailTable(['Employee','Date','Leave items / Jira','Planned taken','Unplanned taken','Total taken'],day.leaveEntries.map(row=>[row.assignee,row.period_day,linkedItems(row),availabilityHours(n(row.planned_taken_hours)),availabilityHours(n(row.unplanned_taken_hours)),availabilityHours(n(row.planned_taken_hours)+n(row.unplanned_taken_hours))]));
  }else target.innerHTML=`<h3>${availabilityIcon('assignment')} ${esc(new Date(day.date+'T12:00:00').toLocaleDateString(undefined,{weekday:'long',day:'numeric',month:'long'}))} · ${availabilityHours(day.work)} logged</h3>`+detailTable(['Subtask / Jira','Title','Worklog author','Start','Due','Original estimate','Hours logged','Actions'],[...day.subtasks.values()].map(row=>[issueLink(row.issueKey),row.title,row.author,row.start,row.due,availabilityHours(row.estimate),availabilityHours(row.hours),htmlValue(`<button type="button" class="availability-delete" data-manage-worklogs="${esc(row.issueKey)}" data-worklog-date="${day.date}" aria-label="Delete my entries on ${esc(row.issueKey)} for ${day.date}">Delete</button>`)]));
  target.scrollIntoView({behavior:window.matchMedia('(prefers-reduced-motion: reduce)').matches?'auto':'smooth',block:'nearest'});
 });
 }
 wireDays();
}
