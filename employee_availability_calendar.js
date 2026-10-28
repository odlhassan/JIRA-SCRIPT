// Pure daily model: subtask worklogs authored by the selected employee(s),
// independent of the summary's logged-hours scope and leave inclusion switch.
function availabilityDays(label){
 const range=days(),keys=detailNameKeys(label),itemMap=new Map(issues.map(x=>[txt(x.issue_key).toUpperCase(),x]));
 const calendar=capacityCalendar(activeProfile(),range).map(day=>({...day,work:0,leave:0,leaveEntries:[],subtasks:new Map()}));
 const byDate=new Map(calendar.map(day=>[day.date,day]));
 for(const row of allWorklogs){
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
  return `<div class="availability-day ${day.type==='Weekend'?'weekend':''} ${isToday?'today':''}" ${isToday?'aria-current="date"':''} data-calendar-date="${day.date}"><div class="availability-date"><span>${Number(day.date.slice(-2))}</span><small>${esc(dayLabel)}</small></div>${day.work>0?`<button type="button" class="availability-hours" data-work-date="${day.date}" aria-label="${day.date}: ${availabilityHours(day.work)} logged; show subtasks">${availabilityIcon('task_alt')}${availabilityHours(day.work)} logged</button>`:`<span class="availability-hours">${availabilityIcon('schedule')}0h logged</span>`}${day.leave>0?`<button type="button" class="availability-hours leave" data-leave-date="${day.date}" aria-label="${day.date}: ${availabilityHours(day.leave)} leave; show leave entries">${availabilityIcon('beach_access')}${availabilityHours(day.leave)} leave</button>`:''}</div>`;
 }).join('');
 return {title:'Availability Details',subtitle:`${label} · ${month.value} · Daily subtask work & leave`,html:`<div class="availability-intro"><div><h3>${esc(first.toLocaleDateString(undefined,{month:'long',year:'numeric'}))}</h3><p>${esc(label)} · A little clarity for every day</p></div>${availabilityIcon('calendar_month')}</div><div class="availability-stats"><div class="availability-stat">Available capacity<strong>${availabilityHours(Math.max(0,totals.capacity-totals.leave))}</strong></div><div class="availability-stat work">Subtask hours logged<strong>${availabilityHours(totals.work)}</strong></div><div class="availability-stat leave">Leave taken<strong>${availabilityHours(totals.leave)}</strong></div></div><div class="availability-calendar-scroll"><div class="availability-weekdays">${['Mon','Tue','Wed','Thu','Fri','Sat','Sun'].map(name=>`<span>${name}</span>`).join('')}</div><div class="availability-grid" role="group" aria-label="Monthly work and leave calendar">${cards}</div></div><div class="availability-legend"><span><i></i>No subtask time logged</span><span><i class="work"></i>Subtask work · click hours for details</span><span><i class="leave"></i>Leave taken · click hours for details</span></div><p class="color-help">All subtask time authored by this employee is shown, including work on subtasks assigned to someone else. Leave stays separate, even when work and leave occur on the same day.</p><section id="availability-day-details" class="availability-details" aria-live="polite"><h3>${availabilityIcon('touch_app')} Explore a day</h3><p class="color-help">Choose green hours for subtask details or red hours for leave entries.</p></section>`};
}
function wireAvailabilityCalendar(label){
 const body=document.querySelector('#ecu-drawer-body'),calendar=availabilityDays(label);
 for(const button of body.querySelectorAll('[data-work-date],[data-leave-date]'))button.addEventListener('click',()=>{
  const isLeave=Boolean(button.dataset.leaveDate),day=calendar.find(x=>x.date===(button.dataset.leaveDate||button.dataset.workDate)),target=body.querySelector('#availability-day-details');
  for(const cell of body.querySelectorAll('.availability-day'))cell.classList.toggle('selected',cell.dataset.calendarDate===day.date);
  for(const badge of body.querySelectorAll('[data-work-date],[data-leave-date]'))badge.setAttribute('aria-pressed',String(badge===button));
  if(isLeave){
   const linkedItems=row=>{
    const ids=txt(row.jira_task_ids).split(',').map(x=>x.trim()).filter(Boolean),urls=txt(row.jira_task_links).split(',').map(x=>x.trim()).filter(Boolean);
    return htmlValue((ids.length?ids:urls).map((id,index)=>detailCell(urls[index]?externalLink(urls[index],ids[index]||'Open leave'):issueLink(id))).join('<br>'));
   };
   target.innerHTML=`<h3>${availabilityIcon('beach_access')} ${esc(day.date)} · ${availabilityHours(day.leave)} leave taken</h3>`+detailTable(['Employee','Date','Leave items / Jira','Planned taken','Unplanned taken','Total taken'],day.leaveEntries.map(row=>[row.assignee,row.period_day,linkedItems(row),availabilityHours(n(row.planned_taken_hours)),availabilityHours(n(row.unplanned_taken_hours)),availabilityHours(n(row.planned_taken_hours)+n(row.unplanned_taken_hours))]));
  }else target.innerHTML=`<h3>${availabilityIcon('assignment')} ${esc(new Date(day.date+'T12:00:00').toLocaleDateString(undefined,{weekday:'long',day:'numeric',month:'long'}))} · ${availabilityHours(day.work)} logged</h3>`+detailTable(['Subtask / Jira','Title','Worklog author','Start','Due','Original estimate','Hours logged'],[...day.subtasks.values()].map(row=>[issueLink(row.issueKey),row.title,row.author,row.start,row.due,availabilityHours(row.estimate),availabilityHours(row.hours)]));
  target.scrollIntoView({behavior:window.matchMedia('(prefers-reduced-motion: reduce)').matches?'auto':'smooth',block:'nearest'});
 });
}
