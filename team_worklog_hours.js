/* The matrix shares the capacity report's filters and canonical data. */
let matrixExportRows=[];
function render(){
 if(!month.value)return;
 updateScopeExplanation();updateFilterSummary();
 const range=days(), calendar=capacityCalendar(activeProfile(),range);
 const names=new Set([...Object.keys(resources),...issues.map(x=>x.assignee),...allWorklogs.map(x=>x.worklog_author||x.issue_assignee),...leaves.map(x=>x.assignee)].map(key));
 const rows=[];
 for(const nk of names){
  if(!nk||!teamAllows(nk)||!nk.includes(key(employeeSearch.value)))continue;
  const resourceKey=Object.keys(resources).find(x=>key(x)===nk);
  if(!resigned.checked&&resources[resourceKey]?.resigned)continue;
  const label=resourceKey||issues.find(x=>key(x.assignee)===nk)?.assignee||allWorklogs.find(x=>key(x.worklog_author||x.issue_assignee)===nk)?.worklog_author||allWorklogs.find(x=>key(x.issue_assignee)===nk)?.issue_assignee||leaves.find(x=>key(x.assignee)===nk)?.assignee||'';
  const dayValues=calendar.map(day=>{
   const work=logs.filter(x=>key(x.worklog_author||x.issue_assignee)===nk&&txt(x.worklog_date)===day.date&&(scope.value==='any'||(key(x.item_assignee)===nk&&/sub-?task/i.test(txt(x.item_issue_type))))) .reduce((sum,x)=>sum+n(x.hours_logged),0);
   const leave=leaves.filter(x=>key(x.assignee)===nk&&txt(x.period_day)===day.date).reduce((sum,x)=>sum+n(x.planned_taken_hours)+n(x.unplanned_taken_hours),0);
   return {work,leave};
  });
  rows.push({label,dayValues,total:dayValues.reduce((s,v)=>s+v.work,0)});
 }
 rows.sort((a,b)=>a.label.localeCompare(b.label));
 const fmt=value=>value.toFixed(2);
 matrixExportRows=[['Employee',...calendar.flatMap(d=>[d.date+' logged hours',d.date+' RLT leave hours']),'Total logged hours'],...rows.map(row=>[row.label,...row.dayValues.flatMap(v=>[v.work,v.leave]),row.total])];
 const header=calendar.map(d=>`<th title="${esc(d.type)}">${esc(d.date.slice(8))}<small>${esc(d.type)}</small></th>`).join('');
 const htmlRows=rows.map(row=>`<tr><td class="employee">${esc(row.label)}${supportToggle.checked&&support.has(key(row.label))?'<span class="support">Support</span>':''}</td>${row.dayValues.map((v,i)=>`<td class="matrix-day ${v.work?'has-hours':'no-hours'}" title="${esc(calendar[i].date)}: ${fmt(v.work)} logged hours; ${fmt(v.leave)} RLT leave hours"><span>${fmt(v.work)}h</span>${v.leave?`<small>Leave ${fmt(v.leave)}h</small>`:''}</td>`).join('')}<td class="num">${fmt(row.total)}h</td></tr>`).join('');
 table.innerHTML=rows.length?`<table class="matrix"><thead><tr><th>Employee</th>${header}<th>Total</th></tr></thead><tbody>${htmlRows}</tbody></table>`:'<div class="empty">No employees match the selected filters.</div>';
 statusEl.textContent=`${rows.length} employees · ${loggedHoursScope().label.toLowerCase()} · filters apply automatically`;
 statusEl.className='status';snapshot.textContent=data.generated_at?`Canonical snapshot ${data.canonical_run_id||''} · ${new Date(data.generated_at).toLocaleString()}`:'';
}
async function downloadMatrix(){
 try{const response=await fetch('/api/team-worklog-hours/export',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({rows:matrixExportRows})});if(!response.ok)throw new Error('Could not export matrix');const url=URL.createObjectURL(await response.blob()),a=document.createElement('a');a.href=url;a.download=`team_worklog_hours_${month.value}.xlsx`;a.click();URL.revokeObjectURL(url)}catch(error){alert(error.message)}
}
