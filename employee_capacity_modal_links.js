/* Shareable view state. URLs never contain descriptions, review tokens or credentials. */
(() => {
 let restoring=false,ready=false,lastRestored='';
 const read=()=>new URLSearchParams(location.hash.slice(1));
 window.eprUpdateModalUrl=(patch,replace=false)=>{
  if(restoring)return;
  const params=read();
  for(const [name,value]of Object.entries(patch)){if(value==null||value==='')params.delete(name);else params.set(name,String(value))}
  const url=new URL(location.href);
  if(url.searchParams.get('availability')==='1'){url.searchParams.delete('availability');url.searchParams.delete('employee')}
  url.hash=params.toString();
  if(url.href!==location.href)history[replace?'replaceState':'pushState'](null,'',url);
  lastRestored=location.href;
 };
 window.eprSetDetailRoute=(label,metric)=>{
  window.eprUpdateModalUrl({modal:metric,employee:label,month:month.value,scope:scope.value,profile:profile.value,
   leaves:includeLeaves.checked?'1':'0',resigned:resigned.checked?'1':'0',support:supportToggle.checked?'1':'0',
   teams:JSON.stringify([...team.querySelectorAll('input:checked')].map(input=>input.value)),
   day:null,kind:null,panel:null,issue:null,date:null,historyMonth:null});
  const head=document.querySelector('.ecu-drawer-head');
  if(!head.querySelector('[data-copy-modal]')){
   const button=document.createElement('button');button.type='button';button.className='ecu-copy-modal';button.dataset.copyModal='1';button.textContent='Copy link';button.setAttribute('aria-label','Copy link to details');
   head.insertBefore(button,document.querySelector('#ecu-drawer-close'));
   button.addEventListener('click',async()=>{try{await navigator.clipboard.writeText(location.href);button.textContent='Link copied'}catch(_error){window.prompt('Copy this details URL',location.href)}});
  }
 };
 window.eprClearDetailRoute=()=>window.eprUpdateModalUrl({modal:null,employee:null,day:null,kind:null,panel:null,issue:null,date:null,historyMonth:null});
 window.eprRestoreModalRoute=()=>{
  ready=true;lastRestored=location.href;const params=read(),metric=params.get('modal');restoring=true;
  try{
   window.eprCloseWorklogManager?.();
   const period=params.get('month');if(/^\d{4}-(0[1-9]|1[0-2])$/.test(period||''))month.value=period;
   if(['any','assigned'].includes(params.get('scope')))scope.value=params.get('scope');
   if(params.has('profile')&&[...profile.options].some(option=>option.value===params.get('profile')))profile.value=params.get('profile');
   for(const [name,input]of [['leaves',includeLeaves],['resigned',resigned],['support',supportToggle]])if(params.has(name))input.checked=params.get(name)==='1';
   if(params.has('teams')){try{const selected=JSON.parse(params.get('teams'));if(Array.isArray(selected))for(const input of team.querySelectorAll('input[type=checkbox]'))input.checked=selected.includes(input.value)}catch(_error){/* Ignore an invalid filter, not the view. */}}
   logs=includeLeaves.checked?allWorklogs:allWorklogs.filter(row=>!isLeaveWorklog(row));data.worklogs=logs;updateTeamSummary();render();
   if(DETAIL_METRICS.includes(metric)&&params.get('employee')){
    window.openEcuDrawer(params.get('employee'),metric,month);
    const day=params.get('day');
    if(metric==='availability'&&/^\d{4}-\d{2}-\d{2}$/.test(day||'')){
     const selector=params.get('kind')==='leave'?'data-leave-date':'data-work-date';
     document.querySelector('['+selector+'="'+day+'"]')?.click();
    }
    if(metric==='availability'&&params.get('panel')==='log'&&/^\d{4}-\d{2}-\d{2}$/.test(day||''))openCalendarWorklog(params.get('employee'),day);
    if(metric==='availability'&&params.get('panel')==='leave'&&/^\d{4}-\d{2}-\d{2}$/.test(day||'')&&/^RLT-\d+$/.test(params.get('issue')||''))openCalendarWorklog(params.get('employee'),day,{mode:'leave',issue:params.get('issue')});
   }else if(!metric){document.querySelector('#ecu-drawer-close')?.click()}
   if(params.get('panel')==='hours'){
    const trigger=document.createElement('button');trigger.hidden=true;trigger.dataset.manageWorklogs=params.get('issue')||'';trigger.dataset.worklogDate=params.get('date')||'';trigger.dataset.worklogMonth=params.get('historyMonth')||month.value;document.body.append(trigger);trigger.click();trigger.remove();
   }
  }finally{restoring=false}
 };
 window.addEventListener('popstate',()=>{if(ready&&lastRestored!==location.href)window.eprRestoreModalRoute()});
 window.addEventListener('hashchange',()=>{if(ready&&lastRestored!==location.href)window.eprRestoreModalRoute()});
})();
