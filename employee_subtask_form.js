/* Guided creation uses live Jira choices, personal OAuth and one-attempt receipts. */
const availabilityCreatedIssues=new Map();
function availabilitySubtasks(){return [...new Map([...availabilityCreatedIssues.values(),...issues].map(item=>[item.issue_key,item])).values()]}
async function openAvailabilitySubtask(label,selectedDate){
 if(document.querySelector('.subtask-dialog'))return;
 const previous=document.activeElement,dialog=document.createElement('dialog');
 dialog.className='subtask-dialog';dialog.setAttribute('aria-labelledby','subtask-title');
 dialog.innerHTML=`<header class="subtask-head"><div><span class="subtask-kicker">A LITTLE PLAN, A BIG STEP ✦</span><h2 id="subtask-title">Create a new subtask</h2><p>Pick its home. Plan your work. Make it happen.</p></div><button type="button" data-close aria-label="Close subtask creator">✕</button></header><div class="subtask-content"><p role="status" data-status>Connecting to your Jira account…</p><div data-compose hidden><p class="subtask-owner" data-owner></p><form id="availability-subtask-form"><section class="subtask-section subtask-home"><h3><span>1</span> Find its home</h3><p>Every subtask lives in a user story inside an epic.</p><div class="subtask-path">Space → Epic → User story → Your new subtask</div>${[['projects','Jira space'],['epics','Epic'],['stories','User story']].map(([kind,title])=>`<div class="subtask-picker" data-picker="${kind}"><span class="subtask-picker-label" id="subtask-${kind}-label">${title}</span><input name="${kind}" type="hidden"><button type="button" role="combobox" aria-labelledby="subtask-${kind}-label" aria-controls="subtask-${kind}-list" aria-expanded="false" aria-haspopup="listbox" aria-required="true" data-pick-toggle="${kind}" disabled><span data-picked="${kind}">${kind==='projects'?'Loading spaces…':kind==='epics'?'Choose a space first':'Choose an epic first'}</span><span aria-hidden="true">⌄</span></button><div class="subtask-menu" id="subtask-${kind}-menu" popover="manual"><input type="search" data-search="${kind}" aria-label="Search ${title.toLowerCase()}" placeholder="Search by title or key…" autocomplete="off"><div role="listbox" id="subtask-${kind}-list" aria-labelledby="subtask-${kind}-label" data-options="${kind}"></div><p data-menu-status="${kind}" role="status"></p><button type="button" data-more="${kind}" hidden>Load more</button></div><small data-hint="${kind}" role="status"></small></div>`).join('')}<p class="subtask-flow-note" data-flow-note hidden></p></section><section class="subtask-section subtask-what"><h3><span>2</span> Give it a purpose</h3><label><span data-title-label>Subtask title</span><input name="summary" maxlength="255" required placeholder="e.g. Build the customer search filter" autocomplete="off"></label><label>Description <small>(optional)</small><textarea name="description" maxlength="2000" rows="3" placeholder="What needs to be done?"></textarea></label></section><section class="subtask-section subtask-when"><h3><span>3</span> Make a little plan</h3><label data-leave-type-field hidden>Leave type<select name="leave_type" aria-label="Leave type"><option value="">Choose a leave type…</option></select></label><div class="subtask-dates"><label>Start date<input name="start_date" type="date" required></label><label>Due date<input name="due_date" type="date" required aria-describedby="subtask-date-help"></label></div><p id="subtask-date-help" class="subtask-help" data-date-help></p><label>Original estimate<input name="estimate" value="1d" maxlength="100" required placeholder="e.g. 1d or 1h 30m" aria-describedby="subtask-estimate-help"></label><div class="subtask-presets" aria-label="Quick original estimates">${['30m','4h','1d','1w'].map(value=>`<button type="button" data-estimate="${value}" aria-pressed="${value==='1d'}">${value}</button>`).join('')}</div><p id="subtask-estimate-help" class="subtask-help">A day or week follows your Jira working-time settings. This is planned effort; it does not log hours.</p></section><footer class="subtask-footer"><span>Assigned to you, automatically ✨</span><button type="submit" class="subtask-primary">Review subtask →</button></footer></form></div><div data-review hidden></div></div>`;
 document.body.append(dialog);dialog.showModal();
 const $=selector=>dialog.querySelector(selector),status=$('[data-status]'),form=$('form');
 let busy=false,sent=false,token=null,review=null,author='',bounds=null,alive=true;
 const kinds=['projects','epics','stories'],pickers=Object.fromEntries(kinds.map(kind=>[kind,{generation:0,query:'',options:[],cursor:null,loaded:false,loading:false,error:'',hint:'',selected:null}]));
 let openKind=null;
 const isLeave=()=>form.elements.projects.value==='RLT';
 const parentReady=kind=>kind==='projects'||Boolean(form.elements.projects.value&&(kind==='epics'||isLeave()||form.elements.epics.value));
 const closeMenus=(restore=false)=>{const prior=openKind;for(const kind of kinds){const menu=$('#subtask-'+kind+'-menu');if(menu.matches(':popover-open'))menu.hidePopover();$('[data-pick-toggle="'+kind+'"]').setAttribute('aria-expanded','false')}openKind=null;if(restore&&prior)$('[data-pick-toggle="'+prior+'"]').focus()};
 const close=()=>{if(busy)return;alive=false;closeMenus();for(const state of Object.values(pickers))clearTimeout(state.timer);window.removeEventListener('resize',positionMenu);dialog.close();dialog.remove();previous?.focus?.()};
 $('[data-close]').addEventListener('click',close);
 dialog.addEventListener('cancel',event=>{event.preventDefault();if(openKind)closeMenus(true);else close()});
 dialog.addEventListener('keydown',event=>{event.stopPropagation();if(event.key==='Escape'&&openKind){event.preventDefault();closeMenus(true)}});
 dialog.addEventListener('click',event=>{if(openKind&&!event.target.closest('[data-picker="'+openKind+'"]'))closeMenus()});
 function syncEnabled(){for(const kind of kinds)$('[data-pick-toggle="'+kind+'"]').disabled=busy||!bounds||!parentReady(kind)}
 const setBusy=value=>{busy=value;if(value)closeMenus();dialog.setAttribute('aria-busy',String(value));for(const control of dialog.querySelectorAll('button'))control.disabled=value;syncEnabled()};
 function positionMenu(){
  if(!openKind)return;const anchor=$('[data-pick-toggle="'+openKind+'"]'),menu=$('#subtask-'+openKind+'-menu'),rect=anchor.getBoundingClientRect();
  const below=innerHeight-rect.bottom-12,above=rect.top-12,up=below<240&&above>below,height=Math.min(360,up?above:below),width=Math.min(Math.max(300,rect.width),innerWidth-24);
  menu.style.width=width+'px';menu.style.maxHeight=Math.max(100,height)+'px';menu.style.left=Math.min(Math.max(12,rect.left),innerWidth-width-12)+'px';
  menu.style.top=up?'auto':(rect.bottom+6)+'px';menu.style.bottom=up?(innerHeight-rect.top+6)+'px':'auto';
 }
 window.addEventListener('resize',positionMenu);dialog.addEventListener('scroll',positionMenu);
 function renderPicker(kind){
  const state=pickers[kind],list=$('[data-options="'+kind+'"]'),hint=$('[data-hint="'+kind+'"]');
  list.innerHTML=state.options.map(item=>`<button type="button" role="option" tabindex="-1" aria-selected="${item.key===form.elements[kind].value}" data-option-key="${esc(item.key)}"><strong>${esc(item.key)}</strong><span>${esc(item.summary)}</span></button>`).join('');
  for(const option of list.querySelectorAll('[role=option]'))option.addEventListener('click',()=>choose(kind,state.options.find(item=>item.key===option.dataset.optionKey)));
  const message=state.loading?'Reading Jira…':state.error||state.hint||'Choose a result.';
  $('[data-menu-status="'+kind+'"]').textContent=message;hint.textContent=state.loading?'Reading Jira…':state.error||state.hint;
  const more=$('[data-more="'+kind+'"]');more.hidden=!state.cursor&&!state.error;more.disabled=busy||state.loading;more.textContent=state.error?'Retry':'Load more';
  syncEnabled();positionMenu();
 }
 function reset(kind,message){
  const state=pickers[kind];state.generation++;clearTimeout(state.timer);Object.assign(state,{query:'',options:[],cursor:null,loaded:false,loading:false,error:'',hint:'',selected:null});
  form.elements[kind].value='';$('[data-search="'+kind+'"]').value='';$('[data-picked="'+kind+'"]').textContent=message;renderPicker(kind);
 }
 function flowChanged(){
  const leave=isLeave();$('[data-picker="epics"]').hidden=leave;
  $('#subtask-stories-label').textContent=leave?'Leave tracker':'User story';
  $('[data-search="stories"]').setAttribute('aria-label',leave?'Search leave tracker':'Search user story');
  $('.subtask-path').textContent=leave?'RLT → Existing leave tracker → Your leave subtask':'Space → Epic → User story → Your new subtask';
  $('#subtask-title').textContent=leave?'Create a leave subtask':'Create a new subtask';
  $('[data-title-label]').textContent=leave?'Leave subtask title':'Subtask title';form.elements.summary.placeholder=leave?'e.g. Annual leave — 18 November':'e.g. Build the customer search filter';
  $('[data-leave-type-field]').hidden=!leave;form.elements.leave_type.required=leave;
  $('[data-flow-note]').hidden=!leave;$('[data-flow-note]').textContent='Leave has its own path: choose an existing RLT task or story. No epic is needed. Recorded leave stays separate from work hours.';
  $('.subtask-home>p').textContent=leave?'Choose the existing leave tracker for this leave.':'Every work subtask lives in a user story inside an epic.';
  syncEnabled();
 }
 function choose(kind,item){
  if(!item||busy)return;const changed=form.elements[kind].value!==item.key;form.elements[kind].value=item.key;pickers[kind].selected=item;
  $('[data-picked="'+kind+'"]').textContent=item.key+' · '+item.summary;$('[data-pick-toggle="'+kind+'"]').removeAttribute('aria-invalid');renderPicker(kind);closeMenus(true);
  if(!changed)return;
  if(kind==='projects'){
   reset('epics','Choose an epic…');reset('stories',isLeave()?'Choose a leave tracker…':'Choose an epic first');
   form.elements.leave_type.innerHTML='<option value="">Choose a leave type…</option>';flowChanged();load(isLeave()?'stories':'epics');
  }else if(kind==='epics'){reset('stories','Choose a user story…');load('stories')}
 }
 async function load(kind,more=false){
  if(!parentReady(kind))return;
  const state=pickers[kind],version=++state.generation,project=form.elements.projects.value,epic=form.elements.epics.value;
  const params=new URLSearchParams({kind:kind==='stories'&&isLeave()?'trackers':kind,q:state.query});
  if(kind!=='projects')params.set('project_key',project);if(kind==='stories'&&!isLeave())params.set('epic_key',epic);
  if(more&&state.cursor)params.set('cursor',state.cursor);if(!more){state.options=[];state.cursor=null}
  state.loading=true;state.error='';renderPicker(kind);
  try{
   const result=await worklogRequest('/subtask/options?'+params);
   if(!alive||state.generation!==version||kind!=='projects'&&(project!==form.elements.projects.value||epic!==form.elements.epics.value))return;
   if(kind==='projects'&&!bounds){
    author=result.author;bounds=result;$('[data-owner]').textContent='✓ Signed in as '+author+' · Your new subtask will be assigned to you.';
    const chosen=selectedDate>=result.first&&selectedDate<=result.last?selectedDate:result.today;
    form.elements.start_date.value=chosen;form.elements.due_date.value=chosen;form.elements.due_date.min=result.first;form.elements.due_date.max=result.last;form.elements.start_date.max=result.last;
    $('[data-date-help]').textContent='Due this month only: '+result.first+' to '+result.last+' (your Jira timezone).';$('[data-compose]').hidden=false;status.textContent='';$('[data-picked="projects"]').textContent='Choose a space…';
   }
   const merged=new Map(state.options.map(item=>[item.key,item]));for(const item of result.options)merged.set(item.key,item);state.options=[...merged.values()];state.cursor=result.next_cursor;state.loaded=true;
   state.hint=result.configuration_warning||result.empty_reason||(state.options.length?state.cursor?'More results available. Search here or load more.':'':state.cursor?'No matching items on this page. Load more.':state.query?'No matches in this dropdown. Clear its search to see all choices.':kind==='epics'?project+' has no visible epics. Choose another space or add an epic and user story in Jira.':kind==='stories'?isLeave()?'No leave trackers are visible in RLT.':'This epic has no visible user stories. Choose another epic.':'No spaces are visible to your Jira account.');
   if(isLeave()&&kind==='stories'&&Array.isArray(result.leave_types)){
    const selected=form.elements.leave_type.value;form.elements.leave_type.innerHTML='<option value="">Choose a leave type…</option>'+result.leave_types.map(item=>`<option value="${esc(item.id)}">${esc(item.label)}</option>`).join('');form.elements.leave_type.value=selected;
   }
  }catch(error){if(!alive||state.generation!==version)return;state.error=error.message;
   if(!bounds){status.textContent=error.message;try{const auth=await fetch('/api/jira-auth/status').then(response=>response.json());if(!alive)return;if(auth.configured===false){status.textContent='Personal Jira sign-in needs administrator setup on this server before you can create a subtask.';return}}catch(_error){}if(!alive)return;const link=document.createElement('a');link.href='/api/jira-auth/login';link.textContent=' Sign in with Jira';status.append(link)}
  }finally{if(alive&&state.generation===version){state.loading=false;renderPicker(kind)}}
 }
 function showPicker(kind){
  if(busy||!parentReady(kind))return;if(openKind===kind){closeMenus(true);return}closeMenus();openKind=kind;
  const state=pickers[kind],search=$('[data-search="'+kind+'"]');
  if(state.query){state.query='';search.value='';state.loaded=false}
  const menu=$('#subtask-'+kind+'-menu');menu.showPopover();$('[data-pick-toggle="'+kind+'"]').setAttribute('aria-expanded','true');positionMenu();search.focus();
  if(!state.loaded&&!state.loading)load(kind);
 }
 for(const kind of kinds){
  $('[data-pick-toggle="'+kind+'"]').addEventListener('click',()=>showPicker(kind));
  $('[data-pick-toggle="'+kind+'"]').addEventListener('keydown',event=>{if(['ArrowDown','ArrowUp'].includes(event.key)){event.preventDefault();showPicker(kind)}});
  $('[data-search="'+kind+'"]').addEventListener('input',event=>{const state=pickers[kind];state.generation++;clearTimeout(state.timer);state.query=event.target.value;state.loading=true;state.options=[];state.cursor=null;renderPicker(kind);state.timer=setTimeout(()=>load(kind),300)});
  $('[data-more="'+kind+'"]').addEventListener('click',()=>load(kind,Boolean(pickers[kind].cursor)));
  $('#subtask-'+kind+'-menu').addEventListener('keydown',event=>{
   if(event.key==='Enter'&&event.target.matches('input')){event.preventDefault();if(pickers[kind].options.length&&!pickers[kind].loading)choose(kind,pickers[kind].options[0]);return}
   if(!['ArrowDown','ArrowUp','Home','End'].includes(event.key))return;
   const options=[...$('[data-options="'+kind+'"]').querySelectorAll('[role=option]')];if(!options.length)return;
   if((event.key==='Home'||event.key==='End')&&event.target.matches('input'))return;
   event.preventDefault();const at=options.indexOf(document.activeElement),next=event.key==='Home'?0:event.key==='End'?options.length-1:event.key==='ArrowDown'?(at+1)%options.length:at<=0?options.length-1:at-1;options[next].focus();
  });
 }
 const estimate=form.elements.estimate;
 function validate(){
  const value=estimate.value.trim().toLowerCase(),parts=[...value.matchAll(/(\d+(?:\.\d+)?)\s*([wdhm])/g)];
  const valid=/^(?:\d+(?:\.\d+)?\s*[wdhm]\s*)+$/.test(value)&&parts.some(p=>Number(p[1])>0)&&new Set(parts.map(p=>p[2])).size===parts.length;
  estimate.setCustomValidity(valid?'':'Use 30m, 4h, 1d, 1w or 1h 30m. Use each unit once.');
  for(const button of dialog.querySelectorAll('[data-estimate]'))button.setAttribute('aria-pressed',String(button.dataset.estimate===value));
  const start=form.elements.start_date,due=form.elements.due_date;
  due.setCustomValidity(bounds&&due.value&&(due.value<bounds.first||due.value>bounds.last)?'Due date must be in the current month.':start.value&&due.value&&start.value>due.value?'Due date must be on or after the start date.':'');
 }
 for(const input of [estimate,form.elements.start_date,form.elements.due_date])input.addEventListener('input',validate);
 for(const button of dialog.querySelectorAll('[data-estimate]'))button.addEventListener('click',()=>{estimate.value=button.dataset.estimate;validate();estimate.focus()});
 form.addEventListener('submit',async event=>{
  event.preventDefault();if(busy||sent)return;for(const kind of isLeave()?['projects','stories']:kinds){if(!form.elements[kind].value){status.textContent='Choose '+(kind==='projects'?'a Jira space':kind==='epics'?'an epic':isLeave()?'a leave tracker':'a user story')+' first.';$('[data-pick-toggle="'+kind+'"]').setAttribute('aria-invalid','true');$('[data-pick-toggle="'+kind+'"]').focus();showPicker(kind);return}}validate();if(!form.reportValidity())return;
  setBusy(true);status.textContent='Checking the '+(isLeave()?'leave tracker':'story')+', dates and your Jira permissions…';
  try{
   const body={project_key:form.elements.projects.value,epic_key:form.elements.epics.value,story_key:form.elements.stories.value,summary:form.elements.summary.value,description:form.elements.description.value,start_date:form.elements.start_date.value,due_date:form.elements.due_date.value,estimate:estimate.value,leave_type:isLeave()?form.elements.leave_type.value:undefined};
   const result=await worklogRequest('/subtask/review',body);if(!alive)return;
   token=result.token;review=result.review;const r=review;
   $('[data-compose]').hidden=true;$('[data-review]').hidden=false;
   $('[data-review]').innerHTML=`<section class="subtask-check"><span class="subtask-kicker">READY WHEN YOU ARE</span><h3>A quick look before it’s live</h3><h4>${esc(r.summary)}</h4><p>${esc(r.project_key)} → ${r.flow==='leave'?'':esc(r.epic_key)+' · '+esc(r.epic_summary)+' → '}${esc(r.story_key)} · ${esc(r.story_summary)}</p><dl><dt>Assigned to</dt><dd>${esc(r.author)}</dd><dt>Start → Due</dt><dd>${esc(r.start_date)} → ${esc(r.due_date)}</dd><dt>Original estimate</dt><dd>${esc(r.estimate)}</dd>${r.flow==='leave'?`<dt>Leave type</dt><dd>${esc(r.leave_label)}</dd>`:''}</dl>${r.description?`<p class="subtask-description">${esc(r.description)}</p>`:''}<p>${r.flow==='leave'?'This creates a leave subtask. Leave is counted as taken only after leave time is recorded.':'This creates one real Jira subtask. It won’t log any hours.'}</p><div class="subtask-footer"><button type="button" data-back>← Edit details</button><button type="button" class="subtask-primary" data-create>Create subtask in Jira ✦</button></div></section>`;
   $('[data-back]').addEventListener('click',()=>{$('[data-compose]').hidden=false;$('[data-review]').hidden=true;status.textContent='';form.elements.summary.focus()});
   $('[data-create]').addEventListener('click',save);
   status.textContent='';setBusy(false);$('[data-create]').focus();
  }catch(error){status.textContent=error.message;setBusy(false)}
 });
 async function save(){
  if(busy||sent)return;sent=true;setBusy(true);status.textContent='Creating your subtask in Jira…';
  try{
   const saved=await worklogRequest('/subtask/submit',{token,confirmed:true});
   availabilityCreatedIssues.set(saved.issue_key,saved);
   document.dispatchEvent(new CustomEvent('epr-subtask-created',{detail:saved}));
   $('[data-review]').innerHTML=`<section class="subtask-success"><div class="subtask-spark">✦</div><h3>Your next step is ready!</h3><strong>${esc(saved.issue_key)} · ${esc(saved.summary)}</strong><p>Assigned to ${esc(saved.assignee)} inside ${esc(saved.parent_issue_key)}.</p><p>${saved.verified?'Verified in Jira.':'Created in Jira. Open it to verify the saved fields.'} No hours have been logged yet.</p><p class="subtask-help">${saved.flow==='leave'?'Record leave time separately; it will appear as leave taken, not work hours.':'You can log work on it right away.'} Other reports show it after the next full fetch.</p><div class="subtask-footer"><a data-jira target="_blank" rel="noopener noreferrer">Open in Jira ↗</a><button type="button" class="subtask-primary" data-log>${saved.flow==='leave'?'Record leave on this subtask':'Log work on this subtask'} →</button></div></section>`;
   const link=$('[data-jira]'),base=String(data.jira_browse_base||'');if(/^https?:\/\//i.test(base))link.href=base.replace(/\/$/,'')+'/'+encodeURIComponent(saved.issue_key);else link.hidden=true;
   status.textContent='';setBusy(false);
   $('[data-log]').addEventListener('click',()=>{close();month.value=bounds.today.slice(0,7);render();window.openEcuDrawer(author,'availability',month);openCalendarWorklog(author,bounds.today,{issue:saved.issue_key,mode:saved.flow==='leave'?'leave':'work',duration:saved.flow==='leave'?String(activeProfile().standard_hours_per_day||8)+'h':undefined})});
  }catch(error){
   status.textContent=error.message+' Check the selected '+(isLeave()?'leave tracker':'user story')+' in Jira before starting another creation.';setBusy(false);
   $('[data-create]').disabled=true;$('[data-back]').disabled=true;
  }
 }
 await load('projects');
}
document.addEventListener('click',event=>{const trigger=event.target.closest('[data-create-subtask]');if(trigger)openAvailabilitySubtask(trigger.dataset.createSubtask,trigger.dataset.subtaskDate)});
