/* Guided creation uses live Jira choices, personal OAuth and one-attempt receipts. */
const availabilityCreatedIssues=new Map();
function availabilitySubtasks(){return [...new Map([...availabilityCreatedIssues.values(),...issues].map(item=>[item.issue_key,item])).values()]}
async function openAvailabilitySubtask(label,selectedDate){
 if(document.querySelector('.subtask-dialog'))return;
 const previous=document.activeElement,dialog=document.createElement('dialog');
 dialog.className='subtask-dialog';dialog.setAttribute('aria-labelledby','subtask-title');
 dialog.innerHTML=`<header class="subtask-head"><div><span class="subtask-kicker">A LITTLE PLAN, A BIG STEP ✦</span><h2 id="subtask-title">Create a new subtask</h2><p>Pick its home. Plan your work. Make it happen.</p></div><button type="button" data-close aria-label="Close subtask creator">✕</button></header><div class="subtask-content"><p role="status" data-status>Connecting to your Jira account…</p><div data-compose hidden><p class="subtask-owner" data-owner></p><form id="availability-subtask-form"><section class="subtask-section subtask-home"><h3><span>1</span> Find its home</h3><p>Every subtask lives in a user story inside an epic.</p><div class="subtask-path">Space → Epic → User story → Your new subtask</div>${[['projects','Jira space','Search spaces by name or key'],['epics','Epic','Search epics by title or key'],['stories','User story','Search user stories by title or key']].map(([kind,title,placeholder])=>`<div class="subtask-picker" data-picker="${kind}"><label>${title}<input type="search" data-search="${kind}" aria-label="Search ${title.toLowerCase()}" placeholder="${placeholder}" autocomplete="off" ${kind==='projects'?'':'disabled'}><select name="${kind}" required aria-label="${title}" disabled><option value="">${kind==='projects'?'Loading spaces…':kind==='epics'?'Choose a space first':'Choose an epic first'}</option></select></label><button type="button" data-more="${kind}" hidden>Load more</button><small data-hint="${kind}" role="status"></small></div>`).join('')}</section><section class="subtask-section subtask-what"><h3><span>2</span> Give it a purpose</h3><label>Subtask title<input name="summary" maxlength="255" required placeholder="e.g. Build the customer search filter" autocomplete="off"></label><label>Description <small>(optional)</small><textarea name="description" maxlength="2000" rows="3" placeholder="What needs to be done?"></textarea></label></section><section class="subtask-section subtask-when"><h3><span>3</span> Make a little plan</h3><div class="subtask-dates"><label>Start date<input name="start_date" type="date" required></label><label>Due date<input name="due_date" type="date" required aria-describedby="subtask-date-help"></label></div><p id="subtask-date-help" class="subtask-help" data-date-help></p><label>Original estimate<input name="estimate" value="1d" maxlength="100" required placeholder="e.g. 1d or 1h 30m" aria-describedby="subtask-estimate-help"></label><div class="subtask-presets" aria-label="Quick original estimates">${['30m','4h','1d','1w'].map(value=>`<button type="button" data-estimate="${value}" aria-pressed="${value==='1d'}">${value}</button>`).join('')}</div><p id="subtask-estimate-help" class="subtask-help">A day or week follows your Jira working-time settings. This is planned effort; it does not log hours.</p></section><footer class="subtask-footer"><span>Assigned to you, automatically ✨</span><button type="submit" class="subtask-primary">Review subtask →</button></footer></form></div><div data-review hidden></div></div>`;
 document.body.append(dialog);dialog.showModal();
 const $=selector=>dialog.querySelector(selector),status=$('[data-status]'),form=$('form');
 let busy=false,sent=false,token=null,review=null,author='',bounds=null,alive=true;
 const generations={projects:0,epics:0,stories:0},cursors={},timers={};
 const close=()=>{if(busy)return;alive=false;for(const timer of Object.values(timers))clearTimeout(timer);dialog.close();dialog.remove();previous?.focus?.()};
 $('[data-close]').addEventListener('click',close);
 dialog.addEventListener('cancel',event=>{event.preventDefault();close()});
 dialog.addEventListener('keydown',event=>event.stopPropagation());
 const setBusy=value=>{busy=value;dialog.setAttribute('aria-busy',String(value));for(const control of dialog.querySelectorAll('button'))control.disabled=value;$('[data-close]').disabled=value};
 function reset(kind,message){
  generations[kind]++;clearTimeout(timers[kind]);cursors[kind]=null;
  const select=form.elements[kind];select.innerHTML=`<option value="">${esc(message)}</option>`;select.disabled=true;
  $('[data-search="'+kind+'"]').disabled=true;$('[data-more="'+kind+'"]').hidden=true;$('[data-hint="'+kind+'"]').textContent='';
 }
 function resetChildren(kind){if(kind==='projects'){reset('epics','Choose a space first');reset('stories','Choose an epic first')}if(kind==='epics')reset('stories','Choose an epic first')}
 async function load(kind,more=false){
  const version=++generations[kind],select=form.elements[kind],search=$('[data-search="'+kind+'"]'),hint=$('[data-hint="'+kind+'"]'),moreButton=$('[data-more="'+kind+'"]');
  const params=new URLSearchParams({kind,q:search.value});
  if(kind!=='projects')params.set('project_key',form.elements.projects.value);
  if(kind==='stories')params.set('epic_key',form.elements.epics.value);
  if(more&&cursors[kind])params.set('cursor',cursors[kind]);
  if(!more){select.innerHTML='<option value="">Loading…</option>';resetChildren(kind)}
  select.disabled=true;moreButton.hidden=true;hint.textContent='Reading Jira…';
  try{
   const result=await worklogRequest('/subtask/options?'+params);
   if(!alive||generations[kind]!==version)return;
   if(kind==='projects'&&!bounds){
    author=result.author;bounds=result;
    $('[data-owner]').textContent='✓ Signed in as '+author+' · Your new subtask will be assigned to you.';
    const chosen=selectedDate>=result.first&&selectedDate<=result.last?selectedDate:result.today;
    form.elements.start_date.value=chosen;form.elements.due_date.value=chosen;
    form.elements.due_date.min=result.first;form.elements.due_date.max=result.last;
    form.elements.start_date.max=result.last;
    $('[data-date-help]').textContent='Due this month only: '+result.first+' to '+result.last+' (your Jira timezone).';
    $('[data-compose]').hidden=false;status.textContent='';
   }
   if(!more)select.innerHTML='<option value="">Choose '+(kind==='projects'?'a space':kind==='epics'?'an epic':'a user story')+'…</option>';
   for(const item of result.options){if(![...select.options].some(option=>option.value===item.key)){const option=new Option(item.key+' · '+item.summary,item.key);select.add(option)}}
   cursors[kind]=result.next_cursor;select.disabled=select.options.length<=1;search.disabled=false;moreButton.hidden=!result.next_cursor;moreButton.disabled=false;
   hint.textContent=result.next_cursor?'More results available. Search or load more.':select.options.length<=1?'No matches. Try another search or parent.':'Choose one to continue.';
  }catch(error){if(alive&&generations[kind]===version){
   hint.textContent=error.message;search.disabled=false;
   if(!bounds){
    status.textContent=error.message;
    try{const auth=await fetch('/api/jira-auth/status').then(response=>response.json());if(!alive)return;if(auth.configured===false){status.textContent='Personal Jira sign-in needs administrator setup on this server before you can create a subtask.';return}}catch(_error){/* Keep the original sign-in error. */}
    if(!alive)return;const link=document.createElement('a');link.href='/api/jira-auth/login';link.textContent=' Sign in with Jira';status.append(link);
   }else {moreButton.hidden=false;moreButton.textContent='Retry'}
  }}
 }
 for(const kind of ['projects','epics','stories']){
  $('[data-search="'+kind+'"]').addEventListener('input',()=>{generations[kind]++;clearTimeout(timers[kind]);resetChildren(kind);form.elements[kind].value='';form.elements[kind].disabled=true;timers[kind]=setTimeout(()=>load(kind),300)});
  $('[data-more="'+kind+'"]').addEventListener('click',()=>load(kind,Boolean(cursors[kind])));
  form.elements[kind].addEventListener('change',()=>{resetChildren(kind);if(kind==='projects'&&form.elements.projects.value){$('[data-search="epics"]').value='';load('epics')}if(kind==='epics'&&form.elements.epics.value){$('[data-search="stories"]').value='';load('stories')}});
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
  event.preventDefault();if(busy||sent)return;validate();if(!form.reportValidity())return;
  setBusy(true);status.textContent='Checking the story, dates and your Jira permissions…';
  try{
   const body={project_key:form.elements.projects.value,epic_key:form.elements.epics.value,story_key:form.elements.stories.value,summary:form.elements.summary.value,description:form.elements.description.value,start_date:form.elements.start_date.value,due_date:form.elements.due_date.value,estimate:estimate.value};
   const result=await worklogRequest('/subtask/review',body);if(!alive)return;
   token=result.token;review=result.review;const r=review;
   $('[data-compose]').hidden=true;$('[data-review]').hidden=false;
   $('[data-review]').innerHTML=`<section class="subtask-check"><span class="subtask-kicker">READY WHEN YOU ARE</span><h3>A quick look before it’s live</h3><h4>${esc(r.summary)}</h4><p>${esc(r.project_key)} → ${esc(r.epic_key)} · ${esc(r.epic_summary)} → ${esc(r.story_key)} · ${esc(r.story_summary)}</p><dl><dt>Assigned to</dt><dd>${esc(r.author)}</dd><dt>Start → Due</dt><dd>${esc(r.start_date)} → ${esc(r.due_date)}</dd><dt>Original estimate</dt><dd>${esc(r.estimate)}</dd></dl>${r.description?`<p class="subtask-description">${esc(r.description)}</p>`:''}<p>This creates one real Jira subtask. It won’t log any hours.</p><div class="subtask-footer"><button type="button" data-back>← Edit details</button><button type="button" class="subtask-primary" data-create>Create subtask in Jira ✦</button></div></section>`;
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
   $('[data-review]').innerHTML=`<section class="subtask-success"><div class="subtask-spark">✦</div><h3>Your next step is ready!</h3><strong>${esc(saved.issue_key)} · ${esc(saved.summary)}</strong><p>Assigned to ${esc(saved.assignee)} inside ${esc(saved.parent_issue_key)}.</p><p>${saved.verified?'Verified in Jira.':'Created in Jira. Open it to verify the saved fields.'} No hours have been logged yet.</p><p class="subtask-help">You can log work on it right away. Other reports show it after the next full fetch.</p><div class="subtask-footer"><a data-jira target="_blank" rel="noopener noreferrer">Open in Jira ↗</a><button type="button" class="subtask-primary" data-log>Log work on this subtask →</button></div></section>`;
   const link=$('[data-jira]'),base=String(data.jira_browse_base||'');if(/^https?:\/\//i.test(base))link.href=base.replace(/\/$/,'')+'/'+encodeURIComponent(saved.issue_key);else link.hidden=true;
   status.textContent='';setBusy(false);
   $('[data-log]').addEventListener('click',()=>{close();month.value=bounds.today.slice(0,7);render();window.openEcuDrawer(author,'availability',month);openCalendarWorklog(author,bounds.today,{issue:saved.issue_key})});
  }catch(error){
   status.textContent=error.message+' Check the selected user story in Jira before starting another creation.';setBusy(false);
   $('[data-create]').disabled=true;$('[data-back]').disabled=true;
  }
 }
 await load('projects');
}
document.addEventListener('click',event=>{const trigger=event.target.closest('[data-create-subtask]');if(trigger)openAvailabilitySubtask(trigger.dataset.createSubtask,trigger.dataset.subtaskDate)});
