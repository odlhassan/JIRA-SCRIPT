/* Personal Jira history: one shared manager for My Day and Availability. */
(() => {
  const api='/api/employee-capacity-utilization/worklog';
  const esc=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const description=node=>typeof node==='string'?node:node?.text||(node?.content||[]).map(description).join(' ');
  const duration=seconds=>{const minutes=Math.round(seconds/60);return [Math.floor(minutes/60)?Math.floor(minutes/60)+'h':'',minutes%60?minutes%60+'m':''].filter(Boolean).join(' ')||'0m'};
  const today=()=>{const now=new Date();return now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0')};
  async function request(path,body){
    const response=await fetch(api+path,{method:body?'POST':'GET',headers:body?{'Content-Type':'application/json','X-EPR-Worklog':'1'}:{},body:body?JSON.stringify(body):undefined});
    const result=await response.json();if(!response.ok)throw new Error(result.error||'Jira request failed.');return result;
  }
  function updateUrl(patch,replace=true){
    if(window.eprUpdateModalUrl)return window.eprUpdateModalUrl(patch,replace);
    const params=new URLSearchParams(location.hash.slice(1));for(const [key,value]of Object.entries(patch)){if(value==null||value==='')params.delete(key);else params.set(key,value)}
    history[replace?'replaceState':'pushState'](null,'',location.pathname+location.search+(params.size?'#'+params:''));
  }
  document.addEventListener('click',event=>{
    const trigger=event.target.closest('[data-manage-worklogs]');if(!trigger||document.querySelector('.wl-manager'))return;
    if(!document.querySelector('link[data-worklog-styles]')){const link=document.createElement('link');link.rel='stylesheet';link.href=api+'/manage.css';link.dataset.worklogStyles='1';document.head.append(link)}
    const route=new URLSearchParams(location.hash.slice(1));
    const initialMonth=trigger.dataset.worklogMonth||route.get('historyMonth')||document.querySelector('#month')?.value||today();
    const dialog=document.createElement('dialog');dialog.className='wl-manager';dialog.setAttribute('aria-label','Manage my logged hours');
    dialog.innerHTML=`<header class="wl-head"><div><span class="wl-kicker">PERSONAL TIMESHEET</span><h2>My logged hours</h2><p data-author>Your own Jira entries, in one place.</p></div><div class="wl-head-actions"><button type="button" data-share>Copy link</button><button type="button" data-close aria-label="Close worklog manager">✕</button></div></header>
      <div class="wl-content"><div class="wl-period"><div><button type="button" data-step="-1" aria-label="Previous history month">‹</button><label>Month<input name="period" type="month" value="${esc(initialMonth)}"></label><button type="button" data-step="1" aria-label="Next history month">›</button></div><div><button type="button" data-refresh>↻ Refresh</button><button type="button" data-export>Export CSV</button></div></div>
      <div class="wl-stats" data-stats></div>
      <section class="wl-filters" aria-label="Filter logged hours"><label class="wl-search">Search your entries<input name="search" type="search" placeholder="Task key, title or work description"></label><label>Task<select name="task"><option value="">All loaded tasks</option></select></label><label>Date (optional)<input name="date" type="date" value="${esc(trigger.dataset.worklogDate||'')}"></label><label>Sort by<select name="sort"><option value="newest">Newest first</option><option value="oldest">Oldest first</option><option value="longest">Longest first</option></select></label></section>
      <form data-search class="wl-direct"><label>Jira task key<input name="issue" pattern="[A-Za-z][A-Za-z0-9_]*-[0-9]+" placeholder="Optional · APP-123" value="${esc(trigger.dataset.manageWorklogs)}"></label><button type="submit">Load my entries</button><button type="button" data-all>All tasks this month</button><small>Browse the month automatically, or open any task directly across all dates.</small></form>
      <p data-status role="status" aria-live="polite"></p><input data-share-url aria-label="Shareable modal URL" readonly hidden><div data-entries class="wl-entries"></div><button type="button" data-more hidden>Load more tasks</button></div>
      <footer class="wl-foot">Only your own entries are shown. Deletion keeps the remaining estimate unchanged. Snapshot reports update after the next Jira refresh.</footer>`;
    document.body.append(dialog);dialog.showModal();
    dialog.addEventListener('keydown',event=>{if(event.key==='Escape')event.stopPropagation()});
    const $=selector=>dialog.querySelector(selector),status=$('[data-status]'),list=$('[data-entries]'),form=$('[data-search]'),period=$('[name="period"]');
    let loadError='',authAvailable=null;
    let entries=[],author='',cursor=null,busy=false,loaded=false,pending=null,failed=new Set(),visible=[],mode=trigger.dataset.manageWorklogs?'issue':'month',generation=0;
    const rowKey=row=>row.issue_key+'|'+row.worklog_id;
    const syncUrl=()=>updateUrl({panel:'hours',issue:mode==='issue'?form.elements.issue.value.trim().toUpperCase():null,date:$('[name="date"]').value,historyMonth:period.value});
    const setBusy=value=>{busy=value;dialog.setAttribute('aria-busy',String(value));for(const control of dialog.querySelectorAll('.wl-period button,.wl-period input,.wl-direct button,.wl-direct input,[data-more]'))control.disabled=value;$('[data-export]').disabled=value||!visible.length};
    const render=()=>{
      const query=$('[name="search"]').value.trim().toLowerCase(),date=$('[name="date"]').value,task=$('[name="task"]').value,sort=$('[name="sort"]').value;
      visible=entries.filter(row=>(!date||row.started.slice(0,10)===date)&&(!task||row.issue_key===task)&&(!query||[row.issue_key,row.summary,description(row.comment),row.worklog_id].join(' ').toLowerCase().includes(query)));
      visible.sort((a,b)=>sort==='longest'?b.seconds-a.seconds:sort==='oldest'?a.started.localeCompare(b.started):b.started.localeCompare(a.started));
      const total=visible.reduce((sum,row)=>sum+row.seconds,0);
      $('[data-stats]').innerHTML=[['Hours shown',duration(total)],['Entries',visible.length],['Tasks',new Set(visible.map(row=>row.issue_key)).size],['Days',new Set(visible.map(row=>row.started.slice(0,10))).size]].map(([label,value])=>`<div><small>${label}</small><strong>${esc(value)}</strong></div>`).join('');
      $('[data-author]').textContent=author?author+' · Personal Jira history':'Your own Jira entries, in one place.';
      list.innerHTML=visible.map(row=>{
        const index=entries.indexOf(row),day=new Date(row.started.slice(0,10)+'T12:00:00'),dateLabel=Number.isNaN(day.getTime())?row.started:day.toLocaleDateString(undefined,{weekday:'short',day:'numeric',month:'short',year:'numeric'});
        const confirm=pending===rowKey(row);
        return `<article class="wl-entry"><div class="wl-entry-top"><div><span class="wl-task">${esc(row.issue_key)}</span><strong>${esc(row.summary||'Jira worklog')}</strong></div><span class="wl-duration">${esc(duration(row.seconds))}</span></div><p class="wl-description">${esc(description(row.comment)||'No work description')}</p><div class="wl-entry-meta"><span>${esc(dateLabel)}${row.started.includes('T')?' · '+esc(row.started.slice(11,16)):''}</span><small>Worklog ${esc(row.worklog_id)}</small><button type="button" class="wl-delete" data-delete="${index}" ${failed.has(rowKey(row))?'disabled':''}>Delete entry</button></div>${confirm?`<div class="wl-confirm"><strong>Delete ${esc(duration(row.seconds))} from ${esc(row.issue_key)}?</strong><p>${esc(dateLabel)} · Worklog ${esc(row.worklog_id)}. This removes the entry from Jira.</p><button type="button" class="wl-danger" data-confirm="${index}">Confirm delete</button><button type="button" data-cancel>Cancel deletion</button></div>`:''}</article>`;
      }).join('')||`<div class="wl-empty"><strong>${loadError?'Jira history is unavailable':loaded?'No entries match this view':'Loading your Jira history…'}</strong><p>${loadError?esc(loadError):loaded?'Try another month, clear filters, or open a specific task.':'Your tasks and descriptions will appear here automatically.'}</p>${loadError?(authAvailable===true?'<a href="/api/jira-auth/login">Sign in with Jira</a>':authAvailable===false?'<p>Personal Jira sign-in needs administrator setup on this server.</p>':''):loaded?'<button type="button" data-clear>Clear filters</button>':''}</div>`;
      $('[data-more]').hidden=!cursor;
      $('[data-export]').disabled=!visible.length||busy;
      if(loaded)status.textContent=`${author} · ${visible.length} entries${mode==='issue'?' on '+form.elements.issue.value.trim().toUpperCase():' in '+period.value}${date?' · '+date:''}${cursor?' · More tasks available; totals cover loaded entries only.':''}`;
    };
    const taskOptions=()=>{const select=$('[name="task"]'),selected=select.value;select.innerHTML='<option value="">All loaded tasks</option>'+[...new Set(entries.map(row=>row.issue_key))].sort().map(task=>`<option>${esc(task)}</option>`).join('');select.value=selected};
    async function load(more=false){
      if(busy)return;loadError='';authAvailable=null;const version=++generation;setBusy(true);pending=null;status.textContent='Reading your entries from Jira…';status.classList.remove('wl-error');
      if(!more){entries=[];cursor=null;loaded=false;failed.clear();render();status.textContent='Reading your entries from Jira…'}
      try{
        const key=form.elements.issue.value.trim().toUpperCase();
        const path=mode==='issue'?'/entries?issue_key='+encodeURIComponent(key):'/history?month='+encodeURIComponent(period.value)+(more&&cursor?'&cursor='+encodeURIComponent(cursor):'');
        const result=await request(path);if(version!==generation||!dialog.isConnected)return;
        author=result.author||'';
        const fresh=(result.entries||[]).map(row=>({...row,issue_key:row.issue_key||result.issue_key||key,seconds:Number(row.seconds)||0}));
        const existing=new Map((more?entries:[]).map(row=>[rowKey(row),row]));for(const row of fresh)existing.set(rowKey(row),row);
        entries=[...existing.values()];cursor=result.next_cursor||null;loaded=true;taskOptions();syncUrl();setBusy(false);render();
      }catch(error){loadError=error.message;try{const auth=await fetch('/api/jira-auth/status').then(response=>response.json());authAvailable=auth.configured}catch(_error){}if(version!==generation||!dialog.isConnected)return;setBusy(false);loaded=true;render();status.textContent=error.message+' Use Refresh to try again.';status.classList.add('wl-error')}
      finally{if(dialog.isConnected)setBusy(false)}
    }
    for(const control of dialog.querySelectorAll('[name="search"],[name="task"],[name="date"],[name="sort"]'))control.addEventListener(control.name==='search'?'input':'change',()=>{if(!busy){pending=null;render();syncUrl()}});
    form.addEventListener('submit',event=>{event.preventDefault();mode=form.elements.issue.value.trim()?'issue':'month';load()});
    $('[data-all]').addEventListener('click',()=>{mode='month';form.elements.issue.value='';$('[name="date"]').value='';$('[name="task"]').value='';load()});
    $('[data-refresh]').addEventListener('click',()=>load());$('[data-more]').addEventListener('click',()=>load(true));
    period.addEventListener('change',()=>{if(!period.value)return;mode='month';form.elements.issue.value='';$('[name="date"]').value='';load()});
    for(const button of dialog.querySelectorAll('[data-step]'))button.addEventListener('click',()=>{const date=new Date(period.value+'-01T12:00:00');date.setMonth(date.getMonth()+Number(button.dataset.step));if(date.getFullYear()<1||date.getFullYear()>9999)return;period.value=date.getFullYear()+'-'+String(date.getMonth()+1).padStart(2,'0');period.dispatchEvent(new Event('change'))});
    list.addEventListener('click',async event=>{
      if(busy)return;const button=event.target.closest('button');if(!button)return;
      if(button.hasAttribute('data-clear')){for(const name of ['search','task','date'])$('[name="'+name+'"]').value='';render();syncUrl();return}
      if(button.hasAttribute('data-cancel')){pending=null;render();return}
      if(button.hasAttribute('data-delete')){pending=rowKey(entries[Number(button.dataset.delete)]);render();$('[data-confirm]')?.focus();return}
      if(!button.hasAttribute('data-confirm'))return;
      const entry=entries[Number(button.dataset.confirm)];setBusy(true);for(const control of list.querySelectorAll('button'))control.disabled=true;status.textContent='Deleting from Jira…';
      try{
        await request('/delete',{issue_key:entry.issue_key,worklog_id:entry.worklog_id,confirmed:true});
        entries=entries.filter(row=>rowKey(row)!==rowKey(entry));pending=null;
        let remaining;
        try{const result=await request('/entries?issue_key='+encodeURIComponent(entry.issue_key));remaining=result.entries.filter(row=>row.worklog_id!==entry.worklog_id)}catch(_error){/* The deletion already succeeded; snapshot reconciliation can wait. */}
        document.dispatchEvent(new CustomEvent('epr-worklog-deleted',{detail:{issue_key:entry.issue_key,author,entries:remaining,deleted:entry}}));
        setBusy(false);taskOptions();render();status.textContent=`Deleted worklog ${entry.worklog_id} from Jira.`;
      }catch(error){failed.add(rowKey(entry));pending=null;setBusy(false);render();status.textContent=error.message+' Check Jira and load entries again before retrying.'}
      finally{setBusy(false)}
    });
    $('[data-export]').addEventListener('click',()=>{
      const cell=value=>{let text=String(value??'');if(/^[=+@\-\t\r]/.test(text))text="'"+text;return '"'+text.replace(/"/g,'""')+'"'};
      const rows=[['Task','Title','Started','Hours','Description','Worklog ID'],...visible.map(row=>[row.issue_key,row.summary,row.started,row.seconds/3600,description(row.comment),row.worklog_id])];
      const url=URL.createObjectURL(new Blob(['\ufeff'+rows.map(row=>row.map(cell).join(',')).join('\r\n')],{type:'text/csv;charset=utf-8'}));const link=document.createElement('a');link.href=url;link.download='my-logged-hours-'+period.value+'.csv';link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
    });
    $('[data-share]').addEventListener('click',async()=>{syncUrl();try{await navigator.clipboard.writeText(location.href);status.textContent='Link copied. Recipients will see their own Jira entries.'}catch(_error){$('[data-share-url]').hidden=false;$('[data-share-url]').value=location.href;$('[data-share-url]').select();status.textContent='Copy this link. Recipients will see their own Jira entries.'}});
    $('[data-close]').addEventListener('click',()=>dialog.close());
    dialog.addEventListener('close',()=>{generation++;dialog.remove();if(!dialog.dataset.routeClosing)updateUrl({panel:null,issue:null,date:null,historyMonth:null});if(trigger.isConnected)trigger.focus()});
    window.eprCloseWorklogManager=()=>{dialog.dataset.routeClosing='1';dialog.close();dialog.remove()};
    updateUrl({panel:'hours'},false);syncUrl();render();load();
  });
  function openFromLink(){const route=new URLSearchParams(location.hash.slice(1));if((new URLSearchParams(location.search).get('manageWorklogs')==='1'||route.get('panel')==='hours')&&!window.eprRestoreModalRoute){const trigger=document.querySelector('[data-manage-worklogs]');if(trigger){trigger.dataset.manageWorklogs=route.get('issue')||'';trigger.dataset.worklogDate=route.get('date')||'';trigger.click()}}}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',openFromLink,{once:true});else openFromLink();
})();
