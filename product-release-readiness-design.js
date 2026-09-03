const PEOPLE = ["Hassan","Maya","Areeba","Tom","Customer Success","Product Team","Process Team","Implementation Team","PGL / APL / Aramco","Unassigned"];
const STATUSES = [
  {key:"done",label:"Done",icon:"check_circle"},
  {key:"planned",label:"Planned",icon:"pending"},
  {key:"skipped",label:"Skipped",icon:"remove_circle"},
  {key:"need_confirmation",label:"Need confirmation",icon:"contact_support"}
];
const DEMO_RELEASES = [
  {id:"demo-v935",is_demo:true,project_key:"DIGITALLOG",release_number:"v9.3.5",release_date:"2026-09-05",release_status:"scheduled",notes:"",epics:[
    {epic_row_id:"demo-1",epic_key:"DL-482",epic_name:"Log Chat enhancements"},
    {epic_row_id:"demo-2",epic_key:"DL-517",epic_name:"Default reports"}
  ]},
  {id:"demo-v312",is_demo:true,project_key:"OMNICONNECT",release_number:"v3.12.0",release_date:"2026-08-31",release_status:"scheduled",notes:"",epics:[
    {epic_row_id:"demo-3",epic_key:"OC-913",epic_name:"Price write-back (fusion controller)"},
    {epic_row_id:"demo-4",epic_key:"OC-927",epic_name:"Manual dip-based stock reconciliation"}
  ]},
  {id:"demo-v331",is_demo:true,project_key:"OMNICONNECT",release_number:"v3.31.0",release_date:"2026-08-19",release_status:"released",notes:"",epics:[
    {epic_row_id:"demo-5",epic_key:"OC-944",epic_name:"Nested calculated tag"}
  ]}
];
let releases = [];
let epicPool = [];
let projectNames = {};
let selectedProductKey = "";
let selectedReleaseId = "";
let boardByRelease = {};
let pickerSelectedIds = new Set();
let detailTargetRef = "";
let detailReleaseId = "";
let pendingReleaseCompletionId = "";

function esc(value) {
  return String(value == null ? "" : value)
    .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}
function uid(prefix) { return prefix + "-" + Date.now().toString(36) + "-" + Math.random().toString(36).slice(2,6); }
function isDemo(release) { return Boolean(release && (release.is_demo || String(release.id).startsWith("demo-"))); }
function selectedRelease() { return releases.find(function(release){ return release.id === selectedReleaseId; }) || null; }
function findRelease(releaseId) { return releases.find(function(release){ return release.id === releaseId; }) || null; }
function storageKey(releaseId) { return "release-readiness-navigator-v4:" + releaseId; }
function formatDate(value) {
  if (!value) return "No release date";
  const date = new Date(value + "T00:00:00");
  return Number.isNaN(date.getTime()) ? value : date.toLocaleDateString(undefined,{day:"numeric",month:"short",year:"numeric"});
}
function optionsHtml(selected) {
  return PEOPLE.map(function(name){ return '<option value="' + esc(name) + '"' + (name === selected ? " selected" : "") + ">" + esc(name) + "</option>"; }).join("");
}
function normalizeEntity(entity) {
  entity.status = entity.status || "planned";
  entity.delayed = entity.status === "planned" && Boolean(entity.delayed);
  entity.owner = entity.owner || "Unassigned";
  entity.confirm_by = entity.confirm_by || "Hassan";
  entity.confirm_from = entity.confirm_from || "Product Team";
  entity.link = entity.link || "";
  entity.notes = entity.notes || "";
  return entity;
}
function normalizeScopes(item) {
  if (!Array.isArray(item.scopes)) item.scopes = [item.scope || "release"];
  item.scopes = item.scopes.filter(Boolean);
  if (!item.scopes.length) item.scopes = ["release"];
  delete item.scope;
}
function seedBoard(release,index) {
  const epics = (release.epics || []).map(function(epic,epicIndex){
    return normalizeEntity({
      id:String(epic.epic_row_id || epic.id || uid("epic")),
      epic_key:epic.epic_key || "EPIC",
      title:epic.epic_name || epic.epic_key || "Unnamed epic",
      status:epicIndex === 0 ? "done" : "planned",
      delayed:false,
      owner:epicIndex === 0 ? "Hassan" : "Areeba"
    });
  });
  const firstScope = epics[0] ? epics[0].id : "release";
  const secondScope = epics[1] ? epics[1].id : firstScope;
  return normalizeBoard({
    status:release.release_status === "released" ? "done" : "planned",
    delayed:false,
    archived:false,
    lifecycle_status:release.release_status || "scheduled",
    confirm_by:"Hassan",
    confirm_from:"Product Team",
    link:"",
    notes:"",
    epics:epics,
    checklists:[
      {id:uid("documentation"),title:"Documentation",status:index % 2 ? "done" : "planned",delayed:false,owner:"Maya",confirm_by:"Hassan",confirm_from:"Product Team",link:"",notes:"",items:[
        {id:uid("doc"),title:"Release notes / user guide write-back",status:"done",delayed:false,owner:"Maya",scopes:["release"],confirm_by:"Hassan",confirm_from:"Product Team",link:"<Link here>",notes:""},
        {id:uid("doc"),title:"Epic-specific documentation impact",status:"planned",delayed:false,owner:"Hassan",scopes:[firstScope],confirm_by:"Hassan",confirm_from:"Product Team",link:"",notes:""}
      ]},
      {id:uid("stakeholder"),title:"Stakeholder buy-in",status:"need_confirmation",delayed:false,owner:"Customer Success",confirm_by:"Hassan",confirm_from:"Implementation Team",link:"",notes:"Coordinate confirmation in the working group.",items:[
        {id:uid("stake"),title:"Customer Success / Tom / Process Team / PGL / APL / Aramco",status:"need_confirmation",delayed:false,owner:"Customer Success",scopes:["release"],confirm_by:"Hassan",confirm_from:"PGL / APL / Aramco",link:"",notes:""}
      ]},
      {id:uid("videos"),title:"Videos",status:"planned",delayed:false,owner:"Maya",confirm_by:"Hassan",confirm_from:"Product Team",link:"",notes:"",items:[
        {id:uid("video"),title:(epics[0] ? epics[0].title : "Feature") + " — feature video",status:"done",delayed:false,owner:"Maya",scopes:[firstScope],confirm_by:"Hassan",confirm_from:"Product Team",link:"Feature video link",notes:""},
        {id:uid("video"),title:(epics[1] ? epics[1].title : "Release") + " — feature video",status:"planned",delayed:false,owner:"Maya",scopes:[secondScope],confirm_by:"Hassan",confirm_from:"Product Team",link:"",notes:""}
      ]},
      {id:uid("email"),title:"Announcement email",status:"planned",delayed:index === 0,owner:"Customer Success",confirm_by:"Hassan",confirm_from:"Product Team",link:"",notes:"Coordinate staged activation with Customer Success.",items:[
        {id:uid("email"),title:"Draft and approve customer announcement",status:"planned",delayed:index === 0,owner:"Customer Success",scopes:["release"],confirm_by:"Hassan",confirm_from:"Product Team",link:"",notes:""}
      ]}
    ]
  });
}
function normalizeBoard(board) {
  normalizeEntity(board);
  board.archived = Boolean(board.archived);
  board.lifecycle_status = String(board.lifecycle_status || "");
  board.epics = (board.epics || []).map(normalizeEntity);
  board.checklists = (board.checklists || []).map(function(check){
    normalizeEntity(check);
    check.title = check.title || "Checklist";
    check.items = (check.items || []).map(function(item){ normalizeEntity(item); normalizeScopes(item); item.title = item.title || "Checklist content"; return item; });
    return check;
  });
  return board;
}
function syncBoardReleaseLifecycle(release,board) {
  if (isDemo(release)) return;
  const lifecycle = release.release_status || "scheduled";
  if (board.lifecycle_status === lifecycle) return;
  if (lifecycle === "released") {
    board.status = "done";
    board.delayed = false;
  } else if (board.lifecycle_status === "released" || !board.lifecycle_status) {
    board.status = "planned";
    board.delayed = false;
  }
  board.lifecycle_status = lifecycle;
}
function syncBoardEpics(release,board) {
  const existing = new Map(board.epics.map(function(epic){ return [String(epic.id),epic]; }));
  board.epics = (release.epics || []).map(function(epic){
    const id = String(epic.epic_row_id || epic.id);
    const current = existing.get(id);
    if (current) { current.epic_key = epic.epic_key || current.epic_key; current.title = epic.epic_name || current.title; return normalizeEntity(current); }
    return normalizeEntity({id:id,epic_key:epic.epic_key || "EPIC",title:epic.epic_name || epic.epic_key || "Unnamed epic",status:"planned",delayed:false,owner:"Unassigned"});
  });
  const valid = new Set(board.epics.map(function(epic){ return epic.id; }));
  board.checklists.forEach(function(check){
    check.items.forEach(function(item){
      item.scopes = item.scopes.filter(function(scope){ return scope === "release" || valid.has(String(scope)); });
      if (!item.scopes.length) item.scopes = ["release"];
    });
  });
}
function getBoard(release) {
  if (!release) return null;
  if (!boardByRelease[release.id]) {
    const saved = localStorage.getItem(storageKey(release.id));
    try { boardByRelease[release.id] = saved ? normalizeBoard(JSON.parse(saved)) : seedBoard(release,releases.indexOf(release)); }
    catch (error) { boardByRelease[release.id] = seedBoard(release,releases.indexOf(release)); }
  }
  syncBoardEpics(release,boardByRelease[release.id]);
  syncBoardReleaseLifecycle(release,boardByRelease[release.id]);
  return boardByRelease[release.id];
}
function saveBoard(releaseId) { localStorage.setItem(storageKey(releaseId),JSON.stringify(boardByRelease[releaseId])); }
function resolveTarget(ref,releaseId) {
  const board = getBoard(findRelease(releaseId));
  if (!board || !ref) return null;
  const parts = ref.split(":");
  if (parts[0] === "release") return board;
  if (parts[0] === "epic") return board.epics.find(function(epic){ return epic.id === parts[1]; }) || null;
  const checklist = board.checklists.find(function(check){ return check.id === parts[1]; });
  if (parts[0] === "check") return checklist || null;
  if (parts[0] === "item" && checklist) return checklist.items.find(function(item){ return item.id === parts[2]; }) || null;
  return null;
}
function targetLabel(ref,releaseId) {
  const release = findRelease(releaseId);
  const target = resolveTarget(ref,releaseId);
  if (!release || !target) return "";
  if (ref === "release") return release.project_key + " · " + release.release_number;
  return target.title || target.epic_key || "Board item";
}
function statusMeta(status) { return STATUSES.find(function(entry){ return entry.key === status; }) || STATUSES[1]; }
function statusBadge(target) {
  const meta = statusMeta(target.status);
  return '<span class="tag' + (target.status === "planned" && target.delayed ? ' demo-tag' : "") + '">' + esc(meta.label) + (target.status === "planned" && target.delayed ? " · delayed" : "") + "</span>";
}
function statusPickerHtml(ref,target,releaseId) {
  const meta = statusMeta(target.status);
  const options = STATUSES.map(function(status){
    return '<button type="button" class="status-option' + (target.status === status.key ? " selected" : "") + '" data-status-value="' + status.key + '"><span>' + esc(status.label) + '</span><span class="material-symbols-rounded">' + status.icon + "</span></button>";
  }).join("");
  const delay = target.status === "planned" ? '<button type="button" class="status-option status-delay" data-toggle-delay="true"><span>' + (target.delayed ? "Clear delayed flag" : "Mark Planned as delayed") + '</span><span class="material-symbols-rounded">schedule</span></button>' : "";
  return '<div class="status-picker" data-status-picker="' + esc(ref) + '" data-release-id="' + esc(releaseId) + '"><button class="status-trigger' + (target.status === "planned" && target.delayed ? " delayed" : "") + '" type="button" data-status-trigger="true"><span class="material-symbols-rounded">' + meta.icon + "</span><span>" + esc(meta.label) + '</span><span class="material-symbols-rounded">expand_more</span></button><div class="status-menu">' + options + delay + "</div></div>";
}
function confirmationHtml(ref,target,releaseId) {
  if (target.status !== "need_confirmation") return "";
  return '<div class="confirmation-line"><span class="material-symbols-rounded">contact_support</span><strong>Confirmation taken by</strong><select class="confirm-select" data-confirm-field="confirm_by" data-target-ref="' + esc(ref) + '" data-release-id="' + esc(releaseId) + '">' + optionsHtml(target.confirm_by) + '</select><strong>from</strong><select class="confirm-select" data-confirm-field="confirm_from" data-target-ref="' + esc(ref) + '" data-release-id="' + esc(releaseId) + '">' + optionsHtml(target.confirm_from) + "</select></div>";
}
function ownerSelectHtml(ref,target,releaseId,label) {
  return '<label class="control-field owner-control"><span class="control-label">' + esc(label || "Responsible") + '</span><select class="owner-select" data-owner-ref="' + esc(ref) + '" data-release-id="' + esc(releaseId) + '">' + optionsHtml(target.owner) + "</select></label>";
}
function scopeSummary(target,board) {
  if (target.scopes.includes("release")) return "Whole release";
  const matched = board.epics.filter(function(epic){ return target.scopes.includes(epic.id); });
  if (!matched.length) return "Whole release";
  if (matched.length === 1) return matched[0].epic_key + " · " + matched[0].title;
  return matched.length + " epics";
}
function scopePickerHtml(ref,target,releaseId) {
  const board = getBoard(findRelease(releaseId));
  const releaseChecked = target.scopes.includes("release");
  let options = '<label class="scope-option" data-scope-search-text="whole release"><input type="checkbox" value="release"' + (releaseChecked ? " checked" : "") + '><span><strong>Whole release</strong><small>Applies once to the complete release</small></span></label>';
  board.epics.forEach(function(epic){
    options += '<label class="scope-option" data-scope-search-text="' + esc((epic.epic_key + " " + epic.title).toLowerCase()) + '"><input type="checkbox" value="' + esc(epic.id) + '"' + (target.scopes.includes(epic.id) ? " checked" : "") + '><span><strong>' + esc(epic.epic_key + " · " + epic.title) + "</strong><small>Specific epic</small></span></label>";
  });
  return '<div class="scope-picker" data-scope-picker="' + esc(ref) + '" data-release-id="' + esc(releaseId) + '"><button class="scope-trigger" type="button" data-scope-trigger="true"><span>' + esc(scopeSummary(target,board)) + '</span><span class="material-symbols-rounded">expand_more</span></button><div class="scope-menu"><input class="scope-search" type="search" placeholder="Search epic name"><div class="scope-options">' + options + "</div></div></div>";
}
function statusControlHtml(ref,target,releaseId) {
  return '<div class="control-field status-control"><span class="control-label">Status</span>' + statusPickerHtml(ref,target,releaseId) + "</div>";
}
function scopeControlHtml(ref,target,releaseId) {
  return '<div class="control-field scope-control"><span class="control-label">Scope</span>' + scopePickerHtml(ref,target,releaseId) + "</div>";
}
function rowActionsHtml(ref,releaseId,deleteTitle,deleteIcon) {
  return '<div class="row-actions"><button class="icon-btn" type="button" data-details-ref="' + esc(ref) + '" data-release-id="' + esc(releaseId) + '" title="Optional details"><span class="material-symbols-rounded">more_horiz</span></button><button class="icon-btn danger" type="button" data-delete-ref="' + esc(ref) + '" data-release-id="' + esc(releaseId) + '" title="' + esc(deleteTitle) + '"><span class="material-symbols-rounded">' + esc(deleteIcon || "close") + "</span></button></div>";
}
function latestReleasedAction(release) {
  return (release.actions || []).find(function(action){ return action.action === "released"; }) || null;
}
function releaseLifecycleHtml(release) {
  if (isDemo(release)) return '<span class="sync-note"><span class="material-symbols-rounded">science</span>Design data stays local</span>';
  if (release.release_status === "released") {
    const action = latestReleasedAction(release);
    const details = action ? [action.actual_date ? formatDate(action.actual_date) : "",action.actor ? "by " + action.actor : ""].filter(Boolean).join(" · ") : "";
    return '<span class="sync-note synced"><span class="material-symbols-rounded">sync</span><span><strong>Synced with Product Releases</strong>' + (details ? " · " + esc(details) : "") + "</span></span>";
  }
  return '<span class="sync-note"><span class="material-symbols-rounded">sync</span>Completion will update Product Releases</span>';
}
function completionPanelHtml(release) {
  if (pendingReleaseCompletionId !== release.id) return "";
  return '<section class="completion-panel" aria-label="Complete release"><div class="completion-copy"><span class="material-symbols-rounded">task_alt</span><div><strong>Complete this release</strong><small>This records the same Released action and history used by Product Releases.</small></div></div><div class="completion-fields"><label class="field"><span>Actual release date</span><input id="completion-actual-date" type="date" value="' + esc(release.release_date) + '"></label><label class="field"><span>Completed by</span><input id="completion-actor" type="text" placeholder="Required"></label><label class="field completion-notes"><span>Comments</span><input id="completion-notes" type="text" placeholder="Optional release note"></label><div class="completion-actions"><button class="btn" id="cancel-release-completion" type="button">Cancel</button><button class="btn primary" id="confirm-release-completion" type="button"><span class="material-symbols-rounded">check_circle</span> Mark released</button></div></div></section>';
}
function toast(message) {
  const element = document.getElementById("toast");
  element.textContent = message;
  element.classList.add("show");
  clearTimeout(toast.timer);
  toast.timer = setTimeout(function(){ element.classList.remove("show"); },2000);
}
function productLabel(key) { return projectNames[key] || key; }
function combineRecognizedReleases(liveReleases) {
  const recognizedProjectKeys = new Set(Object.keys(projectNames));
  liveReleases.forEach(function(release){ if (release.project_key) recognizedProjectKeys.add(release.project_key); });
  const recognizedDemos = DEMO_RELEASES.filter(function(release){ return recognizedProjectKeys.has(release.project_key); });
  return recognizedDemos.concat(liveReleases);
}
function productKeys() {
  const keys = [];
  releases.forEach(function(release){ if (!keys.includes(release.project_key)) keys.push(release.project_key); });
  return keys;
}
function activeReleasesForProduct(key) {
  return releases.filter(function(release){ return release.project_key === key && !getBoard(release).archived; });
}
function renderProductTabs() {
  document.getElementById("product-tabs").innerHTML = productKeys().map(function(key){
    const count = activeReleasesForProduct(key).length;
    return '<button class="product-tab' + (key === selectedProductKey ? " active" : "") + '" type="button" data-product-key="' + esc(key) + '"><span class="product-dot"></span><span>' + esc(productLabel(key)) + '</span><span class="product-count">' + count + "</span></button>";
  }).join("");
}
function renderReleaseList() {
  const list = activeReleasesForProduct(selectedProductKey);
  document.getElementById("release-list-title").textContent = productLabel(selectedProductKey) + " releases";
  document.getElementById("release-list").innerHTML = list.length ? list.map(function(release){
    const board = getBoard(release);
    return '<button class="release-item' + (release.id === selectedReleaseId ? " active" : "") + '" type="button" data-release-select="' + esc(release.id) + '"><span class="release-item-top"><strong>' + esc(release.release_number || "Unnamed release") + "</strong>" + statusBadge(board) + '</span><span class="release-item-date">' + esc(formatDate(release.release_date)) + '</span><span class="release-item-meta"><span>' + board.epics.length + " epics</span>" + (isDemo(release) ? '<span class="tag demo-tag">Demo</span>' : '<span class="tag">Live</span>') + "</span></button>";
  }).join("") : '<div class="empty">No active releases. Archived releases are shown at the bottom.</div>';
}
function epicRowHtml(epic,release) {
  const ref = "epic:" + epic.id;
  return '<article class="epic-row"><div class="epic-copy"><div class="epic-key">' + esc(epic.epic_key) + '</div><div class="epic-name">' + esc(epic.title) + "</div></div>" + statusControlHtml(ref,epic,release.id) + ownerSelectHtml(ref,epic,release.id,"Responsible") + '<div class="row-actions"><button class="icon-btn danger" type="button" data-remove-epic="' + esc(epic.id) + '" data-release-id="' + esc(release.id) + '" title="Remove epic from release"><span class="material-symbols-rounded">close</span></button></div>' + confirmationHtml(ref,epic,release.id) + "</article>";
}
function checklistHtml(check,release) {
  const ref = "check:" + check.id;
  const items = check.items.map(function(item){
    const itemRef = "item:" + check.id + ":" + item.id;
    return '<div class="item-row"><div class="item-copy"><div class="item-title-line"><span class="bullet">•</span><span class="item-title" contenteditable="true" data-title-ref="' + esc(itemRef) + '" data-release-id="' + esc(release.id) + '">' + esc(item.title) + "</span></div>" + (item.link ? '<span class="item-link">↗ ' + esc(item.link) + "</span>" : '<span class="item-link empty-link">No evidence attached</span>') + "</div>" + statusControlHtml(itemRef,item,release.id) + ownerSelectHtml(itemRef,item,release.id,"Responsible") + scopeControlHtml(itemRef,item,release.id) + rowActionsHtml(itemRef,release.id,"Remove content","close") + confirmationHtml(itemRef,item,release.id) + "</div>";
  }).join("");
  return '<article class="check-card"><div class="check-head"><div class="check-copy"><span class="control-label">Checklist</span><span class="check-title" contenteditable="true" data-title-ref="' + esc(ref) + '" data-release-id="' + esc(release.id) + '">' + esc(check.title) + "</span></div>" + statusControlHtml(ref,check,release.id) + ownerSelectHtml(ref,check,release.id,"Owner") + rowActionsHtml(ref,release.id,"Remove checklist","delete") + "</div>" + confirmationHtml(ref,check,release.id) + '<div class="check-items">' + items + '</div><div class="check-foot"><button class="link-button" type="button" data-show-add-item="' + esc(check.id) + '">+ Add content</button><span class="tag">' + check.items.length + " item" + (check.items.length === 1 ? "" : "s") + '</span></div><div class="inline-add" data-add-item-form="' + esc(check.id) + '"><input type="text" placeholder="New checklist content"><button class="btn" type="button" data-cancel-add-item="' + esc(check.id) + '">Cancel</button><button class="btn primary" type="button" data-confirm-add-item="' + esc(check.id) + '">Add</button></div></article>';
}
function renderReleaseDetail() {
  const release = selectedRelease();
  const host = document.getElementById("release-detail");
  if (!release) { host.innerHTML = '<div class="empty">Select an active release or restore one from the archive.</div>'; return; }
  const board = getBoard(release);
  const archiveButton = board.status === "done" ? '<button class="btn" id="archive-release" type="button"><span class="material-symbols-rounded">archive</span> Archive</button>' : "";
  host.innerHTML =
    '<div class="detail-header"><div class="detail-heading"><div><h2>' + esc(productLabel(release.project_key)) + " · " + esc(release.release_number) + '</h2><p>' + (isDemo(release) ? "Design sample — release changes stay local" : "Live release — number, date, epics, and completion use the existing Product Releases APIs") + '</p></div><span class="tag ' + (isDemo(release) ? "demo-tag" : "") + '">' + (isDemo(release) ? "Demo" : "Live data") + '</span></div><div class="release-fields"><div class="field"><label for="release-number-input">Release number</label><input id="release-number-input" type="text" value="' + esc(release.release_number) + '"></div><div class="field"><label for="release-date-input">Release date</label><input id="release-date-input" type="date" value="' + esc(release.release_date) + '"></div><button class="btn primary" id="save-release-fields" type="button"><span class="material-symbols-rounded">save</span> Save release</button></div></div>' +
    '<div class="release-status-row"><div><span class="row-label">Release status</span><small>Readiness and lifecycle</small></div>' + statusPickerHtml("release",board,release.id) + releaseLifecycleHtml(release) + '<div class="release-status-actions">' + archiveButton + '</div>' + confirmationHtml("release",board,release.id) + '</div>' + completionPanelHtml(release) +
    '<div class="section-head"><h3>Epics in this release</h3><button class="btn" id="open-epic-picker" type="button"><span class="material-symbols-rounded">add</span> Add epics from database</button></div><div class="epic-list">' + (board.epics.length ? board.epics.map(function(epic){ return epicRowHtml(epic,release); }).join("") : '<div class="empty">No epics assigned to this release.</div>') + "</div>" +
    '<div class="section-head"><h3>Checklist</h3><button class="btn" id="show-add-checklist" type="button"><span class="material-symbols-rounded">add</span> Add checklist</button></div><div class="checklist-list">' + (board.checklists.length ? board.checklists.map(function(check){ return checklistHtml(check,release); }).join("") : '<div class="empty">No checklists yet.</div>') + '<div class="inline-add" id="add-checklist-form"><input id="new-checklist-title" type="text" placeholder="Checklist name, e.g. Deployment"><button class="btn" id="cancel-add-checklist" type="button">Cancel</button><button class="btn primary" id="confirm-add-checklist" type="button">Add</button></div></div>';
}
function renderArchive() {
  const archived = releases.filter(function(release){ return release.project_key === selectedProductKey && getBoard(release).archived; });
  document.getElementById("archive-list").innerHTML = archived.length ? archived.map(function(release){
    return '<div class="archive-item"><span class="material-symbols-rounded">inventory_2</span><span><strong>' + esc(release.release_number) + '</strong><br>' + esc(formatDate(release.release_date)) + '</span><button class="icon-btn" type="button" data-restore-release="' + esc(release.id) + '" title="Restore to active releases"><span class="material-symbols-rounded">unarchive</span></button></div>';
  }).join("") : '<span class="empty">No archived releases for this product.</span>';
}
function renderAll() {
  renderProductTabs();
  renderReleaseList();
  renderReleaseDetail();
  renderArchive();
  bindDynamicControls();
}
function closePopovers() {
  document.querySelectorAll(".status-picker.open,.scope-picker.open").forEach(function(picker){ picker.classList.remove("open"); });
}
function bindDynamicControls() {
  document.querySelectorAll("[data-product-key]").forEach(function(button){
    button.addEventListener("click",function(){
      selectedProductKey = button.dataset.productKey;
      const first = activeReleasesForProduct(selectedProductKey)[0];
      selectedReleaseId = first ? first.id : "";
      renderAll();
    });
  });
  document.querySelectorAll("[data-release-select]").forEach(function(button){ button.addEventListener("click",function(){ selectedReleaseId = button.dataset.releaseSelect; renderAll(); }); });
  document.querySelectorAll("[data-status-trigger]").forEach(function(button){
    button.addEventListener("click",function(event){ event.stopPropagation(); const picker = button.closest(".status-picker"); const willOpen = !picker.classList.contains("open"); closePopovers(); picker.classList.toggle("open",willOpen); });
  });
  document.querySelectorAll(".status-menu").forEach(function(menu){ menu.addEventListener("click",function(event){ event.stopPropagation(); }); });
  document.querySelectorAll("[data-status-value]").forEach(function(button){
    button.addEventListener("click",async function(){
      const picker = button.closest(".status-picker");
      const target = resolveTarget(picker.dataset.statusPicker,picker.dataset.releaseId);
      if (!target) return;
      await applyStatusSelection(picker.dataset.statusPicker,picker.dataset.releaseId,button.dataset.statusValue);
    });
  });
  document.querySelectorAll("[data-toggle-delay]").forEach(function(button){
    button.addEventListener("click",function(){
      const picker = button.closest(".status-picker");
      const target = resolveTarget(picker.dataset.statusPicker,picker.dataset.releaseId);
      if (!target || target.status !== "planned") return;
      target.delayed = !target.delayed;
      saveBoard(picker.dataset.releaseId);
      renderAll();
      toast(target.delayed ? "Planned marked delayed." : "Delayed flag cleared.");
    });
  });
  document.querySelectorAll("[data-owner-ref]").forEach(function(select){
    select.addEventListener("change",function(){ const target = resolveTarget(select.dataset.ownerRef,select.dataset.releaseId); if (target) { target.owner = select.value; saveBoard(select.dataset.releaseId); } });
  });
  document.querySelectorAll("[data-confirm-field]").forEach(function(select){
    select.addEventListener("change",function(){ const target = resolveTarget(select.dataset.targetRef,select.dataset.releaseId); if (target) { target[select.dataset.confirmField] = select.value; saveBoard(select.dataset.releaseId); } });
  });
  document.querySelectorAll("[data-title-ref]").forEach(function(editable){
    editable.addEventListener("keydown",function(event){ if (event.key === "Enter") { event.preventDefault(); editable.blur(); } });
    editable.addEventListener("blur",function(){ const target = resolveTarget(editable.dataset.titleRef,editable.dataset.releaseId); const value = editable.textContent.trim(); if (!target) return; if (value) { target.title = value; saveBoard(editable.dataset.releaseId); } else { renderAll(); } });
  });
  bindScopePickers();
  document.querySelectorAll("[data-details-ref]").forEach(function(button){ button.addEventListener("click",function(){ openDetails(button.dataset.detailsRef,button.dataset.releaseId); }); });
  document.querySelectorAll("[data-delete-ref]").forEach(function(button){ button.addEventListener("click",function(){ if (window.confirm("Remove this board item?")) deleteTarget(button.dataset.deleteRef,button.dataset.releaseId); }); });
  document.querySelectorAll("[data-show-add-item]").forEach(function(button){ button.addEventListener("click",function(){ const form = document.querySelector('[data-add-item-form="' + button.dataset.showAddItem + '"]'); form.classList.add("open"); form.querySelector("input").focus(); }); });
  document.querySelectorAll("[data-cancel-add-item]").forEach(function(button){ button.addEventListener("click",function(){ document.querySelector('[data-add-item-form="' + button.dataset.cancelAddItem + '"]').classList.remove("open"); }); });
  document.querySelectorAll("[data-confirm-add-item]").forEach(function(button){ button.addEventListener("click",function(){ addChecklistContent(button.dataset.confirmAddItem); }); });
  document.querySelectorAll("[data-remove-epic]").forEach(function(button){ button.addEventListener("click",function(){ if (window.confirm("Remove this epic from the release?")) removeEpic(button.dataset.releaseId,button.dataset.removeEpic); }); });
  document.querySelectorAll("[data-restore-release]").forEach(function(button){ button.addEventListener("click",function(){ const board = getBoard(findRelease(button.dataset.restoreRelease)); board.archived = false; saveBoard(button.dataset.restoreRelease); selectedReleaseId = button.dataset.restoreRelease; renderAll(); toast("Release restored."); }); });
  const saveButton = document.getElementById("save-release-fields");
  if (saveButton) saveButton.addEventListener("click",saveReleaseFields);
  const pickerButton = document.getElementById("open-epic-picker");
  if (pickerButton) pickerButton.addEventListener("click",openEpicPicker);
  const archiveButton = document.getElementById("archive-release");
  if (archiveButton) archiveButton.addEventListener("click",archiveSelectedRelease);
  const confirmCompletionButton = document.getElementById("confirm-release-completion");
  if (confirmCompletionButton) confirmCompletionButton.addEventListener("click",completeLiveRelease);
  const cancelCompletionButton = document.getElementById("cancel-release-completion");
  if (cancelCompletionButton) cancelCompletionButton.addEventListener("click",function(){ pendingReleaseCompletionId = ""; renderAll(); });
  const showChecklist = document.getElementById("show-add-checklist");
  if (showChecklist) showChecklist.addEventListener("click",function(){ const form = document.getElementById("add-checklist-form"); form.classList.add("open"); document.getElementById("new-checklist-title").focus(); });
  const cancelChecklist = document.getElementById("cancel-add-checklist");
  if (cancelChecklist) cancelChecklist.addEventListener("click",function(){ document.getElementById("add-checklist-form").classList.remove("open"); });
  const confirmChecklist = document.getElementById("confirm-add-checklist");
  if (confirmChecklist) confirmChecklist.addEventListener("click",addChecklist);
}
async function applyStatusSelection(ref,releaseId,newStatus) {
  const release = findRelease(releaseId);
  const target = resolveTarget(ref,releaseId);
  if (!release || !target) return;
  if (ref === "release" && !isDemo(release)) {
    if (newStatus === "done" && release.release_status !== "released") {
      pendingReleaseCompletionId = release.id;
      closePopovers();
      renderAll();
      const actor = document.getElementById("completion-actor");
      if (actor) actor.focus();
      return;
    }
    if (newStatus !== "done" && release.release_status === "released") {
      await reopenLiveRelease(release,newStatus);
      return;
    }
  }
  target.status = newStatus;
  if (target.status !== "planned") target.delayed = false;
  saveBoard(releaseId);
  renderAll();
  toast("Status updated to " + statusMeta(target.status).label + ".");
}
async function completeLiveRelease() {
  const release = findRelease(pendingReleaseCompletionId);
  if (!release) return;
  const actualDate = document.getElementById("completion-actual-date").value.trim();
  const actor = document.getElementById("completion-actor").value.trim();
  const notes = document.getElementById("completion-notes").value.trim();
  if (!actualDate) { toast("Actual release date is required."); return; }
  if (!actor) { toast("Completed by is required."); return; }
  const button = document.getElementById("confirm-release-completion");
  button.disabled = true;
  try {
    const response = await fetch("/api/product-releases/" + encodeURIComponent(release.id) + "/actions",{
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body:JSON.stringify({action:"released",actual_date:actualDate,actor:actor,notes:notes})
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Failed to complete release.");
    release.release_status = data.release_status;
    release.actions = data.actions || release.actions;
    if (data.release_date) release.release_date = data.release_date;
    const board = getBoard(release);
    board.status = "done";
    board.delayed = false;
    board.lifecycle_status = data.release_status;
    pendingReleaseCompletionId = "";
    saveBoard(release.id);
    renderAll();
    toast("Release marked Released in Product Releases.");
  } catch (error) { toast(error.message); button.disabled = false; }
}
async function reopenLiveRelease(release,newStatus) {
  try {
    const response = await fetch("/api/product-releases/" + encodeURIComponent(release.id) + "/actions",{
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body:JSON.stringify({action:"reverted",actor:"",notes:"Reopened from release readiness board."})
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Failed to reopen release.");
    release.release_status = data.release_status;
    release.actions = data.actions || release.actions;
    const board = getBoard(release);
    board.status = newStatus;
    board.delayed = false;
    board.lifecycle_status = data.release_status;
    saveBoard(release.id);
    renderAll();
    toast("Release reopened in Product Releases and set to " + statusMeta(newStatus).label + ".");
  } catch (error) { toast(error.message); renderAll(); }
}
function bindScopePickers() {
  document.querySelectorAll("[data-scope-trigger]").forEach(function(button){
    button.addEventListener("click",function(event){ event.stopPropagation(); const picker = button.closest(".scope-picker"); const willOpen = !picker.classList.contains("open"); closePopovers(); picker.classList.toggle("open",willOpen); if (willOpen) picker.querySelector(".scope-search").focus(); });
  });
  document.querySelectorAll(".scope-menu").forEach(function(menu){ menu.addEventListener("click",function(event){ event.stopPropagation(); }); });
  document.querySelectorAll(".scope-search").forEach(function(input){
    input.addEventListener("input",function(){ const query = input.value.toLowerCase().trim(); input.closest(".scope-menu").querySelectorAll(".scope-option").forEach(function(option){ option.style.display = option.dataset.scopeSearchText.includes(query) ? "flex" : "none"; }); });
  });
  document.querySelectorAll(".scope-options input[type=checkbox]").forEach(function(checkbox){
    checkbox.addEventListener("change",function(){
      const picker = checkbox.closest(".scope-picker");
      const target = resolveTarget(picker.dataset.scopePicker,picker.dataset.releaseId);
      const board = getBoard(findRelease(picker.dataset.releaseId));
      if (!target) return;
      if (checkbox.value === "release" && checkbox.checked) {
        target.scopes = ["release"];
      } else {
        target.scopes = target.scopes.filter(function(scope){ return scope !== "release"; });
        if (checkbox.checked && !target.scopes.includes(checkbox.value)) target.scopes.push(checkbox.value);
        if (!checkbox.checked) target.scopes = target.scopes.filter(function(scope){ return scope !== checkbox.value; });
        if (!target.scopes.length) target.scopes = ["release"];
      }
      picker.querySelectorAll('input[type="checkbox"]').forEach(function(box){ box.checked = target.scopes.includes(box.value); });
      picker.querySelector(".scope-trigger span:first-child").textContent = scopeSummary(target,board);
      saveBoard(picker.dataset.releaseId);
    });
  });
}
function deleteTarget(ref,releaseId) {
  const board = getBoard(findRelease(releaseId));
  const parts = ref.split(":");
  if (parts[0] === "check") board.checklists = board.checklists.filter(function(check){ return check.id !== parts[1]; });
  if (parts[0] === "item") { const check = board.checklists.find(function(entry){ return entry.id === parts[1]; }); if (check) check.items = check.items.filter(function(item){ return item.id !== parts[2]; }); }
  saveBoard(releaseId);
  renderAll();
  toast("Board item removed.");
}
function addChecklistContent(checkId) {
  const release = selectedRelease();
  const board = getBoard(release);
  const check = board.checklists.find(function(entry){ return entry.id === checkId; });
  const form = document.querySelector('[data-add-item-form="' + checkId + '"]');
  const title = form.querySelector("input").value.trim();
  if (!check || !title) { toast("Enter the checklist content."); return; }
  check.items.push(normalizeEntity({id:uid("item"),title:title,status:"planned",delayed:false,owner:check.owner,scopes:["release"],link:"",notes:""}));
  normalizeScopes(check.items[check.items.length - 1]);
  saveBoard(release.id);
  renderAll();
  toast("Checklist content added.");
}
function addChecklist() {
  const release = selectedRelease();
  const board = getBoard(release);
  const title = document.getElementById("new-checklist-title").value.trim();
  if (!title) { toast("Enter a checklist name."); return; }
  board.checklists.push(normalizeEntity({id:uid("check"),title:title,status:"planned",delayed:false,owner:"Unassigned",items:[],link:"",notes:""}));
  saveBoard(release.id);
  renderAll();
  toast("Checklist added.");
}
async function saveReleaseFields() {
  const release = selectedRelease();
  if (!release) return;
  const releaseNumber = document.getElementById("release-number-input").value.trim();
  const releaseDate = document.getElementById("release-date-input").value.trim();
  if (!releaseNumber) { toast("Release number is required."); return; }
  if (!releaseDate) { toast("Release date is required."); return; }
  if (isDemo(release)) {
    release.release_number = releaseNumber;
    release.release_date = releaseDate;
    renderAll();
    toast("Design release updated locally.");
    return;
  }
  const button = document.getElementById("save-release-fields");
  button.disabled = true;
  try {
    const response = await fetch("/api/product-releases/" + encodeURIComponent(release.id),{
      method:"PUT",
      headers:{"Content-Type":"application/json"},
      body:JSON.stringify({release_number:releaseNumber,release_date:releaseDate})
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Failed to update release.");
    const epics = release.epics;
    const actions = release.actions;
    Object.assign(release,data.release);
    release.epics = epics;
    release.actions = actions;
    renderAll();
    toast("Release number and date updated.");
  } catch (error) { toast(error.message); button.disabled = false; }
}
function availableEpicsForSelectedRelease() {
  const release = selectedRelease();
  if (!release) return [];
  const assigned = new Set((release.epics || []).map(function(epic){ return String(epic.epic_row_id || epic.id); }));
  return epicPool.filter(function(epic){ return epic.project_key === release.project_key && !assigned.has(String(epic.epic_row_id || epic.id)); });
}
function openEpicPicker() {
  const release = selectedRelease();
  if (!release) return;
  pickerSelectedIds = new Set();
  document.getElementById("epic-picker-context").textContent = productLabel(release.project_key) + " · " + release.release_number;
  document.getElementById("epic-picker-search").value = "";
  renderEpicPickerList();
  openModal("epic-picker");
  document.getElementById("epic-picker-search").focus();
}
function renderEpicPickerList() {
  const query = document.getElementById("epic-picker-search").value.toLowerCase().trim();
  const items = availableEpicsForSelectedRelease().filter(function(epic){ return (String(epic.epic_key) + " " + String(epic.epic_name)).toLowerCase().includes(query); });
  document.getElementById("epic-picker-list").innerHTML = items.length ? items.map(function(epic){
    const id = String(epic.epic_row_id || epic.id);
    const assignedElsewhere = epic.release_id && epic.release_id !== selectedReleaseId;
    return '<label class="epic-picker-option"><input type="checkbox" value="' + esc(id) + '"' + (pickerSelectedIds.has(id) ? " checked" : "") + (assignedElsewhere ? " disabled" : "") + '><span><strong>' + esc(epic.epic_key + " · " + (epic.epic_name || "Unnamed epic")) + "</strong><small>" + esc(epic.product_category || epic.component || "Epic") + (assignedElsewhere ? " · already assigned to another release" : "") + "</small></span></label>";
  }).join("") : '<div class="empty">No available database epics match this product and search.</div>';
  document.querySelectorAll("#epic-picker-list input:not(:disabled)").forEach(function(checkbox){
    checkbox.addEventListener("change",function(){ if (checkbox.checked) pickerSelectedIds.add(checkbox.value); else pickerSelectedIds.delete(checkbox.value); updatePickerCount(); });
  });
  updatePickerCount();
}
function updatePickerCount() { document.getElementById("epic-picker-count").textContent = pickerSelectedIds.size + " selected"; }
async function addSelectedEpics() {
  const release = selectedRelease();
  if (!release || !pickerSelectedIds.size) { toast("Select at least one epic."); return; }
  const selected = Array.from(pickerSelectedIds);
  if (isDemo(release)) {
    selected.forEach(function(id){
      const epic = epicPool.find(function(entry){ return String(entry.epic_row_id || entry.id) === id; });
      if (epic) release.epics.push({epic_row_id:id,epic_key:epic.epic_key,epic_name:epic.epic_name});
    });
    syncBoardEpics(release,getBoard(release));
    saveBoard(release.id);
    closeModal("epic-picker");
    renderAll();
    toast(selected.length + " database epic" + (selected.length === 1 ? "" : "s") + " added to the design release.");
    return;
  }
  const button = document.getElementById("add-selected-epics");
  button.disabled = true;
  try {
    for (const epicId of selected) {
      const response = await fetch("/api/product-releases/" + encodeURIComponent(release.id) + "/epics",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({epic_row_id:epicId,epic_type:"new_feature"})
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Failed to add epic.");
    }
    closeModal("epic-picker");
    await refreshLiveData(release.id);
    toast(selected.length + " epic" + (selected.length === 1 ? "" : "s") + " added.");
  } catch (error) { toast(error.message); button.disabled = false; }
}
async function removeEpic(releaseId,epicId) {
  const release = findRelease(releaseId);
  if (!release) return;
  if (isDemo(release)) {
    release.epics = release.epics.filter(function(epic){ return String(epic.epic_row_id || epic.id) !== String(epicId); });
    syncBoardEpics(release,getBoard(release));
    saveBoard(release.id);
    renderAll();
    toast("Epic removed from the design release.");
    return;
  }
  try {
    const response = await fetch("/api/product-releases/" + encodeURIComponent(release.id) + "/epics/" + encodeURIComponent(epicId),{method:"DELETE"});
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Failed to remove epic.");
    await refreshLiveData(release.id);
    toast("Epic removed from release.");
  } catch (error) { toast(error.message); }
}
function archiveSelectedRelease() {
  const release = selectedRelease();
  if (!release) return;
  const board = getBoard(release);
  if (board.status !== "done") { toast("Only Done releases can be archived."); return; }
  board.archived = true;
  saveBoard(release.id);
  const next = activeReleasesForProduct(selectedProductKey)[0];
  selectedReleaseId = next ? next.id : "";
  renderAll();
  toast("Release moved to archive.");
}
function openDetails(ref,releaseId) {
  const target = resolveTarget(ref,releaseId);
  if (!target) return;
  detailTargetRef = ref;
  detailReleaseId = releaseId;
  document.getElementById("drawer-context").textContent = targetLabel(ref,releaseId);
  document.getElementById("detail-link").value = target.link || "";
  document.getElementById("detail-notes").value = target.notes || "";
  openModal("details-drawer");
}
function saveDrawer() {
  const target = resolveTarget(detailTargetRef,detailReleaseId);
  if (!target) return;
  target.link = document.getElementById("detail-link").value.trim();
  target.notes = document.getElementById("detail-notes").value.trim();
  saveBoard(detailReleaseId);
  closeModal("details-drawer");
  renderAll();
  toast("Optional details saved.");
}
function openModal(id) { const modal = document.getElementById(id); modal.classList.add("open"); modal.setAttribute("aria-hidden","false"); }
function closeModal(id) { const modal = document.getElementById(id); modal.classList.remove("open"); modal.setAttribute("aria-hidden","true"); }
async function fetchLiveReleases() {
  try {
    const response = await fetch("/api/product-releases");
    if (!response.ok) throw new Error("Release API unavailable");
    const data = await response.json();
    return (data.releases || []).filter(function(release){ return release.release_status !== "shelved"; });
  } catch (error) { return []; }
}
async function fetchEpicPool() {
  try {
    const response = await fetch("/api/product-releases/epics/pool");
    if (!response.ok) throw new Error("Epic pool unavailable");
    const data = await response.json();
    return data.epics || [];
  } catch (error) { return []; }
}
async function refreshLiveData(preserveReleaseId) {
  const live = await fetchLiveReleases();
  epicPool = await fetchEpicPool();
  releases = combineRecognizedReleases(live);
  selectedReleaseId = preserveReleaseId && findRelease(preserveReleaseId) ? preserveReleaseId : selectedReleaseId;
  renderAll();
}
async function loadAll() {
  const results = await Promise.all([
    fetchLiveReleases(),
    fetchEpicPool(),
    fetch("/api/projects?include_inactive=0").then(function(response){ return response.ok ? response.json() : {projects:[]}; }).catch(function(){ return {projects:[]}; })
  ]);
  epicPool = results[1];
  (results[2].projects || []).forEach(function(project){ projectNames[project.project_key] = project.project_name || project.display_name || project.project_key; });
  releases = combineRecognizedReleases(results[0]);
  selectedProductKey = releases[0] ? releases[0].project_key : "";
  const first = activeReleasesForProduct(selectedProductKey)[0];
  selectedReleaseId = first ? first.id : "";
  renderAll();
}
document.addEventListener("click",closePopovers);
document.getElementById("epic-picker-search").addEventListener("input",renderEpicPickerList);
document.getElementById("add-selected-epics").addEventListener("click",addSelectedEpics);
document.getElementById("drawer-save").addEventListener("click",saveDrawer);
document.querySelectorAll("[data-close-modal]").forEach(function(button){ button.addEventListener("click",function(){ closeModal(button.dataset.closeModal); }); });
document.querySelectorAll(".modal-backdrop").forEach(function(backdrop){ backdrop.addEventListener("click",function(event){ if (event.target === backdrop) closeModal(backdrop.id); }); });
document.addEventListener("keydown",function(event){ if (event.key === "Escape") document.querySelectorAll(".modal-backdrop.open").forEach(function(modal){ closeModal(modal.id); }); });
loadAll();
