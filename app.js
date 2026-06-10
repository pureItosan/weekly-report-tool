/* =====================================================================
   Weekly Report Hub — app.js
   Single-page tool: import reports, map owners, overlay batches, export Word.
   All data persists in localStorage. No backend.
   ===================================================================== */
'use strict';

/* ---------- storage ---------- */
const LS = {
  members: 'wrt_members_v1',
  tasks:   'wrt_tasks_v1',
  batches: 'wrt_batches_v1',
};
const store = {
  load(k, d){ try{ return JSON.parse(localStorage.getItem(k)) ?? d; }catch(e){ return d; } },
  save(k, v){ localStorage.setItem(k, JSON.stringify(v)); },
};

// when the task key scheme changes between versions, stale tasks would not match new keys and pile up
// as duplicates. Bump DATA_VERSION to drop old tasks/batches once (the member roster is kept).
const DATA_VERSION='13';
try{
  if(localStorage.getItem('wrt_dataver')!==DATA_VERSION){
    localStorage.removeItem(LS.tasks); localStorage.removeItem(LS.batches); localStorage.removeItem('wrt_idcode');
    localStorage.setItem('wrt_dataver',DATA_VERSION);
  }
}catch(e){}

let members = store.load(LS.members, []);   // [{id,name,aliases:[]}]
let tasks   = store.load(LS.tasks, []);     // [task]
let batches = store.load(LS.batches, []);   // [{name,date,new,updated,unchanged}]
let idCodeMap = store.load('wrt_idcode', {}); // persistent ID->code map so project keys stay stable across weeks
let projAliases = store.load('wrt_projalias', []); // [{key,label,tokens:[norm]}] user-defined project groupings
let projMeta = store.load('wrt_projmeta', {});   // {projk:{customer,category,desc}} editable project info
let deletedNames = store.load('wrt_deleted', []); // normalized names the user removed -> never auto-re-add
let projMerge = store.load('wrt_projmerge', {});  // {fromProjk: toProjk} user drag-merged / corrected projects
let pendingReports = [];                            // image-pasted report placeholders, IN MEMORY ONLY (never persisted)
let tableSlides = [];                               // slide numbers with a real table (person-section boundaries)
let catalogMember = '';                            // catalog member filter
let editingProj = '';                              // projk currently being edited
let projCats = store.load('wrt_projcats', []);     // user-added project categories (beyond the built-ins)
let catOpen  = store.load('wrt_catopen', null);    // {cat:true/false} remembered expand/collapse of category groups
let catFilter = '';                                // '' = show all categories; else only show this category (tab filter)
let catalogView = store.load('wrt_catview', 'tree');     // projects panel view: 'tree' | 'matrix'
let treeExpand = {};                               // tree: which Customer Project's task branch is expanded (cp -> true)
let taskGroupBy = store.load('wrt_taskgroup', 'member');   // Tasks view grouping: 'member' | 'project' | 'status'
let autoAdd = store.load('wrt_autoadd', true);        // auto-create members from report owners
let fuzzy   = store.load('wrt_fuzzy', true);          // allow nickname/typo/partial matching (default on)
/* member GROUPS — named rosters that remember member order + roles + aliases.
   The active group always mirrors the live `members` array. */
let memberGroups = store.load('wrt_groups', null);
let activeGroup  = store.load('wrt_active_group', '');
if(!Array.isArray(memberGroups) || !memberGroups.length){
  memberGroups = [{name:'SPD RD3-1', members: members.slice()}];   // seed with the current roster
  activeGroup  = 'SPD RD3-1';
  store.save('wrt_groups', memberGroups); store.save('wrt_active_group', activeGroup);
}
if(!memberGroups.some(g=>g.name===activeGroup)) activeGroup = memberGroups[0].name;
const filters = {q:'', project:'', member:'', status:'', role:'', hideEmpty:false};

/* CLOUD (Firebase) state — only active on the online site, never on file:// */
const CLOUD = {
  on:false, db:null,
  memberEmail:'team@spd-rd3-member.app',  // members log in with the team passcode -> member role
  adminEmail :'team@spd-rd3-admin.app',   // admins log in with the admin passcode -> admin role
  admins:['vito','tom','greg','aaron','zach','john'],   // admin names (lowercased)
  authedAs:null,         // 'member' | 'admin' (which account authenticated)
  me:null, ready:false, applying:false, saveTimer:null,
  upImgs:new Set(),      // image ids already uploaded to cloud (avoid re-upload)
  imgsByTask:{}          // taskId -> [{id,data,w,h}] loaded from cloud
};
function isAdminName(n){ return CLOUD.admins.includes(String(n||'').trim().toLowerCase()); }

function persist(){ persistLocal(); cloudSave(); }
function persistLocal(){
  // members + batches are tiny and must always survive; tasks may be large (images)
  try{ store.save(LS.members, members); }catch(e){ console.warn(e); }
  try{ store.save(LS.batches, batches); }catch(e){ console.warn(e); }
  try{ store.save('wrt_idcode', idCodeMap); }catch(e){}
  try{ store.save('wrt_projalias', projAliases); }catch(e){}
  try{ store.save('wrt_projmeta', (typeof CLOUD!=='undefined'&&CLOUD.on)?projMetaNoImg():projMeta); }catch(e){}
  try{ store.save('wrt_deleted', deletedNames); }catch(e){}
  try{ store.save('wrt_projmerge', projMerge); }catch(e){}
  try{ store.save('wrt_projcats', projCats); }catch(e){}
  try{ const ag=memberGroups.find(g=>g.name===activeGroup); if(ag) ag.members=members.slice();
       store.save('wrt_groups', memberGroups); store.save('wrt_active_group', activeGroup); }catch(e){}
  // in cloud mode, images live in Firestore -> keep localStorage light (text only)
  try{
    const toSave = (typeof CLOUD!=='undefined' && CLOUD.on)
      ? tasks.map(t=>{ const c=Object.assign({},t); delete c.images; return c; })
      : tasks;
    store.save(LS.tasks, toSave);
  }catch(e){ console.warn('tasks persist failed', e);
    toast('⚠ Too many tasks(mostly images)exceeded the browser storage limit; members kept, tasks not saved this time.'); }
}

/* downscale an image dataURL to a JPEG. 1600px keeps schematics / measurement tables
   sharp. In cloud mode images live in Firestore (not localStorage), so size is fine;
   on the offline desktop, localStorage degrades gracefully if it overflows. */
function shrinkImage(dataUrl, max=1600, q=0.85){
  return new Promise(res=>{
    const img=new Image();
    img.onload=()=>{
      let w=img.naturalWidth||img.width, h=img.naturalHeight||img.height;
      if(!w||!h){ res({data:dataUrl,w:480,h:320}); return; }
      if(Math.max(w,h)>max){ const s=max/Math.max(w,h); w=Math.round(w*s); h=Math.round(h*s); }
      const c=document.createElement('canvas'); c.width=w; c.height=h;
      const ctx=c.getContext('2d');
      ctx.fillStyle='#ffffff'; ctx.fillRect(0,0,w,h);           // flatten transparency for JPEG
      ctx.drawImage(img,0,0,w,h);
      try{ res({data:c.toDataURL('image/jpeg',q),w,h}); }
      catch(e){ res({data:dataUrl,w,h}); }
    };
    img.onerror=()=>res({data:dataUrl,w:480,h:320});
    img.src=dataUrl;
  });
}
// Crisp-but-safe: keep an image as high-res as possible while guaranteeing the encoded data URL stays
// under `budget` chars, so a Firestore image doc never exceeds the 1 MB/doc limit (base64 ≈ 1.37× JPEG).
function shrinkImageBudget(dataUrl, maxPx=1800, q0=0.85, budget=900000){
  return new Promise(res=>{
    const img=new Image();
    img.onload=()=>{
      const W=img.naturalWidth||img.width, H=img.naturalHeight||img.height;
      if(!W||!H){ res({data:dataUrl,w:480,h:320}); return; }
      let scale = Math.max(W,H)>maxPx ? maxPx/Math.max(W,H) : 1;
      const render=(sc,q)=>{
        const cw=Math.max(1,Math.round(W*sc)), ch=Math.max(1,Math.round(H*sc));
        const c=document.createElement('canvas'); c.width=cw; c.height=ch;
        const ctx=c.getContext('2d'); ctx.imageSmoothingEnabled=true; ctx.imageSmoothingQuality='high';
        ctx.fillStyle='#ffffff'; ctx.fillRect(0,0,cw,ch); ctx.drawImage(img,0,0,cw,ch);
        let data; try{ data=c.toDataURL('image/jpeg',q); }catch(e){ data=dataUrl; }
        return {data,w:cw,h:ch};
      };
      const qs=[q0,0.8,0.72,0.64,0.56,0.5]; let best=render(scale,q0);
      for(let pass=0; pass<6; pass++){
        for(const q of qs){ best=render(scale,q); if(best.data.length<=budget){ res(best); return; } }
        scale*=0.82;                                            // still over budget -> reduce dimensions, retry
        if(Math.max(W,H)*scale < 480){ res(best); return; }     // already small; accept best effort
      }
      res(best);
    };
    img.onerror=()=>res({data:dataUrl,w:480,h:320});
    img.src=dataUrl;
  });
}

/* ---------- helpers ---------- */
const $  = s => document.querySelector(s);
const $$ = s => Array.from(document.querySelectorAll(s));
const uid = () => Math.random().toString(36).slice(2,10);
const esc = s => String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function toast(msg){ const t=$('#toast'); t.textContent=msg; t.hidden=false; clearTimeout(t._t); t._t=setTimeout(()=>t.hidden=true,2800); }
/* ---- perf diagnostics (temporary): when a click is slow, toast WHERE the time went; click the version badge to copy the log ---- */
const PERF={ev:[], lt:[]};
try{ new PerformanceObserver(l=>l.getEntries().forEach(e=>{ PERF.lt.push(Math.round(e.duration)); if(PERF.lt.length>30) PERF.lt.shift(); })).observe({type:'longtask'}); }catch(e){}
try{ new PerformanceObserver(l=>l.getEntries().forEach(e=>{
    if(e.name!=='click') return;
    const rec={n:PERF.ev.length+1, view:(typeof currentView!=='undefined'?currentView:''),
      wait:Math.round(e.processingStart-e.startTime), js:Math.round(e.processingEnd-e.processingStart),
      paint:Math.round(e.startTime+e.duration-e.processingEnd), total:Math.round(e.duration)};
    PERF.ev.push(rec); if(PERF.ev.length>30) PERF.ev.shift();
    if(rec.total>350) toast(`⏱ slow click #${rec.n}: wait ${rec.wait} · js ${rec.js} · paint ${rec.paint} ms (tap the version badge to copy the log)`);
  })).observe({type:'event', durationThreshold:100});
}catch(e){}
const norm = s => String(s||'').toLowerCase().replace(/[^a-z0-9一-鿿]+/g,'').trim();

/* =====================================================================
   MEMBERS
   ===================================================================== */
const ROLES=['Leader','EE','RF','ANT','SW','PM','QA'];   // disciplines; first 4 are the asked-for ones
function normRole(s){ const n=String(s||'').trim().toUpperCase(); const hit=ROLES.find(r=>r.toUpperCase()===n); return hit||''; }
function parseMemberText(text){
  const out=[];
  text.split(/[\r\n]+/).forEach(line=>{
    line=line.trim(); if(!line) return;
    // optional trailing role tag:  "David [RF]"  or  "David #RF"
    let role=''; const rt=line.match(/[\[#]\s*([A-Za-z]{1,8})\s*\]?\s*$/);
    if(rt && normRole(rt[1])){ role=normRole(rt[1]); line=line.replace(/[\[#]\s*[A-Za-z]{1,8}\s*\]?\s*$/,'').trim(); }
    // Chinese name: a CJK run anywhere on the line  ("Aaron 王小明" / "Aaron 王小明 [EE]")
    let cname=''; const cm=line.match(/[㐀-鿿]{1,12}/);
    if(cm){ cname=cm[0].trim(); line=line.replace(cm[0],' ').replace(/\s+/g,' ').trim(); }
    // CSV style "name,alias,alias" OR "name: alias, alias"
    let name, aliasStr='';
    if(/[::]/.test(line)){ const p=line.split(/[::]/); name=p[0].trim(); aliasStr=p.slice(1).join(':'); }
    else if(line.includes(',')){ const p=line.split(','); name=p[0].trim(); aliasStr=p.slice(1).join(','); }
    else name=line;
    if(!name && cname){ name=cname; cname=''; }   // a Chinese-only line -> that IS the name
    if(!name) return;
    const aliases=aliasStr.split(/[,, , /]/).map(s=>s.trim()).filter(Boolean);
    out.push({name, aliases, role, cname});
  });
  return out;
}
function addMembers(list, opts){
  opts=opts||{};
  list.forEach(m=>{
    if(opts.manual) deletedNames=deletedNames.filter(n=>n!==norm(m.name));  // explicit re-add clears blocklist
    const ex=members.find(x=>norm(x.name)===norm(m.name));
    if(ex){ // merge aliases, do not duplicate
      m.aliases.forEach(a=>{ if(!ex.aliases.some(b=>norm(b)===norm(a))) ex.aliases.push(a); });
      if(m.role && !ex.role) ex.role=m.role;
      if(m.cname && !ex.cname) ex.cname=m.cname;
    } else {
      members.push({id:uid(), name:m.name, aliases:m.aliases||[], role:m.role||'', role2:m.role2||'', cname:m.cname||''});
    }
  });
  reresolveAllTasks();   // membership changed -> re-assign owners on existing tasks
  persist(); renderAll();
}
function deleteMember(id){
  const m=members.find(x=>x.id===id); if(m) deletedNames=[...new Set([...deletedNames, norm(m.name)])]; // remember -> don't auto re-add
  members=members.filter(x=>x.id!==id); reresolveAllTasks(); persist(); renderAll();
}
function setMemberRole(id, role){ const m=members.find(x=>x.id===id); if(m){ m.role=normRole(role); persist(); renderAll(); } }
function setMemberRole2(id, role){ const m=members.find(x=>x.id===id); if(m){ m.role2=normRole(role); persist(); renderAll(); } }
function setMemberCname(id, val){ const m=members.find(x=>x.id===id); if(m){ m.cname=String(val||'').trim(); persist(); renderAll(); } }
function memberDisplay(m){ return m ? (m.cname? m.name+' · '+m.cname : m.name) : ''; }   // "Aaron · 王小明"
function memberRole(id){ const m=members.find(x=>x.id===id); return m?(m.role||''):''; }
function memberRoles(id){ const m=members.find(x=>x.id===id); return m?[m.role,m.role2].filter(Boolean):[]; }

// recompute every task's owner assignment from its raw owner/reporter against current members
function reresolveAllTasks(){
  tasks.forEach(t=>{ if(t.manualOwners) return;   // keep user's manual add/remove
    const r=resolveOwners(t.rawOwner, t.reporter); t.ownerIds=r.ids; t.unmatched=r.unmatched; t.shared=r.shared; });
}

// add every owner/reporter found in the report as a member (only names that don't already match)
function autoAddFromReport(){
  const seen=new Set(), toAdd=[];
  tasks.forEach(t=>{
    [...splitOwners(t.rawOwner), ...splitOwners(t.reporter)].forEach(tok=>{
      if(tok.length<2) return;
      if(!/^[A-Za-z][A-Za-z.\- ]*$/.test(tok)) return;           // names only
      if(/^[A-Z]{2,4}$/.test(tok)) return;                        // skip acronyms like GMI/SW
      if(matchOwner(tok)) return;                                 // already maps to a member
      const k=norm(tok);
      if(deletedNames.includes(k)) return;                        // user removed this person -> don't re-add
      if(seen.has(k)) return; seen.add(k);
      toAdd.push(tok);
    });
  });
  if(toAdd.length) addMembers(toAdd.map(n=>({name:n, aliases:[]})));  // addMembers re-resolves + renders
  return toAdd;
}
function clearMembers(){ if(confirm('Clear the entire member list?(tasks are not deleted)')){ members=[]; deletedNames=[]; persist(); renderAll(); } }

/* =====================================================================
   NAME MATCHING  (owner-priority, multi-owner, fuzzy)
   ===================================================================== */
function splitOwners(raw){
  if(!raw) return [];
  // drop parentheticals like "Jin(SW:Jonas)" -> "Jin"; split on / , & newline ,  + and whitespace
  return String(raw).replace(/\([^)]*\)/g,' ')
    .split(/[\/,&\n, +]|\band\b|\s+/i).map(s=>s.trim()).filter(s=>s.length>=2);
}
function sharedPrefix(a,b){ let i=0; while(i<a.length&&i<b.length&&a[i]===b[i]) i++; return i; }
function lev(a,b){
  const m=a.length,n=b.length,d=Array.from({length:m+1},(_,i)=>[i,...Array(n).fill(0)]);
  for(let j=0;j<=n;j++) d[0][j]=j;
  for(let i=1;i<=m;i++)for(let j=1;j<=n;j++)
    d[i][j]=Math.min(d[i-1][j]+1,d[i][j-1]+1,d[i-1][j-1]+(a[i-1]===b[j-1]?0:1));
  return d[m][n];
}
// returns member id or null for one owner token
function matchOwner(token){
  const t=norm(token); if(!t) return null;
  const cands=members.map(m=>({m,keys:[norm(m.name),...m.aliases.map(norm)].filter(Boolean)}));
  // 1) exact name / alias  (always)
  for(const c of cands) if(c.keys.includes(t)) return c.m.id;
  if(!fuzzy) return null;        // strict mode: only full name or defined alias counts
  // 2) score every member by its best-matching key (nickname / typo)
  const scored=[];
  for(const c of cands){
    let best=-1;
    for(const k of c.keys){
      let s=-1; const short=Math.min(t.length,k.length);
      if(t.length>=3 && (k.startsWith(t)||t.startsWith(k))) s=100+short;          // Sam→Samuel
      else if(short>=3 && (k.includes(t)||t.includes(k))) s=90+short;            // "John"⊂"John Yang"
      else{
        const sp=sharedPrefix(t,k);
        if(sp>=3 && sp>=Math.ceil(short*0.6)) s=60+sp;            // Ravi→Raveendra (shared "rav")
        else if(lev(t,k)<=1 && short>=4) s=58;                    // single typo
        else if(sp>=4) s=56;                                      // share first 4 letters
      }
      if(s>best) best=s;
    }
    scored.push({id:c.m.id, s:best});
  }
  scored.sort((a,b)=>b.s-a.s);
  if(!scored.length || scored[0].s<55) return null;
  // ambiguity guard: don't guess between two near-equal fuzzy candidates
  if(scored[0].s<100 && scored[1] && scored[0].s-scored[1].s<5) return null;
  return scored[0].id;
}
// returns {ids:[], unmatched:[rawToken], shared:bool}
function resolveOwners(rawOwner, reporter){
  const toks=splitOwners(rawOwner);
  const ids=[], unmatched=[];
  toks.forEach(tk=>{ const id=matchOwner(tk); if(id){ if(!ids.includes(id)) ids.push(id);} else unmatched.push(tk); });
  // owner-priority: only fall back to reporter if NO owner tokens at all
  if(ids.length===0 && toks.length===0 && reporter){
    const id=matchOwner(reporter); if(id) ids.push(id);
  }
  return {ids, unmatched, shared: (toks.length>1)};
}

/* =====================================================================
   FILE IMPORT
   ===================================================================== */
const FIELD_MAP = [
  ['reporter','reporter'],
  ['projectname','project'],['project','project'],
  ['currentjobandissue','current'],['currentjob','current'],['thisweek','current'],['currentstatus','current'],
  ['risk','risk'],
  ['duedate','due'],['due','due'],
  ['owner','owner'],['assignee','owner'],
  ['nextweekjobandplan','next'],['nextweek','next'],['nextweekplan','next'],['plan','next'],
  ['nextstepin714days','next'],['nextstep','next'],
  ['mitigationplanoverdue','mitigation'],['mitigationplan','mitigation'],['mitigation','mitigation'],
  ['complexity','complexity'],
  ['progress','progress'],['completion','progress'],
];
function canonField(label){
  const n=norm(label);
  for(const [k,v] of FIELD_MAP) if(n===k) return v;
  for(const [k,v] of FIELD_MAP) if(n.includes(k)) return v;
  return null;
}

async function importFiles(fileList){
  const files=Array.from(fileList);
  for(const f of files){
    const ext=f.name.split('.').pop().toLowerCase();
    let parsed=[];
    try{
      if(ext==='pptx') parsed=await parsePPTX(f);
      else if(ext==='docx') parsed=await parseDOCX(f);
      else if(ext==='xlsx') parsed=await parseXLSX(f);
      else if(ext==='csv'||ext==='txt') parsed=await parseTextTable(await f.text());
      else { toast('Unsupported format: '+ext); continue; }
    }catch(err){ console.error(err); toast('Parse failed: '+f.name+' — '+err.message); continue; }
    if(!parsed.length){ toast('No parsable tasks in '+f.name); continue; }
    // image-pasted report placeholders stay IN MEMORY (not persisted -> no localStorage bloat)
    const reports=parsed.filter(t=>t._imageReport);
    pendingReports.push(...reports);
    overlayTasks(parsed.filter(t=>!t._imageReport), f.name);
  }
  if(autoAdd){ const added=autoAddFromReport(); if(added.length) toast('Auto-added '+added.length+' members from reports'); }
  dedupeTasks(); cleanupGarbledMembers();
  renderAll();
  if(pendingReports.length){
    if(navigator.onLine){
      toast(`Detected ${pendingReports.length} image pages — OCR running in the background…`);
      setTimeout(()=>ocrAllReports(true), 400);          // auto-run in background
    } else {
      toast(`Detected ${pendingReports.length} image pages — go online and click "🔍 OCR image reports" to recognise`);
    }
  }
}

/* ---------- PPTX ---------- */
function localTags(root, local){
  // namespace-safe element collection by localName
  return Array.from(root.getElementsByTagName('*')).filter(e=>e.localName===local);
}
function cellText(tc){
  // preserve paragraph breaks (stacked owners / multi-line jobs are separate <a:p>)
  const paras=localTags(tc,'p').map(p=>localTags(p,'t').map(t=>t.textContent).join(''));
  return paras.join('\n').replace(/[ \t]+/g,' ').replace(/\n{2,}/g,'\n').trim();
}
function projectCode(project){
  // leading identifier used to merge "current" and "next-week" rows of the same project
  let s=String(project||'').split(/[(\n]/)[0].trim();
  const tok=s.split(/\s+/)[0];
  return norm(tok||s);
}
function projTokens(s){ return String(s||'').toLowerCase().split(/[^a-z0-9.]+/).filter(x=>x.length>=2); }
function tokenOverlap(a,b){ const sa=new Set(projTokens(a)); let n=0; projTokens(b).forEach(x=>{ if(sa.has(x)) n++; }); return n; }
const RISK_ORDER={High:3,Medium:2,Low:1,'':0};
function maxByRisk(arr){ return arr.filter(Boolean).sort((a,b)=>RISK_ORDER[b]-RISK_ORDER[a])[0]||''; }
function uniqJoin(arr, sep){ return [...new Set(arr.map(x=>(x||'').trim()).filter(Boolean))].join(sep); }

const CODE_RE=/B\d{2}[A-Z]\d{3}[A-Z]?\d{0,2}|SDX[-\s]?\d{2}(?:[\/\s, ,-]\d{2})?|VB\d{3}/i;
const ID_RE=/\bID\s?\d{2,4}\b/i;
function codeKey(code){
  if(/sdx/i.test(code)){ const nums=(code.match(/\d{2}/g)||[]).map(Number).sort((a,b)=>a-b); return 'c:sdx'+nums.join(''); } // SDX82/85 == SDX85/82
  return 'c:'+norm(code).replace(/t0+$/,'');   // B01W036T00 -> b01w036 ; keeps T01/T03 distinct
}
// Tasks stay GRANULAR (one per report row). These functions only GROUP them by project.
function learnIdCode(list){   // remember ID<->code so B01W043 / ID535 group, stable across weeks
  list.forEach(t=>{ const s=String(t.project||''); const c=(s.match(CODE_RE)||[])[0], id=(s.match(ID_RE)||[])[0];
    if(c&&id){ const ik=norm(id); if(!idCodeMap[ik]) idCodeMap[ik]=codeKey(c); } });
}
// canonical project key: user alias group > code (T00-stripped) > ID(mapped to code) > first token
function projKeyOf(project){
  const s=String(project||''), n=norm(s);
  for(const g of projAliases){ if(g.tokens.some(tok=>tok && n.includes(tok))) return 'a:'+g.key; }
  const c=(s.match(CODE_RE)||[])[0]; if(c) return codeKey(c);          // B-code / SDX / VB
  const id=(s.match(ID_RE)||[])[0]; if(id){ const ik=norm(id); return idCodeMap[ik]||('id:'+ik); }
  if(/patent/i.test(s)) return 'n:patent';                            // keep Patent as its own bucket
  return 'n:general';                                                 // any project without a code -> General
}
// clean project label e.g. "B01W043 / ID535"
function projLabelOf(project){
  const s=String(project||'');
  for(const g of projAliases){ if(g.tokens.some(tok=>tok && norm(s).includes(tok))) return g.label; }
  let c=(s.match(CODE_RE)||[])[0], id=(s.match(ID_RE)||[])[0];
  if(c) c=c.toUpperCase().replace(/T0+$/,'');     // B01W043T00 -> B01W043 (keep T01/T03)
  if(c&&id) return c+' / '+id.toUpperCase().replace(/\s/g,'');
  if(c) return c;
  const lbl=(s.split(/[\n(]/)[0].trim().slice(0,30));
  return lbl||'General';
}
// stable per-task signature so re-import of the same item is Unchanged (not a duplicate)
function descSig(text){
  return norm(String(text||'')
    .replace(/\d+(\.\d+)?\s*%/g,'')
    .replace(/\b(done|pass|fail|ongoing|on going|in progress|wip|pending|hold|closed)\b/gi,''))
    .slice(0,46);
}
// parse a single <a:tbl> into an array of field-objects (one per data row)
function parseTable(tbl){
  const rows=localTags(tbl,'tr').map(tr=>localTags(tr,'tc').map(cellText));
  if(!rows.length) return [];
  // vertical 2-column "field | value" layout (my synthetic template / some decks)
  const vertHits=rows.filter(r=>r.length>=2 && canonField(r[0])).length;
  if(rows[0].length===2 && vertHits>=Math.max(2,rows.length*0.5)){
    const f={};
    rows.forEach(r=>{ const k=canonField(r[0]); if(k) f[k]=r[1]; });
    return (f.project||f.current||f.owner) ? [f] : [];
  }
  // horizontal layout: row0 is a header that maps columns -> fields
  const colMap=rows[0].map(canonField);
  if(colMap.filter(Boolean).length<2) return [];
  const out=[];
  for(let i=1;i<rows.length;i++){
    const r=rows[i]; const f={};
    r.forEach((cell,ci)=>{ const k=colMap[ci]; if(k && cell) f[k]=cell; });
    if(f.project||f.current||f.owner||f.next) out.push(f);
  }
  return out;
}
async function parsePPTX(file){
  const zip=await JSZip.loadAsync(await file.arrayBuffer());
  const slidePaths=Object.keys(zip.files)
    .filter(p=>/^ppt\/slides\/slide\d+\.xml$/.test(p))
    .sort((a,b)=>(+a.match(/\d+/))-(+b.match(/\d+/)));
  const out=[];
  tableSlides=[];                          // slide numbers that have a real table (= person section boundary)
  let lastTaskIdxBySlide=-1, lastTaskSlide=-1, lastSectionStart=0;
  const lastSlideNo = slidePaths.length ? +String(slidePaths[slidePaths.length-1]).match(/slide(\d+)/)[1] : 0;
  for(const sp of slidePaths){
    const xml=await zip.file(sp).async('string');
    const doc=new DOMParser().parseFromString(xml,'application/xml');
    const slideNo=+sp.match(/slide(\d+)/)[1];

    // ---- slide-level reporter label ("Reporter: Aaron" — may be split across runs) ----
    const allText=localTags(doc,'t').map(t=>t.textContent.replace(/\s+/g,' '));
    const joinedText=allText.join(' ');
    let slideReporter='';
    const rm=joinedText.match(/reporter\s*[::]\s*([A-Za-z][\w.]*(?:\s*[\/&,]\s*[A-Za-z][\w.]*)*)/i);
    if(rm) slideReporter=rm[1].trim();
    // closing / divider slide (Thank You, 感謝聆聽, The End…) — its decorative image must NOT
    // be attached to the last member's section.
    const isClosing = /thank\s*you|感謝聆聽|感謝指教|敬請指教|簡報結束|the\s+end|q\s*&\s*a/i.test(joinedText)
      || (slideNo===lastSlideNo && joinedText.replace(/\s/g,'').length<30);

    // ---- images on this slide (keep an original-URL for hi-res OCR if needed) ----
    const relPath=sp.replace(/slides\/(slide\d+)\.xml/,'slides/_rels/$1.xml.rels');
    const images=[]; const origs=[];
    if(zip.file(relPath)){
      const relXml=await zip.file(relPath).async('string');
      const relDoc=new DOMParser().parseFromString(relXml,'application/xml');
      for(const rel of Array.from(relDoc.getElementsByTagName('Relationship'))){
        if((rel.getAttribute('Type')||'').includes('/image')){
          let tgt=rel.getAttribute('Target').replace(/^\.\.\//,'ppt/');
          const mf=zip.file(tgt);
          if(mf){
            const b64=await mf.async('base64');
            const mime=tgt.endsWith('.png')?'image/png':tgt.endsWith('.gif')?'image/gif':'image/jpeg';
            const url=`data:${mime};base64,${b64}`;
            const s=await shrinkImageBudget(url, 1800);   // crisper, still cloud-safe
            images.push({id:uid(), data:s.data, w:s.w, h:s.h});
            origs.push(url);
          }
        }
      }
    }

    // ---- all tables on the slide -> field-objects ----
    let rowObjs=[];
    localTags(doc,'tbl').forEach(tbl=>{ rowObjs=rowObjs.concat(parseTable(tbl)); });

    // ---- fallback: "Field: value" text blocks if no tables ----
    if(!rowObjs.length){
      const f={};
      allText.forEach(line=>{ const m=line.match(/^([A-Za-z ]{3,40})[::]\s*(.+)$/); if(m){ const k=canonField(m[1]); if(k) f[k]=m[2].trim(); } });
      if(f.project||f.current||f.owner) rowObjs.push(f);
    }

    // ---- emit ONE task per "Current Job and Issue" cell; only split when sub-items are tagged with DIFFERENT people ----
    const slideTasks=[];
    rowObjs.forEach(f=>{
      if(!(f.project||f.current||f.owner||f.next)) return;
      if(slideReporter && !f.reporter) f.reporter=slideReporter;
      const items = f.current ? splitSubItems(f.current) : [];
      // group sub-items by their tagged owner; untagged items belong to the row owner
      const byOwner=new Map();
      items.forEach(line=>{ const {owner}=extractItemOwner(line); const k=owner||'__row__'; (byOwner.get(k)||byOwner.set(k,[]).get(k)).push(line); });
      const taggedOwners=[...byOwner.keys()].filter(k=>k!=='__row__');
      if(taggedOwners.length>0){
        // different people inside one cell -> one task per person (their lines combined)
        byOwner.forEach((lines,k)=>{ slideTasks.push(makeTaskFromFields({...f, current:lines.join('\n'), owner: k==='__row__'? f.owner : k}, slideNo)); });
      } else {
        // single owner -> ONE task for the whole Current Job and Issue cell (no fragmentation)
        slideTasks.push(makeTaskFromFields(f, slideNo));
      }
    });

    if(slideTasks.length){
      slideTasks.forEach(t=>{ t._images=images.slice(); out.push(t); });
      lastTaskIdxBySlide=out.length-1; lastTaskSlide=slideNo; lastSectionStart=out.length-slideTasks.length;
      tableSlides.push(slideNo);              // a new person's section starts here
    } else if(images.length){
      const big=images.some(im=>(im.w||0)>=360 && (im.h||0)>=150);
      // (a) SMALL detail images right after a member's report -> attach to that member.
      //     Big images are NOT attached here: a big image may be SOMEONE ELSE's pasted report
      //     (another person's report screenshot pasted after a member's table) or a "Thank You"
      //     page — attaching it to the preceding member is wrong. Big images go to OCR (b),
      //     which attributes them correctly.
      if(!isClosing && !big){
        if(lastTaskSlide>=0 && (slideNo-lastTaskSlide)<=4 && out[lastSectionStart]){
          out[lastSectionStart]._images=(out[lastSectionStart]._images||[]).concat(images);
        } else if(lastTaskIdxBySlide>=0){
          out[lastTaskIdxBySlide]._images=(out[lastTaskIdxBySlide]._images||[]).concat(images);
        }
      }
      // (b) big image -> hi-res OCR placeholder. OCR decides: "Reporter: X" -> X's own report;
      //     no Reporter -> a detail image for the member it FOLLOWS (carried in _afterReporter).
      //     Closing/divider ("Thank You") slides are skipped entirely.
      if(big && !isClosing){
        const hi=[];
        // keep a HIGH-res copy for OCR (in-memory only, never persisted) — sharper source = far fewer garbles.
        for(let k=0;k<images.length;k++){ const s=await shrinkImage(origs[k]||images[k].data, 2600, 0.92); hi.push({id:uid(), data:s.data, w:s.w, h:s.h}); }
        const ph=makeTaskFromFields({project:'Image report', current:'image report slide '+slideNo}, slideNo);
        ph._images=hi; ph._imageReport=true;
        ph._afterReporter = (lastTaskSlide>=0 && out[lastSectionStart] && out[lastSectionStart].reporter) || '';
        out.push(ph);
      }
    }
  }
  return out;
}
// a multi-line "Current Job" cell often holds several sub-items, each tagged "(Name, status)".
// split them so each becomes its own task assigned to the tagged person.
const STATUS_WORDS=/^(done|on-?going|ongoing|pending|wip|in[\s-]?progress|pass(?:ed)?|fail(?:ed)?|disqualified|n\/a|hold|closed|complete[d]?|tbd|new|ok|cont(?:inue)?|wait(?:ing)?)$/i;
function splitSubItems(current){
  const raw=String(current||'').split(/\n+/).map(s=>s.trim()).filter(Boolean);
  const items=[];
  raw.forEach(line=>{
    // merge obvious continuation lines into the previous item
    if(items.length && (/^[a-z(\-]/.test(line) || line.length<6 || /^(temp|p\/n|pn|update|063\.)/i.test(line)))
      items[items.length-1]+=' '+line;
    else items.push(line);
  });
  return items;
}
function extractItemOwner(line){
  const m=line.match(/\(([^)]+)\)\s*$/);          // trailing "(Name[, status])"
  if(!m) return {owner:''};
  const cand=(m[1].split(',')[0]||'').trim();     // the name part, before any ", status"
  if(!cand || STATUS_WORDS.test(cand.replace(/\s+/g,''))) return {owner:''};
  // accept only if every token looks like a real person (a known member, or a Capitalized word, no digits)
  const toks=cand.split(/[\/&]|\band\b/i).map(s=>s.trim()).filter(Boolean);
  const ok=toks.length>0 && toks.every(t=>
    !/\d/.test(t) && t.length<=14 &&
    (members.some(m=>norm(m.name)===norm(t)||(m.aliases||[]).some(a=>norm(a)===norm(t))) || /^[A-Z][a-z]{1,}$/.test(t)));
  return ok ? {owner:cand} : {owner:''};
}
function topicLabel(text){  // a short title from a task's text (for items with no project code)
  const s=String(text||'').split('\n')[0].replace(/^\d+[.)]\s*/,'').trim();
  return s.split(/\s+/).slice(0,7).join(' ').slice(0,42) || '';
}
function makeTaskFromFields(f, slideNo){
  const current=(f.current||'').trim();
  const mit=(f.mitigation||'').trim();
  let project=(f.project||'').trim();
  if(!project) project = topicLabel(current) || topicLabel(f.next) || 'Untitled';   // label by topic, not "Untitled Project"
  return {
    project,
    reporter:(f.reporter||'').trim(),
    rawOwner:(f.owner||f.reporter||'').trim(),
    current: mit ? (current+(current?' ':'')+'Mitigation: '+mit) : current,
    risk:normRisk(f.risk)||'Medium',
    due:(f.due||'').trim(),
    next:(f.next||'').trim(),
    complexity:normRisk(f.complexity)||inferComplexity(current),
    progress:parseProgress(f.progress, current),
    _slide:slideNo, _images:[],
  };
}
function inferComplexity(text){
  const t=(text||'').toLowerCase();
  if(/redesign|root.?cause|debug|bring.?up|architecture|migrat/.test(t)) return 'High';
  if(/study|monitor|sync|update|document|review/.test(t)) return 'Low';
  return 'Medium';
}
function normRisk(v){
  if(!v) return '';
  const n=norm(v);
  if(n==='h'||n.includes('high')) return 'High';
  if(n==='m'||n.includes('med')) return 'Medium';
  if(n==='l'||n.includes('low')) return 'Low';
  return '';
}
function parseProgress(p, text){
  if(p){ const m=String(p).match(/\d+/); if(m) return Math.min(100,+m[0]); }
  const t=(text||'').toLowerCase();
  // explicit percentages win — use the LAST one (usually the overall status line). handle decimals (89.47%)
  const pcts=[...t.matchAll(/(\d{1,3})(?:\.\d+)?\s*%/g)].map(m=>Math.min(100,Math.round(+m[0].replace('%',''))));
  if(pcts.length) return pcts[pcts.length-1];
  const hasOpen=/\bongoing\b|on[\s-]?going|in[\s-]?progress|\bwip\b|to\s?do|todo|next|will\b|plan to|preparing|under\b/.test(t);
  if(/\b(done|pass|passed|completed|complete|finished|closed)\b|✓|✔/.test(t) && !hasOpen) return 100;
  if(/\b(block|blocked|not start|n\/a|hold|stuck|fail|failed)\b/.test(t)) return 15;
  if(/\b(pending|waiting|await)\b/.test(t)) return 25;
  if(hasOpen) return 55;
  return 40;
}

/* ---------- DOCX ---------- */
async function parseDOCX(file){
  const zip=await JSZip.loadAsync(await file.arrayBuffer());
  const xml=await zip.file('word/document.xml').async('string');
  const doc=new DOMParser().parseFromString(xml,'application/xml');
  const paras=localTags(doc,'p').map(p=>localTags(p,'t').map(t=>t.textContent).join('').trim());
  // split into blocks by blank line; each block -> task by Field: value
  const blocks=[]; let cur=[];
  paras.forEach(line=>{ if(!line){ if(cur.length){blocks.push(cur);cur=[];} } else cur.push(line); });
  if(cur.length) blocks.push(cur);
  const out=[];
  blocks.forEach(b=>{
    const f={};
    b.forEach(line=>{ const m=line.match(/^([A-Za-z ]{3,40})[::]\s*(.+)$/); if(m){const k=canonField(m[1]); if(k)f[k]=m[2].trim();} });
    if(f.project||f.current||f.owner) out.push(makeTaskFromFields(f,0));
  });
  return out;
}

/* ---------- XLSX ---------- */
async function parseXLSX(file){
  const wb=XLSX.read(await file.arrayBuffer(),{type:'array'});
  const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
  return rowsToTasks(rows);
}
function rowsToTasks(rows){
  const out=[];
  rows.forEach(r=>{
    const f={};
    Object.keys(r).forEach(col=>{ const k=canonField(col); if(k) f[k]=r[col]; });
    if(f.project||f.current||f.owner) out.push(makeTaskFromFields(f,0));
  });
  return out;
}
/* ---------- CSV / TXT ---------- */
async function parseTextTable(text){
  // try CSV with header
  const lines=text.split(/\r?\n/).filter(l=>l.trim());
  if(lines.length>1 && lines[0].includes(',') && lines.some(l=>canonField(l.split(',')[0]))===false){
    const headers=lines[0].split(',').map(s=>s.trim());
    const rows=lines.slice(1).map(l=>{ const c=splitCsv(l); const o={}; headers.forEach((h,i)=>o[h]=c[i]||''); return o; });
    const t=rowsToTasks(rows); if(t.length) return t;
  }
  // else treat as Field: value blocks
  const blocks=[]; let cur=[];
  lines.concat(['']).forEach(line=>{ if(!line.trim()){ if(cur.length){blocks.push(cur);cur=[];} } else cur.push(line); });
  const out=[];
  blocks.forEach(b=>{ const f={}; b.forEach(line=>{ const m=line.match(/^([A-Za-z ]{3,40})[::]\s*(.+)$/); if(m){const k=canonField(m[1]); if(k)f[k]=m[2].trim();} }); if(f.project||f.current||f.owner) out.push(makeTaskFromFields(f,0)); });
  return out;
}
function splitCsv(line){ const out=[]; let cur='',q=false; for(const ch of line){ if(ch==='"'){q=!q;} else if(ch===','&&!q){out.push(cur);cur='';} else cur+=ch; } out.push(cur); return out.map(s=>s.trim().replace(/^"|"$/g,'')); }

/* =====================================================================
   TASK OVERLAY  (New / Updated / Unchanged)
   ===================================================================== */
function taskKey(project, rawOwner, occ){
  return norm(project)+'||'+norm(rawOwner)+'||'+occ;
}
function signature(t){
  return [norm(t.current),norm(t.next),t.risk,t.due,t.progress].join('~');
}
function overlayTasks(parsed, sourceName){
  learnIdCode(parsed);                          // remember ID<->code mappings first
  let nNew=0,nUpd=0,nUnc=0; const touched=[];
  parsed.forEach(p=>{                            // GRANULAR: one task per report item
    const res=resolveOwners(p.rawOwner, p.reporter);
    const projk=projKeyOf(p.project);
    // key must be stable across imports -> use RAW owner tokens, not resolved member ids (which change as roster grows)
    const ownerSetK=splitOwners(p.rawOwner).map(norm).sort().join('+') || norm(p.reporter) || 'unassigned';
    const key=projk+'|'+ownerSetK+'|'+descSig(p.current||p.next);
    const inc={
      key, projk, projectLabel:projLabelOf(p.project),
      project:p.project, reporter:p.reporter, rawOwner:p.rawOwner,
      ownerIds:res.ids, unmatched:res.unmatched, shared:res.ids.length>1,
      current:p.current, next:p.next, risk:p.risk||'Medium', due:p.due,
      complexity:p.complexity, progress:p.progress, nextProgress:(p.nextProgress!=null?p.nextProgress:null),
      images:p._images||p.images||[],
      analysis:generateAnalysis(p), source:sourceName, imageReport:p._imageReport||false,
    };
    const ex=tasks.find(t=>t.key===key);
    if(!ex){
      inc.status='New'; inc.id=uid(); inc.delta=''; inc.history=[];
      tasks.push(inc); nNew++; touched.push(inc);
    } else if(signature(ex)!==signature(inc)){
      ex.history=ex.history||[];
      ex.history.push({current:ex.current,next:ex.next,risk:ex.risk,progress:ex.progress,due:ex.due,at:Date.now()});
      inc.delta=computeDelta(ex, inc);
      inc.prev={current:ex.current,progress:ex.progress,risk:ex.risk};
      inc.id=ex.id; inc.status='Updated'; inc.history=ex.history;
      inc.images=(inc.images.length?inc.images:ex.images);
      if(ex.manualOwners){ inc.ownerIds=ex.ownerIds; inc.unmatched=ex.unmatched; inc.shared=ex.shared; inc.manualOwners=true; } // keep manual owners
      if(ex.manualEdit){ inc.risk=ex.risk; inc.complexity=ex.complexity; inc.progress=ex.progress; inc.nextProgress=ex.nextProgress; inc.manualEdit=true; } // keep manual risk/cx/%
      if(ex.manualText){ inc.current=ex.current; inc.next=ex.next; inc.analysis=ex.analysis; inc.manualText=true; }              // keep manual text edits
      Object.assign(ex, inc);
      nUpd++; touched.push(ex);
    } else {
      ex.status='Unchanged'; ex.source=sourceName; nUnc++; touched.push(ex);
    }
  });
  batches.unshift({name:sourceName,date:new Date().toISOString().slice(0,16).replace('T',' '),
    nnew:nNew,nupd:nUpd,nunc:nUnc});
  persist();
  toast(`Imported ${sourceName}:New ${nNew} · Updated ${nUpd} · Unchanged ${nUnc}`);
  return touched;
}
// remove near-duplicate tasks under the same person (e.g. same item OCR'd twice with different garbling)
function taskCore(t){
  return String(t.current||t.next||'').toLowerCase()
    .replace(/^[\s\d.)]*[8b][\dolwvi.t]{3,14}\s*/i,'')   // drop a leading (possibly garbled) project code
    .replace(/[^a-z0-9 ]+/g,' ').replace(/\s+/g,' ').trim();  // keep word boundaries for tokenizing
}
function scoreTask(t){
  let s=0;
  if(/^c:|^id:/.test(t.projk||'')) s+=20;                 // a real project code beats "Other/General"
  if(t.manualEdit||t.manualOwners) s+=15;                 // keep what the user touched
  if((t.images||[]).length) s+=3;
  s+=Math.min(8,(t.current||'').length/30);
  return s;
}
function headTokens(t){ return (taskCore(t).match(/[a-z0-9]{3,}/g)||[]).slice(0,6); }   // first 6 significant words
function dedupeTasks(){
  const groups={};
  tasks.forEach(t=>{ if(t.imageReport) return; const k=(t.ownerIds||[]).slice().sort().join('+')||'__un'; (groups[k]=groups[k]||[]).push(t); });
  const remove=new Set();
  Object.values(groups).forEach(arr=>{
    for(let i=0;i<arr.length;i++){
      if(remove.has(arr[i].id)) continue;
      const ha=headTokens(arr[i]); if(ha.length<3) continue;
      for(let j=i+1;j<arr.length;j++){
        if(remove.has(arr[j].id)) continue;
        const hb=headTokens(arr[j]); if(hb.length<3) continue;
        const sb=new Set(hb); let m=0; ha.forEach(x=>{ if(sb.has(x)) m++; });
        if(m/Math.min(ha.length,hb.length) >= 0.7){          // same opening words -> same item (different OCR garbling)
          const keep=scoreTask(arr[i])>=scoreTask(arr[j])?arr[i]:arr[j];
          remove.add((keep===arr[i]?arr[j]:arr[i]).id);
          if(keep===arr[j]){ break; }
        }
      }
    }
  });
  if(remove.size){ tasks=tasks.filter(t=>!remove.has(t.id)); persist(); }
  return remove.size;
}
// remove a stray 0-task member that's one typo away from a real member (e.g. OCR "Riek" vs "Rick")
function cleanupGarbledMembers(){
  const {map}=buildBuckets();
  const cnt=id=>(map.get(id)||[]).length;
  const drop=members.filter(m=> cnt(m.id)===0 && norm(m.name).length>=4 &&
    members.some(o=>o.id!==m.id && cnt(o.id)>0 && lev(norm(m.name),norm(o.name))<=1));
  drop.forEach(m=>{ deletedNames=[...new Set([...deletedNames, norm(m.name)])]; });
  if(drop.length){ const ids=new Set(drop.map(m=>m.id)); members=members.filter(m=>!ids.has(m.id)); reresolveAllTasks(); persist(); }
  return drop.length;
}
function computeDelta(oldT,newT){
  const d=[];
  if(oldT.progress!==newT.progress) d.push(`Progress ${oldT.progress}% → ${newT.progress}%`);
  if(oldT.risk!==newT.risk) d.push(`Risk ${oldT.risk} → ${newT.risk}`);
  if(oldT.due!==newT.due && newT.due) d.push(`Due ${oldT.due||'—'} → ${newT.due}`);
  if(norm(oldT.current)!==norm(newT.current)) d.push('Work description updated');
  if(norm(oldT.next)!==norm(newT.next)) d.push('Next week plan updated');
  return d.join('；');
}
function deleteTask(id){
  const t=tasks.find(x=>x.id===id); if(!t) return false;
  if(!confirm('Delete this task?\n'+((t.current||t.next||'(no description)').slice(0,60))+'\nThis cannot be undone.')) return false;
  (t.images||[]).forEach(im=>cloudDeleteImage(im.id));        // sync-remove its cloud images too
  tasks=tasks.filter(x=>x.id!==id); persist(); renderAll(); toast('Task deleted'); return true;
}
function resetTasks(){ if(confirm('Clear all tasks? (members are kept)')){ tasks=[]; batches=[]; persist(); renderAll(); } }
// 一鍵清除內容:清掉所有任務, 批次與圖片(雲端圖片也一併刪除), 成員名單保留.雙重確認避免誤按.
function clearAllContent(){
  if(!tasks.length && !batches.length){ toast('Nothing to clear'); return; }
  if(!confirm('Clear all: this deletes every task and image (members are kept).\nThis cannot be undone. Proceed?')) return;
  if(!confirm('Confirm again: really clear all report content?')) return;
  tasks.forEach(t=>(t.images||[]).forEach(im=>cloudDeleteImage(im.id)));   // 同步刪除雲端圖片
  tasks=[]; batches=[]; pendingReports=[]; updateOcrBtn();
  persist(); renderAll();
  toast('All content cleared (members kept)');
}

/* =====================================================================
   PROFESSIONAL REWRITE + ANALYSIS
   ===================================================================== */
const PHRASES={
  'Debug issue':'Investigated and debugged the reported defect; root-cause analysis is in progress.',
  'Verification':'Performed verification and regression testing to confirm the expected behaviour.',
  'Document study':'Conducted documentation and specification study to align the implementation.',
  'Blocked by vendor':'Currently blocked pending a vendor deliverable; a mitigation plan has been prepared.',
  'Code review':'Completed code review and incorporated the resulting feedback.',
  'Integration test':'Executed integration testing across the affected modules.',
  'Root cause found':'Identified the root cause and implemented the corresponding fix.',
  'Waiting on data':'Awaiting upstream data/inputs before proceeding to the next stage.',
};
function rewriteProfessional(text){
  let s=String(text||'').replace(/\s*\n\s*/g,'; ').trim();
  if(!s) return '';
  // casual -> professional substitutions
  const subs=[
    [/\bwaiting on\b/gi,'pending delivery of'],
    [/\bstill seen\b/gi,'still observed'],
    [/\bfix\b/gi,'resolve'],
    [/\bcan'?t\b/gi,'cannot'],
    [/\bok\b/gi,'verified'],
    [/\bcheck\b/gi,'validate'],
    [/\btodo\b/gi,'pending action'],
  ];
  subs.forEach(([re,r])=>{ s=s.replace(re,r); });
  // capitalize sentences, ensure terminal period
  s=s.split(/(?<=[.;])\s+/).map(seg=>{ seg=seg.trim(); return seg? seg[0].toUpperCase()+seg.slice(1):seg; }).join(' ');
  if(!/[.!?]$/.test(s)) s+='.';
  return s;
}
function generateAnalysis(t){
  const txt=((t.current||'')+' '+(t.next||'')).toLowerCase();
  const bits=[];
  if(/block|blocked|vendor|license|procure/.test(txt))
    bits.push('Progress is gated by an external dependency; escalation and a fallback plan are recommended to protect the schedule.');
  if(/timeout|crash|fail|defect|bug|error|drop|ack/.test(txt))
    bits.push('A technical defect is under active investigation; root-cause isolation and a targeted regression are the priority.');
  if(/verif|test|regression/.test(txt))
    bits.push('The task is in the verification phase; test coverage should confirm stability before sign-off.');
  if(/document|spec|study|pkce|oauth/.test(txt))
    bits.push('Specification alignment is in focus to reduce downstream rework.');
  const riskNote = t.risk==='High'
    ? 'Risk is High — close monitoring and a contingency are advised.'
    : t.risk==='Low' ? 'Risk is Low and the task is tracking to plan.'
    : 'Risk is Medium; standard monitoring applies.';
  bits.push(riskNote);
  if(bits.length<=1) bits.unshift('Work is progressing as planned with no major blockers reported.');
  return bits.join(' ');
}
function isClosed(t){ return (+t.progress||0)>=100; }   // coerce: progress may arrive as a string ("100")
function statusLine(t){
  const p=t.progress;
  if(p>=100) return 'Closed (100%)';
  if(p<=10) return 'Not started / blocked';
  if(t.risk==='High') return `In progress (${p}%) — at risk`;
  return `In progress (${p}%)`;
}
function progClass(v){ return v>=100?'p-done':v>=70?'p-high':v>=34?'p-mid':'p-low'; }
function statusWord(t){
  if(isClosed(t)) return 'Closed';
  const x=((t.current||'')+' '+(t.next||'')).toLowerCase();
  if(/block|stuck|\bhold\b|fail/.test(x)) return 'Blocked';
  if(/pending|waiting|await|vendor/.test(x)) return 'Pending';
  return 'In-progress';
}

// project "progress" = that project's share of the member's total tasks (sums to 100%).
// groups by RESOLVED project key so merged projects (e.g. B01W046 == B01W046/ID506) count as one.
// e.g. Tom: B01V038 1, ID535 1, ID515 1, ID506 2 (total 5) -> 20/20/20/40%
function projShares(list){
  const projkOf=t=>resolveProjk(t.projk||t.key)||(t.projectLabel||shortProj(t.project));
  const cnt={}, order=[], labelOf={};
  (list||[]).forEach(t=>{ const k=projkOf(t); if(!(k in cnt)){ cnt[k]=0; order.push(k); labelOf[k]=t.projectLabel||shortProj(t.project); } cnt[k]++; });
  const total=(list||[]).length||1, pctK={};
  order.forEach(k=>pctK[k]=Math.round(cnt[k]/total*100));
  return {
    list: order.map(k=>({label:labelOf[k], pct:pctK[k]})),
    share: t=>pctK[projkOf(t)]||0
  };
}
/* ---------- Weekly narrative (plain text, used by preview + mirrors Word export) ---------- */
function memberNarrative(name, list){
  const lbl=t=>t.projectLabel||shortProj(t.project);
  let out='*'+name+'\n';
  const cur=list.filter(t=>t.current), nexts=list.filter(t=>t.next);
  if(!cur.length && !nexts.length) return out+'Pending input — no items reported this week.\n\n';
  const sh=projShares(list);
  out+='This week: [ '+sh.list.map(o=>`${o.label} - ${o.pct}%`).join(' | ')+' ]\n';   // per-member project breakdown (kept)
  cur.forEach((t,i)=>{
    out+=`${i+1}. ${lbl(t)}: ${rewriteProfessional(t.current)}\n`;
    out+=`   Status: ${statusWord(t)} | risk ${(t.risk||'M')[0]} | Due Date: ${t.due||'TBC'}\n`;
    if(t.shared) out+=`   Shared owner: ${(t.ownerIds||[]).map(memberName).join(', ')}\n`;
  });
  if(nexts.length){ out+='Next week:\n'; nexts.forEach((t,i)=>{ out+=`${i+1}. ${lbl(t)}: ${rewriteProfessional(t.next)}\n`; }); }
  return out+'\n';
}
function buildNarrative(memberIds){
  const {map,unassigned}=buildBuckets();
  const targets = memberIds&&memberIds.length ? members.filter(m=>memberIds.includes(m.id)) : members.slice();
  let out='Weekly Report — '+new Date().toISOString().slice(0,10)+'\n\n';
  targets.forEach(m=>{ out+=memberNarrative(m.name, map.get(m.id)); });
  if((!memberIds||!memberIds.length) && unassigned.length) out+=memberNarrative('Unassigned', unassigned);
  return out;
}

/* =====================================================================
   RENDER
   ===================================================================== */
function memberName(id){ const m=members.find(x=>x.id===id); return m?m.name:'?'; }

function visibleTasks(){ return tasks.filter(t=>!t.imageReport); }   // hide un-OCR'd image placeholders
function buildBuckets(){
  // member id -> tasks[], plus Unassigned
  const map=new Map(); members.forEach(m=>map.set(m.id,[]));
  const unassigned=[];
  visibleTasks().forEach(t=>{
    if(t.ownerIds && t.ownerIds.length){
      t.ownerIds.forEach(id=>{ if(map.has(id)) map.get(id).push(t); else unassigned.push(t); });
    } else unassigned.push(t);
  });
  return {map, unassigned};
}

function renderAll(){
  renderGroups(); renderMembers(); renderBatches(); renderStats(); renderHighlights(); renderCatalog(); renderCharts(); renderTeam(); renderFilters(); renderMembersArea(); renderWorkbenchSelect(); updateOcrBtn();
}
function renderFilters(){
  const ps=$('#filterProject'); if(ps){
    const groups=projectGroups();                       // canonical (merged) projects, by clean label
    if(filters.project && !groups.some(g=>g.projk===filters.project)) filters.project='';
    ps.innerHTML='<option value="">All projects</option>'+groups.map(g=>`<option value="${esc(g.projk)}" ${filters.project===g.projk?'selected':''}>${esc(g.label)}</option>`).join('');
  }
  const ms=$('#filterMember'); if(ms){
    ms.innerHTML='<option value="">All members</option>'+members.map(m=>`<option value="${m.id}" ${filters.member===m.id?'selected':''}>${esc(m.name)}</option>`).join('')+
      `<option value="__un__" ${filters.member==='__un__'?'selected':''}>Unassigned</option>`;
  }
  const chk=$('#autoAddChk'); if(chk) chk.checked=autoAdd;
  const fz=$('#fuzzyChk'); if(fz) fz.checked=fuzzy;
  const he=$('#hideEmptyChk'); if(he) he.checked=filters.hideEmpty;
  const fs=$('#filterStatus'); if(fs) fs.value=filters.status;
  const fr=$('#filterRole'); if(fr) fr.value=filters.role;
  const tg=$('#taskGroupSel'); if(tg) tg.value=taskGroupBy;
}

function roleOptions(sel){ return '<option value="">—</option>'+ROLES.map(r=>`<option ${sel===r?'selected':''}>${r}</option>`).join(''); }
function roleBadge(role, sub){ return role?`<span class="role-badge r-${role}${sub?' sub':''}">${esc(role)}</span>`:''; }
function memberRoleBadges(id){ return memberRoles(id).map((r,i)=>roleBadge(r,i>0)).join(''); }
function renderMembers(){
  $('#memberCount').textContent=members.length;
  $('#memberChips').innerHTML=members.map(m=>`
    <li class="mchip role-${m.role||'none'}" draggable="true" data-mid="${m.id}">
      <div class="mchip-top">
        <span class="drag-handle" title="Drag to reorder">⠿</span>
        <span class="mname">${esc(m.name)}</span>
        ${m.cname?`<span class="cname" title="Nickname">${esc(m.cname)}</span>`:''}
        ${roleBadge(m.role)}${roleBadge(m.role2,true)}
        ${m.aliases.length?`<span class="alias">${esc(m.aliases.join(', '))}</span>`:''}
        <button class="mdel" title="Delete" data-del-member="${m.id}">✕</button>
      </div>
      <div class="mchip-roles">
        <label>Main<select class="role-mini" data-setrole="${m.id}">${roleOptions(m.role)}</select></label>
        <label>Sub<select class="role-mini sub" data-setrole2="${m.id}">${roleOptions(m.role2)}</select></label>
        <input class="cname-mini" data-setcname="${m.id}" value="${esc(m.cname||'')}" placeholder="Nickname" maxlength="16" title="Nickname">
      </div>
    </li>`).join('');
  wireMemberDrag();
}

/* ---------- MEMBER GROUPS (named rosters: order + roles + aliases) ---------- */
function syncActiveGroup(){ const g=memberGroups.find(x=>x.name===activeGroup); if(g) g.members=members.slice(); }
function renderGroups(){
  const sel=$('#groupSelect'); if(!sel) return;
  sel.innerHTML=memberGroups.map(g=>`<option value="${esc(g.name)}"${g.name===activeGroup?' selected':''}>${esc(g.name)} · ${(g.members||[]).length} members</option>`).join('');
}
function saveGroup(){ syncActiveGroup(); persist(); renderGroups(); toast('Group saved: '+activeGroup+' ('+members.length+' members)'); }
function switchGroup(name){
  if(name===activeGroup || !memberGroups.some(g=>g.name===name)) return;
  syncActiveGroup();                                   // save current roster into its group
  activeGroup=name;
  const g=memberGroups.find(x=>x.name===name);
  members=(g.members||[]).map(m=>({...m, aliases:(m.aliases||[]).slice()}));
  reresolveAllTasks(); persist(); renderAll();
  toast('Switched to group: '+name);
}
function addGroup(name, copyCurrent){
  name=String(name||'').trim(); if(!name) return false;
  if(memberGroups.some(g=>g.name===name)){ toast('A group with that name already exists'); return false; }
  syncActiveGroup();
  memberGroups.push({name, members: copyCurrent? members.slice() : []});
  activeGroup=name; members = copyCurrent? members.slice() : [];
  reresolveAllTasks(); persist(); renderAll(); return true;
}
function createGroup(){
  const name=(prompt('New group name (e.g. SPD RD3-2):','')||'').trim(); if(!name) return;
  const copy=confirm('Copy current members into the new group?\n\nOK = copy current list  |  Cancel = create an empty list');
  if(addGroup(name, copy)) toast('Group created: '+name);
}
function renameGroup(){
  const g=memberGroups.find(x=>x.name===activeGroup); if(!g) return;
  const name=(prompt('Rename group:', g.name)||'').trim(); if(!name||name===g.name) return;
  if(memberGroups.some(x=>x.name===name)){ toast('A group with that name already exists'); return; }
  g.name=name; activeGroup=name; persist(); renderGroups(); toast('Renamed to: '+name);
}
function deleteGroup(){
  if(memberGroups.length<=1){ toast('At least one group must remain'); return; }
  if(!confirm('Delete group "'+activeGroup+'"?\n(removes only this roster setting; task data is unaffected)')) return;
  memberGroups=memberGroups.filter(g=>g.name!==activeGroup);
  activeGroup=memberGroups[0].name;
  members=(memberGroups[0].members||[]).map(m=>({...m, aliases:(m.aliases||[]).slice()}));
  reresolveAllTasks(); persist(); renderAll(); toast('Group deleted');
}

function renderBatches(){
  $('#batchList').innerHTML = batches.length? batches.slice(0,8).map(b=>`
    <li><span>${esc(b.name)}<br><small>${b.date}</small></span>
    <span><b style="color:var(--new)">${b.nnew}N</b> <b style="color:var(--updated)">${b.nupd}U</b> ${b.nunc}=</span></li>`).join('')
    : '<li>No imports yet</li>';
}
function renderStats(){            // KPI counts now live in the top tabs (no separate stat-card row)
  const vt=visibleTasks();
  const master=projectGroups().filter(g=>(projMeta[g.projk]||{}).master).length;
  const projCount=master || new Set(vt.map(t=>resolveProjk(t.projk||t.key))).size;
  const setN=(id,n)=>{ const e=$('#'+id); if(e) e.textContent=n; };
  setN('tnProjects', projCount); setN('tnTasks', vt.length); setN('tnTeam', members.length);
}
function navStat(view, flt){ filters.status=flt||''; setView(view); renderFilters(); renderMembersArea(); renderStats(); }

/* ---------- PROJECT CATALOG (editable) ---------- */
function cleanDesc(s){ return String(s||'').replace(/\s*\n\s*/g,'; ').replace(/\s+/g,' ').trim().slice(0,140); }
function resolveProjk(pk){ let k=pk, seen=0; while(projMerge[k] && seen++<20) k=projMerge[k]; return k; }
function projectGroups(){
  const map=new Map();
  visibleTasks().forEach(t=>{ const k=resolveProjk(t.projk||t.key); let g=map.get(k);
    if(!g){ const lbl = k==='n:general'?'General' : k==='n:patent'?'Patent' : (t.projectLabel||shortProj(t.project));
      g={projk:k, label:lbl, tasks:[], mem:new Set(), closed:0, high:0, projStr:t.project}; map.set(k,g); }
    g.tasks.push(t); (t.ownerIds||[]).forEach(id=>g.mem.add(id));
    if(isClosed(t)) g.closed++; if(t.risk==='High'&&!isClosed(t)) g.high++;
    if((t.project||'').length>(g.projStr||'').length){ g.projStr=t.project; if(k!=='n:general'&&k!=='n:patent') g.label=t.projectLabel||g.label; }
  });
  // include master-list projects (manually added / pasted) so the official catalog shows even with 0 tasks;
  // where a master shares a key with task-derived tasks, prefer the master's official code as the label.
  Object.keys(projMeta).forEach(k=>{ const m=projMeta[k]; if(!m||!m.master) return;
    const rk=resolveProjk(k);
    if(map.has(rk)){ if(m.code) map.get(rk).label=m.code; }
    else map.set(rk, {projk:rk, label:(m.code||k.replace(/^[a-z]+:/,'').toUpperCase()), tasks:[], mem:new Set(), closed:0, high:0, projStr:m.code||'', master:true});
  });
  return [...map.values()].sort((a,b)=>b.tasks.length-a.tasks.length);
}
function mergeProjects(fromProjk, toProjk){
  if(!fromProjk||!toProjk||fromProjk===toProjk) return;
  if(resolveProjk(toProjk)===fromProjk) return;          // avoid cycles
  projMerge[fromProjk]=toProjk; persist(); renderCatalog(); renderStats();
  toast('Merged project into: '+(projectGroups().find(g=>g.projk===resolveProjk(toProjk))||{}).label);
}
function unmergeProject(projk){ delete projMerge[projk]; Object.keys(projMerge).forEach(k=>{ if(projMerge[k]===projk) delete projMerge[k]; }); persist(); renderCatalog(); }
const BASE_CATS=['Module','IDU','ODU','Dongle'];               // built-ins (General is always the catch-all, kept last)
// full category list = built-ins + user-added (projCats), with General always last
function allCats(){
  const seen=new Set(), out=[];
  BASE_CATS.forEach(c=>{ seen.add(c); out.push(c); });
  (projCats||[]).forEach(c=>{ c=String(c||'').trim(); if(c && c!=='General' && !seen.has(c)){ seen.add(c); out.push(c); } });
  out.push('General'); return out;
}
function addCategory(name){
  name=String(name||'').trim(); if(!name) return;
  if(allCats().some(c=>c.toLowerCase()===name.toLowerCase())){ toast('Category "'+name+'" already exists'); return; }
  projCats.push(name); persist(); renderCatalog(); toast('Category added: '+name);
}
function renameCategory(oldName, newName){
  newName=String(newName||'').trim(); if(!newName||newName===oldName) return;
  if(BASE_CATS.includes(oldName)||oldName==='General'){ toast('Built-in categories cannot be renamed'); return; }
  const i=projCats.findIndex(c=>c===oldName); if(i<0) return;
  if(allCats().some(c=>c.toLowerCase()===newName.toLowerCase())){ toast('Category "'+newName+'" already exists'); return; }
  projCats[i]=newName;
  Object.keys(projMeta).forEach(k=>{ if(projMeta[k] && projMeta[k].category===oldName) projMeta[k].category=newName; });  // move projects over
  if(catOpen && oldName in catOpen){ catOpen[newName]=catOpen[oldName]; delete catOpen[oldName]; store.save('wrt_catopen',catOpen); }
  persist(); renderCatalog(); toast('Category renamed to: '+newName);
}
function deleteCategory(name){
  if(BASE_CATS.includes(name)||name==='General'){ toast('Built-in categories cannot be deleted'); return; }
  const used=projectGroups().filter(g=>projCategory(g)===name).length;
  if(!confirm('Delete category "'+name+'"?'+(used?('its '+used+' project(s) will move back to General.'):'')+'\n(no projects or tasks are deleted)')) return;
  projCats=projCats.filter(c=>c!==name);
  Object.keys(projMeta).forEach(k=>{ if(projMeta[k] && projMeta[k].category===name) projMeta[k].category=''; });  // back to auto/General
  persist(); renderCatalog(); toast('Category deleted: '+name);
}
function setProjCategory(projk, cat){
  const m=projMeta[projk]||{}; m.category=(cat==='General'? '' : cat);   // '' = auto-infer (lands in General if nothing matches)
  projMeta[projk]=m; persist(); renderCatalog(); renderStats();
}
function deleteProjectGroup(projk){
  const rk=resolveProjk(projk);
  const g=projectGroups().find(x=>x.projk===rk); if(!g) return;
  const tn=g.tasks.length;
  if(!confirm('Delete project "'+g.label+'"'+(tn?('and its '+tn+' tasks'):'(no tasks)')+'"? This cannot be undone.')) return;
  if(tn && !confirm('Confirm again: really delete "'+g.label+'" and its tasks?')) return;
  const ids=new Set(g.tasks.map(t=>t.id));
  tasks.forEach(t=>{ if(ids.has(t.id)) (t.images||[]).forEach(im=>cloudDeleteImage(im.id)); });  // sync-delete cloud images
  tasks=tasks.filter(t=>!ids.has(t.id));
  delete projMeta[g.projk]; delete projMeta[projk];                 // also drop its master record
  const mo=$('#projEditModal'); if(mo) mo.hidden=true;
  persist(); renderAll(); toast('Project deleted: '+g.label);
}
/* ---------- PROJECT MASTER LIST (manually added / pasted, stored in your cloud — never in the repo) ---------- */
const PHASES=['','Kickoff','ES','EVT','DVT','PVT','MP','Done','Hold'];   // '' = 未設定階段
function projMasterKey(code){                          // key a master project the SAME way tasks are keyed, so they auto-link
  let k=projKeyOf(code);
  if(k==='n:general'||k==='n:patent'){ const nn=norm(code); if(nn) k='c:'+nn; }   // codeless names get their own bucket
  return k;
}
function mergeField(a,b){                               // union of "/"-separated values, de-duped ("A" + "B" -> "A / B")
  a=String(a||'').trim(); b=String(b||'').trim(); if(!b) return a; if(!a) return b;
  const out=[]; (a+' / '+b).split(/\s*\/\s*/).forEach(x=>{ x=x.trim(); if(x && !out.some(y=>y.toLowerCase()===x.toLowerCase())) out.push(x); });
  return out.join(' / ');
}
// Accepts the full project MATRIX (Customer | Product Type | Customer Project | Component | Model | Chipset),
// separated by | or TAB. Merged cells (blank leading columns on continuation rows) are filled down.
// Falls back to the legacy 5-col format (Code | Customer | Category | Chipset | Description).
function parseProjectList(text){
  const lines=String(text||'').split('\n').filter(l=>l.trim());
  if(!lines.length) return [];
  const split=l=>l.split(/\t|\|/).map(s=>s.trim());
  const isModelCode=s=>/^[A-Z]?\d{2,}[A-Z.\d]/i.test(String(s||'').trim());   // B01W025T02 / 95.3823T00
  const HCELL=/^(customer|客戶|code|代號|product\s*type|customer\s*project|product\s*category|model|chipset)$/i;
  const isHeader=line=>{ const p=split(line); return p.length>=3 && HCELL.test((p[0]||'').trim()); };   // exact first-cell match (so "Customer A" data rows survive)
  const firstData=split(lines.find(l=>!isHeader(l))||'');
  const looksMatrix=/product\s*type|customer\s*project|product\s*category/i.test(lines[0])
                    || (firstData.length>=5 && !isModelCode(firstData[0]));
  if(looksMatrix){
    const carry={cust:'',type:'',cp:''}; const out=[];
    lines.forEach(line=>{
      if(isHeader(line)) return;          // header row
      let [cust,type,cp,comp,model,chip]=split(line).map(s=>s||'');
      if(cust) carry.cust=cust; else cust=carry.cust;                                // fill down merged cells
      if(type) carry.type=type; else type=carry.type;
      if(cp&&cp!=='-') carry.cp=cp; else if(cp==='-'){ carry.cp=''; cp=''; } else cp=carry.cp;
      if(!model && comp && isModelCode(comp)){ model=comp; comp=''; }                // tolerate rows with no component col
      model=model.trim(); if(!model) return;
      const code=model.split('/')[0].trim(); if(!code) return;
      out.push({code, customer:cust, category:type, custProj:cp, component:comp, chipset:(chip||'').replace(/^chipset\s*[:：]\s*/i,''), desc:model});
    });
    return out;
  }
  return lines.map(line=>{                                                           // legacy format
    if(/^(代號|code|customer|客戶)\b/i.test(line)) return null;
    const p=split(line); if(p.length<2 || !p[0]) return null;
    return {code:p[0], customer:p[1]||'', category:p[2]||'', chipset:p[3]||'', desc:p[4]||''};
  }).filter(Boolean);
}
function importProjectList(text){
  const rows=parseProjectList(text);
  if(!rows.length){ toast('No projects parsed. Format: Code | Customer | Category | Chipset | Description'); return; }
  // detect variants that share a base code (B01W025.00/.01/.02) so they become SEPARATE cards;
  // a unique code keeps its base key so it still auto-links to the matching tasks.
  const base={}, cnt={}; rows.forEach(r=>{ const b=projMasterKey(r.code); base[r.code]=b; cnt[b]=(cnt[b]||0)+1; });
  const keys=new Set();
  rows.forEach(r=>{
    const k = cnt[base[r.code]]>1 ? ('c:'+norm(r.code)) : base[r.code];   // split variants by full code
    keys.add(k);
    let cat=r.category||'';
    if(cat && cat!=='General' && !allCats().some(c=>c.toLowerCase()===cat.toLowerCase())){
      if(!projCats.includes(cat)) projCats.push(cat);          // auto-create unknown categories (e.g. LBR)
    } else if(cat){ cat=allCats().find(c=>c.toLowerCase()===cat.toLowerCase())||cat; }   // normalise case
    const ex=projMeta[k]||{};
    projMeta[k]=Object.assign({}, ex, { master:true,
      code:mergeField(ex.code,r.code), customer:mergeField(ex.customer,r.customer),
      category:cat||ex.category||'', chipset:mergeField(ex.chipset,r.chipset), desc:mergeField(ex.desc,r.desc),
      custProj:r.custProj||ex.custProj||'', component:r.component||ex.component||'' });
  });
  // when a base code is split into variants, drop any stale MERGED master left at the base key (unless real tasks use it)
  Object.keys(cnt).forEach(b=>{ if(cnt[b]>1 && projMeta[b] && projMeta[b].master && !keys.has(b)
      && !visibleTasks().some(t=>resolveProjk(t.projk||t.key)===b)) delete projMeta[b]; });
  persist(); renderCatalog(); renderStats();
  $('#projListModal').hidden=true;
  toast('Imported '+rows.length+' lines, created '+keys.size+' projects');
}
function addBlankProject(){
  const code=prompt('New project code (e.g. B01W050.00):'); if(!code||!code.trim()) return;
  const k=projMasterKey(code);
  projMeta[k]=Object.assign({master:true, code:code.trim()}, projMeta[k]||{}, {master:true});
  persist(); renderCatalog(); openProjEdit(k);
}
function setProjPhase(projk, phase){ const m=projMeta[projk]||{}; m.phase=phase||''; projMeta[projk]=m; persist(); renderCatalog(); }

/* ---------- PROJECT SCHEDULE (milestone timeline -> polished SVG, like the dev-schedule slides) ---------- */
const SCHED_LANES_DEFAULT=['Milestone','HW','Enclosure/Tooling','SW'];
function _sd(s){ if(!s) return null; const m=String(s).match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/); return m? new Date(+m[1],+m[2]-1,+m[3]) : null; }
function _mon(d){ const x=new Date(d); const k=(x.getDay()+6)%7; x.setDate(x.getDate()-k); x.setHours(0,0,0,0); return x; }
function _md(d){ return (d.getMonth()+1)+'/'+d.getDate(); }
function _sx(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
// build a polished schedule SVG from {lanes, items:[{lane,label,date,end,type:dot|bar|star,note,color}], start, end, note}
function scheduleSVG(sched){
  sched=sched||{};
  const lanes=(sched.lanes&&sched.lanes.length)?sched.lanes:SCHED_LANES_DEFAULT;
  const items=(sched.items||[]).map(it=>Object.assign({},it,{_d:_sd(it.date),_e:_sd(it.end)})).filter(it=>it._d);
  const today=new Date(); today.setHours(0,0,0,0);
  const ds=[today]; items.forEach(it=>{ ds.push(it._d); if(it._e) ds.push(it._e); });
  let minD=_sd(sched.start), maxD=_sd(sched.end);
  if(!minD) minD = ds.length? new Date(Math.min(...ds.map(d=>+d))) : _mon(today);
  if(!maxD) maxD = ds.length? new Date(Math.max(...ds.map(d=>+d))) : new Date(+today+120*864e5);
  const start=_mon(minD), end=_mon(new Date(+maxD+10*864e5));
  const weeks=Math.max(6, Math.round((end-start)/(7*864e5)));
  const weekW=26, gut=126, padT=66, laneH=96, padB=14;
  const W=gut+weeks*weekW+20, H=padT+lanes.length*laneH+padB;
  const X=d=>gut+((+d-+start)/864e5)/7*weekW;
  let s=`<rect x="0" y="0" width="${W}" height="${H}" fill="#ffffff"/>`;
  // quarter bands
  let i=0; while(i<weeks){ const wd=new Date(+start+i*7*864e5); const q=wd.getFullYear()+' Q'+(Math.floor(wd.getMonth()/3)+1);
    let j=i; while(j<weeks){ const wj=new Date(+start+j*7*864e5); if((wj.getFullYear()+' Q'+(Math.floor(wj.getMonth()/3)+1))!==q) break; j++; }
    const x0=X(new Date(+start+i*7*864e5)), w=(j-i)*weekW;
    s+=`<rect x="${x0}" y="8" width="${w}" height="20" fill="#c7d2e6" stroke="#ffffff"/><text x="${x0+w/2}" y="22" font-size="11" font-weight="700" fill="#1e293b" text-anchor="middle">${q}</text>`; i=j; }
  // week gridlines + dates
  for(let k=0;k<weeks;k++){ const wd=new Date(+start+k*7*864e5); const x=X(wd);
    s+=`<line x1="${x}" y1="30" x2="${x}" y2="${H-padB}" stroke="#eef2f7"/><text x="${x+2}" y="42" font-size="8" fill="#64748b">${_md(wd)}</text>`; }
  // lanes
  lanes.forEach((ln,li)=>{ const yTop=padT+li*laneH, base=yTop+laneH*0.6;
    s+=`<line x1="${gut}" y1="${yTop}" x2="${W-8}" y2="${yTop}" stroke="#cbd5e1"/><text x="10" y="${yTop+laneH*0.52}" font-size="13" font-weight="700" fill="#0f172a">${_sx(ln)}</text><line x1="${gut}" y1="${base}" x2="${W-8}" y2="${base}" stroke="#94a3b8" stroke-width="1.4"/>`; });
  s+=`<line x1="${gut}" y1="${H-padB}" x2="${W-8}" y2="${H-padB}" stroke="#cbd5e1"/>`;
  const baseOf=ln=>{ let li=lanes.indexOf(ln); if(li<0) li=0; return padT+li*laneH+laneH*0.6; };
  // items
  items.forEach(it=>{ const x=X(it._d), y=baseOf(it.lane||lanes[0]);
    if(it.type==='bar' && it._e){ const x2=X(it._e), col=it.color||'#bfdbfe';
      s+=`<rect x="${x}" y="${y-8}" width="${Math.max(6,x2-x)}" height="16" rx="4" fill="${col}" stroke="#60a5fa"/><text x="${(x+x2)/2}" y="${y-12}" font-size="9" fill="#1e3a8a" text-anchor="middle">${_sx(it.label)}</text>`;
    } else if(it.type==='star'){
      s+=`<text x="${x}" y="${y+7}" font-size="22" fill="${it.color||'#ec4899'}" text-anchor="middle">★</text><text x="${x}" y="${y-15}" font-size="9.5" font-weight="600" fill="#0f172a" text-anchor="middle">${_sx(it.label)}</text><text x="${x}" y="${y+24}" font-size="9" fill="#475569" text-anchor="middle">${_sx(it.date)}</text>`;
    } else {
      s+=`<circle cx="${x}" cy="${y}" r="7" fill="${it.color||'#facc15'}" stroke="#b8860b"/><text x="${x}" y="${y-15}" font-size="9.5" font-weight="600" fill="#0f172a" text-anchor="middle">${_sx(it.label)}</text><text x="${x}" y="${y+21}" font-size="9" fill="#475569" text-anchor="middle">${_sx(it.date)}</text>`;
    }
    if(it.note) s+=`<text x="${x}" y="${y+33}" font-size="8" fill="#7c3aed" text-anchor="middle">${_sx(it.note)}</text>`;
  });
  // now line
  const nx=X(today);
  s+=`<line x1="${nx}" y1="30" x2="${nx}" y2="${H-padB}" stroke="#dc2626" stroke-width="2" stroke-dasharray="5 4"/><rect x="${nx-18}" y="29" width="36" height="15" rx="3" fill="#f97316"/><text x="${nx}" y="40" font-size="9" font-weight="700" fill="#fff" text-anchor="middle">Now</text>`;
  if(sched.note) s+=`<text x="${gut+4}" y="60" font-size="10" fill="#92400e">📌 ${_sx(sched.note)}</text>`;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="system-ui,Segoe UI,sans-serif">${s}</svg>`;
}
let _schedEdit=null;
function openSchedule(projk){
  const g=projectGroups().find(x=>x.projk===resolveProjk(projk));
  const meta=projMeta[projk]||{};
  const def={lanes:SCHED_LANES_DEFAULT.slice(), items:[], start:'', end:'', note:''};
  _schedEdit={ projk, title:(meta.code||(g&&g.label)||projk), sched:JSON.parse(JSON.stringify(meta.schedule||def)) };
  if(!_schedEdit.sched.lanes||!_schedEdit.sched.lanes.length) _schedEdit.sched.lanes=SCHED_LANES_DEFAULT.slice();
  renderScheduleEditor(); $('#scheduleModal').hidden=false;
}
function renderScheduleEditor(){
  const {title,sched}=_schedEdit;
  const rows=(sched.items||[]).map((it,i)=>`<tr>
    <td><select data-si="${i}" data-sf="lane">${sched.lanes.map(l=>`<option ${it.lane===l?'selected':''}>${esc(l)}</option>`).join('')}</select></td>
    <td><select data-si="${i}" data-sf="type">${['dot','bar','star'].map(t=>`<option ${it.type===t?'selected':''}>${t}</option>`).join('')}</select></td>
    <td><input data-si="${i}" data-sf="label" value="${esc(it.label||'')}" placeholder="Label"></td>
    <td><input type="date" data-si="${i}" data-sf="date" value="${esc(it.date||'')}"></td>
    <td><input type="date" data-si="${i}" data-sf="end" value="${esc(it.end||'')}" title="bar end date"></td>
    <td><input data-si="${i}" data-sf="note" value="${esc(it.note||'')}" placeholder="Note"></td>
    <td><button class="btn xs danger" data-sdel="${i}">✕</button></td></tr>`).join('');
  const refs=((projMeta[_schedEdit.projk]||{}).images||[]).filter(im=>/schedule/i.test(im.title||''));
  const refStrip=refs.length?`<div class="sched-refs">${refs.map(im=>`<span class="pimg-item"><img src="${im.data}" data-light="${im.id}"><div class="pimg-cap">Original schedule attachment</div><button class="img-del" data-pdelimg="${esc(_schedEdit.projk)}|${im.id}" title="Delete">✕</button></span>`).join('')}</div>`:'';
  $('#scheduleInner').innerHTML=`
    <div class="modal-head"><h2>📅 Schedule · ${esc(title)}</h2><button class="icon-btn" data-close>✕</button></div>
    <div class="modal-body">
      <div class="sched-preview">${(sched.items&&sched.items.length)?scheduleSVG(sched):'<p class="hint">Upload your original schedule, or click "Apply default template" — milestones appear below and the timeline draws live here.</p>'}</div>
      <div class="sched-attach">
        <label class="btn sm">📎 Upload schedule file<input type="file" accept="image/*" hidden data-saddimg="1"></label>
        <button class="btn sm" data-stemplate="1">Apply default template</button>
        <span class="hint">Upload your existing schedule as reference; a default milestone set is added for you to edit.</span>
      </div>
      ${refStrip}
      <div class="sched-controls">
        <label>From <input type="date" data-smeta="start" value="${esc(sched.start||'')}"></label>
        <label>To <input type="date" data-smeta="end" value="${esc(sched.end||'')}"></label>
        <label class="wide">Lanes (comma-separated)<input data-smeta="lanes" value="${esc(sched.lanes.join(', '))}"></label>
        <label class="wide">Title note <input data-smeta="note" value="${esc(sched.note||'')}" placeholder="Caption shown above the timeline (optional)"></label>
      </div>
      <table class="sched-table"><thead><tr><th>Lane</th><th>Type</th><th>Label</th><th>Date</th><th>End (bar)</th><th>Note</th><th></th></tr></thead><tbody>${rows}</tbody></table>
      <button class="btn sm" data-sadd="1">＋ Add milestone</button>
      <p class="hint">Type: dot = milestone, star = key sample (pink), bar = multi-week activity (needs an end date). Editing any field updates the timeline live.</p>
    </div>
    <div class="modal-foot">
      <button class="btn" data-sppt="1">⬇ PowerPoint</button>
      <button class="btn" data-sxlsx="1">⬇ Excel</button>
      <button class="btn" data-sdownload="1">⬇ PNG</button>
      <button class="btn" data-close>Cancel</button>
      <button class="btn primary" data-ssave="1">Save schedule</button>
    </div>`;
}
function addSchedItem(){ _schedEdit.sched.items.push({lane:_schedEdit.sched.lanes[0]||'Milestone', type:'dot', label:'', date:'', end:'', note:''}); renderScheduleEditor(); }
function delSchedItem(i){ _schedEdit.sched.items.splice(+i,1); renderScheduleEditor(); }
function updSchedItem(i,f,v){ _schedEdit.sched.items[+i][f]=v; renderScheduleEditor(); }
function updSchedMeta(f,v){ if(f==='lanes') _schedEdit.sched.lanes=v.split(/[,, ]/).map(x=>x.trim()).filter(Boolean); else _schedEdit.sched[f]=v; renderScheduleEditor(); }
function saveScheduleEdit(){ const m=projMeta[_schedEdit.projk]||{}; m.schedule=_schedEdit.sched; projMeta[_schedEdit.projk]=m; persist(); renderCatalog(); $('#scheduleModal').hidden=true; toast('Schedule saved'); }
function downloadSchedulePNG(){
  const svg=scheduleSVG(_schedEdit.sched), blob=new Blob([svg],{type:'image/svg+xml;charset=utf-8'}), url=URL.createObjectURL(blob);
  const img=new Image(); img.onload=()=>{ const c=document.createElement('canvas'); c.width=img.naturalWidth*2; c.height=img.naturalHeight*2;
    const ctx=c.getContext('2d'); ctx.scale(2,2); ctx.fillStyle='#fff'; ctx.fillRect(0,0,img.naturalWidth,img.naturalHeight); ctx.drawImage(img,0,0); URL.revokeObjectURL(url);
    c.toBlob(b=>{ const a=document.createElement('a'); a.href=URL.createObjectURL(b); a.download=((_schedEdit.title||'project')+'_schedule.png').replace(/[\\/:*?"<>|]/g,'_'); a.click(); }); };
  img.onerror=()=>{ URL.revokeObjectURL(url); toast('PNG export failed'); }; img.src=url;
}
function seedDefaultSchedule(){                         // a standard HW/SW milestone template the user then adjusts
  if(_schedEdit.sched.items && _schedEdit.sched.items.length) return false;
  const base=new Date(); base.setHours(0,0,0,0);
  const d=n=>{ const x=new Date(base); x.setDate(x.getDate()+n); return x.toISOString().slice(0,10); };
  _schedEdit.sched.items=[
    {lane:'Milestone',type:'star',label:'Kickoff',date:d(0),end:'',note:''},
    {lane:'HW',type:'dot',label:'SCH & BOM release',date:d(14),end:'',note:''},
    {lane:'HW',type:'dot',label:'EVT build',date:d(35),end:'',note:''},
    {lane:'HW',type:'dot',label:'DVT build',date:d(70),end:'',note:''},
    {lane:'HW',type:'dot',label:'PVT build',date:d(105),end:'',note:''},
    {lane:'Milestone',type:'star',label:'MP',date:d(140),end:'',note:''},
    {lane:'SW',type:'dot',label:'Alpha Release',date:d(42),end:'',note:''},
    {lane:'SW',type:'dot',label:'Beta Release',date:d(80),end:'',note:''},
    {lane:'SW',type:'dot',label:'GA Release',date:d(125),end:'',note:''}
  ];
  return true;
}
function applyDefaultTemplate(){ if(seedDefaultSchedule()){ renderScheduleEditor(); toast('Default milestone template added — adjust to your real dates'); } else toast('Milestones already exist (clear them first to apply the template)'); }
async function uploadScheduleAttach(fileList){
  const n=await addProjectImages(_schedEdit.projk, fileList, 'Schedule attachment', 'Schedule');
  if(n){ const seeded=seedDefaultSchedule(); renderScheduleEditor(); if(seeded) toast('Schedule attached and a default milestone template added — adjust it'); }
}
function exportSchedulePPT(){
  const svg=scheduleSVG(_schedEdit.sched), blob=new Blob([svg],{type:'image/svg+xml;charset=utf-8'}), url=URL.createObjectURL(blob);
  const img=new Image();
  img.onload=()=>{ const sc=2, c=document.createElement('canvas'); c.width=img.naturalWidth*sc; c.height=img.naturalHeight*sc;
    const ctx=c.getContext('2d'); ctx.scale(sc,sc); ctx.fillStyle='#fff'; ctx.fillRect(0,0,img.naturalWidth,img.naturalHeight); ctx.drawImage(img,0,0); URL.revokeObjectURL(url);
    const dataUrl=c.toDataURL('image/png'); const iw=img.naturalWidth, ih=img.naturalHeight;
    try{ const P=new PptxGenJS(); P.defineLayout({name:'WIDE',width:13.33,height:7.5}); P.layout='WIDE'; const s=P.addSlide();
      s.addText((_schedEdit.title||'')+' Development Schedule',{x:0.4,y:0.22,w:12.5,h:0.6,fontSize:24,bold:true,color:'1A3B79'});
      const maxW=12.6,maxH=6.2; let fw=maxW, fh=fw*ih/iw; if(fh>maxH){ fh=maxH; fw=fh*iw/ih; }
      s.addImage({data:dataUrl, x:(13.33-fw)/2, y:1.0, w:fw, h:fh});
      P.writeFile({fileName:((_schedEdit.title||'project')+'_schedule.pptx').replace(/[\\/:*?"<>|]/g,'_')}); toast('PowerPoint downloaded');
    }catch(e){ console.warn(e); toast('PPT export failed'); }
  };
  img.onerror=()=>{ URL.revokeObjectURL(url); toast('PPT export failed'); }; img.src=url;
}
// Excel with the full-colour schedule image embedded. The free SheetJS can't embed images,
// so we build the .xlsx (OOXML zip) by hand with JSZip and float the PNG over the sheet.
function exportScheduleExcel(){
  const svg=scheduleSVG(_schedEdit.sched), b=new Blob([svg],{type:'image/svg+xml;charset=utf-8'}), url=URL.createObjectURL(b);
  const img=new Image();
  img.onload=()=>{ const sc=2, c=document.createElement('canvas'); c.width=img.naturalWidth*sc; c.height=img.naturalHeight*sc;
    const ctx=c.getContext('2d'); ctx.scale(sc,sc); ctx.fillStyle='#fff'; ctx.fillRect(0,0,img.naturalWidth,img.naturalHeight); ctx.drawImage(img,0,0); URL.revokeObjectURL(url);
    const b64=c.toDataURL('image/png').split(',')[1]; const cx=Math.round(img.naturalWidth*9525), cy=Math.round(img.naturalHeight*9525);
    try{
      const NS='xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
      const zip=new JSZip();
      zip.file('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>');
      zip.file('_rels/.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships ${NS}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
      zip.file('xl/workbook.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Schedule" sheetId="1" r:id="rId1"/></sheets></workbook>');
      zip.file('xl/_rels/workbook.xml.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships ${NS}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
      zip.file('xl/worksheets/sheet1.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rId1"/></worksheet>');
      zip.file('xl/worksheets/_rels/sheet1.xml.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships ${NS}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`);
      zip.file('xl/drawings/drawing1.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Schedule"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`);
      zip.file('xl/drawings/_rels/drawing1.xml.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships ${NS}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`);
      zip.file('xl/media/image1.png', b64, {base64:true});
      zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}).then(out=>{
        downloadBlob(out, ((_schedEdit.title||'project')+'_schedule.xlsx').replace(/[\\/:*?"<>|]/g,'_')); toast('Excel (with the colour timeline) downloaded');
      }).catch(e=>{ console.warn(e); toast('Excel export failed'); });
    }catch(e){ console.warn(e); toast('Excel export failed'); }
  };
  img.onerror=()=>{ URL.revokeObjectURL(url); toast('Excel export failed'); }; img.src=url;
}
/* ---------- DESIGN DOCUMENTS (block diagram / GPIO table / architecture …), per project, cloud-synced ---------- */
// stored per project in projMeta[projk].images (each image may carry a `title`). Two-level browser: pick a project -> its docs.
let _ddProjk=null;
function openDesignDocs(){ _ddProjk=null; renderDesignDocs(); $('#projImgModal').hidden=false; }
function renderDesignDocs(){
  if(_ddProjk){ renderDesignDocsProject(); return; }
  // LEVEL 1 — a card per project; click to open that project's design documents
  const groups=projectGroups();
  const cards=groups.map(g=>{ const m=projMeta[g.projk]||{}, imgs=m.images||[], thumb=imgs[0];
    return `<div class="dd-card" data-ddopen="${esc(g.projk)}">
      <div class="dd-thumb">${thumb?`<img src="${thumb.data}">`:'<span class="dd-noimg">📐</span>'}</div>
      <div class="dd-cardbody"><div class="dd-cardtitle">${esc(m.code||g.label)}</div>
        <div class="dd-cardsub"><span class="cat-name cat-${esc(projCategory(g))}">${esc(projCategory(g))}</span> · ${imgs.length} docs</div></div></div>`;
  }).join('');
  $('#projImgInner').innerHTML=`
    <div class="modal-head"><h2>📐 Design Documents</h2><button class="icon-btn" data-close>✕</button></div>
    <div class="modal-body">
      <p class="hint">Open a project to upload / view its design documents (block diagram, GPIO table, architecture, schematic…).</p>
      <div class="dd-cards">${cards||'<p class="hint">No projects yet — create some in the catalog first.</p>'}</div>
    </div>
    <div class="modal-foot"><button class="btn" data-close>Close</button></div>`;
}
const DOC_FOLDERS=['Block Diagram','GPIO Table','Schematic','Layout','Spec','Other'];
function renderDesignDocsProject(){
  const projk=_ddProjk, m=projMeta[projk]||{}, imgs=m.images||[];
  const g=projectGroups().find(x=>x.projk===resolveProjk(projk));
  // group docs into folders (by their `folder`), list-view rows inside each folder
  const byFolder={}; imgs.forEach(im=>{ const f=im.folder||'Other'; (byFolder[f]=byFolder[f]||[]).push(im); });
  const folders=[...DOC_FOLDERS.filter(f=>byFolder[f]), ...Object.keys(byFolder).filter(f=>!DOC_FOLDERS.includes(f))];
  const body= imgs.length ? folders.map(f=>`
      <div class="dd-folder">
        <div class="dd-folder-head">📁 ${esc(f)} <span class="ct-n">${byFolder[f].length}</span></div>
        <div class="dd-list">${byFolder[f].map(im=>`<div class="dd-row"><img class="dd-rowthumb" src="${im.data}" data-light="${im.id}"><span class="dd-rowtitle">${esc(im.title||'Untitled')}</span><button class="img-del" data-pdelimg="${esc(projk)}|${im.id}" title="Delete">✕</button></div>`).join('')}</div>
      </div>`).join('') : '<p class="hint">No design documents yet — pick a folder, enter a name, and upload the first one.</p>';
  $('#projImgInner').innerHTML=`
    <div class="modal-head"><h2><button class="btn xs" data-ddback="1">← Back</button> &nbsp;📐 ${esc(m.code||(g&&g.label)||projk)} · Design documents</h2><button class="icon-btn" data-close>✕</button></div>
    <div class="modal-body">
      <div class="dd-upload">
        <select id="ddFolder" title="Folder / category">${DOC_FOLDERS.map(f=>`<option>${esc(f)}</option>`).join('')}</select>
        <input id="ddTitle" placeholder="Document name (e.g. RX Block, Top GPIO)">
        <label class="btn sm primary">＋ Upload<input type="file" accept="image/*" multiple hidden data-ddaddimg="1"></label>
      </div>
      <p class="hint">Pick a folder (Block Diagram / GPIO Table / Schematic…) → enter a name → upload. Click a thumbnail to enlarge; delete anytime; synced to the cloud.</p>
      ${body}
    </div>
    <div class="modal-foot"><button class="btn" data-close>Close</button></div>`;
}
async function addProjectImages(projk, fileList, title, folder){
  const all=Array.from(fileList||[]);
  const files=all.filter(f=>!f.type || /^image\//.test(f.type));   // input already限定圖片；手機某些格式 type 會是空字串, 也放行
  if(!files.length){ toast(all.length?'This file is not an image':'No file selected'); return 0; }
  toast('Processing image… please wait');
  const m=projMeta[projk]||{}; m.images=m.images||[]; let n=0;
  for(const f of files){ try{ const s=await shrinkImageBudget(await fileToDataURL(f), 2200); m.images.push({id:uid(), data:s.data, w:s.w, h:s.h, title:title||'', folder:folder||''}); n++; }catch(e){ console.warn('project image failed', e); } }
  if(!n){ toast('Image processing failed — try another file'); return 0; }
  projMeta[projk]=m; persist(); renderCatalog();
  if(!$('#projImgModal').hidden) renderDesignDocs();
  if(!$('#scheduleModal').hidden) renderScheduleEditor();
  toast('Added '+n+' image(s)'); return n;
}
function removeProjectImage(projk, id){
  const m=projMeta[projk]; if(!m||!m.images) return;
  if(!confirm('Delete this image? This cannot be undone.')) return;
  m.images=m.images.filter(x=>x.id!==id); cloudDeleteImage(id); persist(); renderCatalog();
  if(!$('#projImgModal').hidden) renderDesignDocs();
  if(!$('#scheduleModal').hidden) renderScheduleEditor();
  toast('Image removed');
}
function inferCategory(projStr){
  const s=String(projStr||'').toLowerCase();
  if(/dongle|\bdg\d|redcap/.test(s)) return 'Dongle';
  if(/\bodu\b/.test(s)) return 'ODU';
  if(/\bid\s?\d{3}\b|\bidu\b/.test(s)) return 'IDU';
  if(/module|\bfr1\b|\bfr2\b|sdx7\d/.test(s)) return 'Module';
  return 'General';
}
function projCategory(g){ const m=projMeta[g.projk]||{}; return m.category && allCats().includes(m.category)? m.category : (m.category||inferCategory(g.projStr||g.label)); }
function renderCatalog(){
  const sel=$('#catalogMember');
  if(sel) sel.innerHTML='<option value="">All members</option>'+
    members.map(m=>`<option value="${m.id}" ${catalogMember===m.id?'selected':''}>${esc(m.name)}</option>`).join('');
  let groups=projectGroups();
  if(catalogMember) groups=groups.map(g=>({...g, tasks:g.tasks.filter(t=>(t.ownerIds||[]).includes(catalogMember))}))
                                  .filter(g=>g.tasks.length);
  const cont=$('#projectCatalog'); if(!cont) return;
  if(catalogView==='cards') catalogView='tree';                // Cards view retired -> Tree is the content
  const views=[['tree','🌳 Tree'],['matrix','▦ Matrix']];
  const bar=`<div class="cat-viewtabs">
      ${views.map(([v,l])=>`<button class="cv-tab ${catalogView===v?'active':''}" data-catview="${v}">${l}</button>`).join('')}
      <span class="cv-hint">✎ edit · ＋ add a task · drag a task chip onto a model card to move it.</span>
    </div>`;
  if(!groups.length){ cont.innerHTML=bar+`<p class="hint">No projects yet — import reports, or use "＋ Add project / 📋 Paste project list" above.</p>`; return; }
  if(catalogView==='matrix'){ cont.innerHTML=bar+renderMatrixHTML(groups); return; }
  cont.innerHTML=bar+renderTreeHTML(groups);                  // default: Tree
  wireTreeTaskDrag();
}
/* tree: drag a task chip onto a model card to move the task into that project */
let _dragTaskId=null;
function wireTreeTaskDrag(){
  const cont=$('#projectCatalog'); if(!cont) return;
  cont.querySelectorAll('.ptree-task[draggable]').forEach(ch=>{
    ch.addEventListener('dragstart',e=>{ _dragTaskId=ch.dataset.opentask; try{ e.dataTransfer.setData('text/plain',_dragTaskId); }catch(err){} ch.classList.add('dragging'); });
    ch.addEventListener('dragend',()=>{ _dragTaskId=null; ch.classList.remove('dragging'); });
  });
  cont.querySelectorAll('.ptree-leaf').forEach(leaf=>{
    leaf.addEventListener('dragover',e=>{ if(_dragTaskId){ e.preventDefault(); leaf.classList.add('drop-target'); } });
    leaf.addEventListener('dragleave',()=>leaf.classList.remove('drop-target'));
    leaf.addEventListener('drop',e=>{ e.preventDefault(); leaf.classList.remove('drop-target');
      if(!_dragTaskId) return;
      const projk=leaf.dataset.editProj, code=(projMeta[projk]||{}).code||'project';
      editTaskField(_dragTaskId,'project',projk);
      toast('Task moved to '+code);
      _dragTaskId=null;
    });
  });
}
function renderCatalogCardsHTML(groups){
  const byCat={}; groups.forEach(g=>{ const c=projCategory(g); (byCat[c]=byCat[c]||[]).push(g); });
  const order=[...allCats(), ...Object.keys(byCat).filter(c=>!allCats().includes(c))];
  const present=order.filter(c=>byCat[c]);
  if(catFilter && !byCat[catFilter]) catFilter='';            // selected tab emptied -> back to all
  const tab=(c,lbl,n,manage)=>`<button class="cat-tab cat-${esc(c)} ${catFilter===c?'active':''}" data-catfilter="${esc(c)}">${esc(lbl)} <span class="ct-n">${n}</span>${manage?`<span class="cc-x" data-renamecat="${esc(c)}" title="Rename">✎</span><span class="cc-x" data-delcat="${esc(c)}" title="Delete category">✕</span>`:''}</button>`;
  const tabBar=`<div class="cat-tabs">
      ${tab('','All',groups.length,false)}
      ${present.map(c=>tab(c,c,byCat[c].length,(!BASE_CATS.includes(c)&&c!=='General'))).join('')}
      <button class="btn xs primary cat-addtab" data-addcat="1" title="Add a custom category">＋ Add category</button>
    </div>`;
  const showCats = catFilter ? [catFilter] : present;
  const sections = showCats.map(c=>{
    const list=byCat[c]; const nTasks=list.reduce((s,g)=>s+g.tasks.length,0);
    const head = catFilter ? '' : `<div class="cat-sec-head"><span class="cat-name cat-${esc(c)}">${esc(c)}</span><span class="cat-meta">${list.length} projects · ${nTasks} tasks</span></div>`;
    return head + `<div class="cat-sec">${list.map(catalogCard).join('')}</div>`;
  }).join('');
  return tabBar + sections;
}
// ---- Customer-Project (ID525) + component (Mainboard/Module/ME) derivation, until the richer matrix import fills them ----
function projCustProj(g){
  const m=projMeta[g.projk]||{};
  if(m.custProj) return String(m.custProj).trim();
  const hay=[m.code,g.label,m.desc,g.projStr].filter(Boolean).join('  ');
  const t=hay.match(/\b(ID|DG|CM)[\s-]?(\d{2,4})\b/i);
  return t? (t[1].toUpperCase()+t[2]) : '';
}
function projComponent(g){
  const m=projMeta[g.projk]||{};
  if(m.component) return String(m.component).trim();
  const code=String(m.code||g.label||'').toUpperCase();
  if(/\.\d{2}\b/.test(code)) return 'ME';                       // B01W036.00 -> mechanical
  if(/T0+\b/.test(code)) return 'Mainboard';                    // ...T00 -> mainboard
  if(/T\d+\b/.test(code)) return 'Module';                      // ...T02 -> module
  return '';
}
// <option>s for a project picker (value = projk) — constrains entry to known projects (no more free-typed junk)
function projectOptionsHTML(curProjk){
  const all=projectGroups().map(g=>{ const m=projMeta[g.projk]||{}; const cp=m.custProj||projCustProj(g), comp=m.component||projComponent(g);
    return {projk:g.projk, master:!!m.master, label:m.master? `${m.code||g.label}${cp?' · '+cp:''}${comp?' · '+comp:''}` : g.label}; });
  const tag=o=>`<option value="${esc(o.projk)}" ${curProjk===o.projk?'selected':''}>${esc(o.label)}</option>`;
  const master=all.filter(o=>o.master).sort((a,b)=>a.label.localeCompare(b.label));
  const other=all.filter(o=>!o.master).sort((a,b)=>a.label.localeCompare(b.label));
  const hasCur=all.some(o=>o.projk===curProjk);
  return `${(!hasCur&&curProjk)?`<option value="${esc(curProjk)}" selected>(current)</option>`:''}<option value="">— select project —</option>`+
    (master.length?`<optgroup label="Official projects">${master.map(tag).join('')}</optgroup>`:'')+
    (other.length?`<optgroup label="Other (from reports)">${other.map(tag).join('')}</optgroup>`:'');
}
function projTree(groups){                                      // Product Type -> Customer Project -> [project leaves]
  const types={};
  groups.forEach(g=>{
    const ty=projCategory(g)||'General', cp=projCustProj(g), cust=(projMeta[g.projk]||{}).customer||'';
    const key=cp?(cust+'|'+cp):(cust?('·'+cust+'·'+g.label):g.label);   // same project name under different customers (e.g. DG500) stays separate
    const T=(types[ty]=types[ty]||{order:[],map:{}});
    if(!T.map[key]){ T.map[key]={cp,cust,label:cp||cust||g.label,items:[]}; T.order.push(key); }
    T.map[key].items.push(g);
  });
  return types;
}
function projTypeOrder(types){
  const cats=allCats();
  return Object.keys(types).sort((a,b)=>{ const ia=cats.indexOf(a), ib=cats.indexOf(b); return (ia<0?99:ia)-(ib<0?99:ib); });
}
// Matrix & Tree show the OFFICIAL master projects (the imported matrix); task-only groups stay in Cards.
function masterOnly(groups){ const mg=groups.filter(g=>(projMeta[g.projk]||{}).master); return mg.length?mg:groups; }
function renderMatrixHTML(groups){
  groups=masterOnly(groups);
  const types=projTree(groups);
  let html='<div class="pmatrix-wrap">';
  projTypeOrder(types).forEach(ty=>{
    const block=types[ty], nProj=block.order.reduce((s,k)=>s+block.map[k].items.length,0);
    let rows='';
    block.order.forEach(key=>{
      const grp=block.map[key], multi=grp.items.length>1;
      grp.items.forEach((g,idx)=>{
        const m=projMeta[g.projk]||{};
        rows+=`<tr class="${multi&&idx===0?'pm-grp-start':''}">
          <td class="pm-cust">${idx===0?esc(grp.cust||'—'):''}</td>
          <td class="pm-cp">${idx===0?esc(grp.cp||(grp.cust?'':'—')):''}</td>
          <td class="pm-comp">${esc(projComponent(g)||'—')}</td>
          <td class="pm-model"><b>${esc(m.code||g.label)}</b></td>
          <td class="pm-chip">${esc(m.chipset||'—')}</td>
          <td class="pm-tasks">${g.tasks.length||'·'}</td>
          <td class="pm-act">
            <button class="btn xs" data-addtaskproj="${esc(g.projk)}" title="Add a task to this project">＋</button>
            <button class="btn xs" data-sched="${esc(g.projk)}" title="Schedule timeline">📅</button>
            <button class="btn xs" data-edit-proj="${esc(g.projk)}" title="Edit details">✎</button>
            <button class="btn xs danger" data-delproj="${esc(g.projk)}" title="Delete project">🗑</button>
          </td></tr>`;
      });
    });
    html+=`<div class="pm-type"><span class="pm-type-name cat-${esc(ty)}">${esc(ty)}</span><span class="pm-type-n">${nProj} model${nProj>1?'s':''}</span></div>
      <table class="pmatrix"><thead><tr><th>Customer</th><th>Customer&nbsp;Project</th><th>Component</th><th>Model</th><th>Chipset</th><th>Tasks</th><th></th></tr></thead><tbody>${rows}</tbody></table>`;
  });
  return html+'</div>';
}
function renderTreeHTML(groups){
  const types=projTree(masterOnly(groups));
  // weekly-report tasks are tagged by Customer Project (e.g. "B01W043 / ID535"), so link tasks to the
  // project node by derived Customer Project; tasks with no project go under a "Cross-project" node.
  const tasksByCp={}, noCp=[];
  groups.forEach(g=>{ if(!(g.tasks||[]).length) return;
    if((projMeta[g.projk]||{}).master) return;             // tasks assigned to a model live under that leaf, not in the CP-level branch
    const cp=projCustProj(g);
    if(cp)(tasksByCp[cp]=tasksByCp[cp]||[]).push(...g.tasks); else noCp.push(...g.tasks); });
  const usedCp={}, CAP=6;
  const taskChip=t=>{ const txt=(t.current||t.next||'(no description)').replace(/\s+/g,' ').trim().slice(0,48);
    return `<button class="ptree-task st-${esc(t.status||'')}" draggable="true" data-opentask="${esc(t.id)}" title="Open task · drag onto a model card to move it"><span class="ptt-dot"></span><span class="ptt-txt">${esc(txt)}</span></button>`; };
  const taskBranch=(cp,ts)=>{ if(!ts||!ts.length) return '';
    const exp=treeExpand[cp], list=exp?ts:ts.slice(0,CAP);
    const chips=list.map(taskChip).join('');
    const toggle=ts.length>CAP?`<button class="ptree-more" data-treeexp="${esc(cp)}">${exp?'▾ Show less':'▸ +'+(ts.length-CAP)+' more'}</button>`:'';
    return `<div class="ptree-tasks"><div class="ptt-head">📋 ${ts.length} task${ts.length>1?'s':''} · no model assigned — drag one onto a model card</div>${chips}${toggle}</div>`; };
  let html='<div class="ptree-wrap">';
  projTypeOrder(types).forEach(ty=>{
    const block=types[ty];
    html+=`<div class="ptree-type cat-${esc(ty)}">${esc(ty)}</div>`;
    block.order.forEach(key=>{
      const grp=block.map[key];
      const leaves=grp.items.map(g=>{
        const m=projMeta[g.projk]||{};
        const n=(g.tasks||[]).length, lk='leaf:'+g.projk, lexp=treeExpand[lk];
        const ownTasks = (n&&lexp) ? `<div class="ptree-leaftasks">${g.tasks.map(taskChip).join('')}</div>` : '';
        return `<div class="ptree-leafcol">
          <div class="ptree-leafrow">
            <button class="ptree-leaf" data-edit-proj="${esc(g.projk)}" title="Edit ${esc(m.code||g.label)} · drop a task here to move it">
              <span class="ptl-comp">${esc(projComponent(g)||'Model')}</span>
              <span class="ptl-code">${esc(m.code||g.label)}</span>
              ${m.chipset?`<span class="ptl-chip">${esc(m.chipset)}</span>`:''}
            </button>
            ${n?`<button class="ptree-leafexp" data-treeexp="${esc(lk)}" title="Show this model's tasks">${lexp?'▾':'▸'} ${n}</button>`:''}
            <button class="ptree-addtask" data-addtaskproj="${esc(g.projk)}" title="Add a task to ${esc(m.code||g.label)}">＋</button>
          </div>
          ${ownTasks}
        </div>`;
      }).join('');
      const cp=grp.cp; let branch=''; if(cp && !usedCp[cp]){ usedCp[cp]=1; branch=taskBranch(cp, tasksByCp[cp]); }
      html+=`<div class="ptree-row">
        <div class="ptree-node ptn-cp">${esc(grp.cp||grp.label)}${(grp.cust&&grp.cp)?`<span class="ptn-sub">${esc(grp.cust)}</span>`:''}</div>
        <div class="ptree-conn"></div>
        <div class="ptree-leaves">${leaves}${branch}</div>
      </div>`;
    });
  });
  if(noCp.length){
    html+=`<div class="ptree-type">Cross-project / unassigned</div>
      <div class="ptree-row"><div class="ptree-node ptn-cp">Other<span class="ptn-sub">no project</span></div>
        <div class="ptree-conn"></div><div class="ptree-leaves">${taskBranch('__other__', noCp)}</div></div>`;
  }
  return html+'</div>';
}
function optTags(arr,val){ return arr.map(x=>`<option ${x===val?'selected':''}>${x}</option>`).join(''); }
function catalogCard(g){
  const meta=projMeta[g.projk]||{};
  const desc=(meta.desc!=null&&meta.desc!=='')?meta.desc:cleanDesc(g.projStr);
  const ed=editingProj===g.projk;
  const sub=`${g.tasks.length} tasks · ${g.mem.size} members · ${g.high} high risk · ${g.closed} closed`;
  const curCat=projCategory(g);
  const curPhase=meta.phase||'';
  const phaseSel=`<select class="pc-phase-sel${curPhase?' has-phase':''}" data-setphase="${esc(g.projk)}" title="Project phase">${PHASES.map(p=>`<option value="${esc(p)}" ${curPhase===p?'selected':''}>${p||'Phase…'}</option>`).join('')}</select>`;
  // read view: a category dropdown is ALWAYS shown — change it to recategorize instantly (no Edit mode needed)
  const catSelectRead=`<select class="pc-cat-sel cat-${esc(curCat)}" data-setcat="${esc(g.projk)}" title="Change category">${allCats().map(c=>`<option ${curCat===c?'selected':''}>${esc(c)}</option>`).join('')}</select>`;
  // edit mode: category is part of the form, saved together with the other fields on 儲存 (no premature re-render)
  const catSelectEdit=`<select class="pc-in" data-mf="category">${['',...allCats()].map(c=>`<option value="${esc(c)}" ${(meta.category||curCat)===c?'selected':''}>${c||'(category)'}</option>`).join('')}</select>`;
  const metaRow = ed
    ? `<div class="pc-meta editing">
         <input class="pc-in" data-mf="code" placeholder="Code" value="${esc(meta.code||'')}">
         <input class="pc-in" data-mf="customer" placeholder="Customer" value="${esc(meta.customer||'')}">
         ${catSelectEdit}
         <input class="pc-in" data-mf="custProj" placeholder="Customer Project (e.g. ID515)" value="${esc(meta.custProj||'')}">
         <input class="pc-in" data-mf="component" placeholder="Component (e.g. Mainboard)" value="${esc(meta.component||'')}">
         <input class="pc-in" data-mf="chipset" placeholder="Chipset" value="${esc(meta.chipset||'')}">
         <input class="pc-in wide" data-mf="desc" placeholder="Description" value="${esc(desc)}">
         <button class="btn sm primary" data-save-proj="${esc(g.projk)}">Save</button>
       </div>`
    : `<div class="pc-meta">
         <div class="pc-cell">${esc(meta.customer||'—')}</div>
         <div class="pc-cell">${catSelectRead}</div>
         <div class="pc-cell desc">${esc(desc)}</div>
       </div>`;
  const merged=Object.keys(projMerge).filter(k=>resolveProjk(k)===g.projk).length;
  const chipLine=meta.chipset?` · <span class="pc-chipset">Chipset: ${esc(meta.chipset)}</span>`:'';
  const masterTag=(meta.master&&!g.tasks.length)?' <span class="master-tag" title="From your project master list · no tasks reported yet">No tasks</span>':'';
  return `<div class="proj-card cat-b-${esc(curCat)}" draggable="true" data-projk="${esc(g.projk)}">
    <div class="pc-head">
      <div><h3><span class="drag-dot" title="Drag to merge into another project">⠿</span> ${esc(g.label)}${masterTag}${merged?` <span class="merged-tag" title="Merged into another project">＋${merged} merged</span>`:''}</h3><div class="pc-sub">${sub}${chipLine}</div></div>
      <div class="pc-actions">
        ${phaseSel}
        <button class="btn pc-act" data-sched="${esc(g.projk)}" title="Project schedule timeline">📅 Schedule${(meta.schedule&&(meta.schedule.items||[]).length)?' ('+meta.schedule.items.length+')':''}</button>
        ${merged?`<button class="btn pc-act" data-unmerge="${esc(g.projk)}" title="Undo merge">↩ Unmerge</button>`:''}
        <button class="btn pc-act" data-edit-proj="${esc(g.projk)}">${ed?'Cancel':'✎ Edit'}</button>
        <button class="btn pc-act danger" data-delproj="${esc(g.projk)}" title="Delete this project">🗑 Delete</button>
      </div>
    </div>
    ${metaRow}
    <details class="pc-tasks">
      <summary>Tasks and owners (${g.tasks.length})</summary>
      <div class="pc-tasklist">${g.tasks.map(catalogTaskRow).join('')}</div>
    </details>
  </div>`;
}
function memberOptionsByRole(excludeIds){
  const ex=new Set(excludeIds||[]);
  const groups={}; members.filter(m=>!ex.has(m.id)).forEach(m=>{ const r=m.role||'Unassigned'; (groups[r]=groups[r]||[]).push(m); });
  return [...ROLES,'Unassigned'].filter(r=>groups[r]).map(r=>
    `<optgroup label="${r}">${groups[r].map(m=>`<option value="${m.id}">${esc(m.name)}</option>`).join('')}</optgroup>`).join('');
}
function catalogTaskRow(t){
  const owners=(t.ownerIds||[]).map(id=>`<span class="owner-chip">${memberRoleBadges(id)}${esc(memberName(id))}<button data-rmowner="${t.id}|${id}" title="Remove">×</button></span>`).join('')||'<span class="owner-chip none">Unassigned</span>';
  return `<div class="ctask">
    <div class="ctask-desc" data-opentask="${t.id}">${esc((t.current||t.next||'(no description)').slice(0,150))}
      <span class="tag st-${t.status}">${esc(t.status)}</span>${isClosed(t)?' <span class="tag closed">✓ Closed</span>':''}</div>
    <div class="ctask-ctl">
      <span class="owners">${owners}<select class="add-owner" data-addowner="${t.id}"><option value="">＋ Add owner</option>${memberOptionsByRole(t.ownerIds)}</select></span>
      <label>Risk<select data-edit="risk" data-tid="${t.id}">${optTags(['Low','Medium','High'],t.risk)}</select></label>
      <label>Cx<select data-edit="complexity" data-tid="${t.id}">${optTags(['Low','Medium','High'],t.complexity)}</select></label>
      <label>%<input type="number" min="0" max="100" value="${t.progress}" data-edit="progress" data-tid="${t.id}" class="prog-in"></label>
      <button class="ctask-del" data-del-task="${t.id}" title="Delete this task">🗑</button>
    </div>
  </div>`;
}
function editTaskField(id, field, val){
  const t=tasks.find(x=>x.id===id); if(!t) return;
  if(field==='progress') t.progress=Math.max(0,Math.min(100,+val||0));
  else if(field==='nextProgress') t.nextProgress=Math.max(0,Math.min(100,+val||0));
  else if(field==='project'){ const g=projectGroups().find(x=>x.projk===val);
    if(g){ t.projk=g.projk; t.projectLabel=g.label; t.project=(projMeta[g.projk]||{}).code||g.label; }   // picked from dropdown
    else { t.project=val; t.projk=projKeyOf(val); t.projectLabel=projLabelOf(val); } }                    // free-text fallback
  else t[field]=val;
  if(field==='progress'||field==='risk'||field==='complexity'||field==='nextProgress') t.manualEdit=true;  // survive re-import
  if(field==='current'||field==='next'||field==='analysis') t.manualText=true;      // keep user edits
  persist(); renderStats(); renderCatalog(); renderMembersArea();
  // don't re-render the modal for free-text edits (would steal focus mid-typing)
  if(!$('#taskModal').hidden && _openTaskId===id && !['project','current','next','analysis'].includes(field)) openTask(id);
}
function cloudDeleteImage(imgId){                       // also remove from the cloud images collection
  if(imgId && typeof CLOUD!=='undefined' && CLOUD.on && CLOUD.db){
    CLOUD.upImgs.delete(imgId);
    CLOUD.db.collection('images').doc(imgId).delete().catch(e=>console.warn('image delete failed', e));
  }
}
function clearTaskImages(tid){
  const t=tasks.find(x=>x.id===tid); if(!t||!(t.images||[]).length) return;
  if(!confirm('Clear all '+t.images.length+' images on this task?')) return;
  (t.images||[]).forEach(im=>cloudDeleteImage(im.id));
  t.images=[];
  persist(); renderMembersArea(); if(_openTaskId===tid && !$('#taskModal').hidden) openTask(tid);
  toast('All images cleared');
}
function removeTaskImage(tid, imgId){
  const t=tasks.find(x=>x.id===tid); if(!t||!t.images) return;
  if(!confirm('Delete this image? This cannot be undone.')) return;   // confirm to avoid accidental deletion
  t.images=t.images.filter(im=>im.id!==imgId);
  cloudDeleteImage(imgId);
  persist(); renderMembersArea(); if(_openTaskId===tid && !$('#taskModal').hidden) openTask(tid);
  toast('Image removed');
}
async function addTaskImages(tid, fileList){
  const t=tasks.find(x=>x.id===tid); if(!t) return;
  const files=Array.from(fileList||[]).filter(f=>/^image\//.test(f.type)); if(!files.length) return;
  t.images=t.images||[];
  for(const f of files){ const url=await fileToDataURL(f); const s=await shrinkImageBudget(url, 2200); t.images.push({id:uid(), data:s.data, w:s.w, h:s.h}); }
  persist(); renderMembersArea(); if(_openTaskId===tid && !$('#taskModal').hidden) openTask(tid);
  toast('Added '+files.length+' image(s)');
}
function addTaskOwner(id, mid){
  const t=tasks.find(x=>x.id===id); if(!t||!mid) return;
  t.ownerIds=t.ownerIds||[]; if(!t.ownerIds.includes(mid)) t.ownerIds.push(mid);
  t.manualOwners=true; t.shared=t.ownerIds.length>1; persist(); renderAll();
}
function removeTaskOwner(id, mid){
  const t=tasks.find(x=>x.id===id); if(!t) return;
  t.ownerIds=(t.ownerIds||[]).filter(x=>x!==mid);
  t.manualOwners=true; t.shared=t.ownerIds.length>1; persist(); renderAll();
}
function saveProjMeta(projk, card){
  const m=Object.assign({}, projMeta[projk]||{});   // keep existing fields (master/code/phase/chipset/schedule…)
  card.querySelectorAll('.pc-in').forEach(inp=>{ m[inp.dataset.mf]=inp.value.trim(); });
  projMeta[projk]=m; editingProj=''; persist(); renderCatalog();
  const mo=$('#projEditModal'); if(mo) mo.hidden=true; toast('Project saved');
}
// project edit happens in a modal now (the Cards view was retired; Tree/Matrix are the read views)
function openProjEdit(projk){
  const g=projectGroups().find(x=>x.projk===resolveProjk(projk)) || projectGroups().find(x=>x.projk===projk);
  if(!g){ toast('Project not found'); return; }
  editingProj=g.projk;
  const body=$('#projEditBody'); if(body) body.innerHTML=catalogCard(g);
  const mo=$('#projEditModal'); if(mo) mo.hidden=false;
}
function closeProjEdit(){ const mo=$('#projEditModal'); if(mo) mo.hidden=true; editingProj=''; renderCatalog(); }
// Dashboard: high-risk / overdue / in-progress highlight columns (each task clickable)
function renderHighlights(){
  const cont=$('#dashHighlights'); if(!cont) return;
  const vt=visibleTasks();
  const high=vt.filter(t=>t.risk==='High'&&!isClosed(t));
  const today=new Date(); today.setHours(0,0,0,0);
  const overdue=vt.filter(t=>{ const d=_sd(t.due); return d && d<today && !isClosed(t); });
  const wip=vt.filter(t=>!isClosed(t) && (+t.progress||0)>0 && (+t.progress||0)<100);
  const item=t=>`<button class="hl-item" data-opentask="${esc(t.id)}"><span class="hl-proj">${esc(pptPlabel(t))}</span><span class="hl-txt">${esc((t.current||t.next||'(no description)').replace(/\s+/g,' ').trim().slice(0,44))}</span></button>`;
  const col=(icon,title,arr,cls)=>`<div class="hl-col"><div class="hl-head ${cls}">${icon} ${title} <span class="hl-n">${arr.length}</span></div>${arr.slice(0,6).map(item).join('')||'<div class="hl-empty">None 🎉</div>'}${arr.length>6?`<div class="hl-more">+${arr.length-6} more</div>`:''}</div>`;
  cont.innerHTML=col('⚠️','High risk',high,'warn')+col('⏰','Overdue',overdue,'warn')+col('🔧','In progress',wip,'accent');
}
// Team: member cards (role, task load, avg progress) — click to see their tasks
function renderTeam(){
  const cont=$('#teamCards'); if(!cont) return;
  const {map}=buildBuckets();
  const tc=$('#teamCount'); if(tc) tc.textContent=members.length;
  cont.innerHTML=members.map(m=>{
    const ts=map.get(m.id)||[];
    const high=ts.filter(t=>t.risk==='High'&&!isClosed(t)).length;
    const closed=ts.filter(isClosed).length;
    const prog=ts.length? Math.round(ts.reduce((s,t)=>s+(+t.progress||0),0)/ts.length):0;
    const roles=[m.role,m.role2].filter(Boolean).join(' · ')||'—';
    return `<button class="tmember" data-memrow="${esc(m.id)}" title="See ${esc(m.name)}'s tasks">
      <div class="tm-top"><div class="tm-avatar">${esc((m.name||'?').trim().slice(0,1).toUpperCase())}</div>
        <div class="tm-id"><div class="tm-name">${esc(memberDisplay(m))}</div><div class="tm-role">${esc(roles)}</div></div></div>
      <div class="tm-stats"><span class="tm-tasks">${ts.length} task${ts.length===1?'':'s'}</span>${high?`<span class="tm-high">⚠ ${high}</span>`:''}${closed?`<span class="tm-done">✓ ${closed}</span>`:''}</div>
      <div class="tm-bar" title="avg progress ${prog}%"><div class="tm-fill" style="width:${prog}%"></div></div></button>`;
  }).join('') || '<p class="hint">No members yet — add them in “Manage members” below.</p>';
}
function renderCharts(){
  // workload per member — only members with tasks; note how many are pending
  const {map,unassigned}=buildBuckets();
  let wl=members.map(m=>({id:m.id, name:m.name, c:map.get(m.id).length}));   // ALL members in left-list order (0-task shown too)
  if(unassigned.length) wl.push({id:'__un__', name:'Unassigned', c:unassigned.length});
  const active=members.filter(m=>map.get(m.id).length>0).length;
  const maxW=Math.max(1,...wl.map(x=>x.c));
  $('#wlChartSub').textContent = `${active}/${members.length} with tasks · click a name to see theirs`;
  $('#workloadChart').innerHTML = wl.length
    ? wl.map(x=>chartBar(x.name+(x.c===0?' (pending)':''), Math.round(x.c/maxW*100), x.c, 'wl', {mem:x.id})).join('')
    : '<p class="hint">No data</p>';
}
// one chart row. cls 'wl' = workload (purple); else colored by progress. opts.proj/opts.mem make it clickable.
function chartBar(label, pct, valText, cls, opts){
  opts=opts||{};
  const fill = cls==='wl' ? 'wl' : progClass(pct);
  const done = cls!=='wl' && pct>=100;
  const data = opts.proj!==undefined ? `data-proj="${esc(opts.proj)}"` : opts.mem!==undefined ? `data-memrow="${esc(opts.mem)}"` : '';
  return `<div class="prow${data?' clickable':''}" ${data}><span class="pname" title="${esc(label)}">${esc(label)}</span>
    <span class="ptrack"><span class="pfill ${fill}" style="width:${Math.max(0,Math.min(100,pct))}%"></span></span>
    <span class="pval ${done?'done':''}">${esc(valText)}</span></div>`;
}

function renderMembersArea(){
  const {map,unassigned}=buildBuckets();
  const q=filters.q.toLowerCase();
  const matchStatus=t=>{
    const s=filters.status; if(!s) return true;
    if(s==='closed') return isClosed(t);
    if(s==='active') return !isClosed(t);
    if(s==='highrisk') return t.risk==='High' && !isClosed(t);
    return t.status===s;                        // New / Updated
  };
  const matchTask=t=>
    (!filters.project || resolveProjk(t.projk||t.key)===filters.project) &&
    matchStatus(t) &&
    (!q || (t.project+' '+(t.current||'')+' '+(t.next||'')).toLowerCase().includes(q));
  const filtering = filters.project || q || filters.status || filters.role;   // when filtering, hide members with no match
  let html=''; const shown=new Set();

  if(taskGroupBy!=='member'){                                   // group by Project or Status (flat task list)
    const roleOf=id=>{ const mm=members.find(x=>x.id===id); return mm?[mm.role,mm.role2]:[]; };
    const memOk=t=>{ if(!filters.member) return true; if(filters.member==='__un__') return !(t.ownerIds||[]).length; return (t.ownerIds||[]).includes(filters.member); };
    const filt=visibleTasks().filter(t=> matchTask(t) && memOk(t) && (!filters.role || (t.ownerIds||[]).some(id=>roleOf(id).includes(filters.role))) );
    filt.forEach(t=>shown.add(t.id));
    let groups=[];
    if(taskGroupBy==='project'){
      const by={}; filt.forEach(t=>{ const k=pptPlabel(t)||'—'; (by[k]=by[k]||[]).push(t); });
      groups=Object.keys(by).sort().map(k=>[k,by[k]]);
    } else {                                                    // status board: High risk / In progress / Closed
      const b={'⚠️ High risk':[], '🔧 In progress':[], '✅ Closed':[]};
      filt.forEach(t=>{ if(isClosed(t)) b['✅ Closed'].push(t); else if(t.risk==='High') b['⚠️ High risk'].push(t); else b['🔧 In progress'].push(t); });
      groups=Object.keys(b).filter(k=>b[k].length).map(k=>[k,b[k]]);
    }
    html=groups.map(([title,ts])=>groupBlock(title,ts)).join('');
  } else {
    members.forEach(m=>{
      if(filters.member && filters.member!==m.id) return;
      if(filters.role && m.role!==filters.role && m.role2!==filters.role) return;  // primary OR secondary discipline
      const list=map.get(m.id).filter(matchTask); list.forEach(t=>shown.add(t.id));
      if((filters.hideEmpty || filtering) && !list.length) return;
      html+=memberBlock(memberDisplay(m), list, false, m.id);
    });
    if((!filters.member || filters.member==='__un__') && !filters.role){
      const ul=unassigned.filter(matchTask); ul.forEach(t=>shown.add(t.id));
      if(ul.length) html+=memberBlock('Unassigned', ul, true);
    }
  }

  $('#membersArea').innerHTML = html ||
    '<div class="panel"><p class="hint">No matching items (or no members added / reports imported yet).</p></div>';
  const fc=$('#filterCount'); if(fc) fc.textContent='Showing '+shown.size+' tasks';
}
function memberBlock(name, list, isUnassigned, mid){
  const avg=list.length?Math.round(list.reduce((s,t)=>s+(+t.progress||0),0)/list.length):0;
  // aggregate this member's detail/analysis images into a header strip (easy to find, not buried in one card)
  const imgs=[]; const seen=new Set();
  list.forEach(t=>(t.images||[]).forEach(im=>{ if(im&&im.id&&!seen.has(im.id)){ seen.add(im.id); imgs.push(im); } }));
  const strip = imgs.length? `<div class="member-imgs" title="This member's images / analysis">📎 ${imgs.slice(0,10).map(im=>`<img src="${im.data}" data-light="${im.id}">`).join('')}${imgs.length>10?`<span class="more">+${imgs.length-10}</span>`:''}</div>` : '';
  const head=`<div class="member-head ${isUnassigned?'unassigned':''}">
    ${mid?memberRoleBadges(mid):''}<span class="name">${esc(name)}</span>
    <span class="meta">${list.length} tasks · avg ${avg}%</span>
    ${list.length?'':'<span class="pending">Pending input</span>'}
    ${mid?`<button class="btn sm add-task-btn" data-addtask="${mid}">＋ Add task</button>`:''}</div>${strip}`;
  const cards=list.length? `<div class="task-grid">${list.map(taskCard).join('')}</div>` : '';
  return `<div class="member-block" id="mblock-${mid||'unassigned'}">${head}${cards}</div>`;
}
// generic group block (used when Tasks are grouped by Project or Status instead of Member)
function groupBlock(title, list){
  const avg=list.length?Math.round(list.reduce((s,t)=>s+(+t.progress||0),0)/list.length):0;
  const head=`<div class="member-head"><span class="name">${esc(title)}</span><span class="meta">${list.length} task${list.length===1?'':'s'} · avg ${avg}%</span></div>`;
  return `<div class="member-block">${head}<div class="task-grid">${list.map(taskCard).join('')}</div></div>`;
}
// clicking a member in the left list jumps to their report on the right (reveals + highlights)
function jumpToMember(mid){
  if(!mid) return;
  setView('members');
  if(filters.member && filters.member!==mid) filters.member='';
  if(filters.hideEmpty){ filters.hideEmpty=false; const he=$('#hideEmptyChk'); if(he) he.checked=false; }
  renderFilters(); renderMembersArea(); renderStats();
  let block=document.getElementById('mblock-'+mid);
  if(!block){ filters.role=''; filters.status=''; filters.q=''; const fq=$('#filterQ'); if(fq) fq.value=''; renderFilters(); renderMembersArea(); block=document.getElementById('mblock-'+mid); }
  if(!block){ toast('No report found for that member'); return; }
  requestAnimationFrame(()=>{
    const mn=document.querySelector('.main');
    const desktop = mn && getComputedStyle(mn).overflowY!=='visible' && mn.scrollHeight>mn.clientHeight+4;
    if(desktop){                                   // independent-scroll column -> scroll it directly to the top
      const delta=block.getBoundingClientRect().top - mn.getBoundingClientRect().top;
      mn.scrollTop = mn.scrollTop + delta - 12;
    } else {                                        // mobile: normal page scroll
      block.scrollIntoView({block:'start'});
    }
  });
  block.classList.remove('flash'); void block.offsetWidth; block.classList.add('flash');
  setTimeout(()=>{ const b=document.getElementById('mblock-'+mid); if(b) b.classList.remove('flash'); },1600);
}
function progBar(val, closed){
  val=Math.max(0,Math.min(100,val||0));
  return `<div class="prog"><span class="ptrack"><span class="pfill ${progClass(val)}" style="width:${val}%"></span></span><span class="pval ${closed?'done':''}">${val}%</span></div>`;
}
// This-week and Next-week shown as separate rows, each with its own progress bar
function weekRows(t){
  const closed=isClosed(t), hasCur=!!(t.current||'').trim(), hasNext=!!(t.next||'').trim();
  let h='';
  if(hasCur){ h+=`<div class="wk-row"><span class="wk-tag now">This week</span><span class="wk-desc">${esc(t.current)}</span></div>`+progBar(t.progress,closed); }
  if(hasNext){
    const np=(t.nextProgress!=null)?t.nextProgress:(hasCur?0:(t.progress||0));
    h+=`<div class="wk-row"><span class="wk-tag nxt">Next week</span><span class="wk-desc">${esc(t.next)}</span></div>`+progBar(np,false);
  }
  if(!hasCur && !hasNext){ h+=`<div class="wk-row"><span class="wk-desc">(no description)</span></div>`+progBar(t.progress,closed); }
  return h;
}
function taskCard(t){
  const sharedTag=t.shared?`<span class="tag shared">Shared owner</span>`:'';
  const thumbs=(t.images&&t.images.length)?`<div class="img-badge" title="${t.images.length} image(s) (see top of member)">📎 ${t.images.length}</div>`:'';
  const delta=t.delta?`<div class="delta">⟳ Weekly delta:${esc(t.delta)}</div>`:'';
  const closed=isClosed(t);
  const shortP=t.projectLabel || String(t.project||'').split(/[\n(]/)[0].trim();
  return `<div class="card ${closed?'closed':''}" data-task="${t.id}">
    <button class="card-del" data-del-task="${t.id}" title="Delete task">🗑</button>
    <div class="ct"><span class="proj">${esc(shortP)}</span></div>
    <div class="tags">
      ${closed?'<span class="tag closed">✓ Closed</span>':''}
      <span class="tag risk-${t.risk}">Risk ${esc(t.risk)}</span>
      <span class="tag cx">Cx ${esc(t.complexity)}</span>
      ${sharedTag}
      <span class="tag st-${t.status}">${esc(t.status)}</span>
    </div>
    ${weekRows(t)}
    <div class="due">📅 ${esc(t.due||'—')} ${t.reporter?'· Reporter: '+esc(t.reporter):''}</div>
    ${delta}${thumbs}
  </div>`;
}

/* ---------- task detail modal ---------- */
let _openTaskId='';
function openTask(id){
  const t=tasks.find(x=>x.id===id); if(!t) return;
  document.querySelectorAll('#taskImgsBox img').forEach(im=>{ try{ im.src=''; }catch(e){} });   // release the previous task's decoded image bitmaps so memory doesn't pile up on rapid opens
  _openTaskId=id;
  const owners=(t.ownerIds||[]).map(memberName).join(', ')||'—';
  $('#taskModalInner').innerHTML=`
    <div class="modal-head"><h2>${esc(t.projectLabel||t.project)}</h2><button class="icon-btn" data-close>✕</button></div>
    <div class="modal-body detail">
      <div class="tags">
        ${isClosed(t)?'<span class="tag closed">✓ Closed</span>':''}
        <span class="tag risk-${t.risk}">Risk ${esc(t.risk)}</span>
        <span class="tag cx">Complexity ${esc(t.complexity)}</span>
        ${t.shared?'<span class="tag shared">Shared owner</span>':''}
        <span class="tag st-${t.status}">${esc(t.status)}</span>
      </div>
      <div class="edit-row">
        <label>This week %<input type="number" min="0" max="100" value="${t.progress}" data-edit="progress" data-tid="${t.id}" class="prog-in"></label>
        <label>Next week %<input type="number" min="0" max="100" value="${t.nextProgress!=null?t.nextProgress:''}" data-edit="nextProgress" data-tid="${t.id}" class="prog-in" placeholder="—"></label>
        <label>Risk<select data-edit="risk" data-tid="${t.id}">${optTags(['Low','Medium','High'],t.risk)}</select></label>
        <label>Complexity<select data-edit="complexity" data-tid="${t.id}">${optTags(['Low','Medium','High'],t.complexity)}</select></label>
      </div>
      <div class="kv">
        <span class="k">Project</span><span><select class="proj-edit" data-edit="project" data-tid="${t.id}">${projectOptionsHTML(t.projk)}</select></span>
        <span class="k">Status</span><span>${esc(statusLine(t))}</span>
        <span class="k">Due date</span><span>${esc(t.due||'—')}</span>
        <span class="k">Owner (raw)</span><span>${esc(t.rawOwner||'—')}</span>
        <span class="k">Reporter</span><span>${esc(t.reporter||'—')}</span>
        <span class="k">Assigned members</span><span>${esc(owners)}</span>
        ${t.unmatched&&t.unmatched.length?`<span class="k">Unmatched</span><span style="color:var(--warn)">${esc(t.unmatched.join(', '))}</span>`:''}
        <span class="k">Source</span><span>${esc(t.source||'manual')}</span>
      </div>
      <div class="section-title">This week (editable)</div>
      <textarea class="task-edit" data-edit="current" data-tid="${t.id}" rows="3" placeholder="This week's work">${esc(t.current||'')}</textarea>
      <div class="section-title">Next week (editable)</div>
      <textarea class="task-edit" data-edit="next" data-tid="${t.id}" rows="2" placeholder="Next week plan">${esc(t.next||'')}</textarea>
      ${t.delta?`<div class="section-title">Weekly delta</div><div style="color:var(--warn)">${esc(t.delta)}</div>`:''}
      ${t.prev?`<div class="section-title">Previous</div><div class="hint">${esc(t.prev.current||'—')} (was ${t.prev.progress}%, risk ${t.prev.risk})</div>`:''}
      <div class="section-title">Issue analysis (editable)</div>
      <textarea class="task-edit analysis-edit" data-edit="analysis" data-tid="${t.id}" rows="3" placeholder="Issue analysis">${esc(t.analysis||generateAnalysis(t))}</textarea>
      <div class="section-title">Attachments (delete one-by-one or clear all; exports and cloud stay in sync)
        ${(t.images||[]).length?`<button class="btn xs danger clearimg-btn" data-clearimg="${t.id}">🗑 Clear all (${t.images.length})</button>`:''}</div>
      <div class="imgs editable" id="taskImgsBox">
        <label class="img-add" title="Add / replace images">＋ Image<input type="file" accept="image/*" multiple hidden data-addimg="${t.id}"></label>
      </div>
    </div>
    <div class="modal-foot">
      <button class="btn danger" data-del-task="${t.id}">Delete task</button>
      ${(t.images&&t.images.length)?`<button class="btn" data-ocr="${t.id}">🔍 OCR image text</button>`:''}
      <button class="btn primary" data-export-member="${(t.ownerIds||[])[0]||''}">Export this member to Word</button>
      <button class="btn" data-close>Close</button>
    </div>`;
  $('#taskModal').hidden=false;
  loadTaskImages(t);   // decode attachments OFF the main thread (img.decode), then insert — open is instant & the NEXT open never blocks
}
// decode each base64 attachment off-thread, then append the already-decoded image (cheap paint, no main-thread stall)
const _thumbCache={};                                    // image id -> small dataURL; the modal shows THUMBS (full image only in the lightbox)
async function thumbFor(im){
  if(!im||!im.data) return '';
  if(_thumbCache[im.id]) return _thumbCache[im.id];
  try{                                                   // decode + downscale OFF the main thread (createImageBitmap); main thread only draws a 240px tile
    const blob=await (await fetch(im.data)).blob();
    let bmp=await createImageBitmap(blob);
    const r=Math.min(1, 240/Math.max(bmp.width,bmp.height));
    if(r<1){ const small=await createImageBitmap(bmp,{resizeWidth:Math.max(1,Math.round(bmp.width*r)),resizeHeight:Math.max(1,Math.round(bmp.height*r)),resizeQuality:'medium'}); if(bmp.close)bmp.close(); bmp=small; }
    const c=document.createElement('canvas'); c.width=bmp.width; c.height=bmp.height;
    c.getContext('2d').drawImage(bmp,0,0); if(bmp.close)bmp.close();
    const th=c.toDataURL('image/jpeg',.72); _thumbCache[im.id]=th; return th;
  }catch(e){ return im.data; }                           // fallback: full image rather than nothing
}
function loadTaskImages(t){
  const box=$('#taskImgsBox'); if(!box) return; const add=box.querySelector('.img-add');
  (t.images||[]).forEach(im=>{
    thumbFor(im).then(src=>{
      if(_openTaskId!==t.id || !box.isConnected || !box.contains(add) || !src) return;
      const span=document.createElement('span'); span.className='img-edit';
      const img=new Image(); img.decoding='async'; img.setAttribute('data-light', im.id); img.src=src;
      span.appendChild(img);
      const del=document.createElement('button'); del.className='img-del';
      del.setAttribute('data-delimg', t.id+'|'+im.id); del.title='Delete this image'; del.textContent='✕';
      span.appendChild(del); box.insertBefore(span, add);
    });
  });
}
function imageDataById(id){                              // lightbox needs the FULL image for a thumb's id
  for(const t of tasks) for(const im of (t.images||[])) if(im&&im.id===id) return im.data;
  for(const k in projMeta) for(const im of ((projMeta[k]||{}).images||[])) if(im&&im.id===id) return im.data;
  return null;
}

/* ---------- project drill-down: all tasks under a project ---------- */
function openProject(projk){
  const list=tasks.filter(t=>!t.imageReport && resolveProjk(t.projk||t.key)===resolveProjk(projk));
  if(!list.length) return;
  const label=list[0].projectLabel||shortProj(list[0].project);
  const avg=list.length?Math.round(list.reduce((s,t)=>s+(+t.progress||0),0)/list.length):0;
  const closed=list.filter(isClosed).length;
  $('#taskModalInner').innerHTML=`
    <div class="modal-head"><h2>${esc(label)} <span class="ch-sub">${list.length} tasks · avg ${avg}% · ${closed} closed</span></h2>
      <button class="icon-btn" data-close>✕</button></div>
    <div class="modal-body">
      <div class="proj-tasklist">${list.map(t=>`
        <div class="ptask" data-task="${t.id}">
          <div class="pt-top">
            <span class="pt-owner">${esc((t.ownerIds||[]).map(memberName).join(', ')||'Unassigned')}</span>
            <span class="tags">
              ${isClosed(t)?'<span class="tag closed">✓ Closed</span>':`<span class="tag risk-${t.risk}">${esc(t.risk)}</span>`}
              <span class="tag st-${t.status}">${esc(t.status)}</span>
            </span>
          </div>
          <div class="pt-desc">${esc((t.current||t.next||'(no description)').slice(0,170))}</div>
          <div class="prog"><span class="ptrack"><span class="pfill ${progClass(t.progress)}" style="width:${t.progress}%"></span></span><span class="pval ${isClosed(t)?'done':''}">${t.progress}%</span></div>
        </div>`).join('')}
      </div>
    </div>
    <div class="modal-foot"><button class="btn" data-close>Close</button></div>`;
  $('#taskModal').hidden=false;
}

/* ---------- custom project groupings ---------- */
function parseAliasText(text){
  const out=[];
  text.split(/[\r\n]+/).forEach(line=>{ line=line.trim(); if(!line) return;
    let label, rest;
    if(/[::]/.test(line)){ const p=line.split(/[::]/); label=p[0].trim(); rest=p.slice(1).join(':'); }
    else { label=line; rest=line; }
    const tokens=[...new Set([label, ...rest.split(/[,, , \/]/)].map(s=>norm(s)).filter(Boolean))];
    if(label) out.push({key:norm(label), label, tokens});
  });
  return out;
}
function aliasToText(){ return projAliases.map(g=>g.label+': '+g.tokens.join(', ')).join('\n'); }
function applyProjAliases(text){
  projAliases=parseAliasText(text);
  tasks.forEach(t=>{ t.projk=projKeyOf(t.project); t.projectLabel=projLabelOf(t.project); }); // re-group existing
  persist(); renderAll(); toast('Project merge applied ('+projAliases.length+' groups)');
}

/* =====================================================================
   WORKBENCH
   ===================================================================== */
let wbImages=[];
function openWorkbench(preMemberId, preProjk){
  renderWorkbenchSelect();
  const wbm=$('#wbMember');
  if(preMemberId && members.some(m=>m.id===preMemberId)){ const c=wbm.querySelector(`[data-wbm="${preMemberId}"]`); if(c) c.classList.add('on'); }
  // members can only file under their own name
  if(CLOUD.me && !CLOUD.me.admin){
    const me=members.find(x=>x.name.toLowerCase()===CLOUD.me.name.toLowerCase());
    wbm.querySelectorAll('.wbm').forEach(c=>{ const mine=!!me&&c.dataset.wbm===me.id; c.classList.toggle('on',mine); c.disabled=!mine; });
  }
  $('#phraseRow').innerHTML=Object.keys(PHRASES).map(k=>`<button data-phrase="${esc(k)}">${esc(k)}</button>`).join('');
  wbImages=[]; renderWbThumbs();
  $('#wbProject').innerHTML=projectOptionsHTML(preProjk||'');       // project dropdown — known projects only; ＋ on a project pre-selects it
  ['#wbThisWeek','#wbIssue','#wbNext'].forEach(s=>$(s).value='');
  { const d=$('#wbDue'), dp=$('#wbDuePick'); if(d) d.value=''; if(dp) dp.value=''; }
  $('#wbProgress').value=0; $('#wbNextProgress').value=0;
  $('#workbenchModal').hidden=false;
}
function renderWorkbenchSelect(){
  const box=$('#wbMember'); if(!box) return;   // multi-select chips: click to toggle each member
  box.innerHTML=members.map(m=>`<button type="button" class="wbm" data-wbm="${m.id}">${esc(m.name)}</button>`).join('')||'<span class="hint">(add members first)</span>';
}
function renderWbThumbs(){
  $('#wbThumbs').innerHTML=wbImages.map((im,i)=>`<div class="th"><img src="${im.data}" data-light="${im.id}"><button class="rm" data-rm-wb="${i}">✕</button></div>`).join('');
}
function saveWorkbench(){
  const ids=[...document.querySelectorAll('#wbMember .wbm.on')].map(b=>b.dataset.wbm);
  if(!ids.length){ toast('Pick at least one member'); return; }
  const pv=$('#wbProject').value; const pg=projectGroups().find(x=>x.projk===pv);
  const project = pg ? ((projMeta[pg.projk]||{}).code||pg.label) : ((pv||'').trim()||'Untitled Project');
  const current=$('#wbThisWeek').value.trim();
  const names=ids.map(memberName), name=names.join(', ');
  const p={project, reporter:names[0], rawOwner:name, current,
    next:$('#wbNext').value.trim(), risk:$('#wbRisk').value, due:$('#wbDue').value,
    complexity:$('#wbComplexity').value, progress:+$('#wbProgress').value||0,
    nextProgress:+$('#wbNextProgress').value||0,
    _images:wbImages.slice()};
  if($('#wbIssue').value.trim()) p.current += (p.current?' ':'')+'Issue/Note: '+$('#wbIssue').value.trim();
  overlayTasks([p], 'Workbench manual entry');
  renderAll();
  $('#workbenchModal').hidden=true;
  toast('Task saved for '+name);
}

/* =====================================================================
   WORD (DOCX) EXPORT  — builds OOXML zip, embeds images
   ===================================================================== */
function pngSize(dataUrl){
  try{ const b=atob(dataUrl.split(',')[1]); if(b.slice(1,4)!=='PNG') return null;
    const w=(b.charCodeAt(16)<<24)|(b.charCodeAt(17)<<16)|(b.charCodeAt(18)<<8)|b.charCodeAt(19);
    const h=(b.charCodeAt(20)<<24)|(b.charCodeAt(21)<<16)|(b.charCodeAt(22)<<8)|b.charCodeAt(23);
    return {w,h}; }catch(e){ return null; }
}
const EMU=9525; // per pixel @96dpi
function wpDrawing(rid, im){
  const sz=(im && im.w && im.h)?{w:im.w,h:im.h}:(pngSize(im&&im.data)||{w:480,h:320});
  const maxW=4200000; // ~4.4in
  let cx=sz.w*EMU, cy=sz.h*EMU;
  if(cx>maxW){ cy=Math.round(cy*maxW/cx); cx=maxW; }
  const did=Math.floor(Math.random()*1e6);
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">
   <wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${did}" name="img${did}"/>
   <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
   <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
   <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
   <pic:nvPicPr><pic:cNvPr id="${did}" name="img${did}"/><pic:cNvPicPr/></pic:nvPicPr>
   <pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
   <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
   <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
   </a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}
function P(text, opts={}){
  const {bold,size,color,indent}=opts;
  const rpr=`<w:rPr>${bold?'<w:b/>':''}${size?`<w:sz w:val="${size*2}"/>`:''}${color?`<w:color w:val="${color}"/>`:''}</w:rPr>`;
  const ppr=`<w:pPr>${indent?`<w:ind w:left="${indent}"/>`:''}<w:spacing w:after="60"/></w:pPr>`;
  return `<w:p>${ppr}<w:r>${rpr}<w:t xml:space="preserve">${escXml(text)}</w:t></w:r></w:p>`;
}
function escXml(s){ return String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

function shortProj(p){ return String(p||'').split(/[(\n]/)[0].trim().slice(0,40); }
function memberReportXml(name, list, mediaCollector){
  let body=P('*'+name,{bold:true,size:14,color:'1F6FEB'});
  if(!list.length){ body+=P('Pending input — no items reported this week.',{color:'B5790F'}); body+=P('',{}); return body; }
  const plabel=t=>t.projectLabel||shortProj(t.project);
  const cur=list.filter(t=>t.current);
  const sh=projShares(list);
  body+=P('This week: [ '+sh.list.map(o=>`${o.label} - ${o.pct}%`).join(' | ')+' ]',{bold:true});   // per-member project breakdown (kept)
  cur.forEach((t,i)=>{
    body+=P(`${i+1}. ${plabel(t)}: ${rewriteProfessional(t.current)}`,{indent:200});
    body+=P(`Status: ${statusWord(t)} | risk ${(t.risk||'M')[0]} | Due Date: ${t.due||'TBC'}`,{indent:540,color:'444C5C'});
    // lean: only surface analysis when there is a real risk/blocker
    if(t.risk==='High' || /block|defect|timeout|fail|overdue|debug/i.test(t.current||''))
      body+=P('Analysis: '+(t.analysis||generateAnalysis(t)),{indent:540,color:'6B7A92'});
    if(t.shared) body+=P('Shared owner: '+(t.ownerIds||[]).map(memberName).join(', '),{indent:540,color:'1A7A8A'});
    if(t.delta) body+=P('Weekly delta: '+t.delta,{indent:540,color:'B5790F'});
  });
  // Next week
  const nexts=list.filter(t=>t.next);
  if(nexts.length){
    body+=P('Next week:',{bold:true});
    nexts.forEach((t,i)=>{ body+=P(`${i+1}. ${plabel(t)}: ${rewriteProfessional(t.next)}`,{indent:200}); });
  }
  body+=P('',{});
  return body;
}

/* =====================================================================
   PPTX EXPORT — fixed-layout deck styled like the company 6G deck
   (navy + teal, white cards w/ cyan accent, Arial). Same content model
   as the Word export so PPT and Word stay aligned.
   ===================================================================== */
const PPT = {
  // company palette pulled from the 6G deck
  navy:'1A3B79', dark:'1E2C3A', deep:'0C2340', blue:'2E5AAC',
  teal:'0E8597', cyan:'15B5CC', gray:'5E6B7A', tint:'EAF1FB',
  white:'FFFFFF', card:'F8FAFC', track:'E6ECF5', line:'D8E0EC',
  green:'10B981', amber:'F59E0B', font:'Arial', fontB:'Arial Black'
};
function pptPlabel(t){
  const m=projMeta[resolveProjk(t.projk||t.key)]||{};
  if(m.master && m.code) return (m.custProj && m.custProj!==m.code) ? (m.custProj+' - '+m.code) : m.code;   // e.g. "ID535 - B01W043T00"
  return t.projectLabel||shortProj(t.project);
}
function progColor(p){ return p>=100?PPT.green : p>=70?PPT.blue : p>=34?PPT.cyan : PPT.amber; }
function collectImages(list){
  const out=[], seen=new Set();
  (list||[]).forEach(t=>(t.images||[]).forEach(im=>{
    const k=im&&(im.id||im.data); if(im&&im.data&&!seen.has(k)){ seen.add(k); out.push(Object.assign({_lbl:pptPlabel(t)}, im)); }
  }));
  return out;
}
/* keep technical RD3 text VERBATIM — only fix mojibake & strip leading bullet/number */
function cleanRptLine(s){
  return String(s||'')
    .replace(/à/g,'→').replace(/â€™/g,"'").replace(/â†'/g,'→')
    .replace(/^\s*(?:[-–•*]|\d+[.)])\s*/,'')
    .replace(/[ \t]+/g,' ').trim();
}
function splitRptLines(text){
  return String(text||'').split(/\r?\n/).map(cleanRptLine).filter(Boolean);
}
function rptMarkerColor(inner){
  const s=String(inner).toLowerCase();
  if(/(fail|disqualif|block|crash|overdue|error|reject|defect|\bn\/?g\b)/.test(s)) return 'E5484D';
  if(/(done|closed|pass|qualif|verified|finish|complete|\bok\b)/.test(s))          return PPT.green;
  if(/(on-?going|ongoing|pending|progress|wip|tbd|wait)/.test(s))                  return PPT.amber;
  return null;
}
/* split a line into runs so status markers like (Done)/(On-going) get colored */
function rptMarkerRuns(line){
  const parts=[]; let last=0, mm; const re=/\(([^)]{1,40})\)/g;
  while((mm=re.exec(line))){
    if(mm.index>last) parts.push({text:line.slice(last,mm.index)});
    const col=rptMarkerColor(mm[1]);
    parts.push({text:mm[0], options: col?{color:col,bold:true}:null});
    last=mm.index+mm[0].length;
  }
  if(last<line.length) parts.push({text:line.slice(last)});
  return parts.length?parts:[{text:line}];
}
function rptEstLines(s){ return Math.max(1, Math.ceil(s.length/118)); }

function assemblePptx(memberIds){
  if(typeof PptxGenJS==='undefined'){ toast('PPTX library not loaded'); return null; }
  const {map,unassigned}=buildBuckets();
  const targets = memberIds&&memberIds.length
    ? members.filter(m=>memberIds.includes(m.id)) : members.slice();
  if(!targets.length){ toast('No members to export'); return null; }
  const date=new Date().toISOString().slice(0,10);

  const pptx=new PptxGenJS();
  pptx.defineLayout({name:'WIDE', width:13.33, height:7.5});
  pptx.layout='WIDE';
  pptx.author='Weekly Report Hub'; pptx.company='RD'; pptx.title='Weekly Report '+date;
  const R=pptx.ShapeType.rect, RR=pptx.ShapeType.roundRect;

  // ---- shared chrome --------------------------------------------------
  function header(s, name, role, tag){
    s.background={color:PPT.white};
    s.addShape(R,{x:0.7,y:0.34,w:0.14,h:0.14,fill:{color:PPT.cyan}});
    s.addText(`WEEKLY REPORT  ｜  ${date}`,
      {x:0.95,y:0.26,w:9,h:0.3,fontSize:11,bold:true,color:PPT.teal,fontFace:PPT.font,charSpacing:2});
    s.addText(name,{x:0.7,y:0.6,w:10.6,h:0.64,fontSize:28,bold:true,color:PPT.dark,fontFace:PPT.fontB,valign:'middle'});
    if(tag) s.addText(tag,{x:10.8,y:0.72,w:1.83,h:0.36,fontSize:12,bold:true,color:PPT.gray,fontFace:PPT.font,align:'right',valign:'middle'});
    s.addShape(R,{x:0.7,y:1.4,w:11.93,h:0.022,fill:{color:PPT.line}});
    s.addText('Weekly Report Hub · '+date,{x:0.7,y:7.06,w:11.93,h:0.3,fontSize:8,color:PPT.gray,fontFace:PPT.font,align:'right'});
  }
  function sectionLabel(s, text, y){
    s.addShape(R,{x:0.7,y:y+0.04,w:0.12,h:0.22,fill:{color:PPT.cyan}});
    s.addText(text,{x:0.92,y,w:11.7,h:0.3,fontSize:13,bold:true,color:PPT.navy,fontFace:PPT.font,charSpacing:1});
  }
  // collect ALL of a member's tasks for one project into a single task-card model
  function groupAllByProject(list){
    const order=[], g=new Map();
    (list||[]).forEach(t=>{
      const key=pptPlabel(t);
      if(!g.has(key)){ g.set(key,{label:key,cur:[],next:[],progs:[],high:false,due:'',imgs:[],seen:new Set()}); order.push(key); }
      const o=g.get(key);
      if(String(t.current||'').trim()) splitRptLines(t.current).forEach(l=>o.cur.push(l));
      if(String(t.next||'').trim()) splitRptLines(t.next).forEach(l=>o.next.push(l));
      if(typeof t.progress==='number') o.progs.push(t.progress);
      if(!o.due && String(t.due||'').trim()) o.due=String(t.due).trim();
      if(t.risk==='High') o.high=true;
      (t.images||[]).forEach(im=>{ const k=im&&(im.id||im.data); if(im&&im.data&&!o.seen.has(k)){ o.seen.add(k); o.imgs.push(im); } });
    });
    return order.map(k=>g.get(k));
  }
  function groupStatus(g){
    const txt=g.cur.concat(g.next).join(' ').toLowerCase();
    const maxp=g.progs.length?Math.max.apply(null,g.progs):0;
    if(g.high || /\b(blocked|blocker|stuck|crash|overdue)\b/.test(txt))                return {text:'At risk',color:'E5484D'};
    if(/\b(pending|waiting|tbd)\b/.test(txt) || /on hold/.test(txt))                   return {text:'Pending',color:PPT.amber};
    if(/\b(on-?going|ongoing|in progress|wip)\b/.test(txt))                            return {text:'In-progress',color:PPT.teal};
    if(!g.next.length && (maxp>=100 || /\b(done|closed|completed)\b/.test(txt)))        return {text:'Done',color:PPT.green};
    return {text:'In-progress',color:PPT.teal};
  }
  // shared table geometry so EVERY table (this-week / next-week / every member)
  // has identical column widths and styling — neat & consistent (等寬).
  const COLW=[1.95,6.78,1.3,1.9];          // Project | Job & Issue | Due date | Status  (=11.93)
  const TFONT=11, HFONT=12, CPCELL=88;     // bigger, clearer font
  const MAXY=7.0;                          // bottom limit for content before a new slide
  // table cell → bulleted runs, status markers coloured, technical text verbatim
  function cellRuns(lines){
    const arr=(lines&&lines.length?lines:['—']).slice(0,10).map(l=>l.length>200?l.slice(0,197)+'…':l);
    const runs=[];
    arr.forEach(line=>{ const parts=rptMarkerRuns(line);
      parts.forEach((p,i)=>runs.push({text:(i===0?'• ':'')+p.text, options:{
        color:(p.options&&p.options.color)||PPT.dark, bold:!!(p.options&&p.options.bold),
        breakLine:i===parts.length-1 }})); });
    return runs;
  }
  function estRowH(lines){
    const c=(lines&&lines.length?lines:['—']).slice(0,10)
      .reduce((a,l)=>a+Math.max(1,Math.ceil(Math.min(l.length,200)/CPCELL)),0);
    return c*0.19 + 0.12;   // calibrated to the tighter cell margins below
  }
  // one image card (white card, cyan top bar, contained image, scaled caption) — works at any cell size
  function imgCard(s,x,y,w,h,im,n){
    const topBar=0.06, pad=Math.min(0.15, w*0.05), capH=Math.min(0.30, Math.max(0.20, h*0.12));
    s.addShape(RR,{x,y,w,h,rectRadius:0.05,fill:{color:PPT.white},line:{color:PPT.line,width:1}});
    s.addShape(R,{x,y,w,h:topBar,fill:{color:PPT.cyan}});
    const iw=w-2*pad, ih=h-topBar-pad-capH;
    try{ s.addImage({data:im.data,x:x+pad,y:y+topBar+pad*0.5,w:iw,h:ih,sizing:{type:'contain',w:iw,h:ih}}); }catch(e){}
    const fs=Math.max(8, Math.min(11, Math.round(h*3.6)));
    const cap='Fig '+n+(im._lbl? '  ·  '+String(im._lbl).slice(0,24) : '');
    s.addText(cap,{x:x+pad,y:y+h-capH,w:iw,h:capH-0.02,fontSize:fs,color:PPT.gray,fontFace:PPT.font,valign:'middle'});
  }
  // a member's images → ADAPTIVE grid per slide, sized by image count + aspect ratio.
  // chooseGrid maximises how big the images render while keeping every cell above a legible
  // minimum (MINW/MINH); when more images won't fit at that minimum they spill onto extra slides.
  function avgAspect(chunk){
    let a=0,n=0; chunk.forEach(im=>{ const r=(im.w&&im.h)?im.w/im.h:1.4; if(r>0&&isFinite(r)){a+=r;n++;} });
    return n? a/n : 1.4;
  }
  function chooseGrid(k, ar, W, H, GAP, CAPH, MINW, MINH){
    let best=null;
    for(let cols=1; cols<=k; cols++){
      const rows=Math.ceil(k/cols), cw=(W-(cols-1)*GAP)/cols, ch=(H-(rows-1)*GAP)/rows;
      if(cw<MINW || ch<MINH) continue;                 // cell too small to read -> reject (forces a spill)
      const aw=cw-0.30, ah=ch-0.30-CAPH;               // space left for the image after padding + caption
      let dw=aw, dh=aw/ar; if(dh>ah){ dh=ah; dw=ah*ar; }   // contain the aspect ratio in the cell
      const area=Math.max(0,dw)*Math.max(0,dh);
      if(!best || area>best.area) best={cols,rows,area};
    }
    return best;
  }
  function attachPages(name, role, imgs){
    if(!imgs||!imgs.length) return;
    const X0=0.7, Y0=1.98, W=11.93, H=4.95, GAP=0.18, CAPH=0.30;
    const MINW=2.5, MINH=1.45, MAXPP=2;                 // at most 2 imgs/slide; ratio is NEVER distorted (contain) — if 2 won't fit at true ratio, drop to 1 + new page
    let i=0, page=0;
    while(i<imgs.length){
      let take=Math.min(MAXPP, imgs.length-i), grid=null;
      while(take>=1){                                   // shrink the chunk until a legible grid fits
        grid=chooseGrid(take, avgAspect(imgs.slice(i,i+take)), W,H,GAP,CAPH,MINW,MINH);
        if(grid) break; take--;
      }
      if(!grid){ take=1; grid={cols:1,rows:1}; }
      const chunk=imgs.slice(i,i+take);
      const cw=(W-(grid.cols-1)*GAP)/grid.cols, ch=(H-(grid.rows-1)*GAP)/grid.rows;
      const s=pptx.addSlide(); header(s,name,role, page?'Attachments (cont.)':'Attachments');
      sectionLabel(s,'ATTACHMENTS  ｜  Issue images / measurements',1.55);
      chunk.forEach((im,kk)=>{
        const r=Math.floor(kk/grid.cols), c=kk%grid.cols;
        const inRow=(r===grid.rows-1)?(chunk.length-(grid.rows-1)*grid.cols):grid.cols;
        const rowW=inRow*cw+(inRow-1)*GAP, xs=X0+(W-rowW)/2;     // center a partial last row
        imgCard(s, xs+c*(cw+GAP), Y0+r*(ch+GAP), cw, ch, im, i+kk+1);
      });
      i+=take; page++;
    }
  }
  // build ONE table (header + the given rows) at vertical position `top`
  function buildTable(s, top, key, chunk){
    const cols=['Project', key==='cur'?'Job & Issue':'Plan', 'Due date', 'Status'];
    const head=cols.map(t=>({text:t,options:{bold:true,color:'FFFFFF',fill:{color:PPT.navy},fontSize:HFONT,valign:'middle',align:(t==='Due date'||t==='Status')?'center':'left'}}));
    const rows=[head];
    if(!chunk.length){                          // keep the table shape even with no content
      const w={color:'FFFFFF'};
      rows.push([
        {text:'—', options:{color:PPT.gray,fontSize:TFONT,valign:'middle',fill:w}},
        {text:(key==='next'?'(no next week plan)':'(no this week work)'), options:{italic:true,color:PPT.gray,fontSize:TFONT,valign:'middle',fill:w}},
        {text:'—', options:{color:PPT.gray,fontSize:TFONT,align:'center',valign:'middle',fill:w}},
        {text:'—', options:{color:PPT.gray,fontSize:TFONT,align:'center',valign:'middle',fill:w}}
      ]);
    }
    chunk.forEach((gp,idx)=>{ const st=groupStatus(gp), bg={color: idx%2?'F1F5FB':'FFFFFF'};
      rows.push([
        {text:gp.label, options:{bold:true,color:PPT.navy,fontSize:TFONT,valign:'top',fill:bg}},
        {text:cellRuns(gp[key]), options:{valign:'top',fill:bg}},
        {text:gp.due||'—', options:{color:PPT.gray,fontSize:TFONT,align:'center',valign:'middle',fill:bg}},
        {text:st.text, options:{bold:true,color:st.color,fontSize:TFONT,align:'center',valign:'middle',fill:bg}}
      ]);
    });
    s.addTable(rows,{x:0.7,y:top,w:11.93,colW:COLW,border:{type:'solid',color:'D8E0EC',pt:0.5},
      fontFace:PPT.font,fontSize:TFONT,valign:'top',autoPage:false,margin:[3,5,3,5]});
  }

  // ---- render one member: This-week table, then Next-week table BELOW it on the
  //      same slide (flowing to a new slide only when it runs out of room), then
  //      one-big-image attachment pages.
  function renderMember(name, role, list){
    const groups=groupAllByProject(list);
    const thisRows=groups.filter(g=>g.cur.length);
    const nextRows=groups.filter(g=>g.next.length);
    const imgs=collectImages(list);
    if(!thisRows.length && !nextRows.length && !imgs.length){
      const s=pptx.addSlide(); header(s,name,role,'');
      s.addShape(RR,{x:3.0,y:2.95,w:7.33,h:1.5,rectRadius:0.1,fill:{color:PPT.tint},line:{color:PPT.line,width:1}});
      s.addText('Pending input',{x:3.0,y:3.12,w:7.33,h:0.5,fontSize:20,bold:true,color:PPT.navy,fontFace:PPT.font,align:'center'});
      s.addText('No work provided for this week',{x:3.0,y:3.66,w:7.33,h:0.5,fontSize:13,color:PPT.gray,fontFace:PPT.font,align:'center'});
      return;
    }
    const HEAD_H=0.32, LBL=0.40, GAP=0.20;                  // compact spacing -> more fits per slide
    let s=null, y=0, started=false;
    const newSlide=()=>{ s=pptx.addSlide(); header(s,name,role, started?'(cont.)':''); started=true; y=1.5; };
    function section(label, rowsG, key){
      let recorded=false;
      const rec=()=>{ if(!recorded){ recorded=true; if(window.__pgdbg) window.__pgdbg.push({name, sec:/THIS/.test(label)?'this':'next', slide:pptx.slides.length}); } };
      if(!rowsG.length){                                      // always show the section (empty = placeholder)
        if(!s || y+LBL+HEAD_H+0.42 > MAXY) newSlide();
        sectionLabel(s,label,y); y+=LBL; rec();
        buildTable(s, y, key, []);
        y += HEAD_H + 0.40 + GAP;
        return;
      }
      let i=0;
      while(i<rowsG.length){
        const firstH=estRowH(rowsG[i][key]);
        if(!s || y+LBL+HEAD_H+firstH > MAXY) newSlide();     // room for label + header + 1 row
        sectionLabel(s,label,y); y+=LBL; rec();
        const top=y; let used=HEAD_H; const chunk=[];
        while(i<rowsG.length){
          const h=estRowH(rowsG[i][key]);
          if(chunk.length && top+used+h > MAXY) break;
          chunk.push(rowsG[i]); used+=h; i++;
        }
        buildTable(s, top, key, chunk);
        y = top + used + GAP;                                 // gap before next section
      }
    }
    newSlide();
    section('THIS WEEK  ｜  This week', thisRows, 'cur');
    section('NEXT WEEK  ｜  Next week plan', nextRows, 'next');
    attachPages(name,role,imgs);
  }

  // first page(s): the official project MATRIX (Customer | Type | Customer Project | Component | Model | Chipset)
  function matrixSlides(){
    const mgroups=projectGroups().filter(g=>(projMeta[g.projk]||{}).master);
    if(!mgroups.length) return;
    const types=projTree(mgroups), data=[];
    projTypeOrder(types).forEach(ty=>{ const block=types[ty];
      block.order.forEach(key=>{ const grp=block.map[key];
        grp.items.forEach((g,idx)=>{ const m=projMeta[g.projk]||{};
          data.push([idx===0?(grp.cust||'—'):'', idx===0?ty:'', idx===0?(grp.cp||'—'):'',
            m.component||projComponent(g)||'—', m.code||g.label, m.chipset||'—', String(g.tasks.length||'—')]);
        });
      });
    });
    const head=['Customer','Type','Customer Project','Component','Model','Chipset','Tasks'].map(tx=>({text:tx,options:{bold:true,color:'FFFFFF',fill:{color:PPT.navy},fontSize:10,valign:'middle',fontFace:PPT.font}}));
    const CHUNK=24;
    for(let i=0;i<data.length;i+=CHUNK){
      const s=pptx.addSlide();
      s.addShape(R,{x:0,y:0,w:13.33,h:0.12,fill:{color:PPT.cyan}});
      s.addText('PROJECT MATRIX'+(data.length>CHUNK?` (${Math.floor(i/CHUNK)+1})`:''),{x:0.7,y:0.3,w:9,h:0.6,fontSize:24,bold:true,color:PPT.navy,fontFace:PPT.font});
      s.addText(date,{x:10.2,y:0.42,w:2.4,h:0.4,fontSize:12,color:PPT.gray,fontFace:PPT.font,align:'right'});
      const rows=[head, ...data.slice(i,i+CHUNK).map((r,ri)=>r.map((c,ci)=>({text:String(c),options:{fontSize:9,fontFace:PPT.font,color:ci===4?PPT.navy:PPT.dark,bold:ci===4||(!!c&&ci<=2),fill:{color:(ri%2)?'F1F5FB':'FFFFFF'},valign:'middle',align:ci===6?'center':'left'}})))];
      s.addTable(rows,{x:0.55,y:1.05,w:12.23,colW:[1.35,0.95,1.75,2.75,2.15,2.65,0.63],rowH:0.24,border:{type:'solid',color:'D8E0EC',pt:0.5},autoPage:false,margin:[2,4,2,4]});
    }
  }
  matrixSlides();
  targets.forEach(m=>renderMember(memberDisplay(m),[m.role,m.role2].filter(Boolean).join(' · '),map.get(m.id)||[]));
  if((!memberIds||!memberIds.length) && unassigned.length)
    renderMember('Unassigned','',unassigned);

  const fname=(memberIds&&memberIds.length===1?memberName(memberIds[0]):'AllMembers')+
    '_WeeklyReport_'+date+'.pptx';
  return {pptx, fname};
}

async function buildPptx(memberIds){
  const r=assemblePptx(memberIds); if(!r) return;
  try{ await r.pptx.writeFile({fileName:r.fname}); }
  catch(e){ const blob=await r.pptx.write({outputType:'blob'}); downloadBlob(blob,r.fname); }
  toast('Exported '+r.fname);
}

async function exportWord(memberIds){
  const {map,unassigned}=buildBuckets();
  const targets = memberIds && memberIds.length
    ? members.filter(m=>memberIds.includes(m.id))
    : members.slice();
  if(!targets.length){ toast('No members to export'); return; }

  // media collector: dedupe images, assign rIds
  const media=[]; const relParts=[]; const seen=new Map();
  const collect=dataUrl=>{
    if(seen.has(dataUrl)) return seen.get(dataUrl);
    const ext=dataUrl.includes('image/png')?'png':dataUrl.includes('image/gif')?'gif':'jpeg';
    const n=media.length+1; const rid='rIdImg'+n; const fname=`media/image${n}.${ext}`;
    media.push({fname, b64:dataUrl.split(',')[1]});
    relParts.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${fname}"/>`);
    seen.set(dataUrl,rid); return rid;
  };

  let bodyXml='';
  targets.forEach(m=>{ bodyXml+=memberReportXml(memberDisplay(m), map.get(m.id), collect); });
  if((!memberIds||!memberIds.length) && unassigned.length)
    bodyXml+=memberReportXml('Unassigned', unassigned, collect);

  const documentXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
<w:body>${P('Weekly Report — '+new Date().toISOString().slice(0,10),{bold:true,size:20,color:'1F6FEB'})}${bodyXml}
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:bottom="1134" w:left="1134" w:right="1134"/></w:sectPr>
</w:body></w:document>`;

  const zip=new JSZip();
  zip.file('[Content_Types].xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Default Extension="jpeg" ContentType="image/jpeg"/>
<Default Extension="gif" ContentType="image/gif"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
  zip.file('_rels/.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
  zip.file('word/document.xml',documentXml);
  zip.file('word/_rels/document.xml.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relParts.join('')}</Relationships>`);
  media.forEach(m=>zip.file('word/'+m.fname, m.b64, {base64:true}));

  const blob=await zip.generateAsync({type:'blob',
    mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const fname = (memberIds&&memberIds.length===1? memberName(memberIds[0]) : 'AllMembers')+
    '_WeeklyReport_'+new Date().toISOString().slice(0,10)+'.docx';
  downloadBlob(blob, fname);
  toast('Exported '+fname);
  return blob;
}
function downloadBlob(blob, name){
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name;
  document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},1000);
}

/* =====================================================================
   OCR — read text out of image-based reports (Tesseract.js, lazy-loaded)
   ===================================================================== */
let _tessPromise=null;
function ensureTesseract(){
  if(window.Tesseract) return Promise.resolve(window.Tesseract);
  if(_tessPromise) return _tessPromise;
  _tessPromise=new Promise((res,rej)=>{
    const s=document.createElement('script');
    s.src='https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';
    s.onload=()=>res(window.Tesseract);
    s.onerror=()=>{ _tessPromise=null; rej(new Error('Could not load the OCR library (first OCR use needs an internet connection)')); };
    document.head.appendChild(s);
  });
  return _tessPromise;
}
function cleanOcr(txt){
  return String(txt||'')
    .replace(/¢/g,'c').replace(/€/g,'e').replace(/[‘’]/g,"'").replace(/[“”]/g,'"').replace(/[—–]/g,'-')  // common OCR mojibake
    .replace(/[ \t]+/g,' ').replace(/\n{2,}/g,'\n').replace(/^\s+|\s+$/g,'')
    .split('\n').map(l=>l.trim()).filter(l=>l.length>1).join('\n');
}
// upscale + grayscale + percentile contrast-stretch + sharpen -> Tesseract reads small/blurry table text much better
function preprocessForOcr(dataUrl){
  return new Promise(res=>{
    const img=new Image();
    img.onload=()=>{
      let w0=img.naturalWidth||img.width, h0=img.naturalHeight||img.height; if(!w0||!h0){ res(dataUrl); return; }
      // aim for ~2200px on the long edge so text lines are tall enough for the LSTM engine (≈300 DPI feel)
      const long=Math.max(w0,h0);
      const scale = long<2200 ? Math.min(3.2, 2200/long) : (long>3400 ? 3400/long : 1);
      const w=Math.round(w0*scale), h=Math.round(h0*scale);
      const c=document.createElement('canvas'); c.width=w; c.height=h;
      const ctx=c.getContext('2d'); ctx.imageSmoothingEnabled=true; ctx.imageSmoothingQuality='high'; ctx.drawImage(img,0,0,w,h);
      try{
        const d=ctx.getImageData(0,0,w,h), px=d.data, n=w*h;
        // grayscale + histogram
        const gray=new Uint8ClampedArray(n), hist=new Uint32Array(256);
        for(let i=0,j=0;i<px.length;i+=4,j++){ const g=(px[i]*0.299+px[i+1]*0.587+px[i+2]*0.114)|0; gray[j]=g; hist[g]++; }
        // robust contrast stretch using 1st/99th percentiles (ignores stray pure-black/white pixels)
        const lo=Math.floor(n*0.01), hi=Math.floor(n*0.99); let acc=0, mn=0, mx=255;
        for(let v=0;v<256;v++){ acc+=hist[v]; if(acc>=lo){ mn=v; break; } }
        acc=0; for(let v=0;v<256;v++){ acc+=hist[v]; if(acc>=hi){ mx=v; break; } }
        const rng=Math.max(8,mx-mn);
        const lut=new Uint8ClampedArray(256);
        for(let v=0;v<256;v++){ let t=(v-mn)*255/rng; t=255*Math.pow(Math.max(0,Math.min(1,t/255)),1.25); lut[v]=t; }
        const st=new Uint8ClampedArray(n);
        for(let j=0;j<n;j++) st[j]=lut[gray[j]];
        // light sharpen (unsharp 3x3) so anti-aliased glyph edges stay crisp after upscaling
        const out=new Uint8ClampedArray(n);
        for(let y=0;y<h;y++){ for(let x=0;x<w;x++){ const j=y*w+x;
          if(x===0||y===0||x===w-1||y===h-1){ out[j]=st[j]; continue; }
          const s=5*st[j]-st[j-1]-st[j+1]-st[j-w]-st[j+w];
          out[j]=s<0?0:s>255?255:s;
        }}
        for(let i=0,j=0;i<px.length;i+=4,j++){ px[i]=px[i+1]=px[i+2]=out[j]; }
        ctx.putImageData(d,0,0);
      }catch(e){}
      res(c.toDataURL('image/png'));
    };
    img.onerror=()=>res(dataUrl);
    img.src=dataUrl;
  });
}
let _tessWorker=null, _tessWorkerP=null;
async function getWorker(){
  if(_tessWorker) return _tessWorker;
  if(_tessWorkerP) return _tessWorkerP;
  _tessWorkerP=(async()=>{
    await ensureTesseract();
    const w=await Tesseract.createWorker('eng');
    try{ await w.setParameters({ tessedit_pageseg_mode:'6', preserve_interword_spaces:'1', user_defined_dpi:'300', tessedit_do_invert:'0' }); }catch(e){}
    _tessWorker=w; return w;
  })();
  return _tessWorkerP;
}
async function ocrImages(images, onStep){
  const w=await getWorker();
  const parts=[];
  for(let i=0;i<images.length;i++){
    if(onStep) onStep(i+1, images.length);
    let pre; try{ pre=await preprocessForOcr(images[i].data); }catch(e){ pre=images[i].data; }
    const {data}=await w.recognize(pre);
    const t=cleanOcr(data.text);
    if(t) parts.push(t);
  }
  return parts.join('\n');
}
// OCR confuses 0/O, 1/I/l, B/8 and inserts stray chars in codes. Match loosely, then normalize.
const OCR_CODE_RE=/[8B][\dOolI]{1,4}\s?[WV]\s?[\dOolI]{2,5}(?:\s?[.\-]\s?\d{1,2})?(?:\s?T[\dOolI]{1,2})?/i;
function fixCode(s){
  let t=s.toUpperCase().replace(/\s+/g,'').replace(/^8/,'B').replace(/[OQ]/g,'0').replace(/[ILl]/g,'1');
  // canonical B-code = B + 2 digits + 1 letter + 3 digits; OCR often inserts extra leading 0s -> keep last 2 / last 3
  const m=t.match(/^B(\d+)([A-Z])(\d+)(.*)$/);
  if(m){ const d1=m[1].length>=2?m[1].slice(-2):m[1].padStart(2,'0'); const d3=m[3].length>=3?m[3].slice(-3):m[3].padStart(3,'0');
    return 'B'+d1+m[2]+d3+(m[4]||''); }
  return t;
}
// turn OCR text of a pasted weekly-report table back into task rows.
// Reporter-driven: only images that carry a "Reporter:" label become tasks (so diagrams/master-lists are skipped).
// A pasted weekly-report table looks like a report even when OCR garbles it: it has the
// standard column header (Project Name / Current Job / Due Date) and/or canonical project codes.
// Used to make sure SOMEONE ELSE's report (a report screenshot pasted after a member's own table)
// is never silently dumped onto the preceding member as if it were their detail image.
function looksLikeReportText(text){
  if(/report[a-z]{0,3}r\s*[:;:]/i.test(text)) return true;           // "Reporter:" even mis-OCR'd as "Reportar:"
  if(/報告人\s*[::]/.test(text)) return true;
  if(/project\s*name/i.test(text) && /(current\s*job|due\s*date|next\s*week)/i.test(text)) return true;
  const codes=text.match(new RegExp(OCR_CODE_RE.source,'ig'))||[];
  return codes.length>=2;                                            // 2+ project codes = a project table
}
function ocrReportToTasks(text){
  // fuzzy "Reporter:" — OCR turns it into Reportar/Reporler/Repoter etc. Snap the captured name to a real member.
  const rm=text.match(/report[a-z]{0,3}r\s*[:;:]?\s*([A-Za-z][A-Za-z.]+)/i) || text.match(/報告人\s*[::]?\s*([A-Za-z一-鿿]+)/);
  if(!rm) return [];                       // not a personal weekly report -> skip
  const reporter=snapName(rm[1]);          // snap a garbled OCR name to the nearest real member
  const lines=text.split('\n').map(l=>l.replace(/\s+/g,' ').trim()).filter(l=>l.length>2);
  const isHeader=l=>/current job|project name|due date|^owner$|^risk|next week job|product category|customer|reporter/i.test(l)&&l.length<70;
  const out=[]; let section='current', last=null;
  lines.forEach(line=>{
    if(/next\s*(step|week)/i.test(line)){ section='next'; last=null; return; }
    if(isHeader(line)) return;
    let desc=line, project='', owner=reporter, due='';
    const dueM=desc.match(/20\d{2}\s*[\/.\-]\s*\d{1,2}\s*[\/.\-]\s*[A-Za-z0-9]+|\b\d{1,2}\/[A-Za-z]\b/);
    if(dueM){ due=dueM[0].replace(/\s+/g,''); desc=desc.replace(dueM[0],' '); }
    const cm=desc.match(OCR_CODE_RE);
    // a line with neither a project code nor a due date is a continuation of the previous item
    if(!cm && !dueM && last){
      if(section==='next') last.next=(last.next+' '+line).trim(); else last.current=(last.current+' '+line).trim();
      return;
    }
    if(cm){ project=fixCode(cm[0]); desc=desc.replace(cm[0],' '); }
    const tok=desc.trim().split(' ');
    if(tok.length>1){ const lw=tok[tok.length-1];
      // accept trailing word as owner ONLY if it's a real name: ALL-CAPS (RICK), the reporter, or an existing member
      const isName=/^[A-Za-z][A-Za-z.\/]{1,15}$/.test(lw) &&
        (lw.toUpperCase()===lw || norm(lw)===norm(reporter) || members.some(m=>norm(m.name)===norm(lw)||(m.aliases||[]).some(a=>norm(a)===norm(lw))));
      if(isName){ owner=lw; tok.pop(); desc=tok.join(' '); } }
    desc=desc.replace(/[\s|:.\-]+$/,'').replace(/^[\s|:.\-]+/,'').trim();
    if(!project && desc.length<4) return;
    if(!project) project=desc.split(' ').slice(0,3).join(' ').slice(0,24);
    last={project, current:section==='current'?(desc||line):'', next:section==='next'?(desc||line):'', owner, due, reporter};
    out.push(last);
  });
  return out;
}
// snap garbled OCR codes to the REAL projects from the deck's tables; greedy-unique so 2 rows don't claim the same one
// snap an OCR-garbled person name (e.g. "Riek") to the nearest real member ("Rick") -- always, regardless of fuzzy toggle
function snapName(name){
  if(!name) return name;
  const n=norm(name); if(!n) return name;
  for(const m of members){ if(norm(m.name)===n || (m.aliases||[]).some(a=>norm(a)===n)) return m.name; }
  let best=null,bd=99;
  for(const m of members){ const d=lev(n, norm(m.name)); if(d<bd){ bd=d; best=m; } }
  return (best && bd<=1 && n.length>=3)? best.name : name;
}
function convertOcrRows(rows, images){
  const known={}; tasks.forEach(t=>{ if(!t.imageReport && t.projk) known[t.projk]=t.projectLabel||t.projk; });
  const keys=Object.keys(known), used=new Set();
  rows.forEach(r=>{
    r.owner=snapName(r.owner); r.reporter=snapName(r.reporter);                 // Riek -> Rick
    r.current=(r.current||'').replace(/^\s*\d?\s*T\d{2}\s+/i,'');   // strip OCR code remnant like "3T00 "
    r.next=(r.next||'').replace(/^\s*\d?\s*T\d{2}\s+/i,'');
    const target=projKeyOf(r.project).replace(/^c:|^id:|^n:/,'');
    if(target.length<5) return;
    let best=null,bd=99;
    keys.forEach(k=>{ const kk=k.replace(/^c:|^id:|^n:/,''); const d=lev(target,kk)+(used.has(k)?1.5:0); if(d<bd){ bd=d; best=k; } });
    if(best && bd<=2.5){ r.project=known[best]; used.add(best); }
  });
  return rows.map(r=>makeTaskFromFields(r,0));   // images attached by caller (downscaled, storage-light)
}
async function ocrTask(id){
  const t=tasks.find(x=>x.id===id);
  if(!t || !t.images || !t.images.length){ toast('This task has no images to recognise'); return; }
  toast('Starting OCR… the model downloads on first use, please wait');
  try{
    const text=await ocrImages(t.images, (i,n)=>toast(`OCR… ${i}/${n}`));
    if(!text){ toast('No text recognised'); return; }
    if(t.imageReport){
      const rows=ocrReportToTasks(text);
      if(rows.length){
        const parsed=convertOcrRows(rows, t.images);
        tasks=tasks.filter(x=>x.id!==id);
        overlayTasks(parsed, 'Image report OCR');
        if(autoAdd) autoAddFromReport();
        persist(); renderAll(); $('#taskModal').hidden=true;
        toast('✅ Image report recognised — extracted '+parsed.length+' tasks');
        return;
      }
      t.current=text; t.imageReport=false; t.project='OCR report'; t.projectLabel='OCR report';
    } else {
      t.current=(t.current? t.current+' ; ' : '')+'[OCR] '+text;
    }
    t.ocrDone=true; persist(); renderAll(); openTask(id);
    toast('✅ OCR done');
  }catch(e){ toast(e.message||'OCR failed'); }
}
let _ocrRunning=false;
async function downscaleImgs(imgs){    // shrink hi-res OCR images to storage-light thumbnails before attaching
  const out=[]; for(const im of (imgs||[])){ try{ const s=await shrinkImageBudget(im.data, 1200, 0.82); out.push({id:uid(), data:s.data, w:s.w, h:s.h}); }catch(e){} }
  return out;
}
async function ocrAllReports(silent){
  if(_ocrRunning) return;
  const reps=pendingReports.slice().sort((a,b)=>(a._slide||0)-(b._slide||0));   // process in slide order
  if(!reps.length){ if(!silent) toast('No image pages to recognise'); return; }
  _ocrRunning=true;
  let made=0, kept=0, errs=0, lastReportTasks=null, lastReportSlide=-99;
  for(let i=0;i<reps.length;i++){
    const t=reps[i]; const slide=t._slide||0; toast(`🔍 OCR image page ${i+1}/${reps.length}… (running in background, you can keep working)`);
    try{
      const text=await ocrImages(t._images||t.images, ()=>{});
      const rows=ocrReportToTasks(text);
      const isReport = rows.length>=1;                    // fuzzy "Reporter:" -> rows only exist for a real personal report
      if(isReport){
        const parsed=convertOcrRows(rows, null);
        const small=await downscaleImgs(t._images||t.images);
        if(parsed[0]) parsed[0]._images=small;            // keep the report screenshot on the first task
        const created=overlayTasks(parsed,'Image report OCR');
        made+=parsed.length; kept++;
        lastReportTasks=created; lastReportSlide=slide;
      } else if(looksLikeReportText(text)){
        // Clearly SOMEONE's pasted report table, but OCR too garbled to extract the reporter/rows.
        // It is NOT the preceding member's detail image -> DROP it (do not attach to _afterReporter),
        // so a member never ends up showing someone else's report screenshot.
        errs++;
      } else if(t._afterReporter){
        // not a report -> it's the detail/measurement image of the TABLE member whose section it sits in.
        // This takes priority over lastReportTasks: a pasted report may have been OCR'd just above inside
        // this member's section, but the member's OWN measurements (slides after it) still belong to them.
        const owner=tasks.filter(x=>!x.imageReport && norm(x.reporter)===norm(t._afterReporter));
        if(owner.length){
          const small=await downscaleImgs(t._images||t.images);
          owner[0].images=(owner[0].images||[]).concat(small);
        }
      } else if(lastReportTasks && lastReportTasks.length && (slide-lastReportSlide)<=8
                && !tableSlides.some(s=>s>lastReportSlide && s<slide)){
        // fallback (no preceding table member): image page right after a pasted report with no table
        // in between = that report's own test data -> attach to its tasks.
        const small=await downscaleImgs(t._images||t.images);
        lastReportTasks[0].images=(lastReportTasks[0].images||[]).concat(small);
      }
    }catch(e){ console.warn('OCR skip', e); errs++; }
    pendingReports=pendingReports.filter(x=>x!==t);
    updateOcrBtn();
  }
  if(autoAdd) autoAddFromReport();
  const dropped=dedupeTasks(); cleanupGarbledMembers();
  _ocrRunning=false; persist(); renderAll();
  toast(`✅ Image pages recognised: ${kept} reports → ${made} tasks${dropped?` (deduped ${dropped})`:''}`);
}
function updateOcrBtn(){
  const n=pendingReports.length;
  const b=$('#ocrReportsBtn'); if(b) b.textContent='🔍 OCR image reports'+(n?` (${n})`:'');
}
async function ocrToWorkbench(){
  if(!wbImages.length){ toast('Upload an image first'); return; }
  toast('Starting OCR… the model downloads on first use, please wait');
  try{
    const text=await ocrImages(wbImages, (i,n)=>toast(`OCR… ${i}/${n}`));
    if(!text){ toast('No text recognised'); return; }
    const ta=$('#wbThisWeek'); ta.value=(ta.value? ta.value+'\n':'')+text;
    toast('✅ OCR done — filled into This week');
  }catch(e){ toast(e.message||'OCR failed'); }
}

/* =====================================================================
   EVENTS
   ===================================================================== */
function fileToDataURL(f){ return new Promise(res=>{ const r=new FileReader(); r.onload=()=>res(r.result); r.readAsDataURL(f); }); }

function wireEvents(){
  $('#reportInput').addEventListener('change', e=>{ if(e.target.files.length) importFiles(e.target.files); e.target.value=''; });
  $('#addMembersBtn').addEventListener('click', ()=>{ const t=$('#memberPaste').value; if(t.trim()){ addMembers(parseMemberText(t),{manual:true}); $('#memberPaste').value=''; toast('Member list updated'); } });
  $('#memberFileInput').addEventListener('change', async e=>{ const f=e.target.files[0]; if(f){ addMembers(parseMemberText(await f.text()),{manual:true}); toast('Members added from file'); } e.target.value=''; });
  $('#clearMembersBtn').addEventListener('click', clearMembers);
  $('#resetTasksBtn').addEventListener('click', resetTasks);
  { const cb=$('#clearAllBtn'); if(cb) cb.addEventListener('click', clearAllContent); }
  $('#loadFromReportBtn').addEventListener('click', ()=>{ if(!tasks.length){ toast('Import reports first'); return; } const a=autoAddFromReport(); toast(a.length?('Added '+a.length+' members from reports'):'All report owners are already in the list'); });
  $('#groupSelect').addEventListener('change', e=>switchGroup(e.target.value));
  $('#saveGroupBtn').addEventListener('click', saveGroup);
  $('#newGroupBtn').addEventListener('click', createGroup);
  $('#renameGroupBtn').addEventListener('click', renameGroup);
  $('#delGroupBtn').addEventListener('click', deleteGroup);
  $('#autoAddChk').addEventListener('change', e=>{ autoAdd=e.target.checked; store.save('wrt_autoadd',autoAdd); });
  $('#fuzzyChk').addEventListener('change', e=>{ fuzzy=e.target.checked; store.save('wrt_fuzzy',fuzzy); reresolveAllTasks(); persist(); renderAll(); toast(fuzzy?'Fuzzy matching on':'Back to exact-name matching'); });
  // filter controls
  $('#filterQ').addEventListener('input', e=>{ filters.q=e.target.value; renderMembersArea(); });
  $('#filterProject').addEventListener('change', e=>{ filters.project=e.target.value; renderMembersArea(); });
  $('#filterMember').addEventListener('change', e=>{ filters.member=e.target.value; renderMembersArea(); });
  $('#filterStatus').addEventListener('change', e=>{ filters.status=e.target.value; renderMembersArea(); });
  $('#hideEmptyChk').addEventListener('change', e=>{ filters.hideEmpty=e.target.checked; renderMembersArea(); });
  $('#clearFiltersBtn').addEventListener('click', ()=>{ filters.q='';filters.project='';filters.member='';filters.status='';filters.role='';filters.hideEmpty=false; $('#filterQ').value=''; renderAll(); });
  $('#openWorkbenchBtn').addEventListener('click', openWorkbench);
  $('#exportAllBtn').addEventListener('click', ()=>exportWord(null));
  $('#exportPptxBtn').addEventListener('click', ()=>buildPptx(null));
  $('#ocrReportsBtn').addEventListener('click', ocrAllReports);
  $('#previewBtn').addEventListener('click', openNarrative);
  { const b=$('#pptxPreviewBtn'); if(b) b.addEventListener('click', openPptxPreview); }
  { const s=$('#pptxPrevMember'); if(s) s.addEventListener('change', renderPptxPreview); }
  { const e=$('#pptxPrevExportBtn'); if(e) e.addEventListener('click', ()=>{ const mid=$('#pptxPrevMember').value; buildPptx(mid?[mid]:null); }); }
  $('#narrMember').addEventListener('change', renderNarrative);
  $('#narrCopyBtn').addEventListener('click', ()=>{ navigator.clipboard&&navigator.clipboard.writeText($('#narrativeText').textContent); toast('Narrative text copied'); });
  $('#narrExportBtn').addEventListener('click', ()=>{ const mid=$('#narrMember').value; exportWord(mid?[mid]:null); });
  $('#applyAliasBtn').addEventListener('click', ()=>applyProjAliases($('#aliasText').value));
  { const a=$('#addProjBtn'); if(a) a.addEventListener('click', addBlankProject); }
  { const b=$('#importProjListBtn'); if(b) b.addEventListener('click', ()=>{ $('#projListModal').hidden=false; }); }
  { const c=$('#projListImportBtn'); if(c) c.addEventListener('click', ()=>importProjectList($('#projListText').value)); }
  $('#aliasText').value=aliasToText();
  $('#wbSaveBtn').addEventListener('click', saveWorkbench);
  $('#wbOcrBtn').addEventListener('click', ocrToWorkbench);
  { const db=$('#wbDueBtn'), dp=$('#wbDuePick');                       // English due-date: text shows ISO; 📅 opens the native calendar
    if(db&&dp){ db.addEventListener('click', ()=>{ try{ dp.showPicker?dp.showPicker():dp.click(); }catch(e){ toast('Type the date as YYYY-MM-DD'); } });
      dp.addEventListener('change', ()=>{ if(dp.value) $('#wbDue').value=dp.value; }); } }
  $('#wbImages').addEventListener('change', async e=>{ for(const f of e.target.files){ const s=await shrinkImageBudget(await fileToDataURL(f), 2200); wbImages.push({id:uid(), data:s.data, w:s.w, h:s.h}); } renderWbThumbs(); e.target.value=''; });

  // delegated clicks
  document.body.addEventListener('click', e=>{
    const t=e.target;
    if(t.classList && t.classList.contains('build-badge')){ const txt=JSON.stringify({ua:navigator.userAgent, clicks:PERF.ev, longtasks:PERF.lt});
      if(navigator.clipboard&&navigator.clipboard.writeText) navigator.clipboard.writeText(txt).then(()=>toast('Diagnostics copied — paste it to Claude'),()=>prompt('Copy this:',txt));
      else prompt('Copy this:',txt); return; }
    if(t.closest('[data-close]')){ const ov=t.closest('.modal-overlay'); if(ov){ ov.hidden=true;
      if(ov.id==='taskModal'){ _openTaskId=null; const b=$('#taskImgsBox'); if(b) b.querySelectorAll('.img-edit').forEach(s=>{ const im=s.querySelector('img'); if(im) im.src=''; s.remove(); }); } } return; }
    if(t.closest('[data-designdocs]')){ openDesignDocs(); return; }
    { const vt=t.closest('#viewTabs .vt'); if(vt){ setView(vt.dataset.view); renderStats(); return; } }
    const navCard=t.closest('[data-nav]'); if(navCard){ navStat(navCard.dataset.nav, navCard.dataset.flt); return; }
    if(t.closest('.mname') && t.closest('.mchip')){ jumpToMember(t.closest('.mchip').dataset.mid); return; }
    if(t.dataset.delMember){ deleteMember(t.dataset.delMember); return; }
    if(t.dataset.delTask){ if(deleteTask(t.dataset.delTask)) $('#taskModal').hidden=true; return; }
    if(t.dataset.rmWb!==undefined){ wbImages.splice(+t.dataset.rmWb,1); renderWbThumbs(); return; }
    if(t.dataset.phrase){ const ta=$('#wbThisWeek'); ta.value=(ta.value?ta.value+' ':'')+PHRASES[t.dataset.phrase]; return; }
    if(t.dataset.clearimg){ clearTaskImages(t.dataset.clearimg); return; }
    if(t.dataset.delimg){ const [tid,imgId]=t.dataset.delimg.split('|'); removeTaskImage(tid,imgId); return; }
    if(t.dataset.light){ openLight(imageDataById(t.dataset.light)||t.getAttribute('src')); return; }
    if(t.dataset.exportMember!==undefined){ if(t.dataset.exportMember) exportWord([t.dataset.exportMember]); else toast('This task has no matching member'); return; }
    // ----- catalog editing -----
    { const cv=t.closest('[data-catview]'); if(cv){ catalogView=cv.dataset.catview; store.save('wrt_catview',catalogView); renderCatalog(); return; } }
    { const te=t.closest('[data-treeexp]'); if(te){ const c=te.dataset.treeexp; treeExpand[c]=!treeExpand[c]; renderCatalog(); return; } }
    { const ep=t.closest('[data-edit-proj]'); if(ep){ const mo=$('#projEditModal');
      if(mo && !mo.hidden && editingProj===ep.dataset.editProj) closeProjEdit(); else openProjEdit(ep.dataset.editProj); return; } }
    if(t.dataset.saveProj!==undefined){ saveProjMeta(t.dataset.saveProj, t.closest('.proj-card')); return; }
    if(t.dataset.unmerge!==undefined){ unmergeProject(t.dataset.unmerge); return; }
    if(t.dataset.delproj!==undefined){ deleteProjectGroup(t.dataset.delproj); return; }
    if(t.dataset.sched!==undefined){ openSchedule(t.dataset.sched); return; }
    if(t.dataset.pdelimg!==undefined){ const [pk,id]=t.dataset.pdelimg.split('|'); removeProjectImage(pk,id); return; }
    { const dc=t.closest('[data-ddopen]'); if(dc){ _ddProjk=dc.dataset.ddopen; renderDesignDocs(); return; } }
    if(t.dataset.ddback!==undefined){ _ddProjk=null; renderDesignDocs(); return; }
    if(t.dataset.sadd!==undefined){ addSchedItem(); return; }
    if(t.dataset.sdel!==undefined){ delSchedItem(t.dataset.sdel); return; }
    if(t.dataset.ssave!==undefined){ saveScheduleEdit(); return; }
    if(t.dataset.sdownload!==undefined){ downloadSchedulePNG(); return; }
    if(t.dataset.stemplate!==undefined){ applyDefaultTemplate(); return; }
    if(t.dataset.sppt!==undefined){ exportSchedulePPT(); return; }
    if(t.dataset.sxlsx!==undefined){ exportScheduleExcel(); return; }
    if(t.dataset.addcat!==undefined){ const n=prompt('New project category name:'); if(n) addCategory(n); return; }
    if(t.dataset.renamecat!==undefined){ const n=prompt('Rename category "'+t.dataset.renamecat+'" to:', t.dataset.renamecat); if(n) renameCategory(t.dataset.renamecat, n); return; }
    if(t.dataset.delcat!==undefined){ deleteCategory(t.dataset.delcat); return; }
    { const cf=t.closest('[data-catfilter]'); if(cf){ catFilter=cf.dataset.catfilter; renderCatalog(); return; } }
    if(t.dataset.rmowner!==undefined){ const [tid,mid]=t.dataset.rmowner.split('|'); removeTaskOwner(tid,mid); return; }
    if(t.dataset.ocr!==undefined){ ocrTask(t.dataset.ocr); return; }
    { const ot=t.closest('[data-opentask]'); if(ot){ openTask(ot.dataset.opentask); return; } }   // closest(): clicking the TEXT inside the chip must work too
    { const wb=t.closest('.wbm'); if(wb){ if(!wb.disabled) wb.classList.toggle('on'); return; } }
    { const ap=t.closest('[data-addtaskproj]'); if(ap){ openWorkbench(null, ap.dataset.addtaskproj); return; } }
    if(t.dataset.addtask!==undefined){ openWorkbench(t.dataset.addtask); return; }
    if(t.closest('.ctask-ctl')) return;   // don't let control clicks fall through to task-open
    const projRow=t.closest('[data-proj]'); if(projRow){ openProject(projRow.dataset.proj); return; }
    const memRow=t.closest('[data-memrow]'); if(memRow){ filters.member=memRow.dataset.memrow; filters.status=''; setView('members'); renderFilters(); renderMembersArea(); renderStats(); const ma=document.querySelector('.members-area'); if(ma) ma.scrollIntoView({behavior:'smooth',block:'start'}); return; }
    const card=t.closest('[data-task]'); if(card && !t.dataset.delTask){ openTask(card.dataset.task); return; }
  });
  $('#lightbox').addEventListener('click',()=>$('#lightbox').hidden=true);

  // delegated CHANGE events (catalog inline editing + catalog member filter)
  document.body.addEventListener('change', e=>{
    const el=e.target;
    if(el.id==='catalogMember'){ catalogMember=el.value; renderCatalog(); return; }
    if(el.id==='taskGroupSel'){ taskGroupBy=el.value; store.save('wrt_taskgroup',taskGroupBy); renderMembersArea(); return; }
    if(el.id==='filterRole'){ filters.role=el.value; renderMembersArea(); return; }
    if(el.dataset.setrole!==undefined){ setMemberRole(el.dataset.setrole, el.value); return; }
    if(el.dataset.setrole2!==undefined){ setMemberRole2(el.dataset.setrole2, el.value); return; }
    if(el.dataset.setcname!==undefined){ setMemberCname(el.dataset.setcname, el.value); return; }
    if(el.dataset.addimg!==undefined){ addTaskImages(el.dataset.addimg, el.files); el.value=''; return; }
    if(el.dataset.setcat!==undefined){ setProjCategory(el.dataset.setcat, el.value); return; }
    if(el.dataset.setphase!==undefined){ setProjPhase(el.dataset.setphase, el.value); return; }
    if(el.dataset.si!==undefined && el.dataset.sf){ updSchedItem(el.dataset.si, el.dataset.sf, el.value); return; }
    if(el.dataset.smeta!==undefined){ updSchedMeta(el.dataset.smeta, el.value); return; }
    if(el.dataset.ddaddimg!==undefined){ const title=(($('#ddTitle')||{}).value||'').trim(), folder=($('#ddFolder')||{}).value||''; if(_ddProjk) addProjectImages(_ddProjk, el.files, title, folder); el.value=''; return; }
    if(el.dataset.saddimg!==undefined){ uploadScheduleAttach(el.files); el.value=''; return; }
    if(el.dataset.edit && el.dataset.tid){ editTaskField(el.dataset.tid, el.dataset.edit, el.value); return; }
    if(el.dataset.addowner){ addTaskOwner(el.dataset.addowner, el.value); el.value=''; return; }
  });
}
function openLight(src){ $('#lightboxImg').src=src; $('#lightbox').hidden=false; }

/* ---------- view tabs (less scrolling) ---------- */
let currentView=store.load('wrt_view','dashboard');
function setView(v){
  const MIGRATE={catalog:'projects', members:'tasks', workload:'team'};   // old view names -> their new-IA equivalents
  v=MIGRATE[v]||v;
  if(!['dashboard','projects','tasks','team'].includes(v)) v='dashboard';
  currentView=v; store.save('wrt_view',v);
  document.querySelectorAll('[data-pane]').forEach(el=>{
    el.style.display = el.getAttribute('data-pane').split(/\s+/).includes(v) ? '' : 'none';
  });
  $$('#viewTabs .vt').forEach(b=>b.classList.toggle('active', b.dataset.view===v));
}

/* ---------- drag a project card onto another to merge / correct grouping ---------- */
let _dragProjk=null;
function wireProjectDrag(){
  const cont=$('#projectCatalog'); if(!cont) return;
  cont.querySelectorAll('.proj-card[draggable]').forEach(card=>{
    card.addEventListener('dragstart',e=>{ if(e.target.closest('input,select,textarea,button,.pc-tasks')){ e.preventDefault(); return; } _dragProjk=card.dataset.projk; card.classList.add('dragging'); e.stopPropagation(); });
    card.addEventListener('dragend',()=>{ card.classList.remove('dragging'); _dragProjk=null; });
    card.addEventListener('dragover',e=>{ if(_dragProjk&&_dragProjk!==card.dataset.projk){ e.preventDefault(); card.classList.add('drop-target'); } });
    card.addEventListener('dragleave',()=>card.classList.remove('drop-target'));
    card.addEventListener('drop',e=>{ e.preventDefault(); card.classList.remove('drop-target');
      if(_dragProjk && _dragProjk!==card.dataset.projk) mergeProjects(_dragProjk, card.dataset.projk);
    });
  });
}

/* ---------- drag to reorder members ---------- */
let _dragMid=null;
function wireMemberDrag(){
  const ul=$('#memberChips'); if(!ul) return;
  ul.querySelectorAll('li[draggable]').forEach(li=>{
    li.addEventListener('dragstart',e=>{ if(e.target.closest('select,button')){ e.preventDefault(); return; } _dragMid=li.dataset.mid; li.classList.add('dragging'); });
    li.addEventListener('dragend',()=>{ li.classList.remove('dragging'); _dragMid=null; });
    li.addEventListener('dragover',e=>{ e.preventDefault(); li.classList.add('drop-hint'); });
    li.addEventListener('dragleave',()=>li.classList.remove('drop-hint'));
    li.addEventListener('drop',e=>{ e.preventDefault(); li.classList.remove('drop-hint');
      const from=members.findIndex(m=>m.id===_dragMid), to=members.findIndex(m=>m.id===li.dataset.mid);
      if(from<0||to<0||from===to) return;
      const [m]=members.splice(from,1); members.splice(to,0,m);
      persist(); renderAll();
    });
  });
}
function openNarrative(){
  const sel=$('#narrMember');
  sel.innerHTML='<option value="">All members</option>'+members.map(m=>`<option value="${m.id}">${esc(m.name)}</option>`).join('');
  renderNarrative();
  $('#narrativeModal').hidden=false;
}
function renderNarrative(){
  const mid=$('#narrMember')?$('#narrMember').value:'';
  $('#narrativeText').textContent=buildNarrative(mid?[mid]:null);
}
// PPTX preview: one card per member (the deck's per-member slide content) before exporting
function openPptxPreview(){
  const sel=$('#pptxPrevMember'); if(sel) sel.innerHTML='<option value="">All members</option>'+members.map(m=>`<option value="${m.id}">${esc(m.name)}</option>`).join('');
  renderPptxPreview(); $('#pptxPreviewModal').hidden=false;
}
function renderPptxPreview(){
  const mid=$('#pptxPrevMember')?$('#pptxPrevMember').value:'';
  $('#pptxPreviewBody').innerHTML=pptxPreviewHTML(mid?[mid]:null);
}
function pptxPreviewHTML(memberIds){
  const {map,unassigned}=buildBuckets();
  const targets = memberIds&&memberIds.length ? members.filter(m=>memberIds.includes(m.id)) : members.slice();
  const slide=(name,role,list)=>{
    const byP={}, order=[];
    (list||[]).forEach(t=>{ const k=pptPlabel(t)||'—'; if(!byP[k]){ byP[k]={cur:[],next:[],imgs:0}; order.push(k); }
      if((t.current||'').trim()) byP[k].cur.push({txt:t.current,p:t.progress});
      if((t.next||'').trim()) byP[k].next.push({txt:t.next,p:(t.nextProgress!=null?t.nextProgress:null)});
      byP[k].imgs+=(t.images||[]).length; });
    const sec=(title,key)=>{ const ks=order.filter(k=>byP[k][key].length); if(!ks.length) return '';
      return `<div class="pv-sec"><div class="pv-sectitle">${title}</div>${ks.map(k=>`<div class="pv-proj"><b>${esc(k)}</b>${byP[k][key].map(o=>`<div class="pv-line">${esc(String(o.txt||'').replace(/\s+/g,' ').slice(0,140))}${o.p!=null?` <span class="pv-pct">${o.p}%</span>`:''}</div>`).join('')}</div>`).join('')}</div>`; };
    const totalImgs=order.reduce((s,k)=>s+byP[k].imgs,0);
    return `<div class="pv-slide"><div class="pv-head"><span class="pv-name">${esc(name)}</span>${role?`<span class="pv-role">${esc(role)}</span>`:''}<span class="pv-meta">${(list||[]).length} task${(list||[]).length===1?'':'s'}${totalImgs?` · 📎 ${totalImgs}`:''}</span></div>${sec('THIS WEEK','cur')||'<div class="pv-empty">No this-week work</div>'}${sec('NEXT WEEK','next')}</div>`;
  };
  let html=targets.map(m=>slide(memberDisplay(m),[m.role,m.role2].filter(Boolean).join(' · '),map.get(m.id)||[])).join('');
  if((!memberIds||!memberIds.length)&&unassigned.length) html+=slide('Unassigned','',unassigned);
  return html||'<p class="hint">No members / tasks yet.</p>';
}

/* expose for testing */
window.WRT={ get members(){return members;}, get tasks(){return tasks;}, get batches(){return batches;},
  parsePPTX, importFiles, addMembers, parseMemberText, overlayTasks, exportWord, buildPptx, assemblePptx, resolveOwners, matchOwner,
  autoAddFromReport, reresolveAllTasks, buildNarrative, projKeyOf, applyProjAliases, openProject, openNarrative,
  projectGroups, editTaskField, addTaskOwner, removeTaskOwner, removeTaskImage, addTaskImages, openTask, ensureTesseract, ocrImages, ocrReportToTasks, ocrTask, setMemberRole, setMemberRole2,
  navStat, mergeProjects, resolveProjk, ocrAllReports, dedupeTasks, cleanupGarbledMembers, snapName, get pendingReports(){return pendingReports;}, convertOcrRows,
  switchGroup, addGroup, renameGroup, deleteGroup, get memberGroups(){return memberGroups;}, get activeGroup(){return activeGroup;},
  get deletedNames(){return deletedNames;}, get projMerge(){return projMerge;}, get projMeta(){return projMeta;},
  setCatalogMember:(v)=>{catalogMember=v;renderCatalog();},
  get filters(){return filters;}, setFilter:(k,v)=>{filters[k]=v;renderMembersArea();},
  resetTasks:()=>{tasks=[];batches=[];persist();renderAll();},
  clearAll:()=>{members=[];tasks=[];batches=[];persist();renderAll();} };

/* =====================================================================
   CLOUD SYNC (Firebase) — team-passcode login + Firestore live sync.
   Online only (file:// stays purely local). Syncs state + tasks (workspace doc)
   and images (separate `images` collection, one doc per image, live). Loop-safe
   via the `applying` flag + snapshot `hasPendingWrites`.
   ===================================================================== */
function projMetaNoImg(){ const o={}; Object.keys(projMeta).forEach(k=>{ const c=Object.assign({},projMeta[k]); delete c.images; o[k]=c; }); return o; }
let _cloudErrAt=0;
function cloudErrToast(e){                              // surface cloud-sync failures (throttled) so uploads never fail silently
  if(e && /permission|insufficient/i.test(e.message||'')){ const now=Date.now(); if(now-_cloudErrAt>20000){ _cloudErrAt=now;
    toast('⚠ Cloud sync failed: make sure you are signed in and the Firebase rule is allow read,write: if request.auth!=null'); } }
}
function cloudSave(){
  if(!CLOUD.on || !CLOUD.ready || CLOUD.applying) return;
  CLOUD.dirty=true;                                     // we have unsaved local edits -> snapshots must not clobber them until saved
  clearTimeout(CLOUD.saveTimer);
  CLOUD.saveTimer=setTimeout(()=>{
    const tasksLite=tasks.map(t=>{ const c=Object.assign({},t); delete c.images; c._imgN=(t.images||[]).length; return c; });
    CLOUD.db.collection('workspace').doc('main').set({
      members, groups:memberGroups, activeGroup, tasks:tasksLite, batches,
      projAliases, projMeta:projMetaNoImg(), projMerge, projCats, idCodeMap, deletedNames, _ts:Date.now()
    }).then(()=>{ CLOUD.dirty=false; }).catch(e=>{ CLOUD.dirty=false; console.warn('cloud save failed', e); cloudErrToast(e); });
    cloudSaveImages();                                  // upload any new images to the images collection
  }, 700);
}
function cloudSaveImages(){
  if(!CLOUD.on || !CLOUD.db) return;
  tasks.forEach(t=>(t.images||[]).forEach(im=>{
    if(im && im.id && im.data && !CLOUD.upImgs.has(im.id)){
      CLOUD.upImgs.add(im.id);
      CLOUD.db.collection('images').doc(im.id)
        .set({id:im.id, taskId:t.id, data:im.data, w:im.w||0, h:im.h||0, _ts:Date.now()})
        .catch(e=>{ CLOUD.upImgs.delete(im.id); console.warn('image upload failed', e); cloudErrToast(e); });
    }
  }));
  Object.keys(projMeta).forEach(pk=>((projMeta[pk]&&projMeta[pk].images)||[]).forEach(im=>{   // project images (block diagrams / schedules)
    if(im && im.id && im.data && !CLOUD.upImgs.has(im.id)){
      CLOUD.upImgs.add(im.id);
      CLOUD.db.collection('images').doc(im.id)
        .set({id:im.id, projk:pk, title:im.title||'', folder:im.folder||'', data:im.data, w:im.w||0, h:im.h||0, _ts:Date.now()})
        .catch(e=>{ CLOUD.upImgs.delete(im.id); console.warn('proj image upload failed', e); cloudErrToast(e); });
    }
  }));
}
function cloudApplyDoc(d){
  if(!d) return;
  if(Array.isArray(d.members)) members=d.members;
  if(Array.isArray(d.groups)) memberGroups=d.groups;
  if(d.activeGroup) activeGroup=d.activeGroup;
  if(Array.isArray(d.batches)) batches=d.batches;
  if(Array.isArray(d.tasks)){
    const imgById={}; tasks.forEach(t=>{ if((t.images||[]).length) imgById[t.id]=t.images; });  // keep local images
    tasks=d.tasks.map(t=>{
      const m={};                                               // merge local + cloud images by id
      (imgById[t.id]||[]).forEach(im=>{ if(im&&im.id) m[im.id]=im; });
      (CLOUD.imgsByTask[t.id]||[]).forEach(im=>{ if(im&&im.id) m[im.id]=im; });
      return Object.assign({}, t, {images: Object.values(m)});
    });
  }
  if(Array.isArray(d.projAliases)) projAliases=d.projAliases;
  if(d.projMeta){                                             // merge local + cloud project images back in (doc carries none)
    const localImg={}; Object.keys(projMeta).forEach(k=>{ if(((projMeta[k]||{}).images||[]).length) localImg[k]=projMeta[k].images; });
    const inc=d.projMeta;
    Object.keys(inc).forEach(k=>{ const m={};
      (localImg[k]||[]).forEach(im=>{ if(im&&im.id) m[im.id]=im; });
      (((CLOUD.imgsByProj||{})[k])||[]).forEach(im=>{ if(im&&im.id) m[im.id]=im; });
      if(Object.keys(m).length) inc[k]=Object.assign({}, inc[k], {images:Object.values(m)});
    });
    projMeta=inc;
  }
  if(d.projMerge) projMerge=d.projMerge;
  if(Array.isArray(d.projCats)) projCats=d.projCats;
  if(d.idCodeMap) idCodeMap=d.idCodeMap;
  if(Array.isArray(d.deletedNames)) deletedNames=d.deletedNames;
}
function cloudLoadImagesInto(snap){
  CLOUD.imgsByTask={}; CLOUD.imgsByProj={};
  snap.forEach(dd=>{ const im=dd.data(); if(!im||!im.id) return;
    CLOUD.upImgs.add(im.id);
    if(im.projk) (CLOUD.imgsByProj[im.projk]=CLOUD.imgsByProj[im.projk]||[]).push({id:im.id,data:im.data,w:im.w,h:im.h,title:im.title||'',folder:im.folder||''});
    else (CLOUD.imgsByTask[im.taskId]=CLOUD.imgsByTask[im.taskId]||[]).push({id:im.id,data:im.data,w:im.w,h:im.h}); });
}
async function cloudEnter(){                         // load once + subscribe to live updates
  const ref=CLOUD.db.collection('workspace').doc('main');
  // load cloud images first so they attach when the workspace doc is applied
  try{ cloudLoadImagesInto(await CLOUD.db.collection('images').get()); }catch(e){ console.warn('image load failed', e); }
  let existed=false;
  try{ const snap=await ref.get(); if(snap.exists){ existed=true; CLOUD.applying=true; try{ cloudApplyDoc(snap.data()); } finally { CLOUD.applying=false; } persistLocal(); } }
  catch(e){ console.warn('cloud load failed', e); }
  ref.onSnapshot(snap=>{
    // ignore our own writes; and never overwrite unsaved local edits (dirty) — that was reverting imports/deletes
    if(!snap.exists || snap.metadata.hasPendingWrites || CLOUD.dirty) return;
    CLOUD.applying=true;
    try{ cloudApplyDoc(snap.data()); } finally { CLOUD.applying=false; }   // try/finally: a throw must NOT leave applying stuck (would block all saves)
    persistLocal(); renderAll();
  }, e=>console.warn('snapshot error', e));
  // real-time images: add when a member uploads, REMOVE when anyone deletes (cross-device)
  CLOUD.db.collection('images').onSnapshot(snap=>{
    CLOUD.applying=true;
    try{
    snap.docChanges().forEach(ch=>{
      const im=ch.doc.data(); if(!im||!im.id) return;
      if(im.projk){                                      // PROJECT image (block diagram / schedule)
        CLOUD.imgsByProj=CLOUD.imgsByProj||{}; const m=projMeta[im.projk];
        if(ch.type==='removed'){ CLOUD.upImgs.delete(im.id);
          if(CLOUD.imgsByProj[im.projk]) CLOUD.imgsByProj[im.projk]=CLOUD.imgsByProj[im.projk].filter(x=>x.id!==im.id);
          if(m&&m.images) m.images=m.images.filter(x=>x.id!==im.id);
        } else { CLOUD.upImgs.add(im.id); const obj={id:im.id,data:im.data,w:im.w,h:im.h,title:im.title||'',folder:im.folder||''};
          const arr=(CLOUD.imgsByProj[im.projk]=CLOUD.imgsByProj[im.projk]||[]); const ai=arr.findIndex(x=>x.id===im.id); if(ai>=0)arr[ai]=obj; else arr.push(obj);
          if(m){ m.images=m.images||[]; const pi=m.images.findIndex(x=>x.id===im.id); if(pi>=0)m.images[pi]=obj; else m.images.push(obj); }
        }
        return;
      }
      const t=tasks.find(x=>x.id===im.taskId);
      if(ch.type==='removed'){
        CLOUD.upImgs.delete(im.id);
        if(CLOUD.imgsByTask[im.taskId]) CLOUD.imgsByTask[im.taskId]=CLOUD.imgsByTask[im.taskId].filter(x=>x.id!==im.id);
        if(t&&t.images) t.images=t.images.filter(x=>x.id!==im.id);
      } else {                                           // added / modified
        CLOUD.upImgs.add(im.id);
        const obj={id:im.id,data:im.data,w:im.w,h:im.h};
        const arr=(CLOUD.imgsByTask[im.taskId]=CLOUD.imgsByTask[im.taskId]||[]);
        const ai=arr.findIndex(x=>x.id===im.id); if(ai>=0) arr[ai]=obj; else arr.push(obj);
        if(t){ t.images=t.images||[]; const ti=t.images.findIndex(x=>x.id===im.id); if(ti>=0) t.images[ti]=obj; else t.images.push(obj); }
      }
    });
    } finally { CLOUD.applying=false; }
    persistLocal(); renderAll();
  }, e=>console.warn('image snapshot error', e));
  CLOUD.ready=true;
  renderAll();
  if(!existed) cloudSave();                          // empty cloud -> seed from this device
  cloudSaveImages();                                 // push any local-only images up (e.g. uploaded before image-sync existed)
}
function applyRoleUI(){
  document.body.classList.toggle('role-member', !!(CLOUD.me && !CLOUD.me.admin));
  const who=$('#cloudWho'); if(who) who.textContent=CLOUD.me? ('👤 '+CLOUD.me.name+(CLOUD.me.admin?'(admin)':'(member)')) : '';
  const lo=$('#cloudLogout'); if(lo) lo.hidden=!CLOUD.me;
}
function cloudPickName(){
  const sel=$('#cgName'); if(!sel) return;
  let names;
  if(CLOUD.authedAs==='admin'){                         // admin passcode -> only admin names
    names=CLOUD.admins.map(a=>a.charAt(0).toUpperCase()+a.slice(1));
  }else{                                                // team passcode -> members only (NO admin names)
    names=members.map(m=>m.name).filter(n=>!isAdminName(n));
  }
  sel.innerHTML=names.map(n=>`<option>${esc(n)}</option>`).join('')||'<option value="">(roster not set up)</option>';
}
function cloudFinishLogin(name){
  CLOUD.me={name, admin: CLOUD.authedAs==='admin'};     // role comes from WHICH account, not the name
  store.save('wrt_cloud_me', CLOUD.me);
  if(!CLOUD.me.admin){                                  // member -> only their own work
    const m=members.find(x=>x.name.toLowerCase()===String(name).toLowerCase());
    if(m) filters.member=m.id;
    setView('members');
  }
  applyRoleUI();
  $('#cloudGate').hidden=true;
  renderAll();
}
function showGate(step){
  const g=$('#cloudGate'); if(!g) return; g.hidden=false;
  $('#cgPassWrap').hidden = step==='who';
  $('#cgWho').hidden = step!=='who';
  if(step==='who'){ cloudPickName(); }
}
async function cloudInit(){
  CLOUD.on = (location.protocol!=='file:') && typeof firebase!=='undefined' && !!window.FIREBASE_CONFIG;
  if(!CLOUD.on) return;                               // local mode — app already booted normally
  try{ firebase.initializeApp(window.FIREBASE_CONFIG); CLOUD.db=firebase.firestore(); }
  catch(e){ console.warn('firebase init failed', e); CLOUD.on=false; return; }

  $('#cgEnter').addEventListener('click', async ()=>{
    const pass=$('#cgPass').value.trim(); const msg=$('#cgMsg');
    if(!pass){ msg.textContent='Enter the passcode'; return; }
    msg.textContent='Signing in…';
    let role=null;
    try{ await firebase.auth().signInWithEmailAndPassword(CLOUD.memberEmail, pass); role='member'; }
    catch(e1){
      try{ await firebase.auth().signInWithEmailAndPassword(CLOUD.adminEmail, pass); role='admin'; }
      catch(e2){ msg.textContent='Wrong passcode'; return; }
    }
    CLOUD.authedAs=role;
    msg.textContent='';
    await cloudEnter();
    showGate('who');
  });
  $('#cgPass').addEventListener('keydown', e=>{ if(e.key==='Enter') $('#cgEnter').click(); });
  $('#cgGo').addEventListener('click', ()=>{ cloudFinishLogin($('#cgName').value); });
  const logoutBtn=$('#cloudLogout'); if(logoutBtn) logoutBtn.addEventListener('click', ()=>{ store.save('wrt_cloud_me',null); firebase.auth().signOut(); location.reload(); });

  // already signed-in this browser? skip the passcode, go straight in
  firebase.auth().onAuthStateChanged(async user=>{
    if(user && !CLOUD.ready){
      CLOUD.authedAs = (user.email===CLOUD.adminEmail) ? 'admin' : 'member';
      await cloudEnter();
      const saved=store.load('wrt_cloud_me',null);
      if(saved && saved.name){ CLOUD.me={name:saved.name, admin:CLOUD.authedAs==='admin'}; applyRoleUI(); $('#cloudGate').hidden=true; }
      else showGate('who');
    } else if(!user){
      showGate('pass');
    }
  });
}

/* init */
wireEvents();
renderAll();
setView(currentView);
cloudInit();
