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
let autoAdd = store.load('wrt_autoadd', true);        // auto-create members from report owners
let fuzzy   = store.load('wrt_fuzzy', true);          // allow nickname/typo/partial matching (default on)
const filters = {q:'', project:'', member:'', status:'', role:'', hideEmpty:false};

function persist(){
  // members + batches are tiny and must always survive; tasks may be large (images)
  try{ store.save(LS.members, members); }catch(e){ console.warn(e); }
  try{ store.save(LS.batches, batches); }catch(e){ console.warn(e); }
  try{ store.save('wrt_idcode', idCodeMap); }catch(e){}
  try{ store.save('wrt_projalias', projAliases); }catch(e){}
  try{ store.save('wrt_projmeta', projMeta); }catch(e){}
  try{ store.save('wrt_deleted', deletedNames); }catch(e){}
  try{ store.save('wrt_projmerge', projMerge); }catch(e){}
  try{ store.save(LS.tasks, tasks); }
  catch(e){ console.warn('tasks persist failed', e);
    toast('⚠ 任務太多（多為圖片）超過瀏覽器儲存上限，名單已保留，任務本次未存。'); }
}

/* downscale an image dataURL to a small JPEG thumbnail (keeps localStorage/Word small) */
function shrinkImage(dataUrl, max=620, q=0.66){
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

/* ---------- helpers ---------- */
const $  = s => document.querySelector(s);
const $$ = s => Array.from(document.querySelectorAll(s));
const uid = () => Math.random().toString(36).slice(2,10);
const esc = s => String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function toast(msg){ const t=$('#toast'); t.textContent=msg; t.hidden=false; clearTimeout(t._t); t._t=setTimeout(()=>t.hidden=true,2800); }
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
    // CSV style "name,alias,alias" OR "name: alias, alias"
    let name, aliasStr='';
    if(/[:：]/.test(line)){ const p=line.split(/[:：]/); name=p[0].trim(); aliasStr=p.slice(1).join(':'); }
    else if(line.includes(',')){ const p=line.split(','); name=p[0].trim(); aliasStr=p.slice(1).join(','); }
    else name=line;
    if(!name) return;
    const aliases=aliasStr.split(/[,，、/]/).map(s=>s.trim()).filter(Boolean);
    out.push({name, aliases, role});
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
    } else {
      members.push({id:uid(), name:m.name, aliases:m.aliases||[], role:m.role||'', role2:m.role2||''});
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
function clearMembers(){ if(confirm('確定清空整份成員名單？(任務不會被刪除)')){ members=[]; deletedNames=[]; persist(); renderAll(); } }

/* =====================================================================
   NAME MATCHING  (owner-priority, multi-owner, fuzzy)
   ===================================================================== */
function splitOwners(raw){
  if(!raw) return [];
  // drop parentheticals like "Jin(SW:Jonas)" -> "Jin"; split on / , & newline 、 + and whitespace
  return String(raw).replace(/\([^)]*\)/g,' ')
    .split(/[\/,&\n、+]|\band\b|\s+/i).map(s=>s.trim()).filter(s=>s.length>=2);
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
      else { toast('不支援的格式: '+ext); continue; }
    }catch(err){ console.error(err); toast('解析失敗: '+f.name+' — '+err.message); continue; }
    if(!parsed.length){ toast('在 '+f.name+' 找不到可解析的任務'); continue; }
    // image-pasted report placeholders stay IN MEMORY (not persisted -> no localStorage bloat)
    const reports=parsed.filter(t=>t._imageReport);
    pendingReports.push(...reports);
    overlayTasks(parsed.filter(t=>!t._imageReport), f.name);
  }
  if(autoAdd){ const added=autoAddFromReport(); if(added.length) toast('已從報告自動加入 '+added.length+' 位成員'); }
  dedupeTasks(); cleanupGarbledMembers();
  renderAll();
  if(pendingReports.length){
    if(navigator.onLine){
      toast(`偵測到 ${pendingReports.length} 張圖片頁，背景 OCR 辨識中…`);
      setTimeout(()=>ocrAllReports(true), 400);          // auto-run in background
    } else {
      toast(`偵測到 ${pendingReports.length} 張圖片頁，連網後按「🔍 OCR 圖片報告」即可辨識`);
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

const CODE_RE=/B\d{2}[A-Z]\d{3}[A-Z]?\d{0,2}|SDX[-\s]?\d{2}(?:[\/\s、,-]\d{2})?|VB\d{3}/i;
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
    const rm=joinedText.match(/reporter\s*[:：]\s*([A-Za-z][\w.]*(?:\s*[\/&,]\s*[A-Za-z][\w.]*)*)/i);
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
            const s=await shrinkImage(url);
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
      allText.forEach(line=>{ const m=line.match(/^([A-Za-z ]{3,40})[:：]\s*(.+)$/); if(m){ const k=canonField(m[1]); if(k) f[k]=m[2].trim(); } });
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
      // (a) detail/analysis page right after a member's report -> attach to ONE of that member's tasks
      //     (storage-light). Skipped for closing/divider slides so the deck's "Thank You" page
      //     never lands inside the last member's section.
      if(!isClosing){
        if(lastTaskSlide>=0 && (slideNo-lastTaskSlide)<=4 && out[lastSectionStart]){
          out[lastSectionStart]._images=(out[lastSectionStart]._images||[]).concat(images);
        } else if(!big && lastTaskIdxBySlide>=0){
          out[lastTaskIdxBySlide]._images=(out[lastTaskIdxBySlide]._images||[]).concat(images);
        }
      }
      // (b) ANY big single image might be a PASTED REPORT -> keep a hi-res OCR placeholder; OCR decides
      //     (if OCR finds "Reporter: X" -> creates X's tasks; otherwise placeholder is just dropped)
      if(big){
        const hi=[];
        for(let k=0;k<images.length;k++){ const s=await shrinkImage(origs[k]||images[k].data, 1200, 0.82); hi.push({id:uid(), data:s.data, w:s.w, h:s.h}); }
        const ph=makeTaskFromFields({project:'圖片式報告', current:'image report slide '+slideNo}, slideNo);
        ph._images=hi; ph._imageReport=true; out.push(ph);
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
    b.forEach(line=>{ const m=line.match(/^([A-Za-z ]{3,40})[:：]\s*(.+)$/); if(m){const k=canonField(m[1]); if(k)f[k]=m[2].trim();} });
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
  blocks.forEach(b=>{ const f={}; b.forEach(line=>{ const m=line.match(/^([A-Za-z ]{3,40})[:：]\s*(.+)$/); if(m){const k=canonField(m[1]); if(k)f[k]=m[2].trim();} }); if(f.project||f.current||f.owner) out.push(makeTaskFromFields(f,0)); });
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
      complexity:p.complexity, progress:p.progress, images:p._images||p.images||[],
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
      if(ex.manualEdit){ inc.risk=ex.risk; inc.complexity=ex.complexity; inc.progress=ex.progress; inc.manualEdit=true; }       // keep manual risk/cx/%
      Object.assign(ex, inc);
      nUpd++; touched.push(ex);
    } else {
      ex.status='Unchanged'; ex.source=sourceName; nUnc++; touched.push(ex);
    }
  });
  batches.unshift({name:sourceName,date:new Date().toISOString().slice(0,16).replace('T',' '),
    nnew:nNew,nupd:nUpd,nunc:nUnc});
  persist();
  toast(`匯入 ${sourceName}：New ${nNew} · Updated ${nUpd} · Unchanged ${nUnc}`);
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
  if(oldT.progress!==newT.progress) d.push(`進度 ${oldT.progress}% → ${newT.progress}%`);
  if(oldT.risk!==newT.risk) d.push(`風險 ${oldT.risk} → ${newT.risk}`);
  if(oldT.due!==newT.due && newT.due) d.push(`期限 ${oldT.due||'—'} → ${newT.due}`);
  if(norm(oldT.current)!==norm(newT.current)) d.push('工作描述已更新');
  if(norm(oldT.next)!==norm(newT.next)) d.push('下週計畫已更新');
  return d.join('；');
}
function deleteTask(id){ tasks=tasks.filter(t=>t.id!==id); persist(); renderAll(); }
function resetTasks(){ if(confirm('清空所有任務？(成員名單會保留)')){ tasks=[]; batches=[]; persist(); renderAll(); } }

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
function isClosed(t){ return (t.progress||0)>=100; }
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

/* ---------- Weekly narrative (plain text, used by preview + mirrors Word export) ---------- */
function memberNarrative(name, list){
  const lbl=t=>t.projectLabel||shortProj(t.project);
  let out='*'+name+'\n';
  const cur=list.filter(t=>t.current), nexts=list.filter(t=>t.next);
  if(!cur.length && !nexts.length) return out+'Pending input — no items reported this week.\n\n';
  out+='This week: [ '+(cur.length?cur.map(t=>`${lbl(t)} - ${t.progress}%`).join(' | '):'—')+' ]\n';
  cur.forEach((t,i)=>{
    out+=`${i+1}. ${lbl(t)}: ${rewriteProfessional(t.current)}\n`;
    out+=`   Status: ${statusWord(t)} | ${t.progress}% complete | risk ${(t.risk||'M')[0]} | complexity ${t.complexity||'Medium'}\n`;
    if(t.shared) out+=`   Shared owner: ${(t.ownerIds||[]).map(memberName).join(', ')}\n`;
    if(t.images&&t.images.length) out+=`   Attachments: ${t.images.length} image(s)\n`;
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
  renderMembers(); renderBatches(); renderStats(); renderCatalog(); renderCharts(); renderFilters(); renderMembersArea(); renderWorkbenchSelect(); updateOcrBtn();
}
function renderFilters(){
  const ps=$('#filterProject'); if(ps){
    const groups=projectGroups();                       // canonical (merged) projects, by clean label
    if(filters.project && !groups.some(g=>g.projk===filters.project)) filters.project='';
    ps.innerHTML='<option value="">全部專案</option>'+groups.map(g=>`<option value="${esc(g.projk)}" ${filters.project===g.projk?'selected':''}>${esc(g.label)}</option>`).join('');
  }
  const ms=$('#filterMember'); if(ms){
    ms.innerHTML='<option value="">全部成員</option>'+members.map(m=>`<option value="${m.id}" ${filters.member===m.id?'selected':''}>${esc(m.name)}</option>`).join('')+
      `<option value="__un__" ${filters.member==='__un__'?'selected':''}>Unassigned</option>`;
  }
  const chk=$('#autoAddChk'); if(chk) chk.checked=autoAdd;
  const fz=$('#fuzzyChk'); if(fz) fz.checked=fuzzy;
  const he=$('#hideEmptyChk'); if(he) he.checked=filters.hideEmpty;
  const fs=$('#filterStatus'); if(fs) fs.value=filters.status;
  const fr=$('#filterRole'); if(fr) fr.value=filters.role;
}

function roleOptions(sel){ return '<option value="">—</option>'+ROLES.map(r=>`<option ${sel===r?'selected':''}>${r}</option>`).join(''); }
function roleBadge(role, sub){ return role?`<span class="role-badge r-${role}${sub?' sub':''}">${esc(role)}</span>`:''; }
function memberRoleBadges(id){ return memberRoles(id).map((r,i)=>roleBadge(r,i>0)).join(''); }
function renderMembers(){
  $('#memberCount').textContent=members.length;
  $('#memberChips').innerHTML=members.map(m=>`
    <li class="mchip role-${m.role||'none'}" draggable="true" data-mid="${m.id}">
      <div class="mchip-top">
        <span class="drag-handle" title="拖曳調整順序">⠿</span>
        <span class="mname">${esc(m.name)}</span>
        ${roleBadge(m.role)}${roleBadge(m.role2,true)}
        ${m.aliases.length?`<span class="alias">${esc(m.aliases.join(', '))}</span>`:''}
        <button class="mdel" title="刪除" data-del-member="${m.id}">✕</button>
      </div>
      <div class="mchip-roles">
        <label>主<select class="role-mini" data-setrole="${m.id}">${roleOptions(m.role)}</select></label>
        <label>副<select class="role-mini sub" data-setrole2="${m.id}">${roleOptions(m.role2)}</select></label>
      </div>
    </li>`).join('');
  wireMemberDrag();
}
function renderBatches(){
  $('#batchList').innerHTML = batches.length? batches.slice(0,8).map(b=>`
    <li><span>${esc(b.name)}<br><small>${b.date}</small></span>
    <span><b style="color:var(--new)">${b.nnew}N</b> <b style="color:var(--updated)">${b.nupd}U</b> ${b.nunc}=</span></li>`).join('')
    : '<li>尚無匯入紀錄</li>';
}
function renderStats(){
  const vt=visibleTasks();
  const total=vt.length;
  const closed=vt.filter(isClosed).length;
  const active=total-closed;
  const projCount=new Set(vt.map(t=>resolveProjk(t.projk||t.key))).size;  // merged projects count as one
  const avg=total?Math.round(vt.reduce((s,t)=>s+(t.progress||0),0)/total):0;
  const high=vt.filter(t=>t.risk==='High'&&!isClosed(t)).length;
  // each stat card doubles as navigation: [icon, value, label, colorClass, view, statusFilter]
  const cards=[
    ['📁', projCount, '專案 Projects', '', 'catalog', ''],
    ['👥', members.length, '成員 Members', '', 'members', ''],
    ['📋', total, '任務 Tasks', '', 'members', ''],
    ['✅', closed+'/'+total, '已結案 Closed', 'ok', 'members', 'closed'],
    ['📊', avg+'%', '平均進度 Avg', 'accent', 'workload', ''],
    ['⚠️', high, '高風險 High risk', 'warn', 'members', 'highrisk'],
  ];
  $('#statsRow').innerHTML=cards.map(([ic,n,l,c,view,flt])=>{
    const active = view===currentView && (flt||'')===(filters.status||'');
    return `<div class="stat nav ${c} ${active?'active':''}" data-nav="${view}" data-flt="${flt}">
      <div class="st-ic">${ic}</div><div><div class="num">${n}</div><div class="lbl">${l}</div></div></div>`;
  }).join('');
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
  return [...map.values()].sort((a,b)=>b.tasks.length-a.tasks.length);
}
function mergeProjects(fromProjk, toProjk){
  if(!fromProjk||!toProjk||fromProjk===toProjk) return;
  if(resolveProjk(toProjk)===fromProjk) return;          // avoid cycles
  projMerge[fromProjk]=toProjk; persist(); renderCatalog(); renderStats();
  toast('已把專案併入：'+(projectGroups().find(g=>g.projk===resolveProjk(toProjk))||{}).label);
}
function unmergeProject(projk){ delete projMerge[projk]; Object.keys(projMerge).forEach(k=>{ if(projMerge[k]===projk) delete projMerge[k]; }); persist(); renderCatalog(); }
const CATEGORIES=['Module','IDU','ODU','Dongle','General'];
function inferCategory(projStr){
  const s=String(projStr||'').toLowerCase();
  if(/dongle|\bdg\d|redcap/.test(s)) return 'Dongle';
  if(/\bodu\b/.test(s)) return 'ODU';
  if(/\bid\s?\d{3}\b|\bidu\b/.test(s)) return 'IDU';
  if(/module|\bfr1\b|\bfr2\b|sdx7\d/.test(s)) return 'Module';
  return 'General';
}
function projCategory(g){ const m=projMeta[g.projk]||{}; return m.category && CATEGORIES.includes(m.category)? m.category : (m.category||inferCategory(g.projStr||g.label)); }
function renderCatalog(){
  const sel=$('#catalogMember');
  if(sel) sel.innerHTML='<option value="">All members · 全部成員</option>'+
    members.map(m=>`<option value="${m.id}" ${catalogMember===m.id?'selected':''}>${esc(m.name)}</option>`).join('');
  let groups=projectGroups();
  if(catalogMember) groups=groups.map(g=>({...g, tasks:g.tasks.filter(t=>(t.ownerIds||[]).includes(catalogMember))}))
                                  .filter(g=>g.tasks.length);
  const cont=$('#projectCatalog'); if(!cont) return;
  if(!groups.length){ cont.innerHTML='<p class="hint">尚無專案，請先匯入週報。</p>'; return; }
  // group by Product Category (Module / IDU / ODU / Dongle / 其他), 其他 collapsed by default
  const byCat={}; groups.forEach(g=>{ const c=projCategory(g); (byCat[c]=byCat[c]||[]).push(g); });
  const order=[...CATEGORIES, ...Object.keys(byCat).filter(c=>!CATEGORIES.includes(c))];
  cont.innerHTML = order.filter(c=>byCat[c]).map(c=>{
    const list=byCat[c]; const nTasks=list.reduce((s,g)=>s+g.tasks.length,0);
    const collapsed = (c==='General'||c==='其他');
    return `<details class="cat-group" ${collapsed?'':'open'}>
      <summary class="cat-head"><span class="cat-name cat-${esc(c)}">${esc(c)}</span>
        <span class="cat-meta">${list.length} 專案 · ${nTasks} 任務</span></summary>
      ${list.map(catalogCard).join('')}</details>`;
  }).join('');
  wireProjectDrag();
}
function optTags(arr,val){ return arr.map(x=>`<option ${x===val?'selected':''}>${x}</option>`).join(''); }
function catalogCard(g){
  const meta=projMeta[g.projk]||{};
  const desc=(meta.desc!=null&&meta.desc!=='')?meta.desc:cleanDesc(g.projStr);
  const ed=editingProj===g.projk;
  const sub=`${g.tasks.length} tasks · ${g.mem.size} members · ${g.high} high risk · ${g.closed} closed`;
  const curCat=projCategory(g);
  const metaRow = ed
    ? `<div class="pc-meta editing">
         <input class="pc-in" data-mf="customer" placeholder="Customer" value="${esc(meta.customer||'')}">
         <select class="pc-in" data-mf="category">${['',...CATEGORIES].map(c=>`<option value="${esc(c)}" ${(meta.category||curCat)===c?'selected':''}>${c||'(類別)'}</option>`).join('')}</select>
         <input class="pc-in" data-mf="desc" placeholder="Description" value="${esc(desc)}">
         <button class="btn sm primary" data-save-proj="${g.projk}">儲存</button>
       </div>`
    : `<div class="pc-meta">
         <div class="pc-cell">${esc(meta.customer||'—')}</div>
         <div class="pc-cell"><span class="cat-name cat-${esc(curCat)}">${esc(curCat)}</span></div>
         <div class="pc-cell desc">${esc(desc)}</div>
       </div>`;
  const merged=Object.keys(projMerge).filter(k=>resolveProjk(k)===g.projk).length;
  return `<div class="proj-card cat-b-${esc(curCat)}" draggable="true" data-projk="${esc(g.projk)}">
    <div class="pc-head">
      <div><h3><span class="drag-dot" title="拖曳此專案併到另一個">⠿</span> ${esc(g.label)}${merged?` <span class="merged-tag" title="已併入其他專案">＋${merged} 併</span>`:''}</h3><div class="pc-sub">${sub}</div></div>
      <div class="pc-actions">
        ${merged?`<button class="btn sm" data-unmerge="${esc(g.projk)}" title="取消併入">↩ 取消併</button>`:''}
        <button class="btn sm" data-edit-proj="${g.projk}">${ed?'取消':'✎ Edit'}</button>
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
  const groups={}; members.filter(m=>!ex.has(m.id)).forEach(m=>{ const r=m.role||'未分類'; (groups[r]=groups[r]||[]).push(m); });
  return [...ROLES,'未分類'].filter(r=>groups[r]).map(r=>
    `<optgroup label="${r}">${groups[r].map(m=>`<option value="${m.id}">${esc(m.name)}</option>`).join('')}</optgroup>`).join('');
}
function catalogTaskRow(t){
  const owners=(t.ownerIds||[]).map(id=>`<span class="owner-chip">${memberRoleBadges(id)}${esc(memberName(id))}<button data-rmowner="${t.id}|${id}" title="移除">×</button></span>`).join('')||'<span class="owner-chip none">Unassigned</span>';
  return `<div class="ctask">
    <div class="ctask-desc" data-opentask="${t.id}">${esc((t.current||t.next||'(無描述)').slice(0,150))}
      <span class="tag st-${t.status}">${esc(t.status)}</span>${isClosed(t)?' <span class="tag closed">✓ Closed</span>':''}</div>
    <div class="ctask-ctl">
      <span class="owners">${owners}<select class="add-owner" data-addowner="${t.id}"><option value="">＋ 加人</option>${memberOptionsByRole(t.ownerIds)}</select></span>
      <label>Risk<select data-edit="risk" data-tid="${t.id}">${optTags(['Low','Medium','High'],t.risk)}</select></label>
      <label>Cx<select data-edit="complexity" data-tid="${t.id}">${optTags(['Low','Medium','High'],t.complexity)}</select></label>
      <label>%<input type="number" min="0" max="100" value="${t.progress}" data-edit="progress" data-tid="${t.id}" class="prog-in"></label>
    </div>
  </div>`;
}
function editTaskField(id, field, val){
  const t=tasks.find(x=>x.id===id); if(!t) return;
  if(field==='progress') t.progress=Math.max(0,Math.min(100,+val||0));
  else if(field==='project'){ t.project=val; t.projk=projKeyOf(val); t.projectLabel=projLabelOf(val); }
  else t[field]=val;
  if(field==='progress'||field==='risk'||field==='complexity') t.manualEdit=true;  // survive re-import
  persist(); renderStats(); renderCatalog(); renderMembersArea();
  if(!$('#taskModal').hidden && _openTaskId===id && field!=='project') openTask(id);  // refresh Closed/Status live
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
  const m={}; card.querySelectorAll('.pc-in').forEach(inp=>{ m[inp.dataset.mf]=inp.value.trim(); });
  projMeta[projk]=m; editingProj=''; persist(); renderCatalog();
}
function renderCharts(){
  // workload per member — only members with tasks; note how many are pending
  const {map,unassigned}=buildBuckets();
  let wl=members.map(m=>({id:m.id, name:m.name, c:map.get(m.id).length}));   // ALL members in left-list order (0-task shown too)
  if(unassigned.length) wl.push({id:'__un__', name:'Unassigned', c:unassigned.length});
  const active=members.filter(m=>map.get(m.id).length>0).length;
  const maxW=Math.max(1,...wl.map(x=>x.c));
  $('#wlChartSub').textContent = `${active}/${members.length} 位有任務 · 點名字看其任務`;
  $('#workloadChart').innerHTML = wl.length
    ? wl.map(x=>progBar(x.name+(x.c===0?'（待輸入）':''), Math.round(x.c/maxW*100), x.c, 'wl', {mem:x.id})).join('')
    : '<p class="hint">尚無資料</p>';
}
// one chart row. cls 'wl' = workload (purple); else colored by progress. opts.proj/opts.mem make it clickable.
function progBar(label, pct, valText, cls, opts){
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

  members.forEach(m=>{
    if(filters.member && filters.member!==m.id) return;
    if(filters.role && m.role!==filters.role && m.role2!==filters.role) return;  // primary OR secondary discipline
    const list=map.get(m.id).filter(matchTask); list.forEach(t=>shown.add(t.id));
    if((filters.hideEmpty || filtering) && !list.length) return;
    html+=memberBlock(m.name, list, false, m.id);
  });
  if((!filters.member || filters.member==='__un__') && !filters.role){
    const ul=unassigned.filter(matchTask); ul.forEach(t=>shown.add(t.id));
    if(ul.length) html+=memberBlock('Unassigned', ul, true);
  }

  $('#membersArea').innerHTML = html ||
    '<div class="panel"><p class="hint">沒有符合條件的項目（或尚未加入成員 / 匯入週報）。</p></div>';
  const fc=$('#filterCount'); if(fc) fc.textContent='顯示 '+shown.size+' 筆任務';
}
function memberBlock(name, list, isUnassigned, mid){
  const avg=list.length?Math.round(list.reduce((s,t)=>s+(t.progress||0),0)/list.length):0;
  // aggregate this member's detail/analysis images into a header strip (easy to find, not buried in one card)
  const imgs=[]; const seen=new Set();
  list.forEach(t=>(t.images||[]).forEach(im=>{ if(im&&im.id&&!seen.has(im.id)){ seen.add(im.id); imgs.push(im); } }));
  const strip = imgs.length? `<div class="member-imgs" title="本成員的圖片/分析">📎 ${imgs.slice(0,10).map(im=>`<img src="${im.data}" data-light="${im.id}">`).join('')}${imgs.length>10?`<span class="more">+${imgs.length-10}</span>`:''}</div>` : '';
  const head=`<div class="member-head ${isUnassigned?'unassigned':''}">
    ${mid?memberRoleBadges(mid):''}<span class="name">${esc(name)}</span>
    <span class="meta">${list.length} 任務 · 平均 ${avg}%</span>
    ${list.length?'':'<span class="pending">Pending input</span>'}
    ${mid?`<button class="btn sm add-task-btn" data-addtask="${mid}">＋ 新增任務</button>`:''}</div>${strip}`;
  const cards=list.length? `<div class="task-grid">${list.map(taskCard).join('')}</div>` : '';
  return `<div class="member-block">${head}${cards}</div>`;
}
function taskCard(t){
  const sharedTag=t.shared?`<span class="tag shared">Shared owner</span>`:'';
  const thumbs=(t.images&&t.images.length)?`<div class="img-badge" title="${t.images.length} 張圖（見成員上方）">📎 ${t.images.length}</div>`:'';
  const delta=t.delta?`<div class="delta">⟳ Weekly delta：${esc(t.delta)}</div>`:'';
  const closed=isClosed(t);
  const shortP=t.projectLabel || String(t.project||'').split(/[\n(]/)[0].trim();
  return `<div class="card ${closed?'closed':''}" data-task="${t.id}">
    <button class="card-del" data-del-task="${t.id}" title="刪除任務">🗑</button>
    <div class="ct"><span class="proj">${esc(shortP)}</span></div>
    <div class="tags">
      ${closed?'<span class="tag closed">✓ Closed</span>':''}
      <span class="tag risk-${t.risk}">Risk ${esc(t.risk)}</span>
      <span class="tag cx">Cx ${esc(t.complexity)}</span>
      ${sharedTag}
      <span class="tag st-${t.status}">${esc(t.status)}</span>
    </div>
    <div class="desc">${esc(t.current || (t.next?('下週：'+t.next):'(無描述)'))}</div>
    <div class="prog"><span class="ptrack"><span class="pfill ${progClass(t.progress)}" style="width:${t.progress}%"></span></span><span class="pval ${closed?'done':''}">${t.progress}%</span></div>
    <div class="due">📅 ${esc(t.due||'—')} ${t.reporter?'· Reporter: '+esc(t.reporter):''}</div>
    ${delta}${thumbs}
  </div>`;
}

/* ---------- task detail modal ---------- */
let _openTaskId='';
function openTask(id){
  const t=tasks.find(x=>x.id===id); if(!t) return;
  _openTaskId=id;
  const owners=(t.ownerIds||[]).map(memberName).join(', ')||'—';
  const imgs=(t.images||[]).map(im=>`<img src="${im.data}" data-light="${im.id}">`).join('');
  $('#taskModalInner').innerHTML=`
    <div class="modal-head"><h2>${esc(t.projectLabel||t.project)}</h2><button class="icon-btn" data-close>✕</button></div>
    <div class="modal-body detail">
      <div class="tags">
        ${isClosed(t)?'<span class="tag closed">✓ Closed 已完成</span>':''}
        <span class="tag risk-${t.risk}">Risk ${esc(t.risk)}</span>
        <span class="tag cx">Complexity ${esc(t.complexity)}</span>
        ${t.shared?'<span class="tag shared">Shared owner</span>':''}
        <span class="tag st-${t.status}">${esc(t.status)}</span>
      </div>
      <div class="edit-row">
        <label>進度 %<input type="number" min="0" max="100" value="${t.progress}" data-edit="progress" data-tid="${t.id}" class="prog-in"></label>
        <label>Risk<select data-edit="risk" data-tid="${t.id}">${optTags(['Low','Medium','High'],t.risk)}</select></label>
        <label>複雜度<select data-edit="complexity" data-tid="${t.id}">${optTags(['Low','Medium','High'],t.complexity)}</select></label>
      </div>
      <div class="kv">
        <span class="k">Project (可改)</span><span><input class="proj-edit" value="${esc(t.project)}" data-edit="project" data-tid="${t.id}"></span>
        <span class="k">Status</span><span>${esc(statusLine(t))}</span>
        <span class="k">Due date</span><span>${esc(t.due||'—')}</span>
        <span class="k">原始 Owner (raw)</span><span>${esc(t.rawOwner||'—')}</span>
        <span class="k">Reporter</span><span>${esc(t.reporter||'—')}</span>
        <span class="k">分派後成員</span><span>${esc(owners)}</span>
        ${t.unmatched&&t.unmatched.length?`<span class="k">未對應</span><span style="color:var(--warn)">${esc(t.unmatched.join(', '))}</span>`:''}
        <span class="k">來源</span><span>${esc(t.source||'手動')}</span>
      </div>
      <div class="section-title">This week</div>
      <div>${esc(rewriteProfessional(t.current))||'—'}</div>
      <div class="section-title">Next week</div>
      <div>${esc(rewriteProfessional(t.next))||'—'}</div>
      ${t.delta?`<div class="section-title">Weekly delta</div><div style="color:var(--warn)">${esc(t.delta)}</div>`:''}
      ${t.prev?`<div class="section-title">前次描述</div><div class="hint">${esc(t.prev.current||'—')} (was ${t.prev.progress}%, risk ${t.prev.risk})</div>`:''}
      <div class="section-title">Issue analysis</div>
      <div class="analysis">${esc(t.analysis||generateAnalysis(t))}</div>
      ${imgs?`<div class="section-title">圖片 Attachments</div><div class="imgs">${imgs}</div>`:''}
    </div>
    <div class="modal-foot">
      <button class="btn danger" data-del-task="${t.id}">刪除任務</button>
      ${(t.images&&t.images.length)?`<button class="btn" data-ocr="${t.id}">🔍 OCR 圖片文字</button>`:''}
      <button class="btn primary" data-export-member="${(t.ownerIds||[])[0]||''}">匯出此成員 Word</button>
      <button class="btn" data-close>關閉</button>
    </div>`;
  $('#taskModal').hidden=false;
}

/* ---------- project drill-down: all tasks under a project ---------- */
function openProject(projk){
  const list=tasks.filter(t=>!t.imageReport && resolveProjk(t.projk||t.key)===resolveProjk(projk));
  if(!list.length) return;
  const label=list[0].projectLabel||shortProj(list[0].project);
  const avg=Math.round(list.reduce((s,t)=>s+(t.progress||0),0)/list.length);
  const closed=list.filter(isClosed).length;
  $('#taskModalInner').innerHTML=`
    <div class="modal-head"><h2>${esc(label)} <span class="ch-sub">${list.length} 個任務 · 平均 ${avg}% · ${closed} 已結案</span></h2>
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
          <div class="pt-desc">${esc((t.current||t.next||'(無描述)').slice(0,170))}</div>
          <div class="prog"><span class="ptrack"><span class="pfill ${progClass(t.progress)}" style="width:${t.progress}%"></span></span><span class="pval ${isClosed(t)?'done':''}">${t.progress}%</span></div>
        </div>`).join('')}
      </div>
    </div>
    <div class="modal-foot"><button class="btn" data-close>關閉</button></div>`;
  $('#taskModal').hidden=false;
}

/* ---------- custom project groupings ---------- */
function parseAliasText(text){
  const out=[];
  text.split(/[\r\n]+/).forEach(line=>{ line=line.trim(); if(!line) return;
    let label, rest;
    if(/[:：]/.test(line)){ const p=line.split(/[:：]/); label=p[0].trim(); rest=p.slice(1).join(':'); }
    else { label=line; rest=line; }
    const tokens=[...new Set([label, ...rest.split(/[,，、\/]/)].map(s=>norm(s)).filter(Boolean))];
    if(label) out.push({key:norm(label), label, tokens});
  });
  return out;
}
function aliasToText(){ return projAliases.map(g=>g.label+': '+g.tokens.join(', ')).join('\n'); }
function applyProjAliases(text){
  projAliases=parseAliasText(text);
  tasks.forEach(t=>{ t.projk=projKeyOf(t.project); t.projectLabel=projLabelOf(t.project); }); // re-group existing
  persist(); renderAll(); toast('已套用專案歸併（'+projAliases.length+' 組）');
}

/* =====================================================================
   WORKBENCH
   ===================================================================== */
let wbImages=[];
function openWorkbench(preMemberId){
  renderWorkbenchSelect();
  if(preMemberId && members.some(m=>m.id===preMemberId)) $('#wbMember').value=preMemberId;
  $('#phraseRow').innerHTML=Object.keys(PHRASES).map(k=>`<button data-phrase="${esc(k)}">${esc(k)}</button>`).join('');
  wbImages=[]; renderWbThumbs();
  ['#wbProject','#wbThisWeek','#wbIssue','#wbNext'].forEach(s=>$(s).value='');
  $('#wbProgress').value=0;
  $('#workbenchModal').hidden=false;
}
function renderWorkbenchSelect(){
  const sel=$('#wbMember'); if(!sel) return;
  sel.innerHTML=members.map(m=>`<option value="${m.id}">${esc(m.name)}</option>`).join('')||'<option value="">(先加入成員)</option>';
}
function renderWbThumbs(){
  $('#wbThumbs').innerHTML=wbImages.map((im,i)=>`<div class="th"><img src="${im.data}" data-light="${im.id}"><button class="rm" data-rm-wb="${i}">✕</button></div>`).join('');
}
function saveWorkbench(){
  const mid=$('#wbMember').value; if(!mid){ toast('請先加入成員'); return; }
  const project=$('#wbProject').value.trim()||'Untitled Project';
  const current=$('#wbThisWeek').value.trim();
  const name=memberName(mid);
  const p={project, reporter:name, rawOwner:name, current,
    next:$('#wbNext').value.trim(), risk:$('#wbRisk').value, due:$('#wbDue').value,
    complexity:$('#wbComplexity').value, progress:+$('#wbProgress').value||0,
    _images:wbImages.slice()};
  if($('#wbIssue').value.trim()) p.current += (p.current?' ':'')+'Issue/Note: '+$('#wbIssue').value.trim();
  overlayTasks([p], '工作台手動輸入');
  renderAll();
  $('#workbenchModal').hidden=true;
  toast('已儲存任務給 '+name);
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
  // This week summary:  [ P1 - x% | P2 - y% | P3 - z% ]
  const summary='[ '+(cur.length?cur.map(t=>`${plabel(t)} - ${t.progress}%`).join(' | '):'—')+' ]';
  body+=P('This week: '+summary,{bold:true});
  cur.forEach((t,i)=>{
    body+=P(`${i+1}. ${plabel(t)}: ${rewriteProfessional(t.current)}`,{indent:200});
    body+=P(`Status: ${statusWord(t)} | ${t.progress}% complete | risk ${(t.risk||'M')[0]} | complexity ${t.complexity||'Medium'}`,{indent:540,color:'444C5C'});
    // lean: only surface analysis when there is a real risk/blocker
    if(t.risk==='High' || /block|defect|timeout|fail|overdue|debug/i.test(t.current||''))
      body+=P('Analysis: '+(t.analysis||generateAnalysis(t)),{indent:540,color:'6B7A92'});
    if(t.shared) body+=P('Shared owner: '+(t.ownerIds||[]).map(memberName).join(', '),{indent:540,color:'1A7A8A'});
    if(t.delta) body+=P('Weekly delta: '+t.delta,{indent:540,color:'B5790F'});
    if(t.images&&t.images.length){
      body+=P(`Attachments: ${t.images.length} image(s)`,{indent:540,color:'6B7A92'});
      t.images.forEach(im=>{ const rid=mediaCollector(im.data); if(rid) body+=wpDrawing(rid,im); });
    }
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
function pptPlabel(t){ return t.projectLabel||shortProj(t.project); }
function progColor(p){ return p>=100?PPT.green : p>=70?PPT.blue : p>=34?PPT.cyan : PPT.amber; }
function collectImages(list){
  const out=[], seen=new Set();
  (list||[]).forEach(t=>(t.images||[]).forEach(im=>{
    const k=im&&(im.id||im.data); if(im&&im.data&&!seen.has(k)){ seen.add(k); out.push(im); }
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
  if(typeof PptxGenJS==='undefined'){ toast('PPTX 函式庫未載入'); return null; }
  const {map,unassigned}=buildBuckets();
  const targets = memberIds&&memberIds.length
    ? members.filter(m=>memberIds.includes(m.id)) : members.slice();
  if(!targets.length){ toast('沒有成員可匯出'); return null; }
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
    s.addText([
      {text:name,options:{bold:true,color:PPT.dark,fontFace:PPT.fontB}},
      role?{text:'   '+role,options:{color:PPT.teal,fontSize:15,bold:true,fontFace:PPT.font}}:{text:''}
    ],{x:0.7,y:0.6,w:10.6,h:0.64,fontSize:28,fontFace:PPT.fontB,valign:'middle'});
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
  const MAXY=6.9;                          // bottom limit for content before a new slide
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
    return c*0.235 + 0.2;
  }
  // one image card (white card, cyan top, contained image, caption)
  function imgCard(s,x,y,w,h,im,n){
    s.addShape(RR,{x,y,w,h,rectRadius:0.05,fill:{color:PPT.white},line:{color:PPT.line,width:1}});
    s.addShape(R,{x,y,w,h:0.06,fill:{color:PPT.cyan}});
    try{ s.addImage({data:im.data,x:x+0.15,y:y+0.2,w:w-0.3,h:h-0.55,sizing:{type:'contain',w:w-0.3,h:h-0.55}}); }catch(e){}
    s.addText('圖 '+n,{x:x+0.15,y:y+h-0.32,w:w-0.3,h:0.26,fontSize:10,color:PPT.gray,fontFace:PPT.font});
  }
  // member's images → ONE big image per slide
  function attachPages(name, role, imgs){
    imgs.forEach((im,i)=>{
      const s=pptx.addSlide(); header(s,name,role,'附件 attachments');
      sectionLabel(s,'ATTACHMENTS  ｜  Issue 圖片 / 量測',1.55);
      imgCard(s,0.9,1.98,11.53,4.9,im,i+1);
    });
  }
  // build ONE table (header + the given rows) at vertical position `top`
  function buildTable(s, top, key, chunk){
    const cols=['Project', key==='cur'?'Job & Issue':'Plan', 'Due date', 'Status'];
    const head=cols.map(t=>({text:t,options:{bold:true,color:'FFFFFF',fill:{color:PPT.navy},fontSize:HFONT,valign:'middle',align:(t==='Due date'||t==='Status')?'center':'left'}}));
    const rows=[head];
    chunk.forEach((gp,idx)=>{ const st=groupStatus(gp), bg={color: idx%2?'F1F5FB':'FFFFFF'};
      rows.push([
        {text:gp.label, options:{bold:true,color:PPT.navy,fontSize:TFONT,valign:'top',fill:bg}},
        {text:cellRuns(gp[key]), options:{valign:'top',fill:bg}},
        {text:gp.due||'—', options:{color:PPT.gray,fontSize:TFONT,align:'center',valign:'middle',fill:bg}},
        {text:st.text, options:{bold:true,color:st.color,fontSize:TFONT,align:'center',valign:'middle',fill:bg}}
      ]);
    });
    s.addTable(rows,{x:0.7,y:top,w:11.93,colW:COLW,border:{type:'solid',color:'D8E0EC',pt:0.5},
      fontFace:PPT.font,fontSize:TFONT,valign:'top',autoPage:false,margin:[5,6,5,6]});
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
      s.addText('本週尚未提供工作內容',{x:3.0,y:3.66,w:7.33,h:0.5,fontSize:13,color:PPT.gray,fontFace:PPT.font,align:'center'});
      return;
    }
    const HEAD_H=0.34;
    let s=null, y=0, started=false;
    const newSlide=()=>{ s=pptx.addSlide(); header(s,name,role, started?'(續) cont.':''); started=true; y=1.5; };
    function section(label, rowsG, key){
      if(!rowsG.length) return;
      let i=0;
      while(i<rowsG.length){
        const firstH=estRowH(rowsG[i][key]);
        if(!s || y+0.46+HEAD_H+firstH > MAXY) newSlide();    // room for label + header + 1 row
        sectionLabel(s,label,y); y+=0.46;
        const top=y; let used=HEAD_H; const chunk=[];
        while(i<rowsG.length){
          const h=estRowH(rowsG[i][key]);
          if(chunk.length && top+used+h > MAXY) break;
          chunk.push(rowsG[i]); used+=h; i++;
        }
        buildTable(s, top, key, chunk);
        y = top + used + 0.32;                                // gap before next section
      }
    }
    newSlide();
    section('THIS WEEK  ｜  本週工作', thisRows, 'cur');
    section('NEXT WEEK  ｜  下週計畫', nextRows, 'next');
    attachPages(name,role,imgs);
  }

  targets.forEach(m=>renderMember(m.name,[m.role,m.role2].filter(Boolean).join(' · '),map.get(m.id)||[]));
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
  toast('已匯出 '+r.fname);
}

async function exportWord(memberIds){
  const {map,unassigned}=buildBuckets();
  const targets = memberIds && memberIds.length
    ? members.filter(m=>memberIds.includes(m.id))
    : members.slice();
  if(!targets.length){ toast('沒有成員可匯出'); return; }

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
  targets.forEach(m=>{ bodyXml+=memberReportXml(m.name, map.get(m.id), collect); });
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
  toast('已匯出 '+fname);
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
    s.onerror=()=>{ _tessPromise=null; rej(new Error('無法載入 OCR 函式庫（OCR 首次使用需要網路連線）')); };
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
// upscale + grayscale + contrast-stretch -> Tesseract reads small/blurry table text much better
function preprocessForOcr(dataUrl){
  return new Promise(res=>{
    const img=new Image();
    img.onload=()=>{
      let w=img.naturalWidth||img.width, h=img.naturalHeight||img.height; if(!w||!h){ res(dataUrl); return; }
      const scale = w<1900 ? Math.min(2.6, 1900/w) : 1;
      w=Math.round(w*scale); h=Math.round(h*scale);
      const c=document.createElement('canvas'); c.width=w; c.height=h;
      const ctx=c.getContext('2d'); ctx.drawImage(img,0,0,w,h);
      try{
        const d=ctx.getImageData(0,0,w,h), px=d.data; let mn=255,mx=0;
        for(let i=0;i<px.length;i+=4){ const g=(px[i]*0.3+px[i+1]*0.59+px[i+2]*0.11)|0; px[i]=px[i+1]=px[i+2]=g; if(g<mn)mn=g; if(g>mx)mx=g; }
        const rng=Math.max(1,mx-mn);
        for(let i=0;i<px.length;i+=4){ let v=(px[i]-mn)*255/rng; v=255*Math.pow(v/255,1.35); v=v<0?0:v>255?255:v; px[i]=px[i+1]=px[i+2]=v; }
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
    try{ await w.setParameters({ tessedit_pageseg_mode:'6', preserve_interword_spaces:'1' }); }catch(e){}
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
function ocrReportToTasks(text){
  const rm=text.match(/reporter\s*[:;：]?\s*([A-Za-z][A-Za-z.]+)/i);
  if(!rm) return [];                       // not a personal weekly report -> skip
  const reporter=rm[1];
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
  if(!t || !t.images || !t.images.length){ toast('此任務沒有附圖可辨識'); return; }
  toast('OCR 啟動中…首次需下載辨識模型，請稍候');
  try{
    const text=await ocrImages(t.images, (i,n)=>toast(`OCR 辨識中… ${i}/${n}`));
    if(!text){ toast('沒有辨識到文字'); return; }
    if(t.imageReport){
      const rows=ocrReportToTasks(text);
      if(rows.length){
        const parsed=convertOcrRows(rows, t.images);
        tasks=tasks.filter(x=>x.id!==id);
        overlayTasks(parsed, '圖片報告 OCR');
        if(autoAdd) autoAddFromReport();
        persist(); renderAll(); $('#taskModal').hidden=true;
        toast('✅ 圖片報告辨識完成，拆出 '+parsed.length+' 筆任務');
        return;
      }
      t.current=text; t.imageReport=false; t.project='OCR 報告'; t.projectLabel='OCR 報告';
    } else {
      t.current=(t.current? t.current+' ; ' : '')+'[OCR] '+text;
    }
    t.ocrDone=true; persist(); renderAll(); openTask(id);
    toast('✅ OCR 完成');
  }catch(e){ toast(e.message||'OCR 失敗'); }
}
let _ocrRunning=false;
async function downscaleImgs(imgs){    // shrink hi-res OCR images to storage-light thumbnails before attaching
  const out=[]; for(const im of (imgs||[])){ try{ const s=await shrinkImage(im.data, 640, 0.66); out.push({id:uid(), data:s.data, w:s.w, h:s.h}); }catch(e){} }
  return out;
}
async function ocrAllReports(silent){
  if(_ocrRunning) return;
  const reps=pendingReports.slice().sort((a,b)=>(a._slide||0)-(b._slide||0));   // process in slide order
  if(!reps.length){ if(!silent) toast('沒有待辨識的圖片頁'); return; }
  _ocrRunning=true;
  let made=0, kept=0, errs=0, lastReportTasks=null, lastReportSlide=-99;
  for(let i=0;i<reps.length;i++){
    const t=reps[i]; const slide=t._slide||0; toast(`🔍 OCR 圖片頁 ${i+1}/${reps.length}…（背景進行，可繼續操作）`);
    try{
      const text=await ocrImages(t._images||t.images, ()=>{});
      const rows=ocrReportToTasks(text);
      const isReport = rows.length>=2 || (rows.length===1 && /reporter/i.test(text));
      if(isReport){
        const parsed=convertOcrRows(rows, null);
        const small=await downscaleImgs(t._images||t.images);
        if(parsed[0]) parsed[0]._images=small;            // keep the report screenshot on the first task
        const created=overlayTasks(parsed,'圖片報告 OCR');
        made+=parsed.length; kept++;
        lastReportTasks=created; lastReportSlide=slide;
      } else if(lastReportTasks && lastReportTasks.length && (slide-lastReportSlide)<=8
                && !tableSlides.some(s=>s>lastReportSlide && s<slide)){
        // image page AFTER a report, with NO new person's table in between = that report's test data -> attach to its tasks
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
  toast(`✅ 圖片頁辨識完成：${kept} 張報告 → ${made} 筆任務${dropped?`（去重 ${dropped}）`:''}`);
}
function updateOcrBtn(){
  const n=pendingReports.length;
  const b=$('#ocrReportsBtn'); if(b) b.textContent='🔍 OCR 圖片報告'+(n?` (${n})`:'');
}
async function ocrToWorkbench(){
  if(!wbImages.length){ toast('請先上傳圖片'); return; }
  toast('OCR 啟動中…首次需下載辨識模型，請稍候');
  try{
    const text=await ocrImages(wbImages, (i,n)=>toast(`OCR 辨識中… ${i}/${n}`));
    if(!text){ toast('沒有辨識到文字'); return; }
    const ta=$('#wbThisWeek'); ta.value=(ta.value? ta.value+'\n':'')+text;
    toast('✅ OCR 完成，已填入 This week');
  }catch(e){ toast(e.message||'OCR 失敗'); }
}

/* =====================================================================
   EVENTS
   ===================================================================== */
function fileToDataURL(f){ return new Promise(res=>{ const r=new FileReader(); r.onload=()=>res(r.result); r.readAsDataURL(f); }); }

function wireEvents(){
  $('#reportInput').addEventListener('change', e=>{ if(e.target.files.length) importFiles(e.target.files); e.target.value=''; });
  $('#addMembersBtn').addEventListener('click', ()=>{ const t=$('#memberPaste').value; if(t.trim()){ addMembers(parseMemberText(t),{manual:true}); $('#memberPaste').value=''; toast('已更新成員名單'); } });
  $('#memberFileInput').addEventListener('change', async e=>{ const f=e.target.files[0]; if(f){ addMembers(parseMemberText(await f.text()),{manual:true}); toast('已從檔案加入成員'); } e.target.value=''; });
  $('#clearMembersBtn').addEventListener('click', clearMembers);
  $('#resetTasksBtn').addEventListener('click', resetTasks);
  $('#loadFromReportBtn').addEventListener('click', ()=>{ if(!tasks.length){ toast('請先匯入週報'); return; } const a=autoAddFromReport(); toast(a.length?('已補齊 '+a.length+' 位成員'):'報告中的負責人都已在名單'); });
  $('#autoAddChk').addEventListener('change', e=>{ autoAdd=e.target.checked; store.save('wrt_autoadd',autoAdd); });
  $('#fuzzyChk').addEventListener('change', e=>{ fuzzy=e.target.checked; store.save('wrt_fuzzy',fuzzy); reresolveAllTasks(); persist(); renderAll(); toast(fuzzy?'已開啟模糊比對':'已切回完整姓名比對'); });
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
  $('#narrMember').addEventListener('change', renderNarrative);
  $('#narrCopyBtn').addEventListener('click', ()=>{ navigator.clipboard&&navigator.clipboard.writeText($('#narrativeText').textContent); toast('已複製 narrative 文字'); });
  $('#narrExportBtn').addEventListener('click', ()=>{ const mid=$('#narrMember').value; exportWord(mid?[mid]:null); });
  $('#applyAliasBtn').addEventListener('click', ()=>applyProjAliases($('#aliasText').value));
  $('#aliasText').value=aliasToText();
  $('#wbSaveBtn').addEventListener('click', saveWorkbench);
  $('#wbOcrBtn').addEventListener('click', ocrToWorkbench);
  $('#wbImages').addEventListener('change', async e=>{ for(const f of e.target.files){ const s=await shrinkImage(await fileToDataURL(f)); wbImages.push({id:uid(), data:s.data, w:s.w, h:s.h}); } renderWbThumbs(); e.target.value=''; });

  // delegated clicks
  document.body.addEventListener('click', e=>{
    const t=e.target;
    if(t.closest('[data-close]')){ t.closest('.modal-overlay').hidden=true; return; }
    const navCard=t.closest('[data-nav]'); if(navCard){ navStat(navCard.dataset.nav, navCard.dataset.flt); return; }
    if(t.dataset.delMember){ deleteMember(t.dataset.delMember); return; }
    if(t.dataset.delTask){ deleteTask(t.dataset.delTask); $('#taskModal').hidden=true; return; }
    if(t.dataset.rmWb!==undefined){ wbImages.splice(+t.dataset.rmWb,1); renderWbThumbs(); return; }
    if(t.dataset.phrase){ const ta=$('#wbThisWeek'); ta.value=(ta.value?ta.value+' ':'')+PHRASES[t.dataset.phrase]; return; }
    if(t.dataset.light){ openLight(t.getAttribute('src')); return; }
    if(t.dataset.exportMember!==undefined){ if(t.dataset.exportMember) exportWord([t.dataset.exportMember]); else toast('此任務沒有對應成員'); return; }
    // ----- catalog editing -----
    if(t.dataset.editProj!==undefined){ editingProj = editingProj===t.dataset.editProj ? '' : t.dataset.editProj; renderCatalog(); return; }
    if(t.dataset.saveProj!==undefined){ saveProjMeta(t.dataset.saveProj, t.closest('.proj-card')); return; }
    if(t.dataset.unmerge!==undefined){ unmergeProject(t.dataset.unmerge); return; }
    if(t.dataset.rmowner!==undefined){ const [tid,mid]=t.dataset.rmowner.split('|'); removeTaskOwner(tid,mid); return; }
    if(t.dataset.ocr!==undefined){ ocrTask(t.dataset.ocr); return; }
    if(t.dataset.opentask!==undefined){ openTask(t.dataset.opentask); return; }
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
    if(el.id==='filterRole'){ filters.role=el.value; renderMembersArea(); return; }
    if(el.dataset.setrole!==undefined){ setMemberRole(el.dataset.setrole, el.value); return; }
    if(el.dataset.setrole2!==undefined){ setMemberRole2(el.dataset.setrole2, el.value); return; }
    if(el.dataset.edit && el.dataset.tid){ editTaskField(el.dataset.tid, el.dataset.edit, el.value); return; }
    if(el.dataset.addowner){ addTaskOwner(el.dataset.addowner, el.value); el.value=''; return; }
  });
}
function openLight(src){ $('#lightboxImg').src=src; $('#lightbox').hidden=false; }

/* ---------- view tabs (less scrolling) ---------- */
let currentView=store.load('wrt_view','catalog');
function setView(v){
  currentView=v; store.save('wrt_view',v);
  const show=(sel,on)=>{ const e=document.querySelector(sel); if(e) e.style.display=on?'':'none'; };
  show('.catalog-panel', v==='catalog');
  show('.charts', v==='workload');
  show('.filterbar', v==='members');
  show('.members-area', v==='members');
  $$('#viewTabs button').forEach(b=>b.classList.toggle('active', b.dataset.view===v));
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
  sel.innerHTML='<option value="">全部成員 All members</option>'+members.map(m=>`<option value="${m.id}">${esc(m.name)}</option>`).join('');
  renderNarrative();
  $('#narrativeModal').hidden=false;
}
function renderNarrative(){
  const mid=$('#narrMember')?$('#narrMember').value:'';
  $('#narrativeText').textContent=buildNarrative(mid?[mid]:null);
}

/* expose for testing */
window.WRT={ get members(){return members;}, get tasks(){return tasks;}, get batches(){return batches;},
  parsePPTX, importFiles, addMembers, parseMemberText, overlayTasks, exportWord, buildPptx, assemblePptx, resolveOwners, matchOwner,
  autoAddFromReport, reresolveAllTasks, buildNarrative, projKeyOf, applyProjAliases, openProject, openNarrative,
  projectGroups, editTaskField, addTaskOwner, removeTaskOwner, ensureTesseract, ocrImages, ocrReportToTasks, ocrTask, setMemberRole, setMemberRole2,
  navStat, mergeProjects, resolveProjk, ocrAllReports, dedupeTasks, cleanupGarbledMembers, snapName, get pendingReports(){return pendingReports;}, convertOcrRows,
  get deletedNames(){return deletedNames;}, get projMerge(){return projMerge;}, get projMeta(){return projMeta;},
  setCatalogMember:(v)=>{catalogMember=v;renderCatalog();},
  get filters(){return filters;}, setFilter:(k,v)=>{filters[k]=v;renderMembersArea();},
  resetTasks:()=>{tasks=[];batches=[];persist();renderAll();},
  clearAll:()=>{members=[];tasks=[];batches=[];persist();renderAll();} };

/* init */
wireEvents();
renderAll();
setView(currentView);
