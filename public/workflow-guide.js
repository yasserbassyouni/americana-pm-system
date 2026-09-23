(function(){
'use strict';
const lang=()=>window.PMI18N?.getLanguage?.()||localStorage.getItem('pmLang')||'en';
const words={
 en:{manual:'Manual setup',import:'Excel import',line:'Create Line',machine:'Add Machines',task:'Add PM Tasks',review:'Review Annual Plan',years:'Build Future Years',monthly:'Monthly / Execute',auto:'Upload Excel & Auto Build',ready:'Current PM',next:'Next recommended step',guide:'Open Setup Guide',setup:'System Setup',planning:'PM Planning & Operation',continue:'Continue This Line',more:'Add More Machines / Tasks',another:'Create Another Line',complete:'System setup complete'},
 ar:{manual:'إنشاء يدوي',import:'استيراد Excel',line:'إنشاء الخط',machine:'إضافة الماكينات',task:'إضافة مهام PM',review:'مراجعة الخطة السنوية',years:'إنشاء السنوات القادمة',monthly:'الشهري / التنفيذ',auto:'رفع Excel والبناء التلقائي',ready:'PM الحالي',next:'الخطوة التالية المقترحة',guide:'فتح دليل الإعداد',setup:'إعداد النظام',planning:'تخطيط وتشغيل الصيانة الوقائية',continue:'استكمال هذا الخط',more:'إضافة ماكينات / مهام أخرى',another:'إنشاء خط آخر',complete:'اكتمل إعداد النظام'}
};
function w(){return words[lang()]||words.en}
function mode(){return localStorage.getItem('pmSetupMode')||'manual'}
function setMode(m){localStorage.setItem('pmSetupMode',m);render()}
function page(){return location.pathname.split('/').pop()||'index.html'}
function userRole(){try{return String(JSON.parse(localStorage.getItem('pmUser')||'{}').role||'').toLowerCase()}catch(_){return ''}}
function moveLanguageToHeader(){const b=document.querySelector('.pm-lang-control');if(!b)return;const t=document.querySelector('.nav')||document.querySelector('header .bar')||document.querySelector('header .top-inner')||document.querySelector('header');if(t&&b.parentNode!==t)t.appendChild(b)}
function api(url){const token=localStorage.pmToken||'';return fetch(url,{headers:{Authorization:'Bearer '+token}}).then(async r=>{const j=await r.json().catch(()=>({}));if(!r.ok)throw Error(j.error||'Request failed');return j})}
function contextFromPage(master){
 const q=new URLSearchParams(location.search);
 const domLine=document.getElementById('line')?.value||document.getElementById('addLine')?.value||'';
 const saved=localStorage.getItem('pmSetupLine')||'';
 const active=(master?.lines||[]).filter(x=>x&&x.enabled!==false&&x.active!==false).map(x=>String(x.name||''));
 let line=q.get('line')||domLine||saved||active[0]||'';
 if(line&&!active.includes(line)&&active.length)line=active[0];
 const domYear=Number(document.getElementById('year')?.value||document.getElementById('lineYear')?.value||0);
 const year=Number(q.get('year')||domYear||localStorage.getItem('pmSetupYear')||new Date().getFullYear());
 if(line){localStorage.setItem('pmSetupLine',line);localStorage.setItem('pmSetupYear',String(year));}
 return {line,year,active};
}
async function state(){
 const master=await api('/api/pm-master');
 const c=contextFromPage(master);
 const l=(master.lines||[]).find(x=>String(x.name||'')===c.line&&x.enabled!==false&&x.active!==false);
 const hasLine=!!l;
 const machines=(l?.machines||[]).filter(x=>x&&x.enabled!==false);
 let tasks=[]; try{if(hasLine)tasks=(await api(`/api/task-management?line=${encodeURIComponent(c.line)}&year=${c.year}`)).tasks||[]}catch(_){tasks=[]}
 let planner={lines:[]};try{planner=await api('/api/year-planner')}catch(_){planner={lines:[]}}
 const years=(planner.lines||[]).find(x=>String(x.name||'')===c.line)?.years||[];
 const activeTasks=tasks.filter(t=>t&&t.isActive!==false);
 const machineNos=machines.map(m=>String(m.machineNo||m.number||m.no||'').trim()).filter(Boolean);
 function sig(machineNo){
  const list=activeTasks.filter(t=>String(t.machineNo||t.machine||'')===String(machineNo));
  return JSON.stringify(list.map(t=>[t.id||t.taskId||t.pmCode||'',t.part||'',t.maintenance||'',t.frequency||'',t.intervalHours||'',t.plannedWeeks||''].join('|')).sort());
 }
 const reviewedCount=machineNos.filter(m=>localStorage.getItem(`pmAnnualReviewed:${c.line}:${c.year}:${m}`)===sig(m)).length;
 const reviewTotal=machineNos.length, reviewRemaining=Math.max(0,reviewTotal-reviewedCount);
 const reviewed=reviewTotal>0&&reviewedCount===reviewTotal;
 if(reviewed)localStorage.setItem(`pmFlowReviewed:${c.line}:${c.year}`,'1');else localStorage.removeItem(`pmFlowReviewed:${c.line}:${c.year}`);
 return {...c,hasLine,hasMachines:machines.length>0,hasTasks:activeTasks.length>0,reviewed,reviewedCount,reviewTotal,reviewRemaining,hasFutureYear:years.some(y=>Number(y)>c.year)};
}
function href(name,c){return `${name}?line=${encodeURIComponent(c.line||'')}&year=${encodeURIComponent(c.year||new Date().getFullYear())}`}
function statusSequence(s){
 const complete=[s.hasLine,s.hasMachines,s.hasTasks,s.reviewed,s.hasFutureYear,false];
 let firstIncomplete=complete.findIndex(v=>!v);if(firstIncomplete<0)firstIncomplete=complete.length-1;
 return complete.map((v,i)=>v?'completed':(i===firstIncomplete?'current':'pending'));
}
function stepHtml(label,url,status,num){return `<a class="pm-flow-step ${status}" href="${url}"><span class="n">${status==='completed'?'✓':num}</span><span class="txt">${label}</span></a>`}
function paintManual(box,s){
 const x=w();
 const st=statusSequence(s);
 const links=[
  href('pm-master.html',s),
  href('machine-management.html',s),
  href('task-management.html',s),
  href('annual-plans.html',s),
  href('year-planner.html',s),
  href('monthly-task-management.html',s)
 ];
 const reviewLabel=s.reviewTotal>0?`${x.review} — ${s.reviewedCount}/${s.reviewTotal} Reviewed, ${s.reviewRemaining} Remaining`:x.review;
 const labels=[x.line,x.machine,x.task,reviewLabel,x.years,x.monthly];
 const setup=box.querySelector('[data-phase="setup"] .pm-flow-steps');
 const plan=box.querySelector('[data-phase="planning"] .pm-flow-steps');
 if(setup)setup.innerHTML=labels.slice(0,3).map((l,i)=>stepHtml(l,links[i],st[i],i+1)).join('');
 if(plan)plan.innerHTML=labels.slice(3).map((l,j)=>stepHtml(l,links[j+3],st[j+3],j+1)).join('');
 const ni=st.indexOf('current');
 const next=box.querySelector('.pm-flow-next');
 if(next&&ni>=0)next.innerHTML=`${x.next}: <a href="${links[ni]}">${labels[ni]}</a>`;
 const decision=box.querySelector('.pm-flow-decision');
 if(decision){
  decision.style.display=s.hasTasks?'flex':'none';
  if(s.hasTasks){
   decision.innerHTML=`<span class="pm-flow-done">✓ ${x.complete}</span><button data-continue>${x.continue}</button><button data-more>${x.more}</button><button data-another>${x.another}</button>`;
   decision.querySelector('[data-continue]').onclick=()=>location.href=href('annual-plans.html',s);
   decision.querySelector('[data-more]').onclick=()=>location.href=href('machine-management.html',s);
   decision.querySelector('[data-another]').onclick=()=>{localStorage.removeItem('pmSetupLine');localStorage.removeItem('pmSetupYear');location.href='pm-master.html?newLine=1'};
  }
 }
}
function emptyManualState(){
 return {line:'',year:new Date().getFullYear(),hasLine:false,hasMachines:false,hasTasks:false,reviewed:false,reviewedCount:0,reviewTotal:0,reviewRemaining:0,hasFutureYear:false};
}
async function refreshStatus(box){
 if(!box||mode()!=='manual')return;
 // Always paint a safe default first so the workflow never disappears.
 paintManual(box,emptyManualState());
 try{
  const s=await state();
  paintManual(box,s);
 }catch(e){
  console.warn('Workflow status:',e.message);
  // Keep the safe default visible. Never leave empty workflow boxes.
 }
}
function render(){
 if(page()==='login.html')return;
 document.querySelector('.pm-flow')?.remove();
 if(userRole()!=='admin'){moveLanguageToHeader();return}
 const x=w(),m=mode(),box=document.createElement('div');box.className='pm-flow';
 if(m==='import'){
   const p=page(),steps=[['pm-master.html',x.line],['excel-migration.html',x.auto],['index.html',x.ready]],ci=Math.max(0,steps.findIndex(s=>s[0]===p));
   box.innerHTML=`<div class="pm-flow-card"><div class="pm-flow-head"><div class="pm-flow-title">${x.setup}</div><div class="pm-flow-mode"><button data-m="manual">${x.manual}</button><button data-m="import" class="active">${x.import}</button><button data-guide>${x.guide}</button></div></div><div class="pm-flow-steps">${steps.map((s,i)=>stepHtml(s[1],s[0],i<ci?'completed':i===ci?'current':'pending',i+1)).join('')}</div></div>`;
 }else{
   box.innerHTML=`<div class="pm-flow-card"><div class="pm-flow-head"><div class="pm-flow-title">${x.setup}</div><div class="pm-flow-mode"><button data-m="manual" class="active">${x.manual}</button><button data-m="import">${x.import}</button><button data-guide>${x.guide}</button></div></div><div class="pm-flow-phases"><div class="pm-flow-phase" data-phase="setup"><div class="pm-flow-phase-title">${x.setup}</div><div class="pm-flow-steps"></div></div><div class="pm-flow-phase" data-phase="planning"><div class="pm-flow-phase-title">${x.planning}</div><div class="pm-flow-steps"></div></div></div><div class="pm-flow-next">${x.next}: ...</div><div class="pm-flow-decision" style="display:none"></div></div>`;
 }
 const anchor=document.querySelector('header')||document.querySelector('main')||document.body.firstElementChild;if(anchor?.parentNode)anchor.parentNode.insertBefore(box,anchor.nextSibling);else document.body.prepend(box);
 box.querySelectorAll('[data-m]').forEach(b=>b.onclick=()=>setMode(b.dataset.m));box.querySelector('[data-guide]').onclick=()=>location.href='setup-guide.html';const lb=document.querySelector('.pm-lang-control');if(lb)box.querySelector('.pm-flow-mode')?.appendChild(lb);
 if(m==='manual'){
   refreshStatus(box);
   setTimeout(()=>refreshStatus(box),500);
   ['line','year','lineYear','addLine'].forEach(id=>document.getElementById(id)?.addEventListener('change',()=>refreshStatus(box)));
   window.addEventListener('pm-workflow-refresh',()=>refreshStatus(box));
 }
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',render);else render();
window.PMSetupGuide={render,setMode,refresh:()=>{const b=document.querySelector('.pm-flow');if(b)refreshStatus(b)}};
})();
