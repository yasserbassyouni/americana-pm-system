(()=>{
const D={companyName:'Smart PM Maintenance',factoryName:'',headerTitle:'Preventive Maintenance System',theme:{primary:'#c90000',header:'#c90000',background:'#f4f7fb',card:'#ffffff'},logos:[]};
function addStyle(){
  if(document.getElementById('pmBrandingStyle'))return;
  const s=document.createElement('style');s.id='pmBrandingStyle';s.textContent=`
  .pm-header-shell{width:100%;display:flex;flex-direction:column;gap:10px}
  .pm-header-main{display:grid;grid-template-columns:minmax(110px,1fr) minmax(280px,2fr) minmax(110px,1fr);align-items:center;gap:18px;width:100%}
  .pm-logo-slot{min-height:54px;display:flex;align-items:center}.pm-logo-slot.left{justify-content:flex-start}.pm-logo-slot.center{justify-content:center}.pm-logo-slot.right{justify-content:flex-end}
  .pm-header-logo{display:block;object-fit:contain;max-width:100%;max-height:70px}
  .pm-header-text{text-align:center;min-width:0;line-height:1.15}
  .pm-company{font-size:23px;font-weight:800}.pm-factory{font-size:14px;margin-top:3px}.pm-subtitle{font-size:13px;margin-top:3px;opacity:.96}.pm-page-title{font-size:24px;font-weight:800;margin-top:6px}
  .pm-header-nav{display:flex!important;justify-content:flex-end;align-items:center;gap:10px;flex-wrap:wrap;padding-top:9px;border-top:1px solid rgba(255,255,255,.24)}
  .pm-header-nav a,.pm-header-nav button{align-items:center!important;justify-content:center!important;min-height:38px!important;padding:8px 14px!important;border:1px solid rgba(255,255,255,.55)!important;border-radius:9px!important;background:#fff!important;color:#8f0000!important;text-decoration:none!important;font-size:14px!important;font-weight:800!important;line-height:1!important;box-shadow:0 1px 3px rgba(0,0,0,.12)!important;cursor:pointer!important}
  .pm-header-nav a:hover,.pm-header-nav button:hover{background:#f7f7f7!important;transform:translateY(-1px)}
  header.top.pm-clean-header,.top.pm-clean-header{padding:12px 20px!important;min-height:0!important}
  header.top .top-inner.pm-clean-host,header.top .bar.pm-clean-host,.top .top-inner.pm-clean-host,.top .bar.pm-clean-host{max-width:1220px!important;margin:0 auto!important;padding:0!important;display:block!important;width:100%!important}
  @media(max-width:760px){.pm-header-main{grid-template-columns:74px 1fr 74px;gap:8px}.pm-company{font-size:19px}.pm-page-title{font-size:20px}.pm-factory,.pm-subtitle{font-size:12px}.pm-header-nav{justify-content:center}.pm-header-logo{max-height:50px}}
  `;document.head.appendChild(s);
}
function getTitle(host){
  const h=host?.querySelector('h1,h2,h3');
  if(h&&h.textContent.trim())return h.textContent.trim();
  const t=(document.title||'').trim();
  return /^(settings|branding|backup|audit|users|email accounts)$/i.test(t)?t:'';
}
function getActions(host){
  const arr=[...host.querySelectorAll('a,button')];
  return arr.filter((e,i)=>arr.indexOf(e)===i && !e.closest('.pm-header-shell'));
}
function activeLogo(b,slot){
  const x=(b.logos||[])[slot-1];
  return x&&x.enabled&&x.logoUrl?x:null;
}
function logoNode(b,slot,cls){
  const box=document.createElement('div');box.className='pm-logo-slot '+cls;
  const l=activeLogo(b,slot);if(!l)return box;
  const img=document.createElement('img');img.className='pm-header-logo';img.src=l.logoUrl;img.alt='Logo';img.style.width=(l.width||92)+'px';img.style.height=(l.height||48)+'px';img.onerror=()=>img.style.display='none';box.appendChild(img);return box;
}
function buildShell(b,pageTitle,actions){
  const shell=document.createElement('div');shell.className='pm-header-shell';
  const main=document.createElement('div');main.className='pm-header-main';
  main.appendChild(logoNode(b,1,'left'));
  const center=document.createElement('div');center.className='pm-header-text';
  if(b.companyName){const e=document.createElement('div');e.className='pm-company';e.textContent=b.companyName;center.appendChild(e)}
  if(b.factoryName){const e=document.createElement('div');e.className='pm-factory';e.textContent=b.factoryName;center.appendChild(e)}
  if(b.headerTitle){const e=document.createElement('div');e.className='pm-subtitle';e.textContent=b.headerTitle;center.appendChild(e)}
  if(pageTitle && pageTitle!==b.companyName){const e=document.createElement('div');e.className='pm-page-title';e.textContent=pageTitle;center.appendChild(e)}
  main.appendChild(center);main.appendChild(logoNode(b,3,'right'));shell.appendChild(main);
  if(actions.length){const nav=document.createElement('div');nav.className='pm-header-nav';actions.forEach(e=>nav.appendChild(e));shell.appendChild(nav)}
  return shell;
}
function findHost(){
  // Only use a real page header. Never treat content toolbars that happen to use
  // the class "top" (Machine Management / Task Management) as the branding header.
  const dh=document.querySelector('header.top .top-inner');if(dh)return dh;
  const bar=document.querySelector('header.top .bar');if(bar)return bar;
  const top=document.querySelector('header.top');if(top)return top;
  const header=document.querySelector('body > header');if(header)return header;

  // Pages such as Machine Management and Task Management have no <header>.
  // Create a dedicated branding header instead of replacing their controls.
  const h=document.createElement('header');
  h.className='top pm-generated-header pm-clean-header';
  const inner=document.createElement('div');
  inner.className='bar pm-generated-host pm-clean-host';
  h.appendChild(inner);
  document.body.insertBefore(h,document.body.firstChild);
  return inner;
}
function applyTheme(b){const t=b.theme||D.theme;const r=document.documentElement;r.style.setProperty('--pm-primary',t.primary||'#c90000');r.style.setProperty('--pm-header',t.header||t.primary||'#c90000');r.style.setProperty('--pm-bg',t.background||'#f4f7fb');r.style.setProperty('--pm-card',t.card||'#ffffff');let x=document.getElementById('pmThemeOverrides');if(!x){x=document.createElement('style');x.id='pmThemeOverrides';document.head.appendChild(x)}x.textContent=`body{background:var(--pm-bg)!important}header.top,.top.pm-clean-header,.hero{background:var(--pm-header)!important}.nav a.active{color:var(--pm-primary)!important;border-color:var(--pm-primary)!important}.btn.red,button.red{background:var(--pm-primary)!important}.card,.panel,.kpi,.stat,.section,.brand-section,.logo-card{background:var(--pm-card)!important}`;}
function apply(b){
  applyTheme(b);const host=findHost();if(!host)return;
  const generated=host.classList.contains('pm-generated-host');
  const pageTitle=generated ? (document.title||'').trim() : getTitle(host);
  const actions=generated ? [] : getActions(host);
  host.replaceChildren(buildShell(b,pageTitle,actions));host.classList.add('pm-clean-host');
  const top=host.closest('header.top');if(top)top.classList.add('pm-clean-header');
}
function login(b){
  const d=document.getElementById('loginBrand');if(!d)return;
  d.replaceChildren();
  const l=activeLogo(b,2)||activeLogo(b,1)||activeLogo(b,3);
  if(l){const img=document.createElement('img');img.src=l.logoUrl;img.alt='Logo';img.onerror=()=>img.style.display='none';d.appendChild(img)}
  if(b.companyName){const t=document.createElement('div');t.className='pm-login-company';t.textContent=b.companyName;d.appendChild(t)}
  if(b.factoryName){const f=document.createElement('div');f.className='pm-login-factory';f.textContent=b.factoryName;d.appendChild(f)}
  if(b.headerTitle){const h=document.createElement('div');h.className='pm-login-subtitle';h.textContent=b.headerTitle;d.appendChild(h)}
}
async function load(){addStyle();try{const r=await fetch('/api/branding',{cache:'no-store'});const b={...D,...await r.json()};window.pmBranding=b;applyTheme(b);if(document.body.classList.contains('pm-login-page')){login(b)}else{apply(b)}return b}catch(e){console.warn('Branding:',e);if(document.body.classList.contains('pm-login-page'))login(D);return D}}
window.loadPmBranding=load;document.readyState==='loading'?document.addEventListener('DOMContentLoaded',load):load();
})();
