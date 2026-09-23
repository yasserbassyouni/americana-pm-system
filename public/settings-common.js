
const token=localStorage.pmToken;
if(!token) location.href='login.html';
const me=JSON.parse(localStorage.pmUser||'{}');
if(me.role!=='Admin'){ alert('Admin access required'); location.href='index.html'; }
const H={'Authorization':'Bearer '+token,'Content-Type':'application/json'};
async function api(url,opt={}){opt.headers={...(opt.headers||{}),...H};const r=await fetch(url,opt);const j=await r.json();if(!r.ok)throw new Error(j.error||'Error');return j}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
