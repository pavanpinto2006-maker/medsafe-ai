async function api(url,options={}){options.headers={"Content-Type":"application/json",...(options.headers||{})};try{const r=await fetch(url,options);const d=await r.json();if(r.status===401){location="/";return{success:false}}return d}catch(e){console.error(e);return{success:false,message:"Could not connect to MedSafe server."}}}
function showToast(message,type="info"){const t=document.getElementById("toast");if(!t)return;t.textContent=message;t.className=`show ${type}`;clearTimeout(window._toast);window._toast=setTimeout(()=>t.className="",3000)}
function logout(){api("/api/logout",{method:"POST"}).then(()=>location="/")}
function toggleSidebar(){document.getElementById("sidebar")?.classList.toggle("open")}
function openModal(id){document.getElementById(id)?.classList.add("open")}
function closeModal(id){document.getElementById(id)?.classList.remove("open")}
function togglePassword(id){const x=document.getElementById(id);if(x)x.type=x.type==="password"?"text":"password"}
function escapeHtml(v){return String(v??"").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;")}
function escapeAttr(v){return escapeHtml(v)}
function formatDate(v){if(!v)return"—";const d=new Date(v);return isNaN(d)?"—":d.toLocaleDateString(undefined,{day:"2-digit",month:"short",year:"numeric"})}
document.addEventListener("click",e=>{if(e.target.classList.contains("modal"))e.target.classList.remove("open")})

// Premium autocomplete helpers: every searchable entity can surface results after 1 character.
function attachAutocomplete(input, panel, loader, renderItem, onPick){
  if(!input||!panel)return;
  let seq=0;
  async function run(){const q=input.value.trim();if(q.length<1){panel.innerHTML="";return}const token=++seq;const items=await loader(q);if(token!==seq)return;panel.innerHTML=(items||[]).slice(0,10).map((x,i)=>renderItem(x,i)).join("")||'<div class="search-empty">No matching records.</div>'}
  input.addEventListener("input",run); input.addEventListener("keydown",e=>{if(e.key==="Escape")panel.innerHTML=""});
  document.addEventListener("click",e=>{if(!input.contains(e.target)&&!panel.contains(e.target))panel.innerHTML=""});
  window["pick_"+input.id]=x=>{onPick(x);panel.innerHTML=""};
}
async function loadMedicineSuggestions(q){const r=await api("/api/medicines/search?q="+encodeURIComponent(q));return r.medicines||[]}
async function loadPrescriptionSuggestions(q){const r=await api("/api/prescriptions/search?q="+encodeURIComponent(q));return r.prescriptions||[]}
