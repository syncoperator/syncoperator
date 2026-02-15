/* CITITOOL PREMIUM V12 STABLE FIXED */

const DB_KEY = 'QS_DATA_V8';
let db;
let currentIdx = null;

// ---------- SAFE LOAD ----------
function loadDB(){
    try{
        const raw = JSON.parse(localStorage.getItem(DB_KEY));
        db = Array.isArray(raw) ? raw : [];
    }catch{
        db = [];
    }
}
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// ---------- ESCAPE HTML ----------
const esc = s => String(s ?? '')
    .replaceAll('&','&amp;')
    .replaceAll('<','&lt;')
    .replaceAll('>','&gt;')
    .replaceAll('"','&quot;');

// ---------- STYLES ----------
function injectStyles(){
const style=document.createElement('style');
style.textContent=`
:root{--bg:#f2f2f7;--accent:#007aff;--card:#fff;--text:#1c1c1e;--sub:#8e8e93}
body{background:var(--bg);font-family:-apple-system,sans-serif;margin:0;color:var(--text);padding-bottom:120px}
header{background:rgba(255,255,255,.8);backdrop-filter:blur(20px);padding:15px 20px;position:sticky;top:0;z-index:1000;display:flex;justify-content:space-between;align-items:center;border-bottom:.5px solid rgba(0,0,0,.1)}
.h-title{font-weight:800;font-size:22px}
.project-card{background:var(--card);border-radius:20px;margin:12px 16px;padding:20px;display:flex;justify-content:space-between}
.p-num-big{font-size:28px;font-weight:900}
.p-name-sub{font-size:11px;font-weight:700;color:var(--sub)}
.tool-card{background:var(--card);border-radius:16px;margin:8px 16px;padding:16px;display:flex;justify-content:space-between}
.t-name-main{font-size:18px;font-weight:800}
.t-dia-val{color:var(--accent);font-weight:700}
.rev-badge{background:#000;color:#fff;font-size:8px;font-weight:900;padding:2px 6px;border-radius:4px}
.btn-icon{background:#f2f2f7;border:none;border-radius:10px;width:36px;height:36px;color:var(--accent);font-weight:900}
.fab{position:fixed;bottom:30px;right:20px;background:var(--accent);color:#fff;width:60px;height:60px;border-radius:30px;display:flex;align-items:center;justify-content:center;font-size:30px;z-index:2000}
.modal{display:none;position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:3000;align-items:flex-end}
.modal-content{background:#fff;width:100%;border-radius:25px 25px 0 0;padding:25px;max-height:90vh;overflow:auto}
input,textarea{width:100%;padding:14px;margin:8px 0;border:1px solid #eee;border-radius:12px;background:#f9f9fb;font-size:16px;box-sizing:border-box}
.btn-save{background:var(--accent);color:#fff;width:100%;padding:16px;border:none;border-radius:14px;font-weight:700;margin-top:10px}
`;
document.head.appendChild(style);
}

// ---------- APP ----------
function setupApp(){
document.body.innerHTML=`
<header>
<div class="h-title">CitiTool</div>
<button class="btn-icon" onclick="openImport()">JSON</button>
</header>

<div id="home">
<div id="plist"></div>
<div class="fab" onclick="modalP()">+</div>
</div>

<div id="detail" style="display:none">
<div style="padding:20px">
<div id="detName" class="p-name-sub"></div>
<div id="detNum" style="font-size:40px;font-weight:900"></div>
<div style="display:flex;gap:10px;margin-top:15px">
<button class="btn-save" onclick="goHome()">Назад</button>
<button class="btn-save" style="background:#000" onclick="makePDF()">PDF</button>
</div>
</div>
<div id="tlist"></div>
<div class="fab" onclick="modalT()">+</div>
</div>

${modalProjectHTML()}
${modalToolHTML()}
${modalImportHTML()}
`;
}

// ---------- MODALS HTML ----------
function modalProjectHTML(){
return `
<div id="mp" class="modal" onclick="if(event.target==this)this.style.display='none'">
<div class="modal-content">
<input type="hidden" id="pidx">
<input id="pnum" placeholder="NUMMER">
<input id="pname" placeholder="NAME">
<input id="plzf" placeholder="Laufzeit">
<input id="pmat" placeholder="Material">
<input id="psag" placeholder="Sägelänge">
<input id="pabs" placeholder="Abstand">
<input id="pgrf" placeholder="Greifer">
<input id="pstt" placeholder="Soll">
<input id="pstn" placeholder="Ist">
<button class="btn-save" onclick="saveP()">SPEICHERN</button>
</div></div>`;
}

function modalToolHTML(){
return `
<div id="mt" class="modal" onclick="if(event.target==this)this.style.display='none'">
<div class="modal-content">
<input type="hidden" id="tidx">
<input id="tid" placeholder="T-NR">
<input id="tnm" placeholder="NAME">
<input id="tdia" placeholder="Ø">
<button id="revbtn" class="btn-save" style="background:#eee;color:#000" onclick="this.classList.toggle('active');this.textContent=this.classList.contains('active')?'REVOLVER UNTEN':'REVOLVER OBEN'">REVOLVER OBEN</button>
<button class="btn-save" onclick="saveT()">SPEICHERN</button>
<button id="delT" class="btn-save" style="background:none;color:red" onclick="deleteTool()">Löschen</button>
</div></div>`;
}

function modalImportHTML(){
return `
<div id="mi" class="modal" onclick="if(event.target==this)this.style.display='none'">
<div class="modal-content">
<textarea id="imp"></textarea>
<button class="btn-save" onclick="importJSON()">IMPORT</button>
</div></div>`;
}

// ---------- PROJECT LIST ----------
function renderProjects(){
const box=document.getElementById('plist');
box.innerHTML=db.map((p,i)=>`
<div class="project-card" onclick="openProject(${i})">
<div>
<div class="p-name-sub">${esc(p.name||'---')}</div>
<div class="p-num-big">${esc(p.num||'---')}</div>
</div>
<div onclick="event.stopPropagation();delProject(${i})">✕</div>
</div>`).join('');
}

function openProject(i){
if(!db[i]) return;
currentIdx=i;
home.style.display='none';
detail.style.display='block';
detName.textContent=db[i].name||'';
detNum.textContent=db[i].num||'';
renderTools();
}

function goHome(){
currentIdx=null;
home.style.display='block';
detail.style.display='none';
renderProjects();
}

// ---------- PROJECT SAVE ----------
function modalP(i=null){
const edit=i!=null;
pidx.value=edit?i:'';
const p=edit?db[i]:{};
pnum.value=p.num||'';
pname.value=p.name||'';
plzf.value=p.lzf||'';
pmat.value=p.mat||'';
psag.value=p.sag||'';
pabs.value=p.abs||'';
pgrf.value=p.grf||'';
pstt.value=p.stt||'';
pstn.value=p.stn||'';
mp.style.display='flex';
}

function saveP(){
const idx=pidx.value;
const data={
num:pnum.value,
name:pname.value.toUpperCase(),
lzf:plzf.value,
mat:pmat.value,
sag:psag.value,
abs:pabs.value,
grf:pgrf.value,
stt:pstt.value,
stn:pstn.value,
tools: idx===''?[]:(db[idx].tools||[])
};
if(idx==='') db.push(data);
else db[idx]=data;
save();
mp.style.display='none';
renderProjects();
}

// ---------- TOOLS ----------
function renderTools(){
if(currentIdx==null||!db[currentIdx]) return;
const list=db[currentIdx].tools||[];
tlist.innerHTML=list.map((t,i)=>`
<div class="tool-card" onclick="modalT(${i})">
<div>
${t.rev?'<div class="rev-badge">REVOLVER UNTEN</div>':''}
<div>${esc(t.id)}</div>
<div class="t-name-main">${esc(t.nm)}</div>
<div class="t-dia-val">${esc(t.dia)}</div>
</div>
<div>
<button class="btn-icon" onclick="event.stopPropagation();moveTool(${i},-1)">↑</button>
<button class="btn-icon" onclick="event.stopPropagation();moveTool(${i},1)">↓</button>
</div>
</div>`).join('');
}

function modalT(i=null){
const edit=i!=null;
const t=edit?db[currentIdx].tools[i]:{};
tidx.value=edit?i:'';
tid.value=t.id||'';
tnm.value=t.nm||'';
tdia.value=t.dia||'';
revbtn.classList.toggle('active',t.rev);
revbtn.textContent=t.rev?'REVOLVER UNTEN':'REVOLVER OBEN';
delT.style.display=edit?'block':'none';
mt.style.display='flex';
}

function saveT(){
if(currentIdx==null) return;
const i=tidx.value;
const t={
id:tid.value.toUpperCase(),
nm:tnm.value.toUpperCase(),
dia:tdia.value,
rev:revbtn.classList.contains('active')
};
if(i==='') db[currentIdx].tools.push(t);
else db[currentIdx].tools[Number(i)]=t;
save();
mt.style.display='none';
renderTools();
}

function deleteTool(){
const i=Number(tidx.value);
if(!confirm('Удалить инструмент?')) return;
db[currentIdx].tools.splice(i,1);
save();
mt.style.display='none';
renderTools();
}

function moveTool(i,d){
const arr=db[currentIdx].tools;
const n=i+d;
if(n<0||n>=arr.length) return;
[arr[i],arr[n]]=[arr[n],arr[i]];
save();
renderTools();
}

// ---------- DELETE PROJECT ----------
function delProject(i){
if(!confirm('Удалить проект?')) return;
db.splice(i,1);
save();
renderProjects();
}

// ---------- IMPORT ----------
function openImport(){
imp.value=JSON.stringify(db,null,2);
mi.style.display='flex';
}

function importJSON(){
try{
const parsed=JSON.parse(imp.value);
if(!Array.isArray(parsed)) throw 1;
db=parsed;
save();
location.reload();
}catch{
alert('Ошибка JSON');
}
}

// ---------- PDF ----------
function makePDF(){
if(currentIdx==null) return;
const p=db[currentIdx];
const rows=(list)=>list.map(t=>`
<tr>
<td>${esc(t.id)}</td>
<td>${esc(t.nm)}</td>
<td>${esc(t.dia)}</td>
</tr>`).join('');

const oben=(p.tools||[]).filter(x=>!x.rev);
const unten=(p.tools||[]).filter(x=>x.rev);

const html=`
<html><body style="font-family:sans-serif">
<h2>${esc(p.name)}</h2>
<h1>${esc(p.num)}</h1>
<table border=1 width=100%>${rows(oben)}</table>
<br>
<table border=1 width=100%>${rows(unten)}</table>
<script>print()<\/script>
</body></html>`;

const w=window.open('');
if(!w){alert('Popup blocked');return;}
w.document.write(html);
w.document.close();
}

// ---------- INIT ----------
window.onload=()=>{
loadDB();
injectStyles();
setupApp();
renderProjects();
};