const DB_KEY = 'QS_DATA_V10';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'flex'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

// --- ПРОЕКТЫ ---
function saveP() {
    const idx = el('p-idx').value;
    const newP = {
        num: el('p-num').value.toUpperCase(),
        name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value, sag: el('p-sag').value,
        stt: el('p-stt').value, stn: el('p-stn').value,
        abs: el('p-abs').value, grf: el('p-grf').value,
        mat: el('p-mat').value.toUpperCase(),
        tools: (idx !== '' && db[idx]) ? (db[idx].tools || []) : []
    };
    if (idx === '') db.push(newP); else db[idx] = newP;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    hide('m-p'); renderList();
    if(idx !== '') openProject(idx); else goHome();
}

function renderList() {
    const list = el('list-p'); if(!list) return;
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})">
            <div><small>${p.name || '---'}</small><b>${p.num || '---'}</b></div>
            <div style="color:red; font-weight:900; padding:10px;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:100px"></div>';
}

function openProject(i) {
    currentIdx = i;
    el('v-home').classList.remove('active');
    el('v-det').classList.add('active');
    el('h-num').innerText = db[i].num;
    el('h-nam').innerText = db[i].name;
    renderTools();
}

function goHome() { currentIdx = null; el('v-home').classList.add('active'); el('v-det').classList.remove('active'); renderList(); }

// --- ИНСТРУМЕНТЫ ---
function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = tools.map((t, i) => `
        <div class="list-item" draggable="true" ondragstart="dragSrcIdx=${i}" ondragover="event.preventDefault()" ondrop="handleDrop(${i})">
            <div style="color:#ccc; margin-right:15px; font-size:20px;">☰</div>
            <div class="tool-info" onclick="modalT(${i})" style="flex:1;">
                ${t.rev ? '<small style="color:#000; background:#eee; padding:2px 4px; border-radius:3px; display:inline-block; margin-bottom:5px;">REVOLVER UNTEN ↓</small>' : ''}
                <small>T-${t.id || '00'}</small>
                <b>${t.nm || '---'}</b>
            </div>
        </div>`).join('') + '<div style="height:150px"></div>';
}

let dragSrcIdx = null;
function handleDrop(i) {
    const tools = db[currentIdx].tools;
    const item = tools.splice(dragSrcIdx, 1)[0];
    tools.splice(i, 0, item);
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools();
}

// --- PDF ENGINE (ОТКРЫВАЕТСЯ В НОВОМ ОКНЕ ДЛЯ ПЕЧАТИ) ---
function makePDF() {
    const p = db[currentIdx];
    const win = window.open('', '_blank');
    
    let oben = [], unten = [], isU = false;
    (p.tools || []).forEach(t => { if(t.rev) isU = true; if(isU) unten.push(t); else oben.push(t); });

    const content = `
    <html>
    <head>
        <style>
            body { font-family: sans-serif; padding: 15mm; color: #000; }
            .header { display: flex; justify-content: space-between; border-bottom: 6px solid #000; padding-bottom: 10px; margin-bottom: 20px; }
            .header h1 { font-size: 55px; margin: 0; font-weight: 900; letter-spacing: -2px; }
            .info-grid { width: 250px; font-size: 12px; font-weight: bold; }
            .info-row { display: flex; justify-content: space-between; border-bottom: 1px solid #000; padding: 3px 0; }
            .section-title { font-size: 22px; font-weight: 900; border-bottom: 4px solid #000; margin: 25px 0 10px 0; padding-bottom: 5px; }
            .t-table { width: 100%; border-collapse: collapse; }
            .t-table th { text-align: left; font-size: 11px; border-bottom: 2px solid #000; padding: 5px; }
            .t-table td { font-size: 17px; font-weight: bold; border-bottom: 1px solid #ddd; padding: 12px 5px; }
            @media print { .page-break { page-break-before: always; } }
        </style>
    </head>
    <body>
        <div class="header">
            <div><div style="font-size:16px; font-weight:bold;">${p.name}</div><h1>${p.num}</h1></div>
            <div class="info-grid">
                <div class="info-row"><span>MATERIAL</span><span>${p.mat}</span></div>
                <div class="info-row"><span>SÄGELÄNGE</span><span>${p.sag}</span></div>
                <div class="info-row"><span>ABSTAND</span><span>${p.abs}</span></div>
                <div class="info-row"><span>GREIFBACKEN</span><span>${p.grf}</span></div>
                <div class="info-row"><span>STÜCKZAHL</span><span>${p.stt} / ${p.stn}</span></div>
            </div>
        </div>
        <div class="section-title">REVOLVER OBEN</div>
        <table class="t-table">
            <thead><tr><th>T-NR</th><th>WERKZEUGNAME / KOMMENTAR</th><th style="text-align:right">Ø / TOLERANZ</th></tr></thead>
            <tbody>${oben.map(t => `<tr><td style="width:80px">${t.id}</td><td>${t.nm}</td><td style="text-align:right">${t.dia}</td></tr>`).join('')}</tbody>
        </table>
        ${unten.length > 0 ? `
        <div class="page-break"></div>
        <div class="section-title" style="margin-top:0">REVOLVER UNTEN</div>
        <table class="t-table">
            <thead><tr><th>T-NR</th><th>WERKZEUGNAME</th><th style="text-align:right">Ø / TOLERANZ</th></tr></thead>
            <tbody>${unten.map(t => `<tr><td style="width:80px">${t.id}</td><td>${t.nm}</td><td style="text-align:right">${t.dia}</td></tr>`).join('')}</tbody>
        </table>` : ''}
    </body>
    </html>`;

    win.document.write(content);
    win.document.close();
    win.focus();
    setTimeout(() => { win.print(); win.close(); }, 500);
}

// --- ВСПОМОГАТЕЛЬНЫЕ МОДАЛКИ ---
function modalP(edit=false) {
    if(!edit) currentIdx=null; 
    const p=edit?db[currentIdx]:{num:'',name:'',lzf:'',sag:'',stt:'',stn:'',abs:'',grf:'',mat:''};
    el('p-idx').value=edit?currentIdx:''; el('p-num').value=p.num; el('p-nam').value=p.name;
    el('p-lzf').value=p.lzf; el('p-sag').value=p.sag; el('p-stt').value=p.stt; el('p-stn').value=p.stn;
    el('p-abs').value=p.abs; el('p-grf').value=p.grf; el('p-mat').value=p.mat; show('m-p');
}

function modalT(i=null) {
    const edit=i!==null; el('t-idx').value=edit?i:''; 
    const t=edit?db[currentIdx].tools[i]:{id:'',nm:'',dia:'',rev:false};
    el('t-id').value=t.id; el('t-nm').value=t.nm; el('t-dia').value=t.dia;
    const btn=el('btn-rev-toggle'); btn.innerText=t.rev?'ON':'OFF';
    if(t.rev) btn.classList.add('on'); else btn.classList.remove('on');
    btn.onclick=()=>{btn.classList.toggle('on'); btn.innerText=btn.classList.contains('on')?'ON':'OFF';};
    el('btn-del-t').style.display=edit?'block':'none'; show('m-t');
}

function saveT() {
    const i=el('t-idx').value; 
    const t={id:el('t-id').value.toUpperCase(),nm:el('t-nm').value.toUpperCase(),dia:el('t-dia').value,rev:el('btn-rev-toggle').classList.contains('on')};
    if(!db[currentIdx].tools) db[currentIdx].tools=[]; 
    if(i==='') db[currentIdx].tools.push(t); else db[currentIdx].tools[i]=t;
    localStorage.setItem(DB_KEY,JSON.stringify(db)); renderTools(); hide('m-t');
}

function deleteProject(i) { if(confirm('Löschen?')){db.splice(i,1);localStorage.setItem(DB_KEY,JSON.stringify(db));renderList();}}
function delT() { const i=el('t-idx').value; db[currentIdx].tools.splice(i,1); localStorage.setItem(DB_KEY,JSON.stringify(db)); renderTools(); hide('m-t'); }
function exportJSON() { el('imp-area').value=JSON.stringify(db); alert('Daten kopiert!'); }
function importJSON() { db=JSON.parse(el('imp-area').value); localStorage.setItem(DB_KEY,JSON.stringify(db)); renderList(); hide('m-imp'); }

window.onload = renderList;
