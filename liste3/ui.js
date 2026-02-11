const DB_KEY = 'QS_DATA_V9';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'flex'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

// --- PROJEKTE ---
function modalP(edit = false) {
    if (!edit) currentIdx = null; 
    const p = (edit && currentIdx !== null) ? db[currentIdx] : {num:'', name:'', lzf:'', sag:'', stt:'', stn:'', abs:'', grf:'', mat:''};
    el('p-idx').value = edit ? currentIdx : '';
    el('p-num').value = p.num || '';
    el('p-nam').value = p.name || '';
    el('p-lzf').value = p.lzf || '';
    el('p-sag').value = p.sag || '';
    el('p-stt').value = p.stt || '';
    el('p-stn').value = p.stn || '';
    el('p-abs').value = p.abs || '';
    el('p-grf').value = p.grf || '';
    el('p-mat').value = p.mat || '';
    show('m-p');
}

function saveP() {
    const idx = el('p-idx').value;
    const newP = {
        num: el('p-num').value.toUpperCase(),
        name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value, 
        sag: el('p-sag').value,
        stt: el('p-stt').value, 
        stn: el('p-stn').value,
        abs: el('p-abs').value, 
        grf: el('p-grf').value,
        mat: el('p-mat').value.toUpperCase(),
        tools: (idx !== '' && db[idx]) ? (db[idx].tools || []) : []
    };
    if (idx === '') { db.push(newP); currentIdx = db.length - 1; } 
    else { db[idx] = newP; }
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    hide('m-p');
    renderList();
    if(idx !== '') openProject(idx); else goHome();
}

function renderList() {
    const list = el('list-p'); if(!list) return;
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})">
            <div><small>${p.name || 'UNBENANNT'}</small><b>${p.num || '---'}</b></div>
            <div style="color:var(--danger); font-weight:900; padding:10px;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('');
}

function openProject(i) {
    currentIdx = i;
    el('v-home').classList.remove('active');
    el('v-det').classList.add('active');
    el('h-num').innerText = db[i].num || '---';
    el('h-nam').innerText = db[i].name || '---';
    renderTools();
}

function goHome() { currentIdx = null; el('v-home').classList.add('active'); el('v-det').classList.remove('active'); renderList(); }

// --- WERKZEUGE (Drag & Drop) ---
let dragSrcIdx = null;
function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = tools.map((t, i) => `
        <div class="list-item" draggable="true" ondragstart="handleDragStart(${i})" ondragover="event.preventDefault()" ondrop="handleDrop(${i})">
            <div class="handle">☰</div>
            <div style="flex:1" onclick="modalT(${i})">
                ${t.rev ? '<div style="background:#000; color:#fff; font-size:9px; padding:2px 6px; border-radius:4px; margin-bottom:4px; width:fit-content;">UNTEN START ↓</div>' : ''}
                <small style="font-size:11px; color:#8e8e93; font-weight:800;">${t.id || 'T00'}</small>
                <b style="font-size:20px; display:block;">${t.nm || '---'}</b>
            </div>
        </div>`).join('');
}

function handleDragStart(i) { dragSrcIdx = i; }
function handleDrop(i) {
    const tools = db[currentIdx].tools;
    const item = tools.splice(dragSrcIdx, 1)[0];
    tools.splice(i, 0, item);
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools();
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    const btn = el('btn-rev-toggle');
    if(t.rev) btn.classList.add('on'); else btn.classList.remove('on');
    btn.onclick = () => btn.classList.toggle('on');
    el('btn-del-t').style.display = edit ? 'block' : 'none';
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const t = { 
        id: el('t-id').value.toUpperCase(), 
        nm: el('t-nm').value.toUpperCase(), 
        dia: el('t-dia').value, 
        rev: el('btn-rev-toggle').classList.contains('on') 
    };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); hide('m-t');
}

// --- SYSTEM ---
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { const i = el('t-idx').value; db[currentIdx].tools.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }
function exportJSON() { el('imp-area').value = JSON.stringify(db); el('imp-area').select(); document.execCommand('copy'); alert("JSON Kopiert!"); }
function importJSON() { try { const parsed = JSON.parse(el('imp-area').value); if(Array.isArray(parsed)) { db = parsed; localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); } } catch(e) { alert("Fehler!"); } }

// --- PDF ENGINE (Premium Black Line Design) ---
function makePDF() {
    const p = db[currentIdx];
    const header = `
        <div style="display:flex; justify-content:space-between; align-items:flex-end; border-bottom:8px solid #000; padding-bottom:15px; margin-bottom:20px;">
            <div style="flex:1;">
                <div style="font-size:16px; font-weight:900; text-transform:uppercase;">${p.name || ''}</div>
                <div style="font-size:72px; font-weight:900; line-height:0.8; letter-spacing:-3px;">${p.num || '---'}</div>
            </div>
            <div style="width:240px; font-size:13px; font-weight:900; line-height:1.6;">
                <div style="display:flex; justify-content:space-between; border-bottom:2px solid #000;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:2px solid #000;"><span>MATERIAL</span><span>${p.mat || ''}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:2px solid #000;"><span>SÄGELÄNGE</span><span>${p.sag || ''}</span></div>
                <div style="display:flex; justify-content:space-between;"><span>STÜCKZAHL</span><span>${p.stt || ''} / ${p.stn || ''}</span></div>
            </div>
        </div>`;

    const getRow = (t) => `
        <div style="display:flex; align-items:center; border-bottom:2px solid #000; padding:15px 0;">
            <div style="width:100px; font-size:22px; font-weight:900;">${t.id}</div>
            <div style="flex:1; font-size:22px; font-weight:800; text-transform:uppercase;">${t.nm}</div>
            <div style="width:140px; text-align:right; font-size:22px; font-weight:900; white-space:pre-wrap;">${t.dia.split('/').join('\n')}</div>
        </div>`;

    let oben = [], unten = [], isUnten = false;
    (p.tools || []).forEach(t => { if(t.rev) isUnten = true; if(isUnten) unten.push(t); else oben.push(t); });

    let html = `<div style="padding:10mm; font-family:Helvetica, Arial, sans-serif; color:#000;">
        <div style="border:4px solid #000; padding:20px; min-height:270mm; position:relative; page-break-after:always;">
            ${header}
            <div style="font-size:26px; font-weight:900; border-bottom:4px solid #000; margin-bottom:10px;">REVOLVER OBEN</div>
            ${oben.map(getRow).join('')}
            <div style="position:absolute; bottom:20px; width:calc(100% - 40px); border-top:2px solid #000; text-align:center; font-size:10px; font-weight:900; padding-top:10px;">QS CENTRAL ELITE REPORT</div>
        </div>`;

    if(unten.length > 0) {
        html += `<div style="border:4px solid #000; padding:20px; min-height:270mm; position:relative; margin-top:20px;">
            ${header}
            <div style="font-size:26px; font-weight:900; border-bottom:4px solid #000; margin-bottom:10px;">REVOLVER UNTEN</div>
            ${unten.map(getRow).join('')}
            <div style="position:absolute; bottom:20px; width:calc(100% - 40px); border-top:2px solid #000; text-align:center; font-size:10px; font-weight:900; padding-top:10px;">QS CENTRAL ELITE REPORT</div>
        </div>`;
    }
    html += `</div>`;

    el('print-container').innerHTML = html;
    setTimeout(() => { window.print(); }, 300);
}

window.onload = renderList;
