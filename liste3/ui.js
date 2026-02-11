const DB_KEY = 'QS_DATA_V9';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'flex'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

// --- ПРОЕКТЫ ---
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

// --- ИНСТРУМЕНТЫ (Drag & Drop) ---
let dragSrcIdx = null;
function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = tools.map((t, i) => `
        <div class="list-item" draggable="true" ondragstart="handleDragStart(${i})" ondragover="event.preventDefault()" ondrop="handleDrop(${i})">
            <div class="handle">☰</div>
            <div style="flex:1" onclick="modalT(${i})">
                ${t.rev ? '<div style="background:#000; color:#fff; font-size:9px; padding:2px 6px; border-radius:4px; margin-bottom:4px; width:fit-content;">REVOLVER UNTEN ↓</div>' : ''}
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

// --- СИСТЕМА ---
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { const i = el('t-idx').value; db[currentIdx].tools.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }
function exportJSON() { el('imp-area').value = JSON.stringify(db); el('imp-area').select(); document.execCommand('copy'); alert("JSON Kopiert!"); }
function importJSON() { try { const parsed = JSON.parse(el('imp-area').value); if(Array.isArray(parsed)) { db = parsed; localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); } } catch(e) { alert("Fehler!"); } }

// --- НОВЫЙ PDF (ТОЧНО КАК В ОБРАЗЦЕ 230947.pdf) ---
function makePDF() {
    const p = db[currentIdx];
    
    // Стили для печати: жирные линии, фиксированные колонки
    const styles = `
        <style>
            @media print { body { background: #fff; } .no-print { display: none; } }
            .pdf-page { width: 210mm; font-family: Arial, sans-serif; color: #000; padding: 10mm; box-sizing: border-box; }
            .pdf-header { margin-bottom: 20px; }
            .title-name { font-size: 18px; font-weight: bold; text-transform: uppercase; margin-bottom: 5px; }
            .title-num { font-size: 40px; font-weight: 900; margin-bottom: 20px; }
            
            .info-grid { display: grid; grid-template-columns: 140px 1fr; border-top: 3px solid #000; margin-bottom: 30px; }
            .info-label { font-size: 14px; font-weight: bold; border-bottom: 1px solid #000; padding: 4px 0; }
            .info-value { font-size: 14px; border-bottom: 1px solid #000; padding: 4px 0; }
            
            .rev-title { font-size: 20px; font-weight: bold; margin: 20px 0 10px 0; text-transform: uppercase; }
            
            .tool-table { width: 100%; border-collapse: collapse; }
            .tool-table th { text-align: left; font-size: 12px; border-bottom: 3px solid #000; padding: 8px 0; }
            .tool-table td { font-size: 16px; font-weight: bold; border-bottom: 1px solid #000; padding: 10px 0; vertical-align: middle; }
            .col-id { width: 80px; }
            .col-dia { width: 150px; text-align: right; }
        </style>
    `;

    const getHeaderHTML = (pageTitle) => `
        <div class="pdf-header">
            <div class="title-name">${p.name || ''}</div>
            <div class="title-num">${p.num || '---'}</div>
            <div class="info-grid">
                <div class="info-label">LAUFZEIT</div><div class="info-value">${p.lzf || ''}</div>
                <div class="info-label">MATERIAL</div><div class="info-value">${p.mat || ''}</div>
                <div class="info-label">SÄGELÄNGE</div><div class="info-value">${p.sag || ''}</div>
                <div class="info-label">ABSTAND</div><div class="info-value">${p.abs || ''}</div>
                <div class="info-label">GREIFBACKEN</div><div class="info-value">${p.grf || ''}</div>
                <div class="info-label">STÜCKZAHL</div><div class="info-value">${p.stt || ''} / ${p.stn || ''}</div>
            </div>
            <div class="rev-title">${pageTitle}</div>
        </div>
    `;

    const getTableHTML = (tools) => `
        <table class="tool-table">
            <thead>
                <tr>
                    <th class="col-id">T-NR</th>
                    <th>WERKZEUGNAME / KOMMENTAR</th>
                    <th class="col-dia">Ø / TOLERANZ</th>
                </tr>
            </thead>
            <tbody>
                ${tools.map(t => `
                    <tr>
                        <td class="col-id">${t.id}</td>
                        <td>${t.nm}</td>
                        <td class="col-dia">${t.dia.split('/').join('<br>')}</td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
    `;

    let oben = [], unten = [], target = oben;
    (p.tools || []).forEach(t => { if(t.rev) target = unten; target.push(t); });

    let html = styles + '<div class="pdf-page">';
    
    // Page 1: OBEN
    html += getHeaderHTML('REVOLVER OBEN');
    html += getTableHTML(oben);
    
    // Page 2: UNTEN (если есть)
    if (unten.length > 0) {
        html += '<div style="page-break-before: always;"></div>';
        html += getHeaderHTML('REVOLVER UNTEN');
        html += getTableHTML(unten);
    }

    html += '</div>';

    el('print-container').innerHTML = html;
    setTimeout(() => { window.print(); }, 300);
}

window.onload = renderList;
