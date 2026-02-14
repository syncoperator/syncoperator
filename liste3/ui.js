const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png';

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f8f9fb; 
            --accent: #007aff; 
            --card-bg: #ffffff;
            --text-main: #1c1c1e;
            --text-sub: #8e8e93;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, sans-serif !important; 
            margin: 0; padding: 0; 
        }

        /* Брендинг */
        .brand-header { display: flex; flex-direction: column; align-items: center; padding: 25px 0 10px; }
        .logo-img { width: 110px; height: auto; margin-bottom: -8px; }
        .logo-title { font-size: 52px; font-weight: 900; letter-spacing: -3px; line-height: 1; color: #000; }
        .logo-mirror { 
            font-size: 52px; font-weight: 900; letter-spacing: -3px; margin-top: -20px; 
            transform: scaleY(-1); opacity: 0.1;
            background: linear-gradient(to bottom, #000, transparent);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }

        /* Кнопки в ряд */
        .nav-row { display: flex; justify-content: center; gap: 8px; padding: 15px; position: sticky; top: 0; background: rgba(248,249,251,0.9); backdrop-filter: blur(10px); z-index: 100; }
        .btn-premium {
            background: #fff; border: none; border-radius: 12px;
            padding: 10px 15px; font-size: 11px; font-weight: 800; color: #555;
            box-shadow: 0 4px 10px rgba(0,0,0,0.05); text-transform: uppercase;
        }
        .btn-premium.blue { background: var(--accent); color: #fff; }

        /* Карточки */
        .card-item {
            background: #fff; border-radius: 20px; margin: 0 16px 12px; padding: 20px;
            display: flex; align-items: center; justify-content: space-between;
            box-shadow: 0 5px 15px rgba(0,0,0,0.03);
        }
        .p-label { font-size: 10px; font-weight: 700; color: var(--text-sub); text-transform: uppercase; }
        .p-num { font-size: 24px; font-weight: 900; color: #000; }

        /* Модалки */
        .modal { position: fixed; inset: 0; background: rgba(0,0,0,0.5); display: none; align-items: center; justify-content: center; z-index: 2000; }
        .modal.active { display: flex; }
        
        .fab { position: fixed; bottom: 30px; right: 20px; background: var(--accent); color: #fff; width: 60px; height: 60px; border-radius: 30px; font-size: 30px; border: none; box-shadow: 0 8px 20px rgba(0,122,255,0.4); }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// --- ГЛАВНАЯ ---
function renderList() {
    el('v-home').style.display = 'block';
    el('v-det').style.display = 'none';
    
    el('v-home').innerHTML = `
        <div class="nav-row">
            <button class="btn-premium" onclick="exportJSON()">Export</button>
            <button class="btn-premium" onclick="el('m-imp').classList.add('active')">Import</button>
            <button class="btn-premium blue" onclick="modalP()">+ NEU</button>
        </div>
        <div class="brand-header">
            <img src="${LOGO_URL}" class="logo-img">
            <div class="logo-title">CitiTool</div>
            <div class="logo-mirror">CitiTool</div>
        </div>
        <div id="list-p-box"></div>
    `;

    el('list-p-box').innerHTML = db.map((p, i) => `
        <div class="card-item" onclick="openProject(${i})">
            <div>
                <div class="p-label">${p.name || 'UNBENANNT'}</div>
                <div class="p-num">${p.num || '---'}</div>
            </div>
            <button style="border:none; background:none; color:red; opacity:0.3;" onclick="event.stopPropagation(); deleteProject(${i})">✕</button>
        </div>
    `).join('') + '<div style="height:100px"></div>';
}

// --- ИНСТРУМЕНТЫ ---
function renderTools() {
    const p = db[currentIdx];
    el('v-home').style.display = 'none';
    el('v-det').style.display = 'block';

    el('v-det').innerHTML = `
        <div class="nav-row">
            <button class="btn-premium" onclick="goHome()">← Back</button>
            <button class="btn-premium blue" onclick="makePDF()">PDF Report</button>
        </div>
        <div class="brand-header">
            <div class="p-label">PROJEKT</div>
            <div class="logo-title">${p.num}</div>
            <div class="p-label" style="margin-top:5px">${p.name}</div>
        </div>
        <div id="list-t-box"></div>
        <button class="fab" onclick="modalT()">+</button>
    `;

    const box = el('list-t-box');
    const tools = p.tools || [];
    box.innerHTML = tools.map((t, i) => `
        <div class="card-item" onclick="modalT(${i})">
            <div style="flex:1">
                ${t.rev ? `<div style="background:#000; color:#fff; font-size:8px; padding:2px 5px; border-radius:4px; width:fit-content; margin-bottom:5px;">REVOLVER UNTEN</div>` : ''}
                <div class="p-label">${t.id}</div>
                <div style="font-size:18px; font-weight:800;">${t.nm}</div>
                <div style="color:var(--accent); font-weight:700;">${t.dia}</div>
            </div>
            <div style="display:flex; flex-direction:column; gap:5px;">
                <button class="btn-premium" style="padding:5px 10px;" onclick="event.stopPropagation(); moveItem(${i}, -1)">↑</button>
                <button class="btn-premium" style="padding:5px 10px;" onclick="event.stopPropagation(); moveItem(${i}, 1)">↓</button>
            </div>
        </div>
    `).join('') + '<div style="height:180px"></div>';
}

// --- ВСЕ ФУНКЦИИ (ОБЪЕДИНЕНО) ---
function openProject(i) { currentIdx = i; renderTools(); }
function goHome() { currentIdx = null; renderList(); }
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); save(); renderList(); } }

function moveItem(i, dir) {
    const tools = db[currentIdx].tools;
    const target = i + dir;
    if (target >= 0 && target < tools.length) {
        [tools[i], tools[target]] = [tools[target], tools[i]];
        save(); renderTools();
    }
}

function modalP(edit = false) {
    const p = edit ? db[currentIdx] : {num:'', name:'', lzf:'', sag:'', stt:'', stn:'', abs:'', grf:'', mat:''};
    el('p-idx').value = edit ? currentIdx : '';
    el('p-num').value = p.num; el('p-nam').value = p.name;
    el('p-lzf').value = p.lzf; el('p-sag').value = p.sag;
    el('p-stt').value = p.stt; el('p-stn').value = p.stn;
    el('p-abs').value = p.abs; el('p-grf').value = p.grf;
    if(el('p-mat')) el('p-mat').value = p.mat;
    el('m-p').classList.add('active');
}

function saveP() {
    const idx = el('p-idx').value;
    const data = {
        num: el('p-num').value, name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value, sag: el('p-sag').value,
        stt: el('p-stt').value, stn: el('p-stn').value,
        abs: el('p-abs').value, grf: el('p-grf').value,
        mat: el('p-mat') ? el('p-mat').value.toUpperCase() : '',
        tools: (idx !== '' && db[idx]) ? (db[idx].tools || []) : []
    };
    if (idx === '') db.push(data); else db[idx] = data;
    save(); el('m-p').classList.remove('active'); renderList();
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    const btn = el('btn-rev-toggle');
    if(btn) { t.rev ? btn.classList.add('on') : btn.classList.remove('on'); }
    el('m-t').classList.add('active');
}

function saveT() {
    const i = el('t-idx').value;
    const btn = el('btn-rev-toggle');
    const t = { id: el('t-id').value.toUpperCase(), nm: el('t-nm').value.toUpperCase(), dia: el('t-dia').value, rev: btn ? btn.classList.contains('on') : false };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    save(); renderTools(); el('m-t').classList.remove('active');
}

function exportJSON() { el('imp-area').value = JSON.stringify(db); el('imp-area').select(); document.execCommand('copy'); alert("JSON kopiert!"); }
function importJSON() { try { const parsed = JSON.parse(el('imp-area').value); if(Array.isArray(parsed)) { db = parsed; save(); renderList(); el('m-imp').classList.remove('active'); } } catch(e) { alert("Fehler"); } }

function makePDF() {
    const p = db[currentIdx];
    const getPageHead = () => `<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; min-height:90px;"><div style="display:flex; flex-direction:column; justify-content:center;"><div style="font-size:13px; font-weight:900; text-transform:uppercase; color:#666; margin-bottom:2px; line-height:1;">${p.name || ''}</div><div style="font-size:64px; font-weight:900; line-height:0.8; letter-spacing:-2px; margin:0;">${p.num || '---'}</div></div><div style="width:220px; font-size:11px; font-weight:800; line-height:1.5;"><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>MATERIAL</span><span>${p.mat || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>SÄGELÄNGE</span><span>${p.sag || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>ABSTAND</span><span>${p.abs || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>GREIFBACKEN</span><span>${p.grf || ''}</span></div><div style="display:flex; justify-content:space-between;"><span>STÜCKZAHL</span><span>${p.stt || ''} / ${p.stn || ''}</span></div></div></div><div style="border-bottom:5px solid #000; margin-bottom:15px;"></div>`;
    const tableHead = `<div style="display:flex; font-size:10px; font-weight:900; text-transform:uppercase; margin-bottom:6px; padding:0 2px;"><div style="width:75px;">T-NR</div><div style="flex:1;">WERKZEUGNAME / KOMMENTAR</div><div style="width:125px; text-align:right;">Ø / TOLERANZ</div></div><div style="border-bottom:4px solid #000;"></div>`;
    const getRow = (t) => `<div style="display:flex; align-items:baseline; border-bottom:1.5px solid #000; padding:10px 0; width:100%; page-break-inside: avoid;"><div style="width:75px; font-weight:800; font-size:15px;">${t.id}</div><div style="flex:1; font-weight:700; font-size:15px; text-transform:uppercase; padding-right:10px; white-space:pre-wrap;">${t.nm}</div><div style="width:125px; text-align:right; font-weight:800; font-size:14px;">${t.dia.replace(/\//g, '<br>')}</div></div>`;
    let oben = [], unten = [], target = oben;
    (p.tools || []).forEach(t => { if(t.rev) target = unten; target.push(t); });
    let pdfHtml = `<!DOCTYPE html><html><head><style>@page { size: A4; margin: 0; } body { margin: 0; padding: 10mm; background: #fff; font-family: sans-serif; -webkit-print-color-adjust: exact; } .page { width: 210mm; height: 297mm; padding: 15mm; box-sizing: border-box; page-break-after: always; display: flex; flex-direction: column; } .content-border { border: 2.2px solid #000; padding: 20px; flex: 1; display: flex; flex-direction: column; box-sizing: border-box; } .footer { border-top: 1.2px solid #000; padding-top: 5px; font-size: 9px; font-weight: 800; text-align: center; color: #666; margin-top: auto; }</style></head><body><div class="page"><div class="content-border">${getPageHead()}<div style="flex:1;"><div style="margin-bottom:5px; font-size:18px; font-weight:900; text-transform:uppercase;">REVOLVER OBEN</div>${tableHead}${oben.map(getRow).join('')}</div><div class="footer">CITITOOL REPORT</div></div></div>${unten.length > 0 ? `<div class="page"><div class="content-border">${getPageHead()}<div style="flex:1;"><div style="margin-bottom:5px; font-size:18px; font-weight:900; text-transform:uppercase;">REVOLVER UNTEN</div>${tableHead}${unten.map(getRow).join('')}</div><div class="footer">CITITOOL REPORT</div></div></div>` : ''}<script>window.onload = function() { setTimeout(() => { window.print(); }, 400); };</script></body></html>`;
    const win = window.open('', '_blank'); if (win) { win.document.write(pdfHtml); win.document.close(); }
}

window.onload = () => { injectStyles(); renderList(); };
