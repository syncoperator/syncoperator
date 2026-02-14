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
            --text-main: #1c1c1e;
            --text-sub: #8e8e93; 
            --card-bg: #ffffff;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important; 
            margin: 0; color: var(--text-main);
            padding-bottom: 100px; /* Чтобы кнопки не перекрывали контент */
        }
        
        /* НОВАЯ ВЕРХНЯЯ ПАНЕЛЬ */
        .premium-nav {
            position: sticky; top: 0; z-index: 100;
            background: rgba(248, 249, 251, 0.85); backdrop-filter: blur(15px);
            display: flex; justify-content: flex-end; padding: 15px 20px; gap: 10px;
            border-bottom: 0.5px solid rgba(0,0,0,0.05);
        }
        .btn-top {
            background: #fff; border: none; border-radius: 10px;
            padding: 8px 14px; font-size: 11px; font-weight: 800; color: #666;
            box-shadow: 4px 4px 8px #cfd8e3, -4px -4px 8px #fff;
            text-transform: uppercase; cursor: pointer;
        }
        .btn-top.blue { background: var(--accent); color: white; }

        /* БЛОК С ЛОГОТИПОМ */
        .hero { display: flex; flex-direction: column; align-items: center; padding: 20px 0 30px; }
        .logo-huge { width: 200px; height: 200px; object-fit: contain; }
        
        .mirror-wrap { text-align: center; margin-top: 10px; }
        .t-top { font-size: 72px; font-weight: 900; letter-spacing: -4px; color: #1d1d1f; line-height: 0.8; }
        .t-reflect { 
            font-size: 72px; font-weight: 900; letter-spacing: -4px; margin-top: -32px; 
            transform: scaleY(-1); opacity: 0.15;
            background: linear-gradient(to bottom, #000, transparent);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }

        /* Твои оригинальные стили карточек */
        .project-card {
            background: var(--card-bg); border-radius: 24px;
            margin: 0 20px 16px 20px; padding: 20px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.04);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.5);
            transition: all 0.2s ease;
        }
        .p-info { display: flex; flex-direction: column; }
        .p-label { font-size: 11px; font-weight: 700; color: var(--text-sub); text-transform: uppercase; margin-bottom: 4px; }
        .p-title { font-size: 24px; font-weight: 900; color: #000; letter-spacing: -0.5px; }

        .tool-card {
            background: #fff; border-radius: 18px;
            padding: 14px 18px; margin: 0 16px 10px 16px;
            display: flex; align-items: center; justify-content: space-between;
            box-shadow: 0 2px 8px rgba(0,0,0,0.02);
        }
        .btn-delete { color: #ff3b30; font-weight: 800; padding: 10px; font-size: 18px; }
        .fab {
            position: fixed; bottom: 30px; right: 20px;
            background: var(--accent); color: #fff; width: 60px; height: 60px;
            border-radius: 30px; display: flex; align-items: center; justify-content: center;
            font-size: 30px; box-shadow: 0 8px 25px rgba(0,122,255,0.3); z-index: 1000;
        }
        .order-controls { display: flex; flex-direction: column; gap: 4px; margin-left: 12px; }
        .btn-order { background: #f2f2f7; border: none; border-radius: 6px; width: 32px; height: 28px; font-weight: 900; color: var(--accent); }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const show = (id) => { const x = el(id); if(x) x.style.display = 'flex'; };
const hide = (id) => { const x = el(id); if(x) x.style.display = 'none'; };
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// --- ГЛАВНАЯ СТРАНИЦА (С ЛОГОТИПОМ) ---
function renderList() {
    const list = el('list-p'); if(!list) return;
    
    // Вставляем шапку и логотип прямо в контейнер перед списком
    list.innerHTML = `
        <div class="premium-nav">
            <button class="btn-top" onclick="exportJSON()">Export</button>
            <button class="btn-top" onclick="show('m-imp')">Import</button>
            <button class="btn-top blue" onclick="modalP()">+ NEU</button>
        </div>
        <div class="hero">
            <img src="${LOGO_URL}" class="logo-huge">
            <div class="mirror-wrap">
                <div class="t-top">CitiTool</div>
                <div class="t-reflect">CitiTool</div>
            </div>
        </div>
        <div id="project-items"></div>
    `;

    el('project-items').innerHTML = db.map((p, i) => `
        <div class="project-card" onclick="openProject(${i})">
            <div class="p-info">
                <div class="p-label">${p.name || 'UNBENANNT'}</div>
                <div class="p-title">${p.num || '---'}</div>
            </div>
            <div class="btn-delete" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:120px"></div>';
}

// --- ИНСТРУМЕНТЫ (С ЛОГОТИПОМ НОМЕРА) ---
function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const p = db[currentIdx];
    const tools = p.tools || [];
    
    // Вставляем навигацию и заголовок проекта
    list.innerHTML = `
        <div class="premium-nav">
            <button class="btn-top" onclick="goHome()">← Home</button>
            <button class="btn-top blue" onclick="makePDF()">PDF Report</button>
        </div>
        <div class="hero">
            <div class="mirror-wrap">
                <div class="t-top">${p.num}</div>
                <div style="font-weight:800; color:#8e8e93; margin-top:5px;">${p.name}</div>
            </div>
        </div>
        <div id="tool-items"></div>
    `;
    
    const itemsCont = el('tool-items');
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'tool-card';
        const revMark = t.rev ? `<div style="background:#000; color:#fff; font-size:8px; padding:2px 7px; border-radius:4px; margin-bottom:5px; font-weight:900; width:fit-content;">REVOLVER UNTEN</div>` : '';
        item.innerHTML = `
            <div style="flex:1; min-width:0;" onclick="modalT(${i})">
                ${revMark}
                <div style="font-size:10px; font-weight:700; color:var(--text-sub); text-transform:uppercase;">${t.id || 'T0000'}</div>
                <div style="font-size:18px; font-weight:800; color:#000;">${t.nm || '---'}</div>
                <div style="margin-top:4px; font-weight:700; color:var(--accent); font-size:13px;">${t.dia || ''}</div>
            </div>
            <div class="order-controls">
                <button class="btn-order" onclick="event.stopPropagation(); moveItem(${i}, -1)">↑</button>
                <button class="btn-order" onclick="event.stopPropagation(); moveItem(${i}, 1)">↓</button>
            </div>`;
        itemsCont.appendChild(item);
    });
    itemsCont.innerHTML += '<div style="height:180px"></div>';
}

// Остальные функции (moveItem, openProject, goHome, modalP, saveP, modalT, saveT, deleteProject, delT, exportJSON, importJSON, makePDF)
// остаются БЕЗ ИЗМЕНЕНИЙ, как в твоем рабочем коде.

function moveItem(i, direction) {
    const tools = db[currentIdx].tools;
    const target = i + direction;
    if (target >= 0 && target < tools.length) {
        [tools[i], tools[target]] = [tools[target], tools[i]];
        save(); renderTools();
    }
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

function modalP(edit = false) {
    if (!edit) currentIdx = null;
    const p = (edit && currentIdx !== null && db[currentIdx]) ? db[currentIdx] : {num:'', name:'', lzf:'', sag:'', stt:'', stn:'', abs:'', grf:'', mat:''};
    el('p-idx').value = edit ? currentIdx : '';
    el('p-num').value = p.num; el('p-nam').value = p.name;
    el('p-lzf').value = p.lzf; el('p-sag').value = p.sag;
    el('p-stt').value = p.stt; el('p-stn').value = p.stn;
    el('p-abs').value = p.abs; el('p-grf').value = p.grf;
    if(el('p-mat')) el('p-mat').value = p.mat;
    show('m-p');
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
    save(); hide('m-p'); renderList();
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    const btn = el('btn-rev-toggle');
    if(btn) { t.rev ? btn.classList.add('on') : btn.classList.remove('on'); btn.onclick = () => btn.classList.toggle('on'); }
    el('btn-del-t').style.display = edit ? 'block' : 'none';
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const btn = el('btn-rev-toggle');
    const t = { id: el('t-id').value.toUpperCase(), nm: el('t-nm').value.toUpperCase(), dia: el('t-dia').value, rev: btn ? btn.classList.contains('on') : false };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    save(); renderTools(); hide('m-t');
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); save(); renderList(); } }
function delT() { const i = el('t-idx').value; db[currentIdx].tools.splice(i, 1); save(); renderTools(); hide('m-t'); }
function exportJSON() { el('imp-area').value = JSON.stringify(db); el('imp-area').select(); alert("JSON kopiert!"); }
function importJSON() { try { const parsed = JSON.parse(el('imp-area').value); if(Array.isArray(parsed)) { db = parsed; save(); renderList(); hide('m-imp'); } } catch(e) { alert("JSON-Fehler"); } }

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
