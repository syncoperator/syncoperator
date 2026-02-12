const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

// --- СИСТЕМА ДИЗАЙНА "MODERN PRECISION" ---
const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root {
            --bg: #f5f5f7;
            --surface: #ffffff;
            --accent: #007AFF;
            --accent-soft: rgba(0, 122, 255, 0.08);
            --text-primary: #1d1d1f;
            --text-secondary: #86868b;
            --border: rgba(0, 0, 0, 0.04);
            --shadow-sm: 0 2px 8px rgba(0,0,0,0.04);
            --shadow-lg: 0 20px 40px rgba(0,0,0,0.08);
        }

        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "Helvetica Neue", sans-serif !important;
            color: var(--text-primary);
            margin: 0;
            padding: 0;
            overflow-x: hidden;
        }

        /* Навигация и заголовок */
        header {
            background: rgba(255, 255, 255, 0.7) !important;
            backdrop-filter: saturate(180%) blur(20px) !important;
            -webkit-backdrop-filter: saturate(180%) blur(20px) !important;
            border-bottom: 0.5px solid var(--border) !important;
            padding: 16px 20px !important;
            position: sticky;
            top: 0;
            z-index: 1000;
        }

        .header-title {
            font-size: 22px !important;
            font-weight: 700 !important;
            letter-spacing: -0.5px !important;
            color: var(--text-primary) !important;
        }

        /* Контейнеры списков */
        #list-p, #list-t {
            padding: 20px 16px;
            max-width: 600px;
            margin: 0 auto;
        }

        /* Премиальные карточки */
        .list-item {
            background: var(--surface) !important;
            border-radius: 22px !important;
            margin-bottom: 16px !important;
            padding: 20px !important;
            border: 0.5px solid var(--border) !important;
            box-shadow: var(--shadow-sm) !important;
            transition: all 0.4s cubic-bezier(0.16, 1, 0.3, 1);
            position: relative;
            overflow: hidden;
        }

        .list-item:active {
            transform: scale(0.96);
            background: #fafafa !important;
        }

        /* Детализация в карточке инструмента (T-NR сверху, Bold снизу) */
        .t-id-label {
            font-size: 11px;
            font-weight: 600;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 4px;
        }

        .t-name-label {
            font-size: 19px;
            font-weight: 800;
            color: var(--text-primary);
            line-height: 1.25;
            letter-spacing: -0.2px;
        }

        .t-dia-badge {
            margin-top: 12px;
            display: inline-flex;
            background: var(--accent-soft);
            color: var(--accent);
            padding: 6px 12px;
            border-radius: 10px;
            font-size: 14px;
            font-weight: 700;
        }

        /* Кнопки управления */
        .actions-bar {
            position: fixed;
            bottom: 30px;
            left: 50%;
            transform: translateX(-50%);
            display: flex;
            gap: 12px;
            padding: 12px 20px;
            background: rgba(29, 29, 31, 0.8);
            backdrop-filter: blur(20px);
            border-radius: 100px;
            box-shadow: var(--shadow-lg);
            z-index: 900;
        }

        .btn-action {
            background: transparent;
            color: white;
            border: none;
            font-weight: 600;
            padding: 8px 16px;
            font-size: 14px;
            border-radius: 50px;
            transition: background 0.2s;
        }

        .btn-action:active { background: rgba(255,255,255,0.1); }

        /* Модальные окна */
        .modal-content {
            border-radius: 32px !important;
            padding: 30px !important;
            box-shadow: var(--shadow-lg) !important;
        }

        .handle {
            color: #d2d2d7;
            font-size: 18px;
            margin-right: 15px;
        }
    `;
    document.head.appendChild(style);
};

const renameBranding = () => {
    const title = document.querySelector('.header-title') || document.querySelector('h1');
    if(title) title.innerText = 'CitiTool';
};

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'flex'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

// --- ЛОГИКА ПРОЕКТОВ ---
function renderList() {
    renameBranding();
    const list = el('list-p'); if(!list) return;
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})" style="display:flex; justify-content:space-between; align-items:center;">
            <div>
                <div class="t-id-label">${p.name || 'Projekt'}</div>
                <div class="t-name-label">${p.num || '---'}</div>
            </div>
            <div style="background:#F5F5F7; color:#1d1d1f; width:36px; height:36px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:12px; font-weight:700;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:120px"></div>';
}

function openProject(i) {
    if (db[i]) {
        currentIdx = i;
        el('v-home').classList.remove('active');
        el('v-det').classList.add('active');
        el('h-num').innerText = db[i].num || '---';
        el('h-nam').innerText = db[i].name || '---';
        renderTools();
    }
}

function goHome() { currentIdx = null; el('v-home').classList.add('active'); el('v-det').classList.remove('active'); renderList(); }

// --- ЛОГИКА ИНСТРУМЕНТОВ ---
function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'list-item';
        item.setAttribute('data-idx', i);
        item.style.display = 'flex';
        item.style.alignItems = 'center';
        
        const revMark = t.rev ? `<div style="background:#1d1d1f; color:#fff; font-size:9px; padding:3px 10px; border-radius:100px; margin-bottom:8px; font-weight:700; width:fit-content;">UNTERER REVOLVER</div>` : '';
        
        item.innerHTML = `
            <div class="handle">☰</div>
            <div style="flex:1; min-width:0;" onclick="modalT(${i})">
                ${revMark}
                <div class="t-id-label">${t.id || 'T0000'}</div>
                <div class="t-name-label">${t.nm || '---'}</div>
                ${t.dia ? `<div class="t-dia-badge">${t.dia}</div>` : ''}
            </div>`;

        // Drag & Drop
        const handle = item.querySelector('.handle');
        handle.ontouchstart = () => { startIdx = i; item.style.boxShadow = "0 10px 30px rgba(0,0,0,0.1)"; };
        handle.ontouchmove = (e) => {
            e.preventDefault();
            const touch = e.touches[0];
            const target = document.elementFromPoint(touch.clientX, touch.clientY);
            const targetItem = target?.closest('.list-item');
            if (targetItem) {
                const overIdx = parseInt(targetItem.getAttribute('data-idx'));
                if (overIdx !== startIdx) { moveTool(startIdx, overIdx); startIdx = overIdx; }
            }
        };
        handle.ontouchend = () => renderTools();
        list.appendChild(item);
    });
    list.innerHTML += '<div style="height:150px"></div>'; 
}

function moveTool(from, to) {
    const tools = db[currentIdx].tools;
    const item = tools.splice(from, 1)[0];
    tools.splice(to, 0, item);
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools();
}

// --- СТАНДАРТНЫЕ ФУНКЦИИ (ПРОЕКТЫ/ИНСТРУМЕНТЫ) ---
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
    const newP = {
        num: el('p-num').value, name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value, sag: el('p-sag').value,
        stt: el('p-stt').value, stn: el('p-stn').value,
        abs: el('p-abs').value, grf: el('p-grf').value,
        mat: el('p-mat') ? el('p-mat').value.toUpperCase() : '',
        tools: (idx !== '' && db[idx]) ? (db[idx].tools || []) : []
    };
    if (idx === '') { db.push(newP); currentIdx = db.length - 1; } 
    else { db[idx] = newP; }
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    hide('m-p'); renderList();
    if(idx !== '') openProject(idx); else goHome();
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    const btn = el('btn-rev-toggle');
    if(btn) {
        if(t.rev) btn.classList.add('on'); else btn.classList.remove('on');
        btn.onclick = () => btn.classList.toggle('on');
    }
    el('btn-del-t').style.display = edit ? 'block' : 'none';
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const btn = el('btn-rev-toggle');
    const t = { id: el('t-id').value.toUpperCase(), nm: el('t-nm').value.toUpperCase(), dia: el('t-dia').value, rev: btn ? btn.classList.contains('on') : false };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); hide('m-t');
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { const i = el('t-idx').value; db[currentIdx].tools.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }
function exportJSON() { el('imp-area').value = JSON.stringify(db); el('imp-area').select(); alert("JSON kopiert!"); }
function importJSON() { try { const parsed = JSON.parse(el('imp-area').value); if(Array.isArray(parsed)) { db = parsed; localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); } } catch(e) { alert("JSON-Fehler"); } }

// --- PDF ИНИЦИАЛИЗАЦИЯ (НЕ ТРОГАЕМ) ---
function makePDF() {
    const p = db[currentIdx];
    const getPageHead = () => `<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;"><div style="display:flex; flex-direction:column;"><div style="font-size:13px; font-weight:900; color:#666;">${p.name || ''}</div><div style="font-size:64px; font-weight:900; line-height:0.8; letter-spacing:-2px;">${p.num || '---'}</div></div><div style="width:220px; font-size:11px; font-weight:800; line-height:1.5;"><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>MATERIAL</span><span>${p.mat || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>SÄGELÄNGE</span><span>${p.sag || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>ABSTAND</span><span>${p.abs || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>GREIFBACKEN</span><span>${p.grf || ''}</span></div><div style="display:flex; justify-content:space-between;"><span>STÜCKZAHL</span><span>${p.stt || ''} / ${p.stn || ''}</span></div></div></div><div style="border-bottom:5px solid #000; margin-bottom:15px;"></div>`;
    const tableHead = `<div style="display:flex; font-size:10px; font-weight:900; text-transform:uppercase; margin-bottom:6px; padding:0 2px;"><div style="width:75px;">T-NR</div><div style="flex:1;">WERKZEUGNAME / KOMMENTAR</div><div style="width:125px; text-align:right;">Ø / TOLERANZ</div></div><div style="border-bottom:4px solid #000;"></div>`;
    const getRow = (t) => `<div style="display:flex; border-bottom:1.5px solid #000; padding:10px 0; page-break-inside: avoid;"><div style="width:75px; font-weight:800; font-size:15px;">${t.id}</div><div style="flex:1; font-weight:700; font-size:15px; text-transform:uppercase; padding-right:10px;">${t.nm}</div><div style="width:125px; text-align:right; font-weight:800; font-size:14px;">${t.dia.replace(/\//g, '<br>')}</div></div>`;
    let oben = [], unten = [], target = oben;
    (p.tools || []).forEach(t => { if(t.rev) target = unten; target.push(t); });
    let pdfHtml = `<html><head><style>@page { size: A4; margin: 0; } body { margin: 0; padding: 10mm; font-family: sans-serif; } .page { width: 210mm; height: 297mm; padding: 15mm; box-sizing: border-box; page-break-after: always; display: flex; flex-direction: column; } .content-border { border: 2.2px solid #000; padding: 20px; flex: 1; display: flex; flex-direction: column; } .footer { border-top: 1.2px solid #000; padding-top: 5px; font-size: 9px; font-weight: 800; text-align: center; color: #666; margin-top: auto; }</style></head><body><div class="page"><div class="content-border">${getPageHead()}<div style="flex:1;"><div style="margin-bottom:5px; font-size:18px; font-weight:900;">REVOLVER OBEN</div>${tableHead}${oben.map(getRow).join('')}</div><div class="footer">CITITOOL REPORT</div></div></div>${unten.length > 0 ? `<div class="page"><div class="content-border">${getPageHead()}<div style="flex:1;"><div style="margin-bottom:5px; font-size:18px; font-weight:900;">REVOLVER UNTEN</div>${tableHead}${unten.map(getRow).join('')}</div><div class="footer">CITITOOL REPORT</div></div></div>` : ''}<script>window.onload = function() { setTimeout(() => { window.print(); }, 400); };</script></body></html>`;
    const win = window.open('', '_blank');
    if (win) { win.document.write(pdfHtml); win.document.close(); }
}

window.onload = () => {
    injectStyles();
    renameBranding();
    renderList();
};
