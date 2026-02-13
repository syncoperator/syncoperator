const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

// --- СИСТЕМА ДИЗАЙНА "CITITOOL PREMIUM STABLE" ---
const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root {
            --bg: #f2f2f7;
            --surface: #ffffff;
            --accent: #007aff;
            --text-main: #1c1c1e;
            --text-sub: #8e8e93;
            --border: rgba(0,0,0,0.06);
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, BlinkMacSystemFont, sans-serif !important;
            margin: 0; color: var(--text-main);
            overflow-x: hidden;
            -webkit-user-select: none; user-select: none; /* Блокируем выделение везде */
        }
        input, textarea { -webkit-user-select: text; user-select: text; } /* Разрешаем только в полях */

        header {
            background: rgba(255,255,255,0.8) !important;
            backdrop-filter: blur(20px); -webkit-backdrop-filter: blur(20px);
            border-bottom: 0.5px solid rgba(0,0,0,0.1);
            padding: 16px 20px; position: sticky; top: 0; z-index: 1000;
        }
        .header-title { font-weight: 800; font-size: 22px; letter-spacing: -0.5px; }

        .list-item {
            background: var(--surface) !important;
            border-radius: 20px !important;
            padding: 18px !important;
            margin: 0 16px 12px 16px !important;
            box-shadow: 0 4px 12px rgba(0,0,0,0.03) !important;
            display: flex; align-items: center;
            position: relative;
            touch-action: none; /* Отключаем стандартные жесты браузера */
        }
        
        .list-item.dragging {
            z-index: 9999;
            transform: scale(1.03);
            box-shadow: 0 15px 35px rgba(0,0,0,0.1) !important;
            background: #fafafa !important;
            opacity: 0.9;
        }

        .handle {
            padding: 15px 10px; margin-right: 5px;
            color: #c7c7cc; font-size: 22px;
            cursor: grab;
        }

        .t-id-label { font-size: 11px; font-weight: 700; color: var(--text-sub); text-transform: uppercase; margin-bottom: 2px; }
        .t-name-label { font-size: 19px; font-weight: 800; color: #000; line-height: 1.2; }
        .t-dia-label { margin-top: 8px; font-weight: 700; color: var(--accent); font-size: 15px; }

        .btn-main {
            background: var(--accent); color: white; border: none;
            border-radius: 14px; padding: 12px 24px; font-weight: 700;
        }
    `;
    document.head.appendChild(style);
};

// --- УЛУЧШЕННЫЙ DRAG & DROP ДЛЯ МОБИЛОК ---
let dragIdx = null;

function initDrag(item, index) {
    const handle = item.querySelector('.handle');
    
    const onTouchStart = (e) => {
        dragIdx = index;
        item.classList.add('dragging');
        // Вибрация для отклика (если поддерживается)
        if (window.navigator.vibrate) window.navigator.vibrate(10);
    };

    const onTouchMove = (e) => {
        if (dragIdx === null) return;
        e.preventDefault(); // Критично: отключаем скролл страницы
        
        const touch = e.touches[0];
        const target = document.elementFromPoint(touch.clientX, touch.clientY);
        const overItem = target?.closest('.list-item');
        
        if (overItem) {
            const overIdx = parseInt(overItem.getAttribute('data-idx'));
            if (overIdx !== dragIdx) {
                const tools = db[currentIdx].tools;
                const temp = tools[dragIdx];
                tools[dragIdx] = tools[overIdx];
                tools[overIdx] = temp;
                
                dragIdx = overIdx; // Меняем текущий индекс
                save();
                renderTools(); // Мгновенно перерисовываем
            }
        }
    };

    const onTouchEnd = () => {
        dragIdx = null;
        item.classList.remove('dragging');
        renderTools();
    };

    handle.addEventListener('touchstart', onTouchStart, { passive: false });
    handle.addEventListener('touchmove', onTouchMove, { passive: false });
    handle.addEventListener('touchend', onTouchEnd);
}

// --- ОСНОВНАЯ ЛОГИКА ПРИЛОЖЕНИЯ ---
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

function renderList() {
    const t = document.querySelector('.header-title') || document.querySelector('h1');
    if(t) t.innerText = 'CitiTool';
    
    const list = el('list-p'); if(!list) return;
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})">
            <div style="flex:1">
                <div class="t-id-label">${p.name || 'PROJEKT'}</div>
                <div class="t-name-label" style="font-size:22px;">${p.num || '---'}</div>
            </div>
            <div style="color:#ff3b30; font-weight:900; padding:15px;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:100px"></div>';
}

function openProject(i) {
    if (db[i]) {
        currentIdx = i;
        el('v-home').classList.remove('active');
        el('v-det').classList.add('active');
        el('h-num').innerText = db[i].num;
        el('h-nam').innerText = db[i].name;
        renderTools();
    }
}

function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'list-item';
        item.setAttribute('data-idx', i);
        
        const revMark = t.rev ? `<div style="background:#000; color:#fff; font-size:9px; padding:2px 8px; border-radius:6px; margin-bottom:6px; font-weight:800; width:fit-content;">REVOLVER UNTEN ↓</div>` : '';
        
        item.innerHTML = `
            <div class="handle">☰</div>
            <div style="flex:1; min-width:0;" onclick="modalT(${i})">
                ${revMark}
                <div class="t-id-label">${t.id || 'T0000'}</div>
                <div class="t-name-label">${t.nm || '---'}</div>
                <div class="t-dia-label">${t.dia || ''}</div>
            </div>`;
        
        initDrag(item, i);
        list.appendChild(item);
    });
    list.innerHTML += '<div style="height:200px"></div>';
}

// --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---
function goHome() { currentIdx = null; el('v-home').classList.add('active'); el('v-det').classList.remove('active'); renderList(); }
const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'flex'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

function modalP(edit = false) {
    if (!edit) currentIdx = null;
    const p = (edit && db[currentIdx]) ? db[currentIdx] : {num:'', name:'', lzf:'', sag:'', stt:'', stn:'', abs:'', grf:'', mat:''};
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
    if (idx === '') { db.push(newP); currentIdx = db.length - 1; } else { db[idx] = newP; }
    save(); hide('m-p'); renderList();
    if(idx !== '') openProject(idx); else goHome();
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    const btn = el('btn-rev-toggle');
    if(btn) {
        t.rev ? btn.classList.add('on') : btn.classList.remove('on');
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
    save(); renderTools(); hide('m-t');
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); save(); renderList(); } }
function delT() { const i = el('t-idx').value; db[currentIdx].tools.splice(i, 1); save(); renderTools(); hide('m-t'); }
function exportJSON() { el('imp-area').value = JSON.stringify(db); el('imp-area').select(); alert("JSON kopiert!"); }
function importJSON() { try { const parsed = JSON.parse(el('imp-area').value); if(Array.isArray(parsed)) { db = parsed; save(); renderList(); hide('m-imp'); } } catch(e) { alert("JSON-Fehler"); } }

// --- PDF REPORT (НЕ ТРОГАЕМ) ---
function makePDF() {
    const p = db[currentIdx];
    const getPageHead = () => `<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; min-height:90px;"><div style="display:flex; flex-direction:column; justify-content:center;"><div style="font-size:13px; font-weight:900; text-transform:uppercase; color:#666; margin-bottom:2px; line-height:1;">${p.name || ''}</div><div style="font-size:64px; font-weight:900; line-height:0.8; letter-spacing:-2px; margin:0;">${p.num || '---'}</div></div><div style="width:220px; font-size:11px; font-weight:800; line-height:1.5;"><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>MATERIAL</span><span>${p.mat || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>SÄGELÄNGE</span><span>${p.sag || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>ABSTAND</span><span>${p.abs || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>GREIFBACKEN</span><span>${p.grf || ''}</span></div><div style="display:flex; justify-content:space-between;"><span>STÜCKZAHL</span><span>${p.stt || ''} / ${p.stn || ''}</span></div></div></div><div style="border-bottom:5px solid #000; margin-bottom:15px;"></div>`;
    const tableHead = `<div style="display:flex; font-size:10px; font-weight:900; text-transform:uppercase; margin-bottom:6px; padding:0 2px;"><div style="width:75px;">T-NR</div><div style="flex:1;">WERKZEUGNAME / KOMMENTAR</div><div style="width:125px; text-align:right;">Ø / TOLERANZ</div></div><div style="border-bottom:4px solid #000;"></div>`;
    const getRow = (t) => `<div style="display:flex; align-items:baseline; border-bottom:1.5px solid #000; padding:10px 0; width:100%; page-break-inside: avoid;"><div style="width:75px; font-weight:800; font-size:15px;">${t.id}</div><div style="flex:1; font-weight:700; font-size:15px; text-transform:uppercase; padding-right:10px; white-space:pre-wrap;">${t.nm}</div><div style="width:125px; text-align:right; font-weight:800; font-size:14px;">${t.dia.replace(/\//g, '<br>')}</div></div>`;
    let oben = [], unten = [], target = oben;
    (p.tools || []).forEach(t => { if(t.rev) target = unten; target.push(t); });
    let pdfHtml = `<!DOCTYPE html><html><head><style>@page { size: A4; margin: 0; } body { margin: 0; padding: 10mm; background: #fff; font-family: sans-serif; -webkit-print-color-adjust: exact; } .page { width: 210mm; height: 297mm; padding: 15mm; box-sizing: border-box; page-break-after: always; display: flex; flex-direction: column; } .content-border { border: 2.2px solid #000; padding: 20px; flex: 1; display: flex; flex-direction: column; box-sizing: border-box; } .footer { border-top: 1.2px solid #000; padding-top: 5px; font-size: 9px; font-weight: 800; text-align: center; color: #666; margin-top: auto; }</style></head><body><div class="page"><div class="content-border">${getPageHead()}<div style="flex:1;"><div style="margin-bottom:5px; font-size:18px; font-weight:900; text-transform:uppercase;">REVOLVER OBEN</div>${tableHead}${oben.map(getRow).join('')}</div><div class="footer">CITITOOL REPORT</div></div></div>${unten.length > 0 ? `<div class="page"><div class="content-border">${getPageHead()}<div style="flex:1;"><div style="margin-bottom:5px; font-size:18px; font-weight:900; text-transform:uppercase;">REVOLVER UNTEN</div>${tableHead}${unten.map(getRow).join('')}</div><div class="footer">CITITOOL REPORT</div></div></div>` : ''}<script>window.onload = function() { setTimeout(() => { window.print(); }, 400); };</script></body></html>`;
    const win = window.open('', '_blank');
    if (win) { win.document.write(pdfHtml); win.document.close(); }
}

window.onload = () => { injectStyles(); renderList(); };
