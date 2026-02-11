const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;
let startIdx = null;

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'flex'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

// --- УПРАВЛЕНИЕ ПРОЕКТАМИ ---
function renderList() {
    const list = el('list-p'); if(!list) return;
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})">
            <div style="text-align:left;">
                <div class="t-nr">${p.name || 'PROJEKT'}</div>
                <div class="beschreibung" style="margin:5px 0 0 0;">${p.num || '---'}</div>
            </div>
            <div style="position:absolute; right:20px; top:35%; color:#ff3b30; font-weight:900;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div class="ui-padding"></div>';
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

// --- РЕНДЕР ИНСТРУМЕНТОВ (PREMIUM) ---
function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'tool-card';
        item.setAttribute('data-idx', i);

        // Цвет точки: Синий (Нижний револьвер), Зеленый (Верхний)
        const dotColor = t.rev ? '#007aff' : '#62cc71';

        item.innerHTML = `
            <div class="handle">☰</div>
            <div class="rev-dot" style="background:${dotColor}; box-shadow: 0 0 10px ${dotColor};"></div>
            <div onclick="modalT(${i})" style="text-align:center; padding-top:10px;">
                <div class="t-nr">${t.id || 'T-NR'}</div>
                <div class="beschreibung">${t.nm || '---'}</div>
                ${t.dia ? `<div class="dia-val">${t.dia} <small style="font-size:14px; color:var(--sub-text);">mm</small></div>` : ''}
            </div>
        `;
        
        // Drag & Drop
        const handle = item.querySelector('.handle');
        handle.ontouchstart = () => { startIdx = i; };
        handle.ontouchmove = (e) => {
            e.preventDefault();
            const touch = e.touches[0];
            const target = document.elementFromPoint(touch.clientX, touch.clientY);
            const targetItem = target?.closest('.tool-card');
            if (targetItem) {
                const overIdx = parseInt(targetItem.getAttribute('data-idx'));
                if (overIdx !== startIdx) { moveTool(startIdx, overIdx); startIdx = overIdx; }
            }
        };
        handle.ontouchend = () => renderTools();
        item.draggable = true;
        item.ondragstart = () => { startIdx = i; item.classList.add('selected'); };
        item.ondragover = (e) => e.preventDefault();
        item.ondrop = () => { if(startIdx !== i) moveTool(startIdx, i); };
        item.ondragend = () => renderTools();

        list.appendChild(item);
    });

    const footer = document.createElement('div');
    footer.innerHTML = `
        <div style="display:flex; justify-content:center; gap:15px; margin-top:20px;">
            <button onclick="exportJSON(); show('m-imp');">EXPORT JSON</button>
            <button onclick="show('m-imp')">IMPORT JSON</button>
        </div>
        <div class="ui-padding"></div>
    `;
    list.appendChild(footer);
}

function moveTool(from, to) {
    const tools = db[currentIdx].tools;
    const item = tools.splice(from, 1)[0];
    tools.splice(to, 0, item);
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools();
}

// --- МОДАЛЬНЫЕ ОКНА И СОХРАНЕНИЕ ---
function saveP() {
    const idx = el('p-idx').value;
    const newP = {
        num: el('p-num').value,
        name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value, sag: el('p-sag').value,
        stt: el('p-stt').value, stn: el('p-stn').value,
        abs: el('p-abs').value, grf: el('p-grf').value,
        mat: el('p-mat').value.toUpperCase(),
        tools: (idx !== '' && db[idx]) ? (db[idx].tools || []) : []
    };
    if (idx === '') db.push(newP); else db[idx] = newP;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    hide('m-p'); renderList(); goHome();
}

function saveT() {
    const i = el('t-idx').value;
    const btn = el('btn-rev-toggle');
    const t = { 
        id: el('t-id').value.toUpperCase(), 
        nm: el('t-nm').value.toUpperCase(), 
        dia: el('t-dia').value, 
        rev: btn ? btn.classList.contains('on') : false 
    };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); hide('m-t');
}

// --- PDF REPORT (ЗАФИКСИРОВАННАЯ ВЕРСИЯ) ---
function makePDF() {
    const p = db[currentIdx];
    const getPageHead = () => `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <div>
                <div style="font-size:13px; font-weight:900; color:#666;">${p.name || ''}</div>
                <div style="font-size:64px; font-weight:900; line-height:0.8;">${p.num || '---'}</div>
            </div>
            <div style="width:220px; font-size:11px; font-weight:800; line-height:1.6;">
                <div style="border-bottom:1px solid #eee;">LAUFZEIT: ${p.lzf || ''}</div>
                <div style="border-bottom:1px solid #eee;">MATERIAL: ${p.mat || ''}</div>
                <div style="border-bottom:1px solid #eee;">SÄGELÄNGE: ${p.sag || ''}</div>
                <div style="border-bottom:1px solid #eee;">ABSTAND: ${p.abs || ''}</div>
                <div style="border-bottom:1px solid #eee;">STÜCK: ${p.stt || ''}/${p.stn || ''}</div>
            </div>
        </div>
        <div style="border-bottom:5px solid #000; margin-bottom:15px;"></div>`;

    const getRow = (t) => `
        <div style="display:flex; border-bottom:1.5px solid #000; padding:10px 0;">
            <div style="width:75px; font-weight:800;">${t.id}</div>
            <div style="flex:1; font-weight:700; text-transform:uppercase;">${t.nm}</div>
            <div style="width:100px; text-align:right; font-weight:800;">${t.dia}</div>
        </div>`;

    let oben = [], unten = [];
    (p.tools || []).forEach(t => t.rev ? unten.push(t) : oben.push(t));

    let html = `<div style="width:210mm; padding:15mm; background:#fff; color:#000; font-family:sans-serif;">
        <div style="border:3px solid #000; padding:20px; min-height:260mm; page-break-after:always;">
            ${getPageHead()}
            <h2 style="margin:0 0 10px 0;">REVOLVER OBEN</h2>
            ${oben.map(getRow).join('')}
        </div>`;
    if(unten.length > 0) {
        html += `<div style="border:3px solid #000; padding:20px; min-height:260mm; margin-top:20px;">
            ${getPageHead()}
            <h2 style="margin:0 0 10px 0;">REVOLVER UNTEN</h2>
            ${unten.map(getRow).join('')}
        </div>`;
    }
    html += `</div>`;
    el('print-container').innerHTML = html;
    setTimeout(() => { window.print(); }, 150);
}

function exportJSON() { el('imp-area').value = JSON.stringify(db); }
function importJSON() { try { const parsed = JSON.parse(el('imp-area').value); if(Array.isArray(parsed)) { db = parsed; localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); } } catch(e) { alert("Error"); } }
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { const i = el('t-idx').value; db[currentIdx].tools.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }

window.onload = renderList;
