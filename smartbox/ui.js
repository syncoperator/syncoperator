const DB_KEY = 'SMARTBOX_PRO_V2';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const hide = (id) => { el(id).style.display = 'none'; };

function renderList() {
    const list = el('list-p');
    if(!list) return;
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})">
            <div style="flex:1">
                <div class="t-id-label">PROJEKT</div>
                <div class="t-name-large">${p.num || '---'}</div>
                <div style="font-size:13px; font-weight:600; color:#8e8e93;">${p.name || ''}</div>
            </div>
            <div style="font-size:20px; color:#c7c7cc;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:100px;"></div>';
}

function openProject(i) {
    currentIdx = i;
    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    el('v-det').classList.add('active');
    el('h-num').innerText = db[i].num;
    el('h-nam').innerText = db[i].name;
    renderTools();
}

function goHome() {
    currentIdx = null;
    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    el('v-home').classList.add('active');
    renderList();
}

// Прямая сортировка стрелками
function moveTool(idx, dir) {
    const t = db[currentIdx].tools;
    const n = idx + dir;
    if(n >= 0 && n < t.length) {
        [t[idx], t[n]] = [t[n], t[idx]];
        localStorage.setItem(DB_KEY, JSON.stringify(db));
        renderTools();
    }
}

function renderTools() {
    const list = el('list-t');
    const tools = db[currentIdx].tools || [];
    list.innerHTML = tools.map((t, i) => `
        <div class="list-item">
            <div style="display:flex; flex-direction:column; gap:5px; margin-right:10px;">
                <button onclick="moveTool(${i},-1)" style="border:none; background:#f0f0f0; border-radius:6px; padding:8px; font-weight:900;">↑</button>
                <button onclick="moveTool(${i},1)" style="border:none; background:#f0f0f0; border-radius:6px; padding:8px; font-weight:900;">↓</button>
            </div>
            <div style="flex:1" onclick="modalT(${i})">
                <div class="t-id-label">BEMERKUNG</div>
                <div class="t-name-large" style="font-size:18px;">${t.bem || '---'}</div>
                <div style="font-weight:700; font-size:15px;">${t.nm}</div>
                <div class="loc-tag ${t.isStandart ? 'standart' : 'sonder'}">
                    ${t.isStandart ? 'STANDART | ' + (t.loc || '') : 'SONDER | BLAUKISTE'}
                </div>
            </div>
        </div>`).join('') + '<div style="height:150px;"></div>';
}

function saveP() {
    const i = el('p-idx').value;
    const data = {
        num: el('p-num').value, name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value, mat: el('p-mat').value,
        slg: el('p-slg').value, abs: el('p-abs').value,
        grf: el('p-grf').value, stk: el('p-stk').value,
        tools: i !== '' ? db[i].tools : []
    };
    if(i === '') db.push(data); else db[i] = data;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderList(); hide('m-p');
    if(i !== '') openProject(i);
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {bem:'', nm:'', isStandart:false, loc:'', kom:''};
    el('t-bem').value = t.bem; el('t-nm').value = t.nm; el('t-loc').value = t.loc||''; el('t-kom').value = t.kom||'';
    const btn = el('btn-storage-toggle');
    t.isStandart ? btn.classList.add('on') : btn.classList.remove('on');
    btn.onclick = () => btn.classList.toggle('on');
    el('m-t').style.display = 'flex';
}

function saveT() {
    const i = el('t-idx').value;
    const t = {
        bem: el('t-bem').value.toUpperCase(), nm: el('t-nm').value.toUpperCase(),
        isStandart: el('btn-storage-toggle').classList.contains('on'),
        loc: el('t-loc').value.toUpperCase(), kom: el('t-kom').value.toUpperCase()
    };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); hide('m-t');
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i,1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { db[currentIdx].tools.splice(el('t-idx').value,1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }
function openImport() { el('imp-area').value = ''; el('m-imp').style.display = 'flex'; }
function exportJSON() { el('imp-area').value = JSON.stringify(db); }
function importAnyJSON() {
    try { const d = JSON.parse(el('imp-area').value); if(Array.isArray(d)) db = d; localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); } catch(e) { alert('Error'); }
}

// ТОТ САМЫЙ PDF С ЖИРНЫМИ ЛИНИЯМИ И БЕЗ РАЗРЫВОВ
function makePDF() {
    const p = db[currentIdx];
    const sonder = (p.tools || []).filter(x => !x.isStandart);
    const standart = (p.tools || []).filter(x => x.isStandart);
    
    let html = `
    <html><head><style>
        @page { size: A4; margin: 10mm; }
        body { margin: 0; font-family: sans-serif; -webkit-print-color-adjust: exact; }
        
        /* Внешняя рамка - жирная, 3px */
        .pdf-frame { border: 3px solid #000; width: 100%; border-collapse: collapse; }
        
        /* Шапка */
        .header-row td { border-bottom: 3px solid #000; padding: 20px; }
        .meta-box { width: 200px; border-left: 3px solid #000; padding: 12px; font-size: 11px; font-weight: 900; }
        
        /* Разделители категорий (черные плашки) */
        .cat-head { background: #000 !important; color: #fff !important; font-weight: 900; font-size: 14px; padding: 8px 15px !important; text-transform: uppercase; }
        
        /* Строки инструментов */
        .tool-row td { border-bottom: 2px solid #000; padding: 12px 15px; vertical-align: middle; }
        .bem-box { border-left: 2px solid #000; width: 150px; text-align: right; }
        
        .t-name { font-weight: 900; font-size: 18px; display: block; }
        .t-sub { font-size: 10px; font-weight: 800; color: #444; margin-top: 4px; text-transform: uppercase; }
        .b-label { font-size: 8px; color: #888; font-weight: 900; display: block; }
        .b-val { font-size: 16px; font-weight: 900; }
        
        tr { page-break-inside: avoid; }
    </style></head><body>
        <table class="pdf-frame">
            <thead>
                <tr class="header-row">
                    <td>
                        <div style="font-size:14px; font-weight:900;">${p.name || ''}</div>
                        <div style="font-size:55px; font-weight:900; line-height:0.9; letter-spacing:-2px;">${p.num}</div>
                    </td>
                    <td class="meta-box">
                        <div style="display:flex; justify-content:space-between; margin-bottom:2px;"><span>MAT:</span><span>${p.mat||'--'}</span></div>
                        <div style="display:flex; justify-content:space-between; margin-bottom:2px;"><span>LZF:</span><span>${p.lzf||'--'}</span></div>
                        <div style="display:flex; justify-content:space-between; margin-bottom:2px;"><span>SÄGE:</span><span>${p.slg||'--'}</span></div>
                        <div style="display:flex; justify-content:space-between; margin-bottom:2px;"><span>ABST:</span><span>${p.abs||'--'}</span></div>
                        <div style="display:flex; justify-content:space-between; margin-bottom:2px;"><span>BACKEN:</span><span>${p.grf||'--'}</span></div>
                        <div style="display:flex; justify-content:space-between;"><span>STOCK:</span><span>${p.stk||'--'}</span></div>
                    </td>
                </tr>
            </thead>
            <tbody>`;

    const renderGroup = (title, list, locStr) => {
        if(list.length > 0) {
            html += `<tr><td colspan="2" class="cat-head">${title}</td></tr>`;
            list.forEach(t => {
                html += `
                <tr class="tool-row">
                    <td>
                        <span class="t-name">${t.nm}</span>
                        <span class="t-sub">${t.isStandart ? (t.loc || 'STANDART') : 'IN BLAUKISTE'} ${t.kom ? ' | ' + t.kom : ''}</span>
                    </td>
                    <td class="bem-box">
                        <span class="b-label">BEMERKUNG</span>
                        <span class="b-val">${t.bem || '--'}</span>
                    </td>
                </tr>`;
            });
        }
    };

    renderGroup('SONDERWERKZEUGE', sonder);
    renderGroup('STANDARTWERKZEUGE', standart);

    html += `</tbody></table></body></html>`;

    const win = window.open('','_blank');
    win.document.write(html);
    win.document.close();
    setTimeout(() => { win.print(); win.close(); }, 500);
}

window.onload = renderList;
