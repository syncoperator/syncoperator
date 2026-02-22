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
                <div class="t-name-label">${p.num || '---'}</div>
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

// Drag and Drop reordering (Up/Down buttons)
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
                <button onclick="moveTool(${i},-1)" style="border:none; background:#eee; border-radius:4px; padding:5px;">↑</button>
                <button onclick="moveTool(${i},1)" style="border:none; background:#eee; border-radius:4px; padding:5px;">↓</button>
            </div>
            <div style="flex:1" onclick="modalT(${i})">
                <div class="t-id-label">BEMERKUNG</div>
                <div class="t-name-label" style="font-size:18px;">${t.bem || '---'}</div>
                <div style="font-weight:700; font-size:15px;">${t.nm}</div>
                <div class="loc-tag ${t.isStandart ? 'standart' : 'sonder'}">
                    ${t.isStandart ? 'STANDART | ' + (t.loc || '') : 'SONDER | BLAUKISTE'}
                </div>
            </div>
        </div>`).join('') + '<div style="height:120px;"></div>';
}

function modalP() {
    el('p-idx').value = '';
    ['p-num','p-nam','p-lzf','p-mat','p-slg','p-abs','p-grf','p-stk'].forEach(id => el(id).value = '');
    el('m-p').style.display = 'flex';
}

function editCurrentProject() {
    const p = db[currentIdx];
    el('p-idx').value = currentIdx;
    el('p-num').value = p.num; el('p-nam').value = p.name;
    el('p-lzf').value = p.lzf||''; el('p-mat').value = p.mat||'';
    el('p-slg').value = p.slg||''; el('p-abs').value = p.abs||'';
    el('p-grf').value = p.grf||''; el('p-stk').value = p.stk||'';
    el('m-p').style.display = 'flex';
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
    try { const data = JSON.parse(el('imp-area').value); if(Array.isArray(data)) db = data; localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); } catch(e) { alert('Error'); }
}

function makePDF() {
    const p = db[currentIdx];
    const sonder = (p.tools || []).filter(x => !x.isStandart);
    const standart = (p.tools || []).filter(x => x.isStandart);
    
    let html = `
    <html><head><style>
        @page { size: A4; margin: 10mm; }
        body { margin: 0; font-family: sans-serif; -webkit-print-color-adjust: exact; }
        .outer-frame { border: 2.5px solid #000; width: 100%; }
        table { width: 100%; border-collapse: collapse; }
        td { border-bottom: 1.5px solid #000; vertical-align: middle; }
        .head-cell { padding: 20px; border-bottom: 2.5px solid #000; }
        .meta-cell { width: 220px; border-left: 2.5px solid #000; padding: 15px; border-bottom: 2.5px solid #000; }
        .cat-row { background: #000; color: #fff; font-weight: 900; font-size: 14px; padding: 8px 15px; }
        .tool-name { font-weight: 900; font-size: 18px; }
        .tool-info { font-size: 10px; font-weight: 800; color: #444; text-transform: uppercase; }
        .bem-cell { width: 160px; border-left: 1.5px solid #000; padding: 10px 15px; text-align: right; }
        .bem-label { font-size: 8px; font-weight: 900; color: #888; }
        .bem-val { font-weight: 900; font-size: 16px; }
        tr { page-break-inside: avoid; }
    </style></head><body>
        <div class="outer-frame">
            <table>
                <thead>
                    <tr>
                        <td class="head-cell">
                            <div style="font-size:15px; font-weight:900; text-transform:uppercase;">${p.name}</div>
                            <div style="font-size:60px; font-weight:900; line-height:0.8; letter-spacing:-2px;">${p.num}</div>
                        </td>
                        <td class="meta-cell">
                            <div style="display:flex; justify-content:space-between; font-size:11px; font-weight:900; margin-bottom:3px;"><span>MAT:</span><span>${p.mat||'--'}</span></div>
                            <div style="display:flex; justify-content:space-between; font-size:11px; font-weight:900; margin-bottom:3px;"><span>LZF:</span><span>${p.lzf||'--'}</span></div>
                            <div style="display:flex; justify-content:space-between; font-size:11px; font-weight:900; margin-bottom:3px;"><span>SÄGE:</span><span>${p.slg||'--'}</span></div>
                            <div style="display:flex; justify-content:space-between; font-size:11px; font-weight:900; margin-bottom:3px;"><span>ABST:</span><span>${p.abs||'--'}</span></div>
                            <div style="display:flex; justify-content:space-between; font-size:11px; font-weight:900; margin-bottom:3px;"><span>BACKEN:</span><span>${p.grf||'--'}</span></div>
                            <div style="display:flex; justify-content:space-between; font-size:11px; font-weight:900;"><span>STOCK:</span><span>${p.stk||'--'}</span></div>
                        </td>
                    </tr>
                </thead>
                <tbody>`;

    const addGroup = (title, list, labelFn) => {
        if(list.length > 0) {
            html += `<tr><td colspan="2" class="cat-row">${title}</td></tr>`;
            list.forEach(t => {
                html += `
                <tr>
                    <td style="padding:10px 15px;">
                        <div class="tool-name">${t.nm}</div>
                        <div class="tool-info">${labelFn(t)} ${t.kom ? ' | '+t.kom : ''}</div>
                    </td>
                    <td class="bem-cell">
                        <div class="bem-label">BEMERKUNG</div>
                        <div class="bem-val">${t.bem || '--'}</div>
                    </td>
                </tr>`;
            });
        }
    };

    addGroup('SONDERWERKZEUGE', sonder, () => 'IN BLAUKISTE');
    addGroup('STANDARTWERKZEUGE', standart, (t) => t.loc || 'STANDART');

    html += `</tbody></table></div></body></html>`;

    const win = window.open('','_blank');
    win.document.write(html);
    win.document.close();
    setTimeout(() => { win.print(); win.close(); }, 500);
}

window.onload = renderList;
