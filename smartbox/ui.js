const DB_KEY = 'SMARTBOX_PRO_V2';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

function renderList() {
    const list = el('list-p');
    if(!list) return;
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})">
            <div style="flex:1">
                <div class="t-id-label">PROJEKT</div>
                <div class="t-name-label">${p.num || '---'}</div>
                <div style="font-size:11px; font-weight:700; color:#86868b;">${p.name || ''}</div>
            </div>
            <div style="font-size:18px; padding:10px; color:#86868b;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:120px"></div>';
}

function openProject(i) {
    currentIdx = i;
    el('v-home').classList.remove('active');
    el('v-det').classList.add('active');
    el('h-num').innerText = db[i].num;
    el('h-nam').innerText = db[i].name;
    renderTools();
}

function goHome() {
    currentIdx = null;
    el('v-home').classList.add('active');
    el('v-det').classList.remove('active');
    renderList();
}

function renderTools() {
    const list = el('list-t');
    if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = tools.map((t, i) => `
        <div class="list-item" data-idx="${i}">
            <div class="handle">☰</div>
            <div style="flex:1" onclick="modalT(${i})">
                <div class="t-name-label">${t.nm}</div>
                <div class="loc-tag ${t.isStandart ? 'standart' : 'sonder'}">
                    ${t.isStandart ? 'STANDART | ' + (t.loc || '') : 'SONDER | BLAUKISTE'}
                </div>
                <div style="font-size:11px; margin-top:5px; font-weight:700; color:#86868b;">${t.id || ''}</div>
            </div>
        </div>`).join('') + '<div style="height:150px"></div>';
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', isStandart:false, loc:''};
    el('t-id').value = t.id; 
    el('t-nm').value = t.nm; 
    el('t-loc').value = t.loc || '';
    const btn = el('btn-storage-toggle');
    if(t.isStandart) btn.classList.add('on'); else btn.classList.remove('on');
    btn.onclick = () => btn.classList.toggle('on');
    el('btn-del-t').style.display = edit ? 'block' : 'none';
    el('m-t').style.display = 'flex';
}

function saveT() {
    const i = el('t-idx').value;
    const t = { 
        id: el('t-id').value.toUpperCase(), 
        nm: el('t-nm').value.toUpperCase(), 
        isStandart: el('btn-storage-toggle').classList.contains('on'),
        loc: el('t-loc').value.toUpperCase()
    };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); hide('m-t');
}

function modalP() { 
    el('p-idx').value = ''; 
    el('p-num').value = ''; el('p-nam').value = '';
    el('p-lzf').value = ''; el('p-mat').value = ''; el('p-grf').value = '';
    el('m-p').style.display = 'flex'; 
}

function editCurrentProject() {
    const p = db[currentIdx];
    el('p-idx').value = currentIdx;
    el('p-num').value = p.num; el('p-nam').value = p.name;
    el('p-lzf').value = p.lzf || ''; el('p-mat').value = p.mat || ''; el('p-grf').value = p.grf || '';
    el('m-p').style.display = 'flex';
}

function saveP() {
    const i = el('p-idx').value;
    const pData = { 
        num: el('p-num').value, 
        name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value, mat: el('p-mat').value, grf: el('p-grf').value,
        tools: (i !== '' ? db[i].tools : [])
    };
    if(i === '') db.push(pData); else db[i] = pData;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    if(i !== '') { el('h-num').innerText = pData.num; el('h-nam').innerText = pData.name; }
    renderList(); hide('m-p');
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i,1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { db[currentIdx].tools.splice(el('t-idx').value, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }

// --- УМНЫЙ ИМПОРТ ---
function importAnyJSON() {
    try {
        const raw = JSON.parse(el('imp-area').value);
        if (Array.isArray(raw)) { db = raw; } 
        else if (raw.num || raw.name) {
            const imported = {
                num: raw.num || '---',
                name: (raw.name || 'IMPORT').toUpperCase(),
                lzf: raw.lzf || '', mat: raw.mat || '', grf: raw.grf || '',
                tools: (raw.tools || []).map(t => ({
                    nm: (t.name || t.nm || '').toUpperCase(),
                    id: (t.id || '').toUpperCase(),
                    isStandart: false, loc: ''
                }))
            };
            db.push(imported);
        }
        localStorage.setItem(DB_KEY, JSON.stringify(db));
        renderList(); hide('m-imp');
    } catch(e) { alert('Format Error'); }
}

function openImport() { el('imp-area').value = ''; el('m-imp').style.display = 'flex'; }
function exportJSON() { el('imp-area').value = JSON.stringify(db); }

// --- PDF ГЕНЕРАТОР (МАКСИМАЛЬНАЯ СТАБИЛЬНОСТЬ) ---
function makePDF() {
    const p = db[currentIdx];
    const sonder = (p.tools || []).filter(t => !t.isStandart);
    const standart = (p.tools || []).filter(t => t.isStandart);

    const getRow = (t) => `
        <tr style="page-break-inside: avoid;">
            <td style="border-bottom: 1.5px solid #000; padding: 8px 10px;">
                <div style="font-weight: 700; font-size: 15px; text-transform: uppercase;">${t.nm}</div>
                <div style="font-size: 10px; font-weight: 800; color: #555;">${t.isStandart ? (t.loc || '') : 'IN BLAUKISTE'}</div>
            </td>
            <td style="border-bottom: 1.5px solid #000; padding: 8px 10px; text-align: right; font-weight: 800; font-size: 13px;">
                ${t.id || ''}
            </td>
        </tr>`;

    const html = `<html>
    <head>
        <style>
            @page { size: A4; margin: 10mm; }
            body { margin: 0; padding: 0; font-family: Helvetica, Arial, sans-serif; background: #fff; }
            
            table { width: 100%; border-collapse: collapse; border: 2.5px solid #000; table-layout: fixed; }
            
            /* Фикс для повтора шапки в iOS и Chrome */
            thead { display: table-header-group; }
            tfoot { display: table-footer-group; }

            .header-box { padding: 15px; text-align: left; }
            .info-grid { display: flex; justify-content: space-between; font-size: 10px; font-weight: 800; margin-top: 10px; border-top: 1px solid #eee; padding-top: 5px; }
            
            .section-title { 
                background: #000 !important; color: #fff !important; 
                padding: 10px; font-weight: 900; font-size: 13px; 
                text-transform: uppercase; -webkit-print-color-adjust: exact;
            }

            td, th { overflow: hidden; word-wrap: break-word; }
        </style>
    </head>
    <body>
        <table>
            <thead>
                <tr>
                    <th colspan="2" class="header-box">
                        <div style="font-size:12px; font-weight:900; color:#666;">${p.name || ''}</div>
                        <div style="font-size:54px; font-weight:900; line-height:0.9; letter-spacing:-2px;">${p.num || '---'}</div>
                        <div class="info-grid">
                            <span>MAT: ${p.mat || '--'} | LZF: ${p.lzf || '--'}</span>
                            <span>BACKEN: ${p.grf || '--'}</span>
                        </div>
                    </th>
                </tr>
            </thead>
            <tbody>
                ${sonder.length > 0 ? `
                    <tr><td colspan="2" class="section-title">IN BLAUKISTE (SONDER)</td></tr>
                    ${sonder.map(getRow).join('')}
                ` : ''}
                ${standart.length > 0 ? `
                    <tr><td colspan="2" class="section-title">IN SCHUBLADEN (STANDART)</td></tr>
                    ${standart.map(getRow).join('')}
                ` : ''}
            </tbody>
        </table>
        <script>
            window.onload = () => { 
                setTimeout(() => { window.print(); window.close(); }, 600); 
            };
        <\/script>
    </body></html>`;
    
    const win = window.open('', '_blank');
    win.document.write(html);
    win.document.close();
}

window.onload = () => renderList();
