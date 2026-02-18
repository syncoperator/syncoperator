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

function importAnyJSON() {
    try {
        const raw = JSON.parse(el('imp-area').value);
        if (Array.isArray(raw)) { db = raw; } 
        else if (raw.num || raw.name) {
            const imported = {
                num: raw.num || '---', name: (raw.name || 'IMPORT').toUpperCase(),
                lzf: raw.lzf || '', mat: raw.mat || '', grf: raw.grf || '',
                tools: (raw.tools || []).map(t => ({
                    nm: (t.name || t.nm || '').toUpperCase(),
                    id: (t.id || '').toUpperCase(), isStandart: false, loc: ''
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

// --- ГЕНЕРАЦИЯ PDF С ЖЕСТКИМ РАЗДЕЛЕНИЕМ СТРАНИЦ ---
function makePDF() {
    const p = db[currentIdx];
    const allTools = p.tools || [];
    const LIMIT = 13; // Лимит инструментов на страницу

    const createPage = (toolsSegment, isFirst) => {
        const sonder = toolsSegment.filter(t => !t.isStandart);
        const standart = toolsSegment.filter(t => t.isStandart);
        const getRow = (t) => `
            <tr>
                <td style="border-bottom: 2px solid #000; padding: 10px;">
                    <div style="font-weight: 800; font-size: 16px;">${t.nm}</div>
                    <div style="font-size: 11px; font-weight: 700; color: #444;">${t.isStandart ? (t.loc || '') : 'IN BLAUKISTE'}</div>
                </td>
                <td style="border-bottom: 2px solid #000; padding: 10px; text-align: right; font-weight: 900; font-size: 14px;">${t.id || ''}</td>
            </tr>`;

        return `
        <div style="page-break-after: always; padding: 10mm; box-sizing: border-box;">
            <table style="width: 100%; border-collapse: collapse; border: 3px solid #000;">
                <thead>
                    <tr>
                        <th colspan="2" style="padding: 20px; text-align: left; border-bottom: 5px solid #000;">
                            <div style="font-size: 14px; font-weight: 900; color: #666;">${p.name || ''}</div>
                            <div style="font-size: 60px; font-weight: 900; line-height: 0.8; letter-spacing: -2px;">${p.num || '---'}</div>
                            <div style="display: flex; justify-content: space-between; font-size: 11px; font-weight: 900; margin-top: 15px; border-top: 2px solid #000; padding-top: 5px;">
                                <span>MAT: ${p.mat || '--'} | LZF: ${p.lzf || '--'}</span>
                                <span>BACKEN: ${p.grf || '--'}</span>
                            </div>
                        </th>
                    </tr>
                </thead>
                <tbody>
                    ${sonder.length > 0 ? `<tr><td colspan="2" style="background:#000; color:#fff; padding:10px; font-weight:900; font-size:14px; -webkit-print-color-adjust: exact;">IN BLAUKISTE (SONDER)</td></tr>${sonder.map(getRow).join('')}` : ''}
                    ${standart.length > 0 ? `<tr><td colspan="2" style="background:#000; color:#fff; padding:10px; font-weight:900; font-size:14px; -webkit-print-color-adjust: exact;">IN SCHUBLADEN (STANDART)</td></tr>${standart.map(getRow).join('')}` : ''}
                </tbody>
            </table>
        </div>`;
    };

    let fullHtml = "";
    for (let i = 0; i < allTools.length; i += LIMIT) {
        fullHtml += createPage(allTools.slice(i, i + LIMIT), i === 0);
    }

    const win = window.open('', '_blank');
    win.document.write(`<html><head><style>@page { margin: 0; } body { margin: 0; font-family: sans-serif; }</style></head><body>${fullHtml || createPage([], true)}<script>window.onload=()=>{setTimeout(()=>{window.print();window.close();},600);};<\/script></body></html>`);
    win.document.close();
}

window.onload = () => renderList();
