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
    el('p-slg').value = ''; el('p-abs').value = ''; el('p-stk').value = '';
    el('m-p').style.display = 'flex'; 
}

function editCurrentProject() {
    const p = db[currentIdx];
    el('p-idx').value = currentIdx;
    el('p-num').value = p.num || '';
    el('p-nam').value = p.name || '';
    el('p-lzf').value = p.lzf || ''; 
    el('p-mat').value = p.mat || '';
    el('p-grf').value = p.grf || '';
    el('p-slg').value = p.slg || '';
    el('p-abs').value = p.abs || '';
    el('p-stk').value = p.stk || '';
    el('m-p').style.display = 'flex';
}

function saveP() {
    const i = el('p-idx').value;
    const pData = { 
        num: el('p-num').value, 
        name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value, 
        mat: el('p-mat').value, 
        grf: el('p-grf').value,
        slg: el('p-slg').value,
        abs: el('p-abs').value,
        stk: el('p-stk').value,
        tools: (i !== '' ? db[i].tools : [])
    };
    if(i === '') db.push(pData); else db[i] = pData;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    if(i !== '') { el('h-num').innerText = pData.num; el('h-nam').innerText = pData.name; }
    renderList(); hide('m-p');
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i,1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { db[currentIdx].tools.splice(el('t-idx').value, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }

// --- PDF ГЕНЕРАЦИЯ (НОВАЯ ШАПКА) ---
function makePDF() {
    const p = db[currentIdx];
    const allTools = p.tools || [];
    const LIMIT = 13; 

    const createPage = (toolsSegment, pageNum, totalPages) => {
        const sonder = toolsSegment.filter(t => !t.isStandart);
        const standart = toolsSegment.filter(t => t.isStandart);
        const row = (t) => `
            <tr>
                <td style="border-bottom: 2px solid #000; padding: 10px;">
                    <div style="font-weight: 800; font-size: 16px; text-transform: uppercase;">${t.nm}</div>
                    <div style="font-size: 10px; font-weight: 700; color: #444;">${t.isStandart ? (t.loc || '') : 'IN BLAUKISTE'}</div>
                </td>
                <td style="border-bottom: 2px solid #000; padding: 10px; text-align: right; font-weight: 900; font-size: 14px;">${t.id || ''}</td>
            </tr>`;

        return `
        <div style="width: 210mm; height: 297mm; padding: 10mm; box-sizing: border-box; page-break-after: always;">
            <div style="border: 3px solid #000; height: 100%; display: flex; flex-direction: column;">
                <div style="padding: 20px; border-bottom: 5px solid #000;">
                    <div style="font-size: 13px; font-weight: 900; color: #666; display: flex; justify-content: space-between;">
                        <span>${p.name || ''}</span>
                        <span>${totalPages > 1 ? 'PAGE '+pageNum+'/'+totalPages : ''}</span>
                    </div>
                    <div style="font-size: 64px; font-weight: 900; line-height: 0.8; letter-spacing: -2px; margin: 10px 0;">${p.num || '---'}</div>
                    
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 5px; font-size: 10px; font-weight: 900; border-top: 2px solid #000; padding-top: 10px;">
                        <div>MAT: ${p.mat || '--'}</div>
                        <div>LZF: ${p.lzf || '--'}</div>
                        <div>SÄGELÄNGE: ${p.slg || '--'}</div>
                        <div>ABSTAND: ${p.abs || '--'}</div>
                        <div>BACKEN: ${p.grf || '--'}</div>
                        <div>STÜCK: ${p.stk || '--'}</div>
                    </div>
                </div>
                <table style="width: 100%; border-collapse: collapse;">
                    <tbody>
                        ${sonder.length > 0 ? `<tr><td colspan="2" style="background:#000; color:#fff; padding:10px; font-weight:900; font-size:14px; -webkit-print-color-adjust: exact;">IN BLAUKISTE (SONDER)</td></tr>${sonder.map(row).join('')}` : ''}
                        ${standart.length > 0 ? `<tr><td colspan="2" style="background:#000; color:#fff; padding:10px; font-weight:900; font-size:14px; -webkit-print-color-adjust: exact;">IN SCHUBLADEN (STANDART)</td></tr>${standart.map(row).join('')}` : ''}
                    </tbody>
                </table>
            </div>
        </div>`;
    };

    let html = "";
    const total = Math.ceil(allTools.length / LIMIT) || 1;
    for (let i = 0; i < total; i++) {
        html += createPage(allTools.slice(i * LIMIT, (i + 1) * LIMIT), i + 1, total);
    }

    const win = window.open('', '_blank');
    win.document.write('<html><head><style>@page { margin: 0; } body { margin: 0; font-family: sans-serif; }</style></head><body>' + html + '</body></html>');
    win.document.close();
    setTimeout(() => { win.print(); win.close(); }, 750);
}

window.onload = () => renderList();
