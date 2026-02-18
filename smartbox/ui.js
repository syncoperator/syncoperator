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
                <div style="font-size:10px; font-weight:800; color:#888;">PROJEKT</div>
                <div style="font-size:22px; font-weight:900;">${p.num || '---'}</div>
                <div style="font-size:12px; font-weight:700; color:#444;">${p.name || ''}</div>
            </div>
            <div style="padding:10px; color:#ccc;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:120px;"></div>';
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

function renderTools() {
    const list = el('list-t');
    const tools = db[currentIdx].tools || [];
    list.innerHTML = tools.map((t, i) => `
        <div class="list-item" onclick="modalT(${i})">
            <div style="flex:1">
                <div style="font-size:10px; font-weight:800; color:#888;">${t.id || '---'}</div>
                <div style="font-size:16px; font-weight:800;">${t.nm}</div>
                <div class="loc-tag ${t.isStandart ? 'standart' : 'sonder'}">
                    ${t.isStandart ? 'STANDART | ' + (t.loc || '') : 'SONDER | BLAUKISTE'}
                </div>
            </div>
        </div>`).join('') + '<div style="height:150px;"></div>';
}

function modalP() {
    el('p-idx').value = '';
    ['p-num','p-nam','p-lzf','p-mat','p-slg','p-abs','p-grf','p-stk'].forEach(id => el(id).value = '');
    el('m-p').style.display = 'flex';
}

function editCurrentProject() {
    const p = db[currentIdx];
    el('p-idx').value = currentIdx;
    el('p-num').value = p.num || '';
    el('p-nam').value = p.name || '';
    el('p-lzf').value = p.lzf || '';
    el('p-mat').value = p.mat || '';
    el('p-slg').value = p.slg || '';
    el('p-abs').value = p.abs || '';
    el('p-grf').value = p.grf || '';
    el('p-stk').value = p.stk || '';
    el('m-p').style.display = 'flex';
}

function saveP() {
    const i = el('p-idx').value;
    const data = {
        num: el('p-num').value,
        name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value, mat: el('p-mat').value,
        slg: el('p-slg').value, abs: el('p-abs').value,
        grf: el('p-grf').value, stk: el('p-stk').value,
        tools: i !== '' ? db[i].tools : []
    };
    if(i === '') db.push(data); else db[i] = data;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    if(i !== '') { el('h-num').innerText = data.num; el('h-nam').innerText = data.name; }
    renderList(); hide('m-p');
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', isStandart:false, loc:''};
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-loc').value = t.loc || '';
    const btn = el('btn-storage-toggle');
    t.isStandart ? btn.classList.add('on') : btn.classList.remove('on');
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

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i,1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { db[currentIdx].tools.splice(el('t-idx').value,1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }

function openImport() { el('imp-area').value = ''; el('m-imp').style.display = 'flex'; }
function exportJSON() { el('imp-area').value = JSON.stringify(db); }
function importAnyJSON() {
    try { db = JSON.parse(el('imp-area').value); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); } catch(e) { alert('Error'); }
}

function makePDF() {
    const p = db[currentIdx];
    const LIMIT = 13;
    const totalPages = Math.ceil((p.tools || []).length / LIMIT) || 1;
    let html = "";

    for(let i=0; i<totalPages; i++) {
        const segment = (p.tools || []).slice(i*LIMIT, (i+1)*LIMIT);
        const sonder = segment.filter(t => !t.isStandart);
        const standart = segment.filter(t => t.isStandart);
        const row = (t) => `<tr><td style="border-bottom:2px solid #000;padding:10px;"><div style="font-weight:800;font-size:16px;">${t.nm}</div><div style="font-size:10px;font-weight:700;">${t.isStandart?t.loc:'BLAUKISTE'}</div></td><td style="border-bottom:2px solid #000;text-align:right;font-weight:900;">${t.id||''}</td></tr>`;

        html += `
        <div style="width:210mm; height:297mm; padding:10mm; box-sizing:border-box; page-break-after:always;">
            <div style="border:3px solid #000; height:100%; display:flex; flex-direction:column;">
                <div style="padding:20px; border-bottom:5px solid #000;">
                    <div style="font-size:12px; font-weight:900; color:#666;">${p.name || ''}</div>
                    <div style="font-size:64px; font-weight:900; line-height:0.8; margin:10px 0;">${p.num || '---'}</div>
                    <div style="display:grid; grid-template-columns:1fr 1fr; gap:5px; font-size:11px; font-weight:900; border-top:2px solid #000; padding-top:10px;">
                        <div>MAT: ${p.mat||'--'}</div><div>LZF: ${p.lzf||'--'}</div>
                        <div>SÄGE: ${p.slg||'--'}</div><div>ABST: ${p.abs||'--'}</div>
                        <div>BACKEN: ${p.grf||'--'}</div><div>STÜCK: ${p.stk||'--'}</div>
                    </div>
                </div>
                <table style="width:100%; border-collapse:collapse;">
                    <tbody>
                        ${sonder.length?`<tr><td colspan="2" style="background:#000;color:#fff;padding:8px;font-weight:900;">SONDER</td></tr>${sonder.map(row).join('')}`:''}
                        ${standart.length?`<tr><td colspan="2" style="background:#000;color:#fff;padding:8px;font-weight:900;">STANDART</td></tr>${standart.map(row).join('')}`:''}
                    </tbody>
                </table>
            </div>
        </div>`;
    }
    const win = window.open('','_blank');
    win.document.write(`<html><head><style>@page{margin:0;}body{margin:0;font-family:sans-serif;}</style></head><body>${html}</body></html>`);
    win.document.close();
    setTimeout(() => { win.print(); win.close(); }, 700);
}

window.onload = renderList;
