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
            <div style="font-size:20px; color:#c7c7cc; padding-left:15px;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:140px;"></div>';
}

function openProject(i) {
    currentIdx = i;
    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    el('v-det').classList.add('active');
    el('h-num').innerText = db[i].num;
    el('h-nam').innerText = db[i].name;
    renderTools();
    window.scrollTo(0,0);
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
                <div class="t-id-label" style="font-size:9px;">BEMERKUNG</div>
                <div style="font-size:18px; font-weight:900; margin-bottom:2px;">${t.bem || '---'}</div>
                <div style="font-size:16px; font-weight:800;">${t.nm}</div>
                <div class="loc-tag ${t.isStandart ? 'standart' : 'sonder'}">
                    ${t.isStandart ? 'STANDART | ' + (t.loc || '') : 'SONDER | BLAUKISTE'}
                </div>
                ${t.kom ? `<div style="font-size:12px; color:#555; font-weight:600; margin-top:6px; padding:4px 8px; background:#f0f0f5; border-radius:6px;">${t.kom}</div>` : ''}
            </div>
        </div>`).join('') + '<div style="height:160px;"></div>';
}

function modalP() {
    el('p-idx').value = '';
    ['p-num','p-nam','p-lzf','p-mat','p-slg','p-abs','p-grf','p-stk'].forEach(id => el(id).value = '');
    el('m-p').style.display = 'flex';
}

function editCurrentProject() {
    const p = db[currentIdx];
    el('p-idx').value = currentIdx;
    el('p-num').value = p.num || ''; el('p-nam').value = p.name || '';
    el('p-lzf').value = p.lzf || ''; el('p-mat').value = p.mat || '';
    el('p-slg').value = p.slg || ''; el('p-abs').value = p.abs || '';
    el('p-grf').value = p.grf || ''; el('p-stk').value = p.stk || '';
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
    if(i !== '') { el('h-num').innerText = data.num; el('h-nam').innerText = data.name; }
    renderList(); hide('m-p');
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {bem:'', nm:'', isStandart:false, loc:'', kom:''};
    el('t-bem').value = t.bem || ''; el('t-nm').value = t.nm || ''; el('t-loc').value = t.loc || ''; el('t-kom').value = t.kom || '';
    const btn = el('btn-storage-toggle');
    t.isStandart ? btn.classList.add('on') : btn.classList.remove('on');
    btn.onclick = () => btn.classList.toggle('on');
    el('btn-del-t').style.display = edit ? 'block' : 'none';
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
    try { const p = JSON.parse(el('imp-area').value); if(Array.isArray(p)) db = p; localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); } catch(e) { alert('JSON Error'); }
}

function makePDF() {
    const p = db[currentIdx];
    const tools = p.tools || [];
    
    const sonder = tools.filter(t => !t.isStandart);
    const standart = tools.filter(t => t.isStandart);

    const fullList = [];
    if(sonder.length > 0) {
        fullList.push({ type: 'header', label: 'SONDERWERKZEUGE' });
        sonder.forEach(t => fullList.push({ type: 'row', label: 'SONDERWERKZEUGE', data: t }));
    }
    if(standart.length > 0) {
        fullList.push({ type: 'header', label: 'STANDARTWERKZEUGE' });
        standart.forEach(t => fullList.push({ type: 'row', label: 'STANDARTWERKZEUGE', data: t }));
    }

    // ЛИМИТ 20 ПОЗИЦИЙ (строки + плашки)
    const LIMIT = 20; 
    const totalPages = Math.ceil(fullList.length / LIMIT) || 1;
    let html = "";

    for(let i=0; i<totalPages; i++) {
        let segment = fullList.slice(i * LIMIT, (i + 1) * LIMIT);
        
        if (segment.length > 0 && segment[0].type === 'row' && i > 0) {
            segment.unshift({ type: 'header', label: segment[0].label + " (FORTSETZUNG)" });
            if (segment.length > LIMIT) segment.pop();
        }

        const rowsHtml = segment.map(item => {
            if(item.type === 'header') {
                return `<tr style="height:32px;">
                    <td colspan="2" style="background:#000;color:#fff;padding:0 12px;font-weight:900;font-size:13px;text-transform:uppercase;-webkit-print-color-adjust:exact;vertical-align:middle;">
                        ${item.label}
                    </td>
                </tr>`;
            } else {
                const t = item.data;
                return `<tr style="height:42px;">
                    <td style="border-bottom:1.5px solid #000;padding:0 10px;vertical-align:middle;">
                        <div style="font-weight:900;font-size:14px;text-transform:uppercase;line-height:1.1;white-space:nowrap;overflow:hidden;">${t.nm}</div>
                        <div style="font-size:9px;font-weight:800;color:#333;text-transform:uppercase;">${t.isStandart ? t.loc : 'IN BLAUKISTE'}${t.kom ? ' | ' + t.kom : ''}</div>
                    </td>
                    <td style="border-bottom:1.5px solid #000;text-align:right;padding-right:10px;vertical-align:middle;width:80px;">
                        <div style="font-size:7px;font-weight:900;color:#666;line-height:1;">BEMERKUNG</div>
                        <div style="font-weight:900;font-size:13px;line-height:1;">${t.bem || '--'}</div>
                    </td>
                </tr>`;
            }
        }).join('');

        html += `
        <div style="width:210mm;height:297mm;padding:8mm 10mm;box-sizing:border-box;page-break-after:always;overflow:hidden;">
            <div style="border:1.5px solid #000;height:100%;display:flex;flex-direction:column;box-sizing:border-box;">
                <div style="padding:10px 20px;border-bottom:3.5px solid #000;flex-shrink:0;">
                    <div style="font-size:12px;font-weight:900;color:#666;text-transform:uppercase;">${p.name || ''}</div>
                    <div style="font-size:54px;font-weight:900;line-height:0.8;margin:5px 0;letter-spacing:-2px;">${p.num || '---'}</div>
                    <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;font-size:10.5px;font-weight:900;border-top:1.5px solid #000;padding-top:8px;margin-top:5px;">
                        <div>MAT: ${p.mat || '--'}</div><div>LZF: ${p.lzf || '--'}</div>
                        <div>SÄGE: ${p.slg || '--'}</div><div>ABST: ${p.abs || '--'}</div>
                        <div>BACKEN: ${p.grf || '--'}</div><div>STÜCK: ${p.stk || '--'}</div>
                    </div>
                </div>
                <table style="width:100%;border-collapse:collapse;table-layout:fixed;flex-grow:1;">
                    <colgroup><col><col style="width:100px;"></colgroup>
                    <tbody>${rowsHtml}</tbody>
                </table>
            </div>
        </div>`;
    }

    const win = window.open('','_blank');
    win.document.write(`<html><head><style>
        @page { size: A4; margin: 0; }
        body { margin: 0; padding: 0; font-family: -apple-system, sans-serif; -webkit-print-color-adjust: exact; }
        * { box-sizing: border-box; }
    </style></head><body>${html}</body></html>`);
    win.document.close();
    setTimeout(() => { win.print(); win.close(); }, 800);
}


window.onload = renderList;
