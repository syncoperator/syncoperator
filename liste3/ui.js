const DB_KEY = 'QS_DATA_ELITE';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'flex'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

function renderList() {
    const list = el('list-p'); if(!list) return;
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})">
            <div class="indicator"></div>
            <div class="item-info" style="flex:1;">
                <small>${p.name || 'PROJECT'}</small>
                <b>${p.num || '---'}</b>
            </div>
            <div style="color:#ff5f5f; font-weight:900; padding:10px;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('');
}

function openProject(i) {
    currentIdx = i;
    el('v-home').classList.remove('active');
    el('v-det').classList.add('active');
    el('h-num').innerText = db[i].num;
    el('h-nam').innerText = db[i].name;
    renderTools();
}

function goHome() { currentIdx = null; el('v-home').classList.add('active'); el('v-det').classList.remove('active'); renderList(); }

function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'list-item';
        item.setAttribute('data-idx', i);
        item.innerHTML = `
            <div class="handle" style="cursor:grab; color:#c1c9d2; margin-right:15px; font-size:20px; touch-action:none;">☰</div>
            <div class="item-info" style="flex:1;" onclick="modalT(${i})">
                <small>${t.id || 'T0000'}</small>
                <b>${t.nm || '---'}</b>
            </div>
            <div style="font-weight:800; color:var(--primary);">${t.dia || ''}</div>
        `;
        const h = item.querySelector('.handle');
        h.ontouchstart = () => { startIdx = i; item.style.boxShadow = 'var(--shadow-inner)'; };
        h.ontouchmove = (e) => {
            e.preventDefault();
            const touch = e.touches[0];
            const target = document.elementFromPoint(touch.clientX, touch.clientY)?.closest('.list-item');
            if (target) {
                const overIdx = parseInt(target.getAttribute('data-idx'));
                if (overIdx !== startIdx) { moveTool(startIdx, overIdx); startIdx = overIdx; }
            }
        };
        h.ontouchend = () => { renderTools(); };
        list.appendChild(item);
    });
}

function moveTool(from, to) {
    const tools = db[currentIdx].tools;
    const item = tools.splice(from, 1)[0];
    tools.splice(to, 0, item);
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools();
}

function modalP(edit = false) {
    if (!edit) currentIdx = null;
    const p = (edit && currentIdx !== null) ? db[currentIdx] : {num:'', name:'', lzf:'', sag:'', stt:'', stn:'', abs:'', grf:'', mat:''};
    el('p-num').value = p.num; el('p-nam').value = p.name; el('p-mat').value = p.mat || '';
    el('p-lzf').value = p.lzf; el('p-sag').value = p.sag;
    show('m-p');
}

function saveP() {
    const newP = {
        num: el('p-num').value, name: el('p-nam').value.toUpperCase(),
        mat: el('p-mat').value.toUpperCase(), lzf: el('p-lzf').value, sag: el('p-sag').value,
        tools: currentIdx !== null ? (db[currentIdx].tools || []) : []
    };
    if (currentIdx === null) db.push(newP); else db[currentIdx] = newP;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    hide('m-p'); goHome();
}

function modalT(i = null) {
    const edit = i !== null; el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    const toggle = el('t-rev-toggle');
    if(t.rev) toggle.classList.add('on'); else toggle.classList.remove('on');
    toggle.firstChild.style.left = t.rev ? '32px' : '4px';
    toggle.firstChild.style.background = t.rev ? 'var(--primary)' : '#c1c9d2';
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const t = { 
        id: el('t-id').value.toUpperCase(), nm: el('t-nm').value.toUpperCase(), 
        dia: el('t-dia').value, rev: el('t-rev-toggle').classList.contains('on') 
    };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); hide('m-t');
}

function makePDF() {
    const p = db[currentIdx];
    const head = `<div style="display:flex; font-size:10px; font-weight:900; margin-bottom:5px;">
        <div style="width:75px;">T-NR</div><div style="flex:1;">WERKZEUGNAME</div><div style="width:125px; text-align:right;">Ø / TOLERANZ</div>
    </div><div style="border-bottom:4px solid #000;"></div>`;
    const row = (t) => `<div style="display:flex; border-bottom:1.5px solid #000; padding:10px 0; font-weight:700; font-size:15px;">
        <div style="width:75px;">${t.id}</div><div style="flex:1;">${t.nm}</div><div style="width:125px; text-align:right;">${t.dia}</div>
    </div>`;
    let o = [], u = []; (p.tools || []).forEach(t => t.rev ? u.push(t) : o.push(t));
    const html = `<div style="width:210mm; padding:15mm; background:#fff; font-family:sans-serif;">
        <div style="border:2px solid #000; padding:25px; min-height:260mm;">
            <div style="display:flex; justify-content:space-between; margin-bottom:20px;">
                <div><div style="font-size:14px; font-weight:900;">${p.name}</div><div style="font-size:60px; font-weight:900;">${p.num}</div></div>
                <div style="width:200px; font-size:12px; font-weight:800; line-height:1.8;">
                    <div style="border-bottom:1px solid #eee;">MAT: ${p.mat}</div>
                    <div style="border-bottom:1px solid #eee;">LZ: ${p.lzf}</div>
                </div>
            </div>
            <div style="border-bottom:5px solid #000; margin-bottom:20px;"></div>
            <div style="font-size:18px; font-weight:900;">OBEN</div>${head}${o.map(row).join('')}
            ${u.length ? `<div style="margin-top:30px; font-size:18px; font-weight:900;">UNTEN</div>${head}${u.map(row).join('')}` : ''}
        </div>
    </div>`;
    el('print-container').innerHTML = html; window.print();
}

function runImp() {
    const text = el('imp-area').value; const regex = /(T[0O]\d{2,4})/gi; const parts = text.split(regex);
    for (let i = 1; i < parts.length; i += 2) {
        db[currentIdx].tools.push({ id: parts[i].trim().toUpperCase().replace('O', '0'), nm: parts[i+1].trim().toUpperCase(), dia: '', rev: false });
    }
    localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-imp');
}
function exportJSON() { el('imp-area').value = JSON.stringify(db); }
function importJSON() { db = JSON.parse(el('imp-area').value); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); }
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i,1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { db[currentIdx].tools.splice(el('t-idx').value, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }

window.onload = renderList;
