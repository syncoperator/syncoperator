const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'flex'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

// --- PROJEKTE ---
function modalP(edit = false) {
    if (!edit) currentIdx = null; 
    const p = (edit && currentIdx !== null && db[currentIdx]) ? db[currentIdx] : {num:'', name:'', lzf:'', sag:'', stt:'', stn:'', abs:'', grf:'', mat:''};
    el('p-idx').value = edit ? currentIdx : '';
    el('p-num').value = p.num || '';
    el('p-nam').value = p.name || '';
    el('p-lzf').value = p.lzf || '';
    el('p-sag').value = p.sag || '';
    el('p-stt').value = p.stt || '';
    el('p-stn').value = p.stn || '';
    el('p-abs').value = p.abs || '';
    el('p-grf').value = p.grf || '';
    if(el('p-mat')) el('p-mat').value = p.mat || ''; 
    show('m-p');
}

function saveP() {
    const idx = el('p-idx').value;
    const newP = {
        num: el('p-num').value,
        name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value, sag: el('p-sag').value,
        stt: el('p-stt').value, stn: el('p-stn').value,
        abs: el('p-abs').value, grf: el('p-grf').value,
        mat: el('p-mat') ? el('p-mat').value.toUpperCase() : '',
        tools: (idx !== '' && db[idx]) ? (db[idx].tools || []) : []
    };
    if (idx === '') { db.push(newP); currentIdx = db.length - 1; } 
    else { db[idx] = newP; }
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    hide('m-p');
    renderList();
    if(idx !== '') openProject(idx); else goHome();
}

function renderList() {
    const list = el('list-p'); if(!list) return;
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})">
            <div class="item-icon">📄</div>
            <div class="item-content" style="flex:1;">
                <small>${p.name || '---'}</small>
                <b>${p.num || '---'}</b>
            </div>
            <div style="color:var(--danger); font-size:20px; padding:10px;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:100px"></div>';
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

// --- WERKZEUGE (TOUCH DRAG & DROP) ---
let startIdx = null;

function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'list-item';
        item.setAttribute('data-idx', i);
        
        const revMark = t.rev ? `<div style="background:#000; color:#fff; font-size:9px; padding:2px 6px; border-radius:4px; margin-bottom:5px; font-weight:900;">UNTEN</div>` : '';

        item.innerHTML = `
            <div class="handle" style="cursor:grab; color:#ccc; font-size:20px; padding-right:10px; touch-action:none;">☰</div>
            <div class="item-icon">🛠</div>
            <div class="item-content" style="flex:1;" onclick="modalT(${i})">
                ${revMark}
                <small>${t.id || 'T0000'}</small>
                <b>${t.nm || '---'}</b>
            </div>
        `;
        
        const handle = item.querySelector('.handle');
        handle.ontouchstart = () => { startIdx = i; item.style.background = "#f9f9f9"; };
        handle.ontouchmove = (e) => {
            e.preventDefault();
            const touch = e.touches[0];
            const target = document.elementFromPoint(touch.clientX, touch.clientY);
            const targetItem = target?.closest('.list-item');
            if (targetItem) {
                const overIdx = parseInt(targetItem.getAttribute('data-idx'));
                if (overIdx !== startIdx) {
                    moveTool(startIdx, overIdx);
                    startIdx = overIdx;
                }
            }
        };
        handle.ontouchend = () => { item.style.background = ""; renderTools(); };

        list.appendChild(item);
    });
    list.innerHTML += '<div style="height:150px; pointer-events:none;"></div>'; 
}

function moveTool(from, to) {
    const tools = db[currentIdx].tools;
    const item = tools.splice(from, 1)[0];
    tools.splice(to, 0, item);
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools();
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    const btn = el('btn-rev-toggle');
    if(t.rev) { btn.style.background = '#3b82f6'; btn.firstChild.style.left = '24px'; btn.classList.add('on'); }
    else { btn.style.background = '#ddd'; btn.firstChild.style.left = '2px'; btn.classList.remove('on'); }
    btn.onclick = () => {
        const isOn = btn.classList.toggle('on');
        btn.style.background = isOn ? '#3b82f6' : '#ddd';
        btn.firstChild.style.left = isOn ? '24px' : '2px';
    };
    el('btn-del-t').style.display = edit ? 'block' : 'none';
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const isOn = el('btn-rev-toggle').classList.contains('on');
    const t = { id: el('t-id').value.toUpperCase(), nm: el('t-nm').value.toUpperCase(), dia: el('t-dia').value, rev: isOn };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); hide('m-t');
}

// --- IMPORT / EXPORT ---
function runImp() {
    const text = el('imp-area').value; if (!text.trim()) return;
    const regex = /(T[0O]\d{2,4})/gi; const parts = text.split(regex);
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    for (let i = 1; i < parts.length; i += 2) {
        let id = parts[i].trim().toUpperCase().replace('O', '0');
        let name = (parts[i + 1] || '').trim().replace(/[\r\n]+/g, ' ').replace(/\s\s+/g, ' ');
        db[currentIdx].tools.push({ id, nm: name.toUpperCase(), dia: '', rev: false });
    }
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    el('imp-area').value = ''; renderTools(); hide('m-imp');
}
function exportJSON() { el('imp-area').value = JSON.stringify(db); el('imp-area').select(); alert("JSON kopiert!"); }
function importJSON() { try { const parsed = JSON.parse(el('imp-area').value); if(Array.isArray(parsed)) { db = parsed; localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); } } catch(e) { alert("JSON-Fehler"); } }
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { const i = el('t-idx').value; db[currentIdx].tools.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }

// --- PDF ---
function makePDF() {
    const p = db[currentIdx];
    const head = `<div style="display:flex; font-size:10px; font-weight:900; text-transform:uppercase; margin-bottom:6px;">
        <div style="width:75px;">T-NR</div><div style="flex:1;">WERKZEUGNAME / KOMMENTAR</div><div style="width:125px; text-align:right;">Ø / TOLERANZ</div>
    </div><div style="border-bottom:4px solid #000; margin-bottom:0px;"></div>`;
    const getRow = (t) => `<div style="display:flex; border-bottom:1.5px solid #000; padding:10px 0; font-weight:800;">
        <div style="width:75px;">${t.id}</div><div style="flex:1;">${t.nm}</div><div style="width:125px; text-align:right;">${t.dia}</div>
    </div>`;
    let o = [], u = [], trg = o; (p.tools || []).forEach(t => { if(t.rev) trg = u; trg.push(t); });
    const html = `
    <div style="width:210mm; padding:12mm; font-family:sans-serif; background:white;">
        <div style="border:2px solid #000; padding:25px; min-height:265mm;">
            <div style="display:flex; justify-content:space-between; margin-bottom:15px;">
                <div><div style="font-size:13px; font-weight:900; color:#666;">${p.name || ''}</div><div style="font-size:64px; font-weight:900; letter-spacing:-2px;">${p.num || '---'}</div></div>
                <div style="width:220px; font-size:11px; font-weight:800; line-height:1.5;">
                    <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>MATERIAL</span><span>${p.mat || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>SÄGELÄНGE</span><span>${p.sag || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>ABSTAND</span><span>${p.abs || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>GREIFBACKEN</span><span>${p.grf || ''}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>STÜCKZAHL</span><span>${p.stt || ''}/${p.stn || ''}</span></div>
                </div>
            </div>
            <div style="border-bottom:5px solid #000; margin-bottom:15px;"></div>
            <div style="margin-bottom:5px; font-size:18px; font-weight:900;">REVOLVER OBEN</div>${head}${o.map(getRow).join('')}
            ${u.length ? `<div style="margin-top:35px; font-size:18px; font-weight:900;">REVOLVER UNTEN</div>${head}${u.map(getRow).join('')}` : ''}
        </div>
    </div>`;
    el('print-container').innerHTML = html; window.print();
}

window.onload = renderList;
