const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'flex'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

// --- PROJEKTE ---
function modalP(edit = false) {
    if (!edit) currentIdx = null; 
    const p = (edit && currentIdx !== null) ? db[currentIdx] : {num:'', name:'', lzf:'', sag:'', stt:'', stn:'', abs:'', grf:''};
    
    el('p-idx').value = edit ? currentIdx : '';
    el('p-num').value = p.num || '';
    el('p-nam').value = p.name || '';
    el('p-lzf').value = p.lzf || '';
    el('p-sag').value = p.sag || '';
    el('p-stt').value = p.stt || '';
    el('p-stn').value = p.stn || '';
    el('p-abs').value = p.abs || '';
    el('p-grf').value = p.grf || '';
    
    show('m-p');
}

function saveP() {
    const idx = el('p-idx').value;
    const newP = {
        num: el('p-num').value,
        name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value,
        sag: el('p-sag').value,
        stt: el('p-stt').value,
        stn: el('p-stn').value,
        abs: el('p-abs').value,
        grf: el('p-grf').value,
        tools: (idx !== '' && db[idx]) ? db[idx].tools : []
    };

    if (idx === '') {
        db.push(newP);
        currentIdx = db.length - 1;
    } else {
        db[idx] = newP;
    }
    
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    hide('m-p');
    renderList();
    if(idx !== '') openProject(idx);
    else goHome();
}

function renderList() {
    const list = el('list-p');
    if(!list) return;
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})">
            <div><small>${p.name || 'UNBENANNT'}</small><b>${p.num || '---'}</b></div>
            <div style="color:var(--danger); font-weight:800; padding:10px;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>
    `).join('');
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

// --- TOOLS ---
function renderTools() {
    const listT = el('list-t');
    if(!listT || currentIdx === null) return;
    listT.innerHTML = (db[currentIdx].tools || []).map((t, i) => `
        <div class="list-item" onclick="modalT(${i})">
            <div style="flex:1"><b>${t.id}</b><small>${t.nm}</small></div>
            <div style="font-weight:900; color:var(--accent)">${t.dia}</div>
        </div>
    `).join('');
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:''};
    el('t-id').value = t.id;
    el('t-nm').value = t.nm;
    el('t-dia').value = t.dia;
    el('btn-del-t').style.display = edit ? 'block' : 'none';
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const t = { id: el('t-id').value.toUpperCase(), nm: el('t-nm').value.toUpperCase(), dia: el('t-dia').value };
    if(i === '') db[currentIdx].tools.push(t);
    else db[currentIdx].tools[i] = t;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools();
    hide('m-t');
}

function runImp() {
    const text = el('imp-area').value;
    if (!text.trim()) return;
    const regex = /(T[0O]\d{2,4})/gi;
    const parts = text.split(regex);
    for (let i = 1; i < parts.length; i += 2) {
        let id = parts[i].trim().toUpperCase().replace('O', '0');
        let name = (parts[i + 1] || '').trim().replace(/[\r\n]+/g, ' ').replace(/\s\s+/g, ' ');
        db[currentIdx].tools.push({ id: id, nm: name.toUpperCase(), dia: '' });
    }
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    el('imp-area').value = '';
    renderTools();
    hide('m-imp');
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { const i = el('t-idx').value; db[currentIdx].tools.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }

// --- FIXED PDF ---
function makePDF() {
    const p = db[currentIdx];
    const rows = p.tools.map(t => `
        <div style="display: flex; align-items: baseline; border-bottom: 1px solid #eee; padding: 12px 0;">
            <div style="width: 75px; font-weight: 800; font-size: 15px; color: #000; font-family: sans-serif;">${t.id}</div>
            <div style="flex: 1; font-weight: 700; font-size: 15px; text-transform: uppercase; color: #000; font-family: sans-serif;">${t.nm}</div>
            <div style="width: 100px; text-align: right; font-weight: 800; font-size: 15px; color: #000; font-family: sans-serif;">${t.dia}</div>
        </div>
    `).join('');

    const html = `
    <div style="width: 210mm; padding: 12mm; color: #000; box-sizing: border-box; background: #fff;">
        <div style="border: 2px solid #000; padding: 25px; min-height: 265mm; display: flex; flex-direction: column; box-sizing: border-box;">
            <div style="display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 5px;">
                <div style="margin-bottom: -10px;">
                    <div style="font-size: 14px; font-weight: 900; font-family: sans-serif; text-transform: uppercase; margin-bottom: 2px;">${p.name}</div>
                    <div style="font-size: 72px; font-weight: 900; font-family: sans-serif; line-height: 0.8; letter-spacing: -3px;">${p.num}</div>
                </div>
                <div style="width: 195px; font-size: 11px; font-weight: 800; font-family: sans-serif; line-height: 1.6;">
                    <div style="display:flex; justify-content:space-between;"><span>ABSTAND</span><span>${p.abs}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>GREIFBACKEN</span><span>${p.grf}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>LAUFZEIT</span><span>${p.lzf}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>SÄGELÄNGE</span><span>${p.sag}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>STÜCK T</span><span>${p.stt}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>STÜCK N</span><span>${p.stn}</span></div>
                </div>
            </div>
            <div style="border-bottom: 5px solid #000; margin-bottom: 18px;"></div>
            <div style="display: flex; font-size: 10px; font-weight: 900; font-family: sans-serif; text-transform: uppercase; margin-bottom: 5px; padding: 0 2px;">
                <div style="width: 75px;">T-NR</div>
                <div style="flex: 1;">WERKZEUGNAME / KOMMENTAR</div>
                <div style="width: 100px; text-align: right;">Ø / TOLERANZ</div>
            </div>
            <div style="border-bottom: 2px solid #eee; margin-bottom: 0px;"></div>
            <div style="flex: 1;">${rows}</div>
        </div>
    </div>`;

    const container = el('print-container');
    if(container) { 
        container.innerHTML = html; 
        setTimeout(() => { window.print(); }, 150); 
    }
}

window.onload = renderList;
