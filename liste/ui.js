const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'flex'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

// --- ПРОЕКТЫ ---
function modalP(edit = false) {
    if (!edit) currentIdx = null; 
    const p = (edit && currentIdx !== null && db[currentIdx]) ? db[currentIdx] : {num:'', name:'', lzf:'', sag:'', stt:'', stn:'', abs:'', grf:''};
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
        lzf: el('p-lzf').value, sag: el('p-sag').value,
        stt: el('p-stt').value, stn: el('p-stn').value,
        abs: el('p-abs').value, grf: el('p-grf').value,
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
            <div><small>${p.name || '---'}</small><b>${p.num || '---'}</b></div>
            <div style="color:var(--danger); font-weight:900; padding:15px; z-index:20;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
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

// --- ИНСТРУМЕНТЫ (ИНВЕРСИЯ: ИМЯ СЛЕВА, НОМЕР СПРАВА) ---
function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'list-item';
        item.draggable = true;
        item.dataset.index = i;
        // ТУТ ИЗМЕНЕНО: nm сверху, id (T-NR) крупно снизу
        item.innerHTML = `
            <div style="flex:1" onclick="modalT(${i})">
                <small>${t.nm || 'BEZEICHNUNG'}</small>
                <b style="font-size:24px;">${t.id}</b>
            </div>
            <div style="font-weight:900; color:var(--accent); font-size:20px;">${t.dia}</div>
        `;
        
        item.ondragstart = (e) => { e.dataTransfer.setData('text/plain', i); item.style.opacity = '0.4'; };
        item.ondragend = () => { item.style.opacity = '1'; renderTools(); };
        item.ondragover = (e) => e.preventDefault();
        item.ondrop = (e) => {
            e.preventDefault();
            const from = e.dataTransfer.getData('text/plain');
            moveTool(parseInt(from), i);
        };
        list.appendChild(item);
    });
    list.innerHTML += '<div style="height:120px"></div>';
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
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:''};
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    el('btn-del-t').style.display = edit ? 'block' : 'none';
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const t = { id: el('t-id').value.toUpperCase(), nm: el('t-nm').value.toUpperCase(), dia: el('t-dia').value };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); hide('m-t');
}

// --- ИМПОРТ / ЭКСПОРТ ---
function runImp() {
    const text = el('imp-area').value; if (!text.trim()) return;
    const regex = /(T[0O]\d{2,4})/gi; const parts = text.split(regex);
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    for (let i = 1; i < parts.length; i += 2) {
        let id = parts[i].trim().toUpperCase().replace('O', '0');
        let name = (parts[i + 1] || '').trim().replace(/[\r\n]+/g, ' ').replace(/\s\s+/g, ' ');
        db[currentIdx].tools.push({ id, nm: name.toUpperCase(), dia: '' });
    }
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    el('imp-area').value = ''; renderTools(); hide('m-imp');
}

function exportJSON() {
    el('imp-area').value = JSON.stringify(db);
    el('imp-area').select();
    alert("JSON kopiert! Speichere den Text aus dem Feld.");
}

function importJSON() {
    try {
        const data = JSON.parse(el('imp-area').value);
        if(Array.isArray(data)) {
            db = data;
            localStorage.setItem(DB_KEY, JSON.stringify(db));
            renderList();
            alert("Erfolgreich importiert!");
            hide('m-imp');
        }
    } catch(e) { alert("Fehler: Ungültiges JSON"); }
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { const i = el('t-idx').value; db[currentIdx].tools.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }

// --- PDF ---
function makePDF() {
    const p = db[currentIdx];
    const rows = (p.tools || []).map(t => `
        <div style="display:flex; align-items:baseline; border-bottom:1px solid #eee; padding:10px 0;">
            <div style="width:75px; font-weight:800; font-size:15px; font-family:sans-serif;">${t.id}</div>
            <div style="flex:1; font-weight:700; font-size:15px; text-transform:uppercase; font-family:sans-serif;">${t.nm}</div>
            <div style="width:100px; text-align:right; font-weight:800; font-size:15px; font-family:sans-serif;">${t.dia}</div>
        </div>`).join('');

    const html = `
    <div style="width:210mm; padding:12mm; box-sizing:border-box; background:#fff; font-family:sans-serif; color:#000;">
        <div style="border:2px solid #000; padding:25px; min-height:265mm; display:flex; flex-direction:column; box-sizing:border-box;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; min-height:90px;">
                <div style="display:flex; flex-direction:column; justify-content:center;">
                    <div style="font-size:13px; font-weight:900; text-transform:uppercase; color:#666; margin-bottom:2px; line-height:1;">${p.name || ''}</div>
                    <div style="font-size:64px; font-weight:900; line-height:0.8; letter-spacing:-2px; margin:0;">${p.num || '---'}</div>
                </div>
                <div style="width:220px; font-size:11px; font-weight:800; line-height:1.5; display:flex; flex-direction:column; justify-content:center;">
                    <div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>ABSTAND</span><span>${p.abs || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>GREIFBACKEN</span><span>${p.grf || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>SÄGELÄNGE</span><span>${p.sag || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>STÜCK T</span><span>${p.stt || ''}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>STÜCK N</span><span>${p.stn || ''}</span></div>
                </div>
            </div>
            <div style="border-bottom:5px solid #000; margin-bottom:15px;"></div>
            <div style="display:flex; font-size:10px; font-weight:900; text-transform:uppercase; margin-bottom:6px; padding:0 2px;">
                <div style="width:75px;">T-NR</div>
                <div style="flex:1;">WERKZEUGNAME / KOMMENTAR</div>
                <div style="width:100px; text-align:right;">Ø / TOLERANZ</div>
            </div>
            <div style="border-bottom:3px solid #000; margin-bottom:0px;"></div>
            <div style="flex:1;">${rows}</div>
        </div>
    </div>`;
    el('print-container').innerHTML = html;
    setTimeout(() => { window.print(); }, 150);
}

window.onload = renderList;
