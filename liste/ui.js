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
            <div style="color:var(--danger); font-weight:800; padding:15px;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>
    `).join('');
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

function goHome() {
    currentIdx = null;
    el('v-home').classList.add('active');
    el('v-det').classList.remove('active');
    renderList();
}

// --- ИНСТРУМЕНТЫ ---
function renderTools() {
    const listT = el('list-t');
    if(!listT || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    listT.innerHTML = tools.map((t, i) => `
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

function deleteProject(i) {
    // В премиум-дизайне заменим стандартный confirm позже, пока фикс
    if(confirm('Löschen?')) {
        db.splice(i, 1);
        localStorage.setItem(DB_KEY, JSON.stringify(db));
        renderList();
    }
}

function delT() {
    const i = el('t-idx').value;
    db[currentIdx].tools.splice(i, 1);
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools();
    hide('m-t');
}

// PDF функция остается без изменений (как в прошлом сообщении)
function makePDF() { ... } 

window.onload = renderList;
