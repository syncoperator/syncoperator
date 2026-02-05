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
            <div style="color:var(--danger); font-weight:900; padding:15px;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
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

// --- ИНСТРУМЕНТЫ ---
function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'list-item';
        item.draggable = true;
        
        const revLabel = t.rev ? `<div style="background:#000; color:#fff; font-size:10px; padding:2px 6px; border-radius:4px; margin-bottom:5px; width:fit-content; font-weight:900;">REVOLVER UNTEN ↓</div>` : '';

        item.innerHTML = `
            <div style="flex:1; min-width:0;" onclick="modalT(${i})">
                ${revLabel}
                <small style="color:#8e8e93; font-weight:700; font-size:11px;">${t.id || 'T0000'}</small>
                <b style="font-size:20px; font-weight:900; display:block; line-height:1.2;">${t.nm || '---'}</b>
            </div>
        `;
        
        item.ondragstart = (e) => { e.dataTransfer.setData('text/plain', i); item.style.opacity = '0.4'; };
        item.ondragend = () => { item.style.opacity = '1'; };
        item.ondragover = (e) => e.preventDefault();
        item.ondrop = (e) => {
            e.preventDefault();
            const from = e.dataTransfer.getData('text/plain');
            moveTool(parseInt(from), i);
        };
        list.appendChild(item);
    });
    list.innerHTML += '<div style="height:180px;"></div>'; 
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
    el('t-idx').value = (i === null) ? '' : i;
    
    // Создаем дефолтный объект, если i === null
    let t = {id:'', nm:'', dia:'', rev:false};
    if (edit && db[currentIdx].tools[i]) {
        t = db[currentIdx].tools[i];
    }
    
    el('t-id').value = t.id || ''; 
    el('t-nm').value = t.nm || ''; 
    el('t-dia').value = t.dia || '';
    
    // Работа с кнопкой револьвера (без лишних пересозданий)
    let btnRev = el('btn-rev-toggle');
    if(!btnRev) {
        btnRev = document.createElement('div');
        btnRev.id = 'btn-rev-toggle';
        const content = el('m-t').querySelector('.modal-content');
        content.insertBefore(btnRev, el('btn-save-t'));
    }
    
    btnRev.style.cssText = `margin: 10px 0; padding: 15px; border-radius: 12px; text-align: center; font-weight: 900; cursor: pointer;`;
    
    const updateBtn = (state) => {
        btnRev.dataset.state = state;
        btnRev.innerText = state ? "✓ UNTEN (START)" : "SET AS UNTEN START";
        btnRev.style.background = state ? "#000" : "#f0f0f0";
        btnRev.style.color = state ? "#fff" : "#000";
    };

    updateBtn(t.rev === true);
    btnRev.onclick = () => updateBtn(btnRev.dataset.state === 'false');

    el('btn-del-t').style.display = edit ? 'block' : 'none';
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const btn = el('btn-rev-toggle');
    const isUnten = btn ? btn.dataset.state === 'true' : false;

    const t = { 
        id: el('t-id').value.toUpperCase(), 
        nm: el('t-nm').value.toUpperCase(), 
        dia: el('t-dia').value,
        rev: isUnten 
    };

    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    
    if(i === '') { 
        db[currentIdx].tools.push(t); 
    } else { 
        db[currentIdx].tools[parseInt(i)] = t; 
    }

    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); 
    hide('m-t');
}

// --- PDF ---
function makePDF() {
    const p = db[currentIdx];
    let html = '';
    
    const getHeader = (isPage2 = false) => `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <div>
                <div style="font-size:12px; font-weight:900; color:#666; text-transform:uppercase;">${p.name || ''} ${isPage2 ? '(SEITE 2)' : ''}</div>
                <div style="font-size:50px; font-weight:900; line-height:0.8;">${p.num || '---'}</div>
            </div>
            ${!isPage2 ? `
            <div style="width:200px; font-size:10px; font-weight:800; line-height:1.4;">
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>ABSTAND</span><span>${p.abs || ''}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div>
                <div style="display:flex; justify-content:space-between;"><span>STÜCK T</span><span>${p.stt || ''}</span></div>
            </div>` : ''}
        </div>
        <div style="border-bottom:4px solid #000; margin-bottom:10px;"></div>`;

    const getRow = (t) => {
        const isLong = (t.dia || '').length > 15;
        const displayDia = t.dia.includes('/') ? t.dia.split('/').join('<br>') : t.dia;
        return `
        <div style="display:flex; align-items:${isLong ? 'center' : 'baseline'}; border-bottom:1px solid #eee; padding:8px 0;">
            <div style="width:70px; font-weight:800; font-size:14px;">${t.id}</div>
            <div style="flex:1; font-weight:700; font-size:14px; text-transform:uppercase; white-space:pre-wrap;">${t.nm}</div>
            <div style="width:120px; text-align:right; font-weight:800; font-size:13px; line-height:1.1;">${displayDia}</div>
        </div>`;
    };

    let oben = [], unten = [], target = oben;
    (p.tools || []).forEach(t => {
        if(t.rev) target = unten;
        target.push(t);
    });

    html += `<div style="width:210mm; padding:10mm; background:#fff; font-family:sans-serif;">
                <div style="border:2px solid #000; padding:20px; min-height:270mm;">
                    ${getHeader()}
                    <div style="background:#000; color:#fff; padding:5px 10px; font-weight:900; font-size:13px;">REVOLVER OBEN</div>
                    ${oben.map(t => getRow(t)).join('')}
                    ${unten.length > 0 && (oben.length + unten.length < 22) ? `<div style="background:#000; color:#fff; padding:5px 10px; font-weight:900; font-size:13px; margin-top:10px;">REVOLVER UNTEN</div>` + unten.map(t => getRow(t)).join('') : ''}
                </div>
            </div>`;

    if(unten.length > 0 && (oben.length + unten.length >= 22)) {
        html += `<div style="width:210mm; padding:10mm; background:#fff; font-family:sans-serif; page-break-before:always;">
                    <div style="border:2px solid #000; padding:20px; min-height:270mm;">
                        ${getHeader(true)}
                        <div style="background:#000; color:#fff; padding:5px 10px; font-weight:900; font-size:13px;">REVOLVER UNTEN</div>
                        ${unten.map(t => getRow(t)).join('')}
                    </div>
                </div>`;
    }

    el('print-container').innerHTML = html;
    setTimeout(() => { window.print(); }, 200);
}

// --- СЕРВИС ---
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

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { const i = el('t-idx').value; db[currentIdx].tools.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }
function exportJSON() { el('imp-area').value = JSON.stringify(db); alert("JSON Kopiert!"); }
function importJSON() { try { const p = JSON.parse(el('imp-area').value); if(Array.isArray(p)) { db = p; localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); } } catch(e){alert("Error");} }

window.onload = renderList;
