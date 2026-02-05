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

// --- ИНСТРУМЕНТЫ (С МАРКЕРОМ РЕВОЛЬВЕРА) ---
function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'list-item';
        item.draggable = true;
        item.style.padding = '12px 15px';
        
        // Визуальная метка начала нижнего револьвера в приложении
        const revLabel = t.rev ? `<div style="background:#000; color:#fff; font-size:10px; padding:2px 6px; border-radius:4px; margin-bottom:5px; width:fit-content;">UNTEN</div>` : '';

        item.innerHTML = `
            <div style="flex:1; min-width:0;" onclick="modalT(${i})">
                ${revLabel}
                <small style="color:#8e8e93; font-weight:700; font-size:11px;">${t.id || 'T0000'}</small>
                <b style="font-size:20px; font-weight:900; display:block; line-height:1.2;">${t.nm || '---'}</b>
            </div>
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
    el('t-id').value = t.id; 
    el('t-nm').value = t.nm; 
    el('t-dia').value = t.dia;
    // Добавь в HTML модалки <input type="checkbox" id="t-rev"> если хочешь, 
    // либо мы будем использовать поле t-idx для скрытой логики. 
    // Для простоты я добавлю проверку в saveT
    el('btn-del-t').style.display = edit ? 'block' : 'none';
    show('m-t');
    
    // Временная кнопка переключения револьвера прямо в модалке (если нет чекбокса)
    if(edit) {
        const btnRev = document.createElement('button');
        btnRev.id = 'temp-rev-btn';
        btnRev.innerText = t.rev ? "✓ UNTEN" : "SET UNTEN";
        btnRev.className = t.rev ? "btn-sec" : "btn-main";
        btnRev.style.marginTop = "10px";
        btnRev.onclick = () => { t.rev = !t.rev; btnRev.innerText = t.rev ? "✓ UNTEN" : "SET UNTEN"; };
        el('m-t').querySelector('.modal-content').appendChild(btnRev);
    }
}

// При закрытии модалки удаляем временную кнопку
function hideMT() { 
    hide('m-t'); 
    const b = el('temp-rev-btn'); 
    if(b) b.remove(); 
}

function saveT() {
    const i = el('t-idx').value;
    const isUnten = el('temp-rev-btn') ? el('temp-rev-btn').innerText.includes('✓') : false;
    const t = { 
        id: el('t-id').value.toUpperCase(), 
        nm: el('t-nm').value.toUpperCase(), 
        dia: el('t-dia').value,
        rev: isUnten 
    };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); hideMT();
}

// --- СЛУЖЕБНЫЕ ---
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { const i = el('t-idx').value; db[currentIdx].tools.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hideMT(); }
function exportJSON() { el('imp-area').value = JSON.stringify(db); el('imp-area').select(); alert("Kopiert!"); }
function importJSON() { try { const p = JSON.parse(el('imp-area').value); if(Array.isArray(p)) { db = p; localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); } } catch(e) { alert("Error"); } }

// --- PDF С РАЗДЕЛЕНИЕМ РЕВОЛЬВЕРОВ ---
function makePDF() {
    const p = db[currentIdx];
    let html = '';
    
    const header = (isNewPage = false) => `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; min-height:90px;">
            <div style="display:flex; flex-direction:column;">
                <div style="font-size:13px; font-weight:900; text-transform:uppercase; color:#666;">${p.name || ''} ${isNewPage ? '(SEITE 2)' : ''}</div>
                <div style="font-size:64px; font-weight:900; line-height:0.8; letter-spacing:-2px;">${p.num || '---'}</div>
            </div>
            ${!isNewPage ? `
            <div style="width:220px; font-size:11px; font-weight:800; line-height:1.5;">
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>ABSTAND</span><span>${p.abs || ''}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>GREIFBACKEN</span><span>${p.grf || ''}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>SÄGELÄNGE</span><span>${p.sag || ''}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>STÜCK T</span><span>${p.stt || ''}</span></div>
                <div style="display:flex; justify-content:space-between;"><span>STÜCK N</span><span>${p.stn || ''}</span></div>
            </div>` : ''}
        </div>
        <div style="border-bottom:5px solid #000; margin-bottom:10px;"></div>`;

    const row = (t) => {
        const isLong = (t.dia || '').length > 15;
        const align = isLong ? 'center' : 'baseline';
        const displayDia = t.dia.includes('/') ? t.dia.split('/').join('<br>') : t.dia;
        return `
        <div style="display:flex; align-items:${align}; border-bottom:1px solid #eee; padding:8px 0;">
            <div style="width:75px; font-weight:800; font-size:15px;">${t.id}</div>
            <div style="flex:1; font-weight:700; font-size:15px; text-transform:uppercase; padding-right:10px; white-space:pre-wrap;">${t.nm}</div>
            <div style="width:125px; text-align:right; font-weight:800; font-size:14px; line-height:1.2;">${displayDia}</div>
        </div>`;
    };

    const sectionTitle = (title) => `
        <div style="background:#000; color:#fff; padding:4px 10px; font-weight:900; font-size:14px; margin:15px 0 5px 0; display:flex; justify-content:space-between;">
            <span>${title}</span>
        </div>`;

    // Собираем контент
    let oben = [];
    let unten = [];
    let foundUnten = false;

    (p.tools || []).forEach(t => {
        if(t.rev) foundUnten = true;
        if(foundUnten) unten.push(t); else oben.push(t);
    });

    const pageStart = `<div style="width:210mm; padding:12mm; box-sizing:border-box; background:#fff; font-family:sans-serif; color:#000; page-break-after:always;">
        <div style="border:2px solid #000; padding:25px; min-height:265mm; display:flex; flex-direction:column; box-sizing:border-box;">`;
    const pageEnd = `</div></div>`;

    // Первая страница
    html += pageStart + header() + sectionTitle("REVOLVER OBEN") + oben.map(t => row(t)).join('');
    
    if (unten.length > 0) {
        // Если места мало (больше 22 инструментов всего), кидаем на вторую страницу
        const totalTools = (p.tools || []).length;
        const forceBreak = totalTools > 20; 

        if (forceBreak) {
            html += pageEnd + pageStart + header(true) + sectionTitle("REVOLVER UNTEN") + unten.map(t => row(t)).join('') + pageEnd;
        } else {
            html += sectionTitle("REVOLVER UNTEN") + unten.map(t => row(t)).join('') + pageEnd;
        }
    } else {
        html += pageEnd;
    }

    el('print-container').innerHTML = html;
    setTimeout(() => { window.print(); }, 150);
}

window.onload = renderList;
