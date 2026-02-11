const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'flex'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

// --- ПРОЕКТЫ ---
function modalP(edit = false) {
    if (!edit) currentIdx = null; 
    const p = (edit && currentIdx !== null) ? db[currentIdx] : {num:'', name:'', lzf:'', sag:'', stt:'', stn:'', abs:'', grf:'', mat:''};
    
    el('p-idx').value = edit ? currentIdx : '';
    el('p-num').value = p.num || '';
    el('p-nam').value = p.name || '';
    el('p-lzf').value = p.lzf || '';
    el('p-sag').value = p.sag || '';
    el('p-stt').value = p.stt || '';
    el('p-stn').value = p.stn || '';
    el('p-abs').value = p.abs || '';
    el('p-grf').value = p.grf || '';
    el('p-mat').value = p.mat || '';
    
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
        mat: el('p-mat').value.toUpperCase(),
        tools: (idx !== '' && db[idx]) ? (db[idx].tools || []) : []
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
    currentIdx = i;
    el('v-home').classList.remove('active');
    el('v-det').classList.add('active');
    el('h-num').innerText = db[i].num || '---';
    el('h-nam').innerText = db[i].name || '---';
    renderTools();
}

function goHome() { currentIdx = null; el('v-home').classList.add('active'); el('v-det').classList.remove('active'); renderList(); }

// --- DRAG & DROP ---
function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'list-item drag-item';
        item.setAttribute('data-idx', i);
        
        const revMark = t.rev ? `<div style="background:#000; color:#fff; font-size:9px; padding:2px 6px; border-radius:4px; margin-bottom:5px; font-weight:900; width:fit-content;">UNTEN START ↓</div>` : '';

        item.innerHTML = `
            <div class="handle" style="cursor:grab; color:#ccc; font-size:24px; padding:10px; touch-action:none;">☰</div>
            <div style="flex:1; min-width:0;" onclick="modalT(${i})">
                ${revMark}
                <small style="color:#8e8e93; font-weight:700; font-size:11px;">${t.id || 'T0000'}</small>
                <b style="font-size:20px; font-weight:900; display:block;">${t.nm || '---'}</b>
            </div>
            <div style="font-weight:800; color:var(--accent);">${t.dia || ''}</div>
        `;
        
        const handle = item.querySelector('.handle');
        handle.addEventListener('touchstart', () => { item.style.opacity = '0.5'; }, {passive: true});
        handle.addEventListener('touchmove', (e) => {
            e.preventDefault();
            const touch = e.touches[0];
            const afterElement = getDragAfterElement(list, touch.clientY);
            if (afterElement == null) list.appendChild(item);
            else list.insertBefore(item, afterElement);
        }, {passive: false});
        handle.addEventListener('touchend', () => { item.style.opacity = '1'; saveOrder(); });

        list.appendChild(item);
    });
    list.innerHTML += '<div style="height:180px;"></div>'; 
}

function getDragAfterElement(container, y) {
    const draggableElements = [...container.querySelectorAll('.drag-item:not(.dragging)')];
    return draggableElements.reduce((closest, child) => {
        const box = child.getBoundingClientRect();
        const offset = y - box.top - box.height / 2;
        if (offset < 0 && offset > closest.offset) return { offset: offset, element: child };
        else return closest;
    }, { offset: Number.NEGATIVE_INFINITY }).element;
}

function saveOrder() {
    const items = [...el('list-t').querySelectorAll('.drag-item')];
    db[currentIdx].tools = items.map(item => db[currentIdx].tools[item.getAttribute('data-idx')]);
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools();
}

// --- PDF ---
function makePDF() {
    const p = db[currentIdx];
    const footer = `
        <div style="margin-top: auto; padding-top: 15px;">
            <div style="border-top: 1.5px solid #000; width: 100%; margin-bottom: 5px;"></div>
            <div style="text-align: center; font-size: 9px; font-weight: 800; text-transform: uppercase;">QS CENTRAL PREMIUM REPORT</div>
        </div>`;

    const getH = () => `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <div><div style="font-size:13px; font-weight:900; color:#666;">${p.name}</div><div style="font-size:64px; font-weight:900; line-height:0.8;">${p.num}</div></div>
            <div style="width:220px; font-size:11px; font-weight:800; line-height:1.5;">
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>LAUFZEIT</span><span>${p.lzf}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>MATERIAL</span><span>${p.mat}</span></div>
                <div style="display:flex; justify-content:space-between;"><span>STÜCK</span><span>${p.stt}/${p.stn}</span></div>
            </div>
        </div><div style="border-bottom:5px solid #000; margin-bottom:15px;"></div>`;

    const subH = `<div style="display:flex; font-size:10px; font-weight:900; text-transform:uppercase; border-bottom:4px solid #000;">
        <div style="width:75px;">T-NR</div><div style="flex:1;">WERKZEUGNAME</div><div style="width:125px; text-align:right;">Ø / TOLERANZ</div>
    </div>`;

    const getR = (t) => `<div style="display:flex; border-bottom:1.5px solid #000; padding:10px 0; font-size:15px; font-weight:700;">
        <div style="width:75px; font-weight:800;">${t.id}</div><div style="flex:1;">${t.nm}</div><div style="width:125px; text-align:right;">${t.dia}</div>
    </div>`;

    let oben = (p.tools || []).filter(t => !t.rev), unten = (p.tools || []).filter(t => t.rev);
    const split = oben.length > 12 || (oben.length + unten.length > 15);

    let html = `<div style="width:210mm; font-family:sans-serif;">
        <div style="border:2px solid #000; padding:25px; min-height:280mm; display:flex; flex-direction:column; box-sizing:border-box; page-break-after:always;">
            ${getH()}<b>REVOLVER OBEN</b>${subH}${oben.map(getR).join('')}
            ${(!split && unten.length > 0) ? `<br><b>REVOLVER UNTEN</b>${subH}${unten.map(getR).join('')}` : ''}
            ${footer}
        </div>`;

    if (split && unten.length > 0) {
        html += `<div style="border:2px solid #000; padding:25px; min-height:280mm; display:flex; flex-direction:column; box-sizing:border-box;">
            ${getH()}<b>REVOLVER UNTEN</b>${subH}${unten.map(getR).join('')}${footer}
        </div>`;
    }

    el('print-container').innerHTML = html + `</div>`;
    setTimeout(() => { window.print(); }, 250);
}

// Модалка T
function modalT(i = null) {
    const edit = i !== null; el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    const btn = el('btn-rev-toggle');
    btn.classList.toggle('on', t.rev);
    btn.onclick = () => btn.classList.toggle('on');
    el('btn-del-t').style.display = edit ? 'block' : 'none';
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const t = { id: el('t-id').value.toUpperCase(), nm: el('t-nm').value.toUpperCase(), dia: el('t-dia').value, rev: el('btn-rev-toggle').classList.contains('on') };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); hide('m-t');
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { db[currentIdx].tools.splice(el('t-idx').value, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }
function exportJSON() { el('imp-area').value = JSON.stringify(db); }
function importJSON() { db = JSON.parse(el('imp-area').value); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); }

window.onload = renderList;
