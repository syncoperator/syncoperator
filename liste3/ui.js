const DB_KEY = 'QS_DATA_V12';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'flex'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

function renderList() {
    const list = el('list-p'); if(!list) return;
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})">
            <div class="indicator" style="background:#3b82f6"></div>
            <div class="item-info" style="flex:1;">
                <small>${p.name || 'PROJECT'}</small>
                <b>${p.num || '---'}</b>
            </div>
            <div style="color:#d1d1d6; font-size:20px; padding:10px;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
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

let startIdx = null;
function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'list-item';
        item.setAttribute('data-idx', i);
        item.innerHTML = `
            <div class="handle" style="cursor:grab; color:#d1d1d6; font-size:20px; padding-right:15px; touch-action:none;">☰</div>
            <div class="indicator" style="background:${t.rev ? '#000' : '#3b82f6'}"></div>
            <div class="item-info" style="flex:1;" onclick="modalT(${i})">
                <small class="t-nr-small">${t.id || 'T0000'}</small>
                <b>${t.nm || '---'}</b>
            </div>
            <div style="font-weight:800; color:var(--primary);">${t.dia || ''}</div>
        `;
        const h = item.querySelector('.handle');
        h.ontouchstart = () => { startIdx = i; item.style.opacity = '0.5'; };
        h.ontouchmove = (e) => {
            e.preventDefault();
            const touch = e.touches[0];
            const target = document.elementFromPoint(touch.clientX, touch.clientY)?.closest('.list-item');
            if (target) {
                const overIdx = parseInt(target.getAttribute('data-idx'));
                if (overIdx !== startIdx) { moveTool(startIdx, overIdx); startIdx = overIdx; }
            }
        };
        h.ontouchend = () => { item.style.opacity = '1'; renderTools(); };
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
    el('p-idx').value = edit ? currentIdx : '';
    el('p-num').value = p.num; el('p-nam').value = p.name;
    el('p-lzf').value = p.lzf; el('p-sag').value = p.sag;
    el('p-stt').value = p.stt; el('p-stn').value = p.stn;
    el('p-abs').value = p.abs; el('p-grf').value = p.grf;
    el('p-mat').value = p.mat || '';
    show('m-p');
}

function saveP() {
    const idx = el('p-idx').value;
    const newP = {
        num: el('p-num').value, name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value, sag: el('p-sag').value,
        stt: el('p-stt').value, stn: el('p-stn').value,
        abs: el('p-abs').value, grf: el('p-grf').value,
        mat: el('p-mat').value.toUpperCase(),
        tools: (idx !== '' && db[idx]) ? (db[idx].tools || []) : []
    };
    if (idx === '') { db.push(newP); currentIdx = db.length - 1; } 
    else { db[idx] = newP; }
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    hide('m-p'); renderList();
    if(idx !== '') openProject(idx); else goHome();
}

function modalT(i = null) {
    const edit = i !== null; 
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    
    // Прямое обращение к ID полей
    el('t-id').value = t.id; 
    el('t-nm').value = t.nm; 
    el('t-dia').value = t.dia;
    
    const toggle = el('t-rev-toggle');
    if(t.rev) toggle.classList.add('on'); else toggle.classList.remove('on');
    toggle.style.background = t.rev ? '#3b82f6' : '#e5e5ea';
    toggle.firstChild.style.left = t.rev ? '25px' : '3px';
    
    el('btn-del-t').style.display = edit ? 'block' : 'none';
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const tId = el('t-id').value.toUpperCase();
    const tNm = el('t-nm').value.toUpperCase();
    const tDia = el('t-dia').value;
    const isRev = el('t-rev-toggle').classList.contains('on');

    if (!db[currentIdx].tools) db[currentIdx].tools = [];
    
    const toolObj = { id: tId, nm: tNm, dia: tDia, rev: isRev };

    if (i === '') {
        db[currentIdx].tools.push(toolObj);
    } else {
        db[currentIdx].tools[i] = toolObj;
    }

    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); 
    hide('m-t');
}

function makePDF() {
    const p = db[currentIdx];
    const headerRow = `<div style="display:flex; font-size:10px; font-weight:900; text-transform:uppercase; margin-bottom:6px; padding:0 2px;">
        <div style="width:75px;">T-NR</div><div style="flex:1;">WERKZEUGNAME / KOMMENTAR</div><div style="width:125px; text-align:right;">Ø / TOLERANZ</div>
    </div><div style="border-bottom:4px solid #000; margin-bottom:0px;"></div>`;
    
    const getRow = (t) => `<div style="display:flex; align-items:baseline; border-bottom:1.5px solid #000; padding:10px 0; width:100%;">
        <div style="width:75px; font-weight:800; font-size:15px;">${t.id}</div>
        <div style="flex:1; font-weight:700; font-size:15px; text-transform:uppercase;">${t.nm}</div>
        <div style="width:125px; text-align:right; font-weight:800; font-size:14px;">${t.dia}</div>
    </div>`;

    const getFullHeader = () => `
        <div style="display:flex; justify-content:space-between; margin-bottom:15px;">
            <div><div style="font-size:13px; font-weight:900;">${p.name}</div><div style="font-size:64px; font-weight:900;">${p.num}</div></div>
            <div style="width:220px; font-size:11px; font-weight:800; line-height:1.5;">
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>LAUFZEIT</span><span>${p.lzf}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>MATERIAL</span><span>${p.mat}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>SÄGELÄNGE</span><span>${p.sag}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>ABSTAND</span><span>${p.abs}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>GREIFBACKEN</span><span>${p.grf}</span></div>
                <div style="display:flex; justify-content:space-between;"><span>STÜCKZAHL</span><span>${p.stt}/${p.stn}</span></div>
            </div>
        </div>
        <div style="border-bottom:5px solid #000; margin-bottom:15px;"></div>`;

    let o = [], u = [];
    (p.tools || []).forEach(t => t.rev ? u.push(t) : o.push(t));

    const forceBreak = o.length > 12 || (o.length + u.length) > 15;

    let html = `
    <div style="width:210mm; padding:12mm; background:#fff; font-family:sans-serif;">
        <div style="border:2px solid #000; padding:25px; min-height:265mm; display:flex; flex-direction:column; position:relative;">
            ${getFullHeader()}
            <div style="font-size:18px; font-weight:900; margin-top:10px;">REVOLVER OBEN</div>
            ${headerRow}
            ${o.map(getRow).join('')}
            ${(!forceBreak && u.length > 0) ? `<div style="margin-top:30px; font-size:18px; font-weight:900;">REVOLVER UNTEN</div>${headerRow}${u.map(getRow).join('')}` : ''}
            <div style="position:absolute; bottom:20px; left:0; width:100%; text-align:center; font-size:9px; color:#aaa;">QS CENTRAL PREMIUM REPORT</div>
        </div>
        ${(forceBreak && u.length > 0) ? `
        <div class="page-break" style="border:2px solid #000; padding:25px; min-height:265mm; display:flex; flex-direction:column; position:relative; margin-top:20px;">
            ${getFullHeader()}
            <div style="font-size:18px; font-weight:900; margin-top:10px;">REVOLVER UNTEN</div>
            ${headerRow}
            ${u.map(getRow).join('')}
            <div style="position:absolute; bottom:20px; left:0; width:100%; text-align:center; font-size:9px; color:#aaa;">QS CENTRAL PREMIUM REPORT - PAGE 2</div>
        </div>` : ''}
    </div>`;

    el('print-container').innerHTML = html; window.print();
}

function runImp() {
    const text = el('imp-area').value; if (!text.trim()) return;
    const regex = /(T[0O]\d{2,4})/gi; const parts = text.split(regex);
    for (let i = 1; i < parts.length; i += 2) {
        db[currentIdx].tools.push({ id: parts[i].trim().toUpperCase().replace('O', '0'), nm: parts[i + 1].trim().toUpperCase(), dia: '', rev: false });
    }
    el('imp-area').value = ''; renderTools(); hide('m-imp');
}
function exportJSON() { el('imp-area').value = JSON.stringify(db); }
function importJSON() { db = JSON.parse(el('imp-area').value); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); }
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { db[currentIdx].tools.splice(el('t-idx').value, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }

window.onload = renderList;
