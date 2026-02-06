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
        <div class="premium-card" onclick="openProject(${i})">
            <div class="btn-del-p" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
            <div class="card-num">${p.num || '---'}</div>
            <div class="card-name">${p.name || 'UNNAMED'}</div>
        </div>`).join('');
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

// --- WERKZEUGE (DRAG & DROP) ---
let startIdx = null;

function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'tool-card';
        item.setAttribute('data-idx', i);

        const revMark = t.rev ? `<div style="background:var(--text-main); color:#fff; font-size:9px; padding:2px 6px; border-radius:4px; margin-top:5px; font-weight:900; width:fit-content;">UNTEN ↓</div>` : '';

        item.innerHTML = `
            <div class="handle">☰</div>
            <div class="tool-info" onclick="modalT(${i})">
                <div class="t-nr-label">${t.id || 'T-NR'}</div>
                <b class="t-desc-label">${t.nm || 'Keine Beschreibung'}</b>
                ${revMark}
            </div>
        `;
        
        const handle = item.querySelector('.handle');
        handle.ontouchstart = () => { startIdx = i; item.style.boxShadow = "inset 4px 4px 8px var(--neu-lower)"; };
        handle.ontouchmove = (e) => {
            e.preventDefault();
            const touch = e.touches[0];
            const target = document.elementFromPoint(touch.clientX, touch.clientY);
            const targetItem = target?.closest('.tool-card');
            if (targetItem) {
                const overIdx = parseInt(targetItem.getAttribute('data-idx'));
                if (overIdx !== startIdx) { moveTool(startIdx, overIdx); startIdx = overIdx; }
            }
        };
        handle.ontouchend = () => { renderTools(); };

        item.draggable = true;
        item.ondragstart = () => { startIdx = i; item.style.opacity = '0.4'; };
        item.ondragover = (e) => e.preventDefault();
        item.ondrop = () => { if(startIdx !== i) moveTool(startIdx, i); };
        item.ondragend = () => { renderTools(); };

        list.appendChild(item);
    });
    list.innerHTML += '<div class="bottom-safe-area"></div>'; 
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
    if(t.rev) btn.classList.add('on'); else btn.classList.remove('on');
    btn.onclick = () => btn.classList.toggle('on');
    el('btn-del-t').style.display = edit ? 'block' : 'none';
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const t = { 
        id: el('t-id').value.toUpperCase(), 
        nm: el('t-nm').value.toUpperCase(), 
        dia: el('t-dia').value, 
        rev: el('btn-rev-toggle').classList.contains('on') 
    };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); hide('m-t');
}

function runImp() {
    const raw = el('imp-area').value.trim();
    if(!raw) return;
    const lines = raw.split('\n');
    lines.forEach(line => {
        const parts = line.split('\t');
        if(parts.length >= 2) db[currentIdx].tools.push({ id: parts[0].trim(), nm: parts[1].trim(), dia: parts[2] || '', rev: false });
    });
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); hide('m-imp');
}

function exportJSON() { el('imp-area').value = JSON.stringify(db); el('imp-area').select(); }
function importJSON() { try { const parsed = JSON.parse(el('imp-area').value); if(Array.isArray(parsed)) { db = parsed; localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); } } catch(e) { alert("Error"); } }
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { const i = el('t-idx').value; db[currentIdx].tools.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }

// PDF остается как было - центрирование и жирные линии
function makePDF() {
    const p = db[currentIdx];
    const headerRow = `<div style="display:flex; font-size:10px; font-weight:900; text-transform:uppercase; margin-bottom:6px; padding:0 2px;"><div style="width:75px;">T-NR</div><div style="flex:1;">WERKZEUGNAME / KOMMENTAR</div><div style="width:125px; text-align:right;">Ø / TOLERANZ</div></div><div style="border-bottom:4px solid #000; margin-bottom:0px;"></div>`;
    const getRow = (t) => {
        const d = t.dia.includes('/') ? t.dia.split('/').join('<br>') : t.dia;
        return `<div style="display:flex; align-items:baseline; border-bottom:1.5px solid #000; padding:10px 0; width:100%;"><div style="width:75px; font-weight:800; font-size:15px;">${t.id}</div><div style="flex:1; font-weight:700; font-size:15px; text-transform:uppercase; padding-right:10px; white-space:pre-wrap;">${t.nm}</div><div style="width:125px; text-align:right; font-weight:800; font-size:14px; line-height:1.2;">${d}</div></div>`;
    };
    let oben = [], unten = [], target = oben;
    (p.tools || []).forEach(t => { if(t.rev) target = unten; target.push(t); });

    const html = `
    <div style="width:210mm; padding:12mm; box-sizing:border-box; background:#fff; font-family:sans-serif; color:#000;">
        <div style="border:2px solid #000; padding:25px; min-height:265mm; display:flex; flex-direction:column;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
                <div><div style="font-size:13px; font-weight:900; color:#666;">${p.name || ''}</div><div style="font-size:64px; font-weight:900; line-height:0.8; letter-spacing:-2px;">${p.num || '---'}</div></div>
                <div style="width:220px; font-size:11px; font-weight:800; line-height:1.5;">
                    <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>MATERIAL</span><span>${p.mat || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>SÄGELÄNGE</span><span>${p.sag || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>ABSTAND</span><span>${p.abs || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>GREIFBACKEN</span><span>${p.grf || ''}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>STÜCKZAHL</span><span>${p.stt || ''} / ${p.stn || ''}</span></div>
                </div>
            </div>
            <div style="border-bottom:5px solid #000; margin-bottom:15px;"></div>
            <div style="flex:1;">
                <div style="margin-bottom:5px; font-size:18px; font-weight:900;">REVOLVER OBEN</div>${headerRow}${oben.map(getRow).join('')}
                ${unten.length > 0 ? `<div style="margin-top:35px; margin-bottom:5px; font-size:18px; font-weight:900;">REVOLVER UNTEN</div>${headerRow}${unten.map(getRow).join('')}` : ''}
            </div>
        </div>
    </div>`;
    el('print-container').innerHTML = html;
    setTimeout(() => { window.print(); }, 150);
}
window.onload = renderList;
