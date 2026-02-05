const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'flex'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

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

function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'list-item';
        item.draggable = true;
        item.style.padding = '12px 15px';
        item.style.flexDirection = 'column';
        item.style.alignItems = 'flex-start';

        const revMark = t.rev ? `<div style="background:#000; color:#fff; font-size:10px; padding:4px 8px; border-radius:4px; margin-bottom:8px; font-weight:900; letter-spacing:1px;">REVOLVER UNTEN ↓</div>` : '';

        item.innerHTML = `
            ${revMark}
            <div style="flex:1; width:100%;" onclick="modalT(${i})">
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
            const to = i;
            const toolsArr = db[currentIdx].tools;
            const movedItem = toolsArr.splice(from, 1)[0];
            toolsArr.splice(to, 0, movedItem);
            localStorage.setItem(DB_KEY, JSON.stringify(db));
            renderTools();
        };
        list.appendChild(item);
    });
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

function makePDF() {
    const p = db[currentIdx];
    const getRow = (t) => {
        const isLong = (t.dia || '').length > 15;
        const align = isLong ? 'center' : 'baseline';
        const displayDia = t.dia.includes('/') ? t.dia.split('/').join('<br>') : t.dia;
        return `
        <div style="display:flex; align-items:${align}; border-bottom:1.5px solid #000; padding:10px 0; width:100%;">
            <div style="width:75px; font-weight:800; font-size:16px;">${t.id}</div>
            <div style="flex:1; font-weight:700; font-size:16px; text-transform:uppercase; padding-right:10px; white-space:pre-wrap;">${t.nm}</div>
            <div style="width:135px; text-align:right; font-weight:900; font-size:16px; line-height:1.2;">${displayDia}</div>
        </div>`;
    };

    let oben = [], unten = [], target = oben;
    (p.tools || []).forEach(t => {
        if(t.rev) target = unten;
        target.push(t);
    });

    const html = `
    <div style="width:210mm; padding:12mm; box-sizing:border-box; background:#fff; font-family:sans-serif; color:#000;">
        <div style="border:3.5px solid #000; padding:25px; min-height:265mm; display:flex; flex-direction:column; box-sizing:border-box;">
            
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
                <div>
                    <div style="font-size:14px; font-weight:900; text-transform:uppercase; color:#000;">${p.name || ''}</div>
                    <div style="font-size:68px; font-weight:900; line-height:0.8; letter-spacing:-2px;">${p.num || '---'}</div>
                </div>
                <div style="width:230px; font-size:12px; font-weight:900; line-height:1.6;">
                    <div style="display:flex; justify-content:space-between; border-bottom:2px solid #000;"><span>ABSTAND</span><span>${p.abs || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:2px solid #000;"><span>GREIFBACKEN</span><span>${p.grf || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:2px solid #000;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:2px solid #000;"><span>SÄGELÄNGE</span><span>${p.sag || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:2px solid #000;"><span>STÜCK T</span><span>${p.stt || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:2px solid #000;"><span>STÜCK N</span><span>${p.stn || ''}</span></div>
                </div>
            </div>
            
            <div style="border-bottom:10px solid #000; margin-bottom:25px;"></div>
            
            <div style="flex:1;">
                <div style="background:#000; color:#fff; padding:10px 20px; font-weight:900; font-size:20px; width:fit-content; letter-spacing:1px; margin-bottom:5px;">REVOLVER OBEN</div>
                <div style="border-bottom:5px solid #000; margin-bottom:10px;"></div>
                ${oben.map(getRow).join('')}
                
                ${unten.length > 0 ? `
                    <div style="background:#000; color:#fff; padding:10px 20px; font-weight:900; font-size:20px; width:fit-content; letter-spacing:1px; margin-top:40px; margin-bottom:5px;">REVOLVER UNTEN</div>
                    <div style="border-bottom:5px solid #000; margin-bottom:10px;"></div>
                    ${unten.map(getRow).join('')}
                ` : ''}
            </div>
        </div>
    </div>`;

    el('print-container').innerHTML = html;
    setTimeout(() => { window.print(); }, 150);
}

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
function exportJSON() { el('imp-area').value = JSON.stringify(db); el('imp-area').select(); alert("JSON OK"); }
function importJSON() { try { const parsed = JSON.parse(el('imp-area').value); if(Array.isArray(parsed)) { db = parsed; localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); hide('m-imp'); } } catch(e) { alert("JSON-Fehler"); } }
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { const i = el('t-idx').value; db[currentIdx].tools.splice(i, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }

window.onload = renderList;
