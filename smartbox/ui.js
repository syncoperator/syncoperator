const DB_KEY = 'SMARTBOX_DATA_V1';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const show = (id) => el(id).style.display = 'flex';
const hide = (id) => el(id).style.display = 'none';

// --- PROJECTS ---
function renderList() {
    const list = el('list-p');
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})">
            <div style="flex:1">
                <div class="t-id-label">${p.name || 'Projekt'}</div>
                <div class="t-name-label">${p.num || '---'}</div>
            </div>
            <div style="color:red; font-weight:800; padding:10px;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:120px"></div>';
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
    const list = el('list-t');
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'list-item';
        item.setAttribute('data-idx', i);
        item.innerHTML = `
            <div class="handle">☰</div>
            <div style="flex:1" onclick="modalT(${i})">
                ${t.rev ? '<div class="t-loc-badge" style="background:gray">REV UNTEN</div>' : ''}
                <div class="t-id-label">${t.id}</div>
                <div class="t-name-label">${t.nm}</div>
                ${t.loc ? `<div class="t-loc-badge">📍 ${t.loc}</div>` : ''}
                ${t.dia ? `<br><div class="t-dia-badge">${t.dia}</div>` : ''}
            </div>`;
        
        const handle = item.querySelector('.handle');
        handle.ontouchstart = () => { startIdx = i; };
        handle.ontouchmove = (e) => {
            const touch = e.touches[0];
            const target = document.elementFromPoint(touch.clientX, touch.clientY)?.closest('.list-item');
            if (target) {
                const overIdx = parseInt(target.getAttribute('data-idx'));
                if (overIdx !== startIdx) { moveTool(startIdx, overIdx); startIdx = overIdx; }
            }
        };
        list.appendChild(item);
    });
    list.innerHTML += '<div style="height:150px"></div>';
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
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false, loc:''};
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    el('t-loc').value = t.loc || '';
    const btn = el('btn-rev-toggle');
    t.rev ? btn.classList.add('on') : btn.classList.remove('on');
    btn.onclick = () => btn.classList.toggle('on');
    el('btn-del-t').style.display = edit ? 'block' : 'none';
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const t = { 
        id: el('t-id').value.toUpperCase(), nm: el('t-nm').value.toUpperCase(), 
        dia: el('t-dia').value, rev: el('btn-rev-toggle').classList.contains('on'),
        loc: el('t-loc').value.toUpperCase()
    };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); hide('m-t');
}

function modalP() { el('p-idx').value = ''; show('m-p'); }
function saveP() {
    const p = { 
        num: el('p-num').value, name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value, mat: el('p-mat').value,
        sag: el('p-sag').value, abs: el('p-abs').value,
        grf: el('p-grf').value, stt: el('p-stt').value, stn: el('p-stn').value,
        tools: []
    };
    db.push(p); localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderList(); hide('m-p');
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i,1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderList(); } }
function delT() { db[currentIdx].tools.splice(el('t-idx').value, 1); localStorage.setItem(DB_KEY, JSON.stringify(db)); renderTools(); hide('m-t'); }

// --- PDF ГЕНЕРАТОР ---
function makePDF() {
    const p = db[currentIdx];
    const inBox = (p.tools || []).filter(t => !t.loc || t.loc === 'BOX');
    const inArsenal = (p.tools || []).filter(t => t.loc && t.loc !== 'BOX');

    const getRow = (t) => `
        <div style="display:flex; border-bottom:1.5px solid #000; padding:8px 0; align-items:center;">
            <div style="width:70px; font-weight:800; font-size:14px;">${t.id}</div>
            <div style="flex:1;">
                <div style="font-weight:700; font-size:14px;">${t.nm}</div>
                <div style="font-size:9px; font-weight:800; color:#555;">${t.loc ? 'LAGER: '+t.loc : 'IN DER BOX'}</div>
            </div>
            <div style="width:100px; text-align:right; font-weight:800;">${t.dia}</div>
        </div>`;

    const html = `<html><body style="padding:15mm; font-family:sans-serif;">
        <div style="border:2px solid #000; padding:20px;">
            <div style="display:flex; justify-content:space-between;">
                <div><div style="font-size:12px; font-weight:800;">${p.name}</div><div style="font-size:50px; font-weight:900;">${p.num}</div></div>
                <div style="font-size:10px; font-weight:700;">${p.mat}<br>${p.lzf}</div>
            </div>
            <div style="background:#000; color:#fff; padding:5px; margin-top:20px; font-weight:900;">📦 IN DER BOX (Herz)</div>
            ${inBox.map(getRow).join('')}
            <div style="background:#000; color:#fff; padding:5px; margin-top:20px; font-weight:900;">🏗️ ARSENAL (Nachrüsten)</div>
            ${inArsenal.map(getRow).join('')}
        </div>
        <script>window.onload=()=>{window.print();window.close();}<\/script>
    </body></html>`;
    
    const win = window.open('','_blank');
    win.document.write(html);
    win.document.close();
}

window.onload = () => renderList();
