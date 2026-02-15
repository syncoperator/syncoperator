const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { --bg: #f8f9fb; --accent: #007aff; --card-bg: #ffffff; }
        body { background: var(--bg) !important; font-family: -apple-system, sans-serif !important; margin: 0; padding-bottom: 120px; }
        
        /* ЛОГО 250px И СБЛИЖЕННЫЙ ТЕКСТ */
        .brand-block { display: flex; flex-direction: column; align-items: center; padding: 40px 0 10px; }
        .logo-img { width: 250px !important; height: auto; margin-bottom: -45px !important; z-index: 1; }
        .header-title-main { font-weight: 900; font-size: 64px !important; letter-spacing: -4px; margin: 0; z-index: 2; position: relative; color: #000; }

        header { background: rgba(255,255,255,0.85); backdrop-filter: blur(15px); padding: 20px; position: sticky; top: 0; z-index: 100; border-bottom: 0.5px solid rgba(0,0,0,0.05); }
        
        .project-card { background: var(--card-bg); border-radius: 24px; margin: 0 20px 16px 20px; padding: 20px; box-shadow: 0 4px 20px rgba(0,0,0,0.04); display: flex; align-items: center; justify-content: space-between; }
        .tool-card { background: #fff; border-radius: 18px; padding: 14px 18px; margin: 0 16px 10px 16px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 2px 8px rgba(0,0,0,0.02); }

        .order-controls { display: flex; flex-direction: column; gap: 4px; }
        .btn-order { background: #f2f2f7; border: none; border-radius: 6px; width: 32px; height: 28px; font-weight: 900; color: var(--accent); }
        .btn-order.disabled { opacity: 0; pointer-events: none; }

        /* КНОПКА REVOLVER (ТУМБЛЕР) */
        #btn-rev-toggle {
            background: #f2f2f7; border: 2px solid transparent; padding: 14px;
            border-radius: 14px; font-weight: 800; width: 100%; margin: 10px 0;
            cursor: pointer; transition: 0.2s; text-align: center;
        }
        #btn-rev-toggle.on { background: #000 !important; color: #fff !important; }

        .fab { position: fixed; bottom: 30px; right: 20px; background: var(--accent); color: #fff; width: 60px; height: 60px; border-radius: 30px; display: flex; align-items: center; justify-content: center; font-size: 30px; z-index: 1000; border:none; }
        input { width: 100%; padding: 14px; border-radius: 12px; border: 1.5px solid #eee; margin-top: 5px; box-sizing: border-box; font-size: 16px; margin-bottom: 10px; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const show = (id) => { const x = el(id); if(x) x.style.display = 'flex'; };
const hide = (id) => { const x = el(id); if(x) x.style.display = 'none'; };
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// --- ГЛАВНАЯ ---
function renderList() {
    const list = el('list-p'); if(!list) return;
    list.innerHTML = `<div class="brand-block"><img src="https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png" class="logo-img"><div class="header-title-main">CitiTool</div></div>` + 
    db.map((p, i) => `
        <div class="project-card" onclick="openProject(${i})">
            <div><div style="font-size:11px; font-weight:700; color:var(--text-sub);">${p.name || 'PROJEKT'}</div><div style="font-size:24px; font-weight:900;">${p.num || '---'}</div></div>
            <div style="color:#ff3b30; font-weight:800; padding:10px;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:120px"></div>';
}

function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = tools.map((t, i) => `
        <div class="tool-card">
            <div style="flex:1;" onclick="modalT(${i})">
                ${t.rev ? `<div style="background:#000; color:#fff; font-size:8px; padding:2px 7px; border-radius:4px; margin-bottom:5px; font-weight:900; width:fit-content;">REVOLVER UNTEN</div>` : ''}
                <div style="font-size:10px; font-weight:700; color:#8e8e93;">${t.id || 'T0000'}</div>
                <div style="font-size:18px; font-weight:800;">${t.nm || '---'}</div>
                <div style="color:var(--accent); font-weight:700; font-size:13px;">${t.dia || ''}</div>
            </div>
            <div class="order-controls">
                <button class="btn-order ${i === 0 ? 'disabled' : ''}" onclick="event.stopPropagation(); moveItem(${i}, -1)">↑</button>
                <button class="btn-order ${i === tools.length - 1 ? 'disabled' : ''}" onclick="event.stopPropagation(); moveItem(${i}, 1)">↓</button>
            </div>
        </div>`).join('') + '<div style="height:180px"></div>';
}

// --- УПРАВЛЕНИЕ ---
function openProject(i) { currentIdx = i; el('v-home').classList.remove('active'); el('v-det').classList.add('active'); el('h-num').innerText = db[i].num; el('h-nam').innerText = db[i].name; renderTools(); }
function goHome() { currentIdx = null; el('v-home').classList.add('active'); el('v-det').classList.remove('active'); renderList(); }

function saveP() {
    const idx = el('p-idx').value;
    const data = {
        num: el('p-num').value, name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value, sag: el('p-sag').value, stt: el('p-stt').value, stn: el('p-stn').value,
        abs: el('p-abs').value, grf: el('p-grf').value, mat: el('p-mat') ? el('p-mat').value.toUpperCase() : '',
        tools: (idx !== '' && db[idx]) ? (db[idx].tools || []) : []
    };
    if (idx === '') db.push(data); else db[idx] = data;
    save(); hide('m-p'); renderList();
}

// --- ИСПРАВЛЕННЫЙ ПЕРЕКЛЮЧАТЕЛЬ ---
function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    
    const btn = el('btn-rev-toggle');
    if(btn) {
        // Устанавливаем начальное состояние
        btn.classList.toggle('on', t.rev);
        btn.innerText = t.rev ? "REVOLVER UNTEN" : "REVOLVER OBEN";
        
        btn.onclick = () => {
            const isNowOn = btn.classList.toggle('on');
            btn.innerText = isNowOn ? "REVOLVER UNTEN" : "REVOLVER OBEN";
        };
    }
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const isUnten = el('btn-rev-toggle').classList.contains('on');
    const t = { 
        id: el('t-id').value.toUpperCase(), 
        nm: el('t-nm').value.toUpperCase(), 
        dia: el('t-dia').value, 
        rev: isUnten 
    };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    save(); renderTools(); hide('m-t');
}

function moveItem(i, dir) { const t = db[currentIdx].tools; if(i+dir>=0 && i+dir<t.length) { [t[i],t[i+dir]]=[t[i+dir],t[i]]; save(); renderTools(); } }
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); save(); renderList(); } }

// --- PDF С УМНЫМ РАЗРЫВОМ ---
function makePDF() {
    const p = db[currentIdx];
    const getHead = () => `<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; min-height:90px;"><div><div style="font-size:13px; font-weight:900; color:#666;">${p.name || ''}</div><div style="font-size:64px; font-weight:900; line-height:0.8; letter-spacing:-2px; margin:0;">${p.num || '---'}</div></div><div style="width:220px; font-size:11px; font-weight:800; line-height:1.5;"><div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>MATERIAL</span><span>${p.mat || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>SÄGELÄNGE</span><span>${p.sag || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>ABSTAND</span><span>${p.abs || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>GREIFBACKEN</span><span>${p.grf || ''}</span></div><div style="display:flex; justify-content:space-between;"><span>STÜCKZAHL</span><span>${p.stt || ''} / ${p.stn || ''}</span></div></div></div><div style="border-bottom:5px solid #000; margin-bottom:15px;"></div>`;
    const tHead = (t) => `<div style="margin-top:10px; font-size:18px; font-weight:900;">${t}</div><div style="display:flex; font-size:10px; font-weight:900; margin-bottom:6px;"><div style="width:75px;">T-NR</div><div style="flex:1;">WERKZEUGNAME</div><div style="width:125px; text-align:right;">Ø / TOLERANZ</div></div><div style="border-bottom:4px solid #000;"></div>`;
    const getRow = (t) => `<div style="display:flex; align-items:baseline; border-bottom:1.5px solid #000; padding:10px 0; page-break-inside: avoid;"><div style="width:75px; font-weight:800; font-size:15px;">${t.id}</div><div style="flex:1; font-weight:700; font-size:15px; text-transform:uppercase;">${t.nm}</div><div style="width:125px; text-align:right; font-weight:800; font-size:14px;">${t.dia.replace(/\//g, '<br>')}</div></div>`;
    
    let oben = [], unten = [];
    (p.tools || []).forEach(t => t.rev ? unten.push(t) : oben.push(t));
    const shouldBreak = (p.tools || []).length > 10;

    let html = `<!DOCTYPE html><html><head><style>@page { size: A4; margin: 0; } body { margin: 0; padding: 10mm; font-family: sans-serif; } .page { padding: 15mm; min-height: 297mm; box-sizing: border-box; display: flex; flex-direction: column; border: 2.2px solid #000; page-break-after: always; } .footer { border-top: 1.2px solid #000; padding-top: 5px; font-size: 9px; font-weight: 800; text-align: center; margin-top: auto; }</style></head><body><div class="page">${getHead()}${tHead('REVOLVER OBEN')}${oben.map(getRow).join('')}`;
    
    if (unten.length > 0) {
        if (shouldBreak) {
            html += `<div class="footer">CITITOOL</div></div><div class="page">${getHead()}${tHead('REVOLVER UNTEN')}${unten.map(getRow).join('')}`;
        } else {
            html += `${tHead('REVOLVER UNTEN')}${unten.map(getRow).join('')}`;
        }
    }
    html += `<div class="footer">CITITOOL</div></div><script>window.onload=()=>setTimeout(()=>window.print(),400)</script></body></html>`;
    const win = window.open('','_blank'); win.document.write(html); win.document.close();
}

window.onload = () => { injectStyles(); renderList(); };
