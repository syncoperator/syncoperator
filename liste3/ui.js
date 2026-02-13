const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f2f4f7; 
            --neu-light: #ffffff;
            --neu-shadow: #d1d9e6;
            --accent: #007aff;
            --text-main: #1a1a1a;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important; 
            margin: 0; padding-bottom: 80px; color: var(--text-main);
        }

        /* Дорогой зеркальный заголовок */
        .header-container {
            padding: 50px 0 30px 0;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
        }
        .mirror-box {
            position: relative;
            line-height: 1;
            text-align: center;
        }
        .mirror-main {
            font-size: 52px;
            font-weight: 900;
            letter-spacing: -2px;
            color: var(--text-main);
            position: relative;
            z-index: 2;
        }
        .mirror-reflect {
            font-size: 52px;
            font-weight: 900;
            letter-spacing: -2px;
            margin-top: -22px;
            transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.25) 0%, rgba(242,244,247,1) 85%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            opacity: 0.4;
            user-select: none;
        }

        /* Кнопка + NEU справа */
        .top-controls {
            display: flex;
            justify-content: flex-end;
            padding: 0 25px;
            margin-bottom: 10px;
        }
        .btn-neu {
            background: var(--bg);
            border: none;
            border-radius: 14px;
            padding: 10px 18px;
            font-size: 13px;
            font-weight: 800;
            color: var(--accent);
            box-shadow: 5px 5px 10px var(--neu-shadow), -5px -5px 10px var(--neu-light);
            transition: all 0.2s ease;
        }
        .btn-neu:active {
            box-shadow: inset 2px 2px 5px var(--neu-shadow), inset -2px -2px 5px var(--neu-light);
            transform: scale(0.98);
        }

        /* Карточки проектов и инструментов */
        .neu-card {
            background: var(--bg);
            border-radius: 28px;
            padding: 24px;
            margin: 15px 20px;
            box-shadow: 10px 10px 20px var(--neu-shadow), -10px -10px 20px var(--neu-light);
            display: flex;
            align-items: center;
            justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.4);
        }

        .p-label { font-size: 10px; font-weight: 800; color: #99a1ad; text-transform: uppercase; letter-spacing: 0.5px; }
        .p-title { font-size: 24px; font-weight: 900; color: #000; margin-top: 2px; }

        /* Компактные кнопки перемещения */
        .order-wrap { display: flex; flex-direction: column; gap: 6px; }
        .btn-move {
            width: 36px; height: 32px;
            background: var(--bg);
            border-radius: 10px;
            border: none;
            box-shadow: 3px 3px 6px var(--neu-shadow), -3px -3px 6px var(--neu-light);
            color: var(--accent);
            font-size: 14px; font-weight: 900;
            display: flex; align-items: center; justify-content: center;
        }
        .btn-move:active { box-shadow: inset 2px 2px 4px var(--neu-shadow), inset -2px -2px 4px var(--neu-light); }
        .btn-move.disabled { opacity: 0; pointer-events: none; }

        .btn-del { color: #ff3b30; font-size: 18px; font-weight: 900; padding: 10px; }
        
        /* Фикс отступа снизу для скролла */
        .spacer { height: 120px; width: 100%; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'block'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// --- ГЛАВНЫЙ ЭКРАН ---
function renderList() {
    const main = el('v-home');
    if(!main) return;
    
    main.innerHTML = `
        <div class="header-container">
            <div class="mirror-box">
                <div class="mirror-main">CitiTool</div>
                <div class="mirror-reflect">CitiTool</div>
            </div>
        </div>
        <div class="top-controls">
            <button class="btn-neu" onclick="modalP()">+ NEU</button>
        </div>
        <div id="list-p"></div>
    `;
    
    const list = el('list-p');
    list.innerHTML = db.map((p, i) => `
        <div class="neu-card" onclick="openProject(${i})">
            <div>
                <div class="p-label">PROJEKT</div>
                <div class="p-title">${p.num || '---'}</div>
                <div style="font-size:12px; color:#666; font-weight:600; margin-top:4px;">${p.name || ''}</div>
            </div>
            <div class="btn-del" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div class="spacer"></div>';
}

// --- СПИСОК ИНСТРУМЕНТОВ ---
function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'neu-card';
        const revMark = t.rev ? `<div style="font-size:9px; font-weight:900; color:#ff9500; margin-bottom:4px;">REVOLVER UNTEN</div>` : '';
        
        item.innerHTML = `
            <div style="flex:1" onclick="modalT(${i})">
                ${revMark}
                <div class="p-label">${t.id || 'T0000'}</div>
                <div class="p-title" style="font-size:20px;">${t.nm || '---'}</div>
                <div style="margin-top:5px; font-weight:800; color:var(--accent); font-size:14px;">${t.dia || ''}</div>
            </div>
            <div class="order-wrap">
                <button class="btn-move ${i === 0 ? 'disabled' : ''}" onclick="event.stopPropagation(); moveItem(${i}, -1)">↑</button>
                <button class="btn-move ${i === tools.length - 1 ? 'disabled' : ''}" onclick="event.stopPropagation(); moveItem(${i}, 1)">↓</button>
            </div>`;
        list.appendChild(item);
    });
    list.innerHTML += '<div class="spacer"></div>';
}

function moveItem(i, direction) {
    const tools = db[currentIdx].tools;
    const target = i + direction;
    if (target >= 0 && target < tools.length) {
        [tools[i], tools[target]] = [tools[target], tools[i]];
        save(); renderTools();
    }
}

// --- НАВИГАЦИЯ ---
function openProject(i) { 
    currentIdx = i; 
    el('v-home').style.display='none'; 
    el('v-det').style.display='block'; 
    el('h-num').innerText = db[i].num; 
    el('h-nam').innerText = db[i].name; 
    renderTools(); 
}

function goHome() { 
    currentIdx = null; 
    el('v-home').style.display='block'; 
    el('v-det').style.display='none'; 
    renderList(); 
}

// --- МОДАЛЬНЫЕ ОКНА ---
function modalP(edit = false) {
    if (!edit) currentIdx = null;
    const p = (edit && db[currentIdx]) ? db[currentIdx] : {num:'', name:'', lzf:'', sag:'', stt:'', stn:'', abs:'', grf:'', mat:''};
    el('p-idx').value = edit ? currentIdx : '';
    el('p-num').value = p.num; el('p-nam').value = p.name;
    el('p-lzf').value = p.lzf; el('p-sag').value = p.sag;
    el('p-stt').value = p.stt; el('p-stn').value = p.stn;
    el('p-abs').value = p.abs; el('p-grf').value = p.grf;
    if(el('p-mat')) el('p-mat').value = p.mat;
    show('m-p');
}

function saveP() {
    const idx = el('p-idx').value;
    const data = {
        num: el('p-num').value, name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value, sag: el('p-sag').value,
        stt: el('p-stt').value, stn: el('p-stn').value,
        abs: el('p-abs').value, grf: el('p-grf').value,
        mat: el('p-mat') ? el('p-mat').value.toUpperCase() : '',
        tools: (idx !== '' && db[idx]) ? (db[idx].tools || []) : []
    };
    if (idx === '') db.push(data); else db[idx] = data;
    save(); hide('m-p'); renderList();
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    const btn = el('btn-rev-toggle');
    if(btn) { 
        t.rev ? btn.classList.add('on') : btn.classList.remove('on'); 
        btn.onclick = () => btn.classList.toggle('on'); 
    }
    el('btn-del-t').style.display = edit ? 'block' : 'none';
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const btn = el('btn-rev-toggle');
    const t = { 
        id: el('t-id').value.toUpperCase(), 
        nm: el('t-nm').value.toUpperCase(), 
        dia: el('t-dia').value, 
        rev: btn ? btn.classList.contains('on') : false 
    };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    save(); renderTools(); hide('m-t');
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); save(); renderList(); } }
function delT() { const i = el('t-idx').value; db[currentIdx].tools.splice(i, 1); save(); renderTools(); hide('m-t'); }

// --- PDF REPORT ---
function makePDF() {
    const p = db[currentIdx];
    const getPageHead = () => `<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; min-height:90px;"><div style="display:flex; flex-direction:column; justify-content:center;"><div style="font-size:13px; font-weight:900; text-transform:uppercase; color:#666; margin-bottom:2px; line-height:1;">${p.name || ''}</div><div style="font-size:64px; font-weight:900; line-height:0.8; letter-spacing:-2px; margin:0;">${p.num || '---'}</div></div><div style="width:220px; font-size:11px; font-weight:800; line-height:1.5;"><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>MATERIAL</span><span>${p.mat || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>SÄGELÄNGE</span><span>${p.sag || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>ABSTAND</span><span>${p.abs || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>GREIFBACKEN</span><span>${p.grf || ''}</span></div><div style="display:flex; justify-content:space-between;"><span>STÜCKZAHL</span><span>${p.stt || ''} / ${p.stn || ''}</span></div></div></div><div style="border-bottom:5px solid #000; margin-bottom:15px;"></div>`;
    const tableHead = `<div style="display:flex; font-size:10px; font-weight:900; text-transform:uppercase; margin-bottom:6px; padding:0 2px;"><div style="width:75px;">T-NR</div><div style="flex:1;">WERKZEUGNAME / KOMMENTAR</div><div style="width:125px; text-align:right;">Ø / TOLERANZ</div></div><div style="border-bottom:4px solid #000;"></div>`;
    const getRow = (t) => `<div style="display:flex; align-items:baseline; border-bottom:1.5px solid #000; padding:10px 0; width:100%; page-break-inside: avoid;"><div style="width:75px; font-weight:800; font-size:15px;">${t.id}</div><div style="flex:1; font-weight:700; font-size:15px; text-transform:uppercase; padding-right:10px; white-space:pre-wrap;">${t.nm}</div><div style="width:125px; text-align:right; font-weight:800; font-size:14px;">${t.dia.replace(/\//g, '<br>')}</div></div>`;
    let oben = [], unten = [], target = oben;
    (p.tools || []).forEach(t => { if(t.rev) target = unten; target.push(t); });
    let pdfHtml = `<!DOCTYPE html><html><head><style>@page { size: A4; margin: 0; } body { margin: 0; padding: 10mm; background: #fff; font-family: sans-serif; -webkit-print-color-adjust: exact; } .page { width: 210mm; height: 297mm; padding: 15mm; box-sizing: border-box; page-break-after: always; display: flex; flex-direction: column; } .content-border { border: 2.2px solid #000; padding: 20px; flex: 1; display: flex; flex-direction: column; box-sizing: border-box; } .footer { border-top: 1.2px solid #000; padding-top: 5px; font-size: 9px; font-weight: 800; text-align: center; color: #666; margin-top: auto; }</style></head><body><div class="page"><div class="content-border">${getPageHead()}<div style="flex:1;"><div style="margin-bottom:5px; font-size:18px; font-weight:900; text-transform:uppercase;">REVOLVER OBEN</div>${tableHead}${oben.map(getRow).join('')}</div><div class="footer">CITITOOL REPORT</div></div></div>${unten.length > 0 ? `<div class="page"><div class="content-border">${getPageHead()}<div style="flex:1;"><div style="margin-bottom:5px; font-size:18px; font-weight:900; text-transform:uppercase;">REVOLVER UNTEN</div>${tableHead}${unten.map(getRow).join('')}</div><div class="footer">CITITOOL REPORT</div></div></div>` : ''}<script>window.onload = function() { setTimeout(() => { window.print(); }, 400); };</script></body></html>`;
    const win = window.open('', '_blank'); if (win) { win.document.write(pdfHtml); win.document.close(); }
}

window.onload = () => { injectStyles(); renderList(); };
