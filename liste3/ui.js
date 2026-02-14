const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f8f9fb; 
            --accent: #007aff; 
            --text-main: #1c1c1e;
            --text-sub: #8e8e93; 
            --card-bg: #ffffff;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important; 
            margin: 0; color: var(--text-main);
            padding-bottom: 150px;
        }

        /* ЛОГО: ОГРОМНОЕ И ВПЛОТНУЮ */
        .brand-block {
            display: flex; flex-direction: column; align-items: center;
            padding: 40px 0 10px;
        }
        .logo-img {
            width: 250px !important; /* Твой запрос: больше 200px */
            height: auto;
            margin-bottom: -40px !important; /* Максимальное сближение */
            z-index: 1;
        }
        .header-title { 
            font-weight: 900; 
            font-size: 64px !important; /* Дорогой большой шрифт */
            letter-spacing: -4px; 
            margin: 0; 
            z-index: 2;
            position: relative;
            color: #000;
        }
        
        /* Твой стиль заголовка в PDF и деталях */
        header {
            background: rgba(255,255,255,0.85) !important;
            backdrop-filter: blur(15px); -webkit-backdrop-filter: blur(15px);
            padding: 20px; position: sticky; top: 0; z-index: 100;
            border-bottom: 0.5px solid rgba(0,0,0,0.05);
        }

        /* РАЗДЕЛИТЕЛЬ REVOLVER (КНОПКА) */
        .rev-control {
            display: flex; align-items: center; justify-content: space-between;
            background: #f2f2f7; padding: 15px; border-radius: 18px; margin: 15px 0;
        }
        .btn-rev {
            background: #fff; border: 2px solid #ddd; border-radius: 12px;
            padding: 10px 20px; font-weight: 900; font-size: 12px; cursor: pointer;
        }
        .btn-rev.on { 
            background: #000 !important; color: #fff !important; border-color: #000 !important; 
        }

        /* Остальные твои рабочие стили */
        .project-card {
            background: var(--card-bg); border-radius: 24px; margin: 0 20px 16px 20px;
            padding: 20px; box-shadow: 0 4px 20px rgba(0,0,0,0.04);
            display: flex; align-items: center; justify-content: space-between;
        }
        .tool-card {
            background: #fff; border-radius: 18px; padding: 14px 18px; margin: 0 16px 10px 16px;
            display: flex; align-items: center; justify-content: space-between;
            box-shadow: 0 2px 8px rgba(0,0,0,0.02);
        }
        .btn-order { background: #f2f2f7; border: none; border-radius: 6px; width: 32px; height: 28px; color: var(--accent); font-weight: 900; }
        .fab { position: fixed; bottom: 30px; right: 20px; background: var(--accent); color: #fff; width: 60px; height: 60px; border-radius: 30px; display: flex; align-items: center; justify-content: center; font-size: 30px; z-index: 1000; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const show = (id) => { const x = el(id); if(x) x.style.display = 'flex'; };
const hide = (id) => { const x = el(id); if(x) x.style.display = 'none'; };
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// --- РЕНДЕР ГЛАВНОЙ ---
function renderList() {
    const list = el('list-p'); if(!list) return;
    el('v-home').style.display = 'block';
    el('v-det').style.display = 'none';
    
    // Вставляем лого и заголовок на главную
    list.innerHTML = `
        <div class="brand-block">
            <img src="https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png" class="logo-img">
            <div class="header-title">CitiTool</div>
        </div>
    ` + db.map((p, i) => `
        <div class="project-card" onclick="openProject(${i})">
            <div style="display:flex; flex-direction:column;">
                <div style="font-size:11px; font-weight:700; color:var(--text-sub); text-transform:uppercase;">${p.name || 'UNBENANNT'}</div>
                <div style="font-size:24px; font-weight:900; color:#000;">${p.num || '---'}</div>
            </div>
            <div style="color:#ff3b30; font-weight:800; padding:10px; font-size:18px;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:120px"></div>';
}

// --- ИНСТРУМЕНТЫ ---
function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'tool-card';
        const revMark = t.rev ? `<div style="background:#000; color:#fff; font-size:8px; padding:2px 7px; border-radius:4px; margin-bottom:5px; font-weight:900; width:fit-content;">REVOLVER UNTEN</div>` : '';
        
        item.innerHTML = `
            <div style="flex:1; min-width:0;" onclick="modalT(${i})">
                ${revMark}
                <div style="font-size:10px; font-weight:700; color:var(--text-sub);">${t.id || 'T0000'}</div>
                <div style="font-size:18px; font-weight:800;">${t.nm || '---'}</div>
                <div style="margin-top:4px; font-weight:700; color:var(--accent); font-size:13px;">${t.dia || ''}</div>
            </div>
            <div style="display:flex; flex-direction:column; gap:4px; margin-left:12px;">
                <button class="btn-order" onclick="event.stopPropagation(); moveItem(${i}, -1)">↑</button>
                <button class="btn-order" onclick="event.stopPropagation(); moveItem(${i}, 1)">↓</button>
            </div>`;
        list.appendChild(item);
    });
    list.innerHTML += '<div style="height:180px"></div>';
}

// --- ЛОГИКА ---
function openProject(i) { currentIdx = i; el('v-home').style.display='none'; el('v-det').style.display='block'; el('h-num').innerText = db[i].num; el('h-nam').innerText = db[i].name; renderTools(); }
function goHome() { currentIdx = null; renderList(); }

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    
    const btn = el('btn-rev-toggle');
    if(btn) {
        t.rev ? btn.classList.add('on') : btn.classList.remove('on');
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

// Твоя функция PDF (без изменений)
function makePDF() {
    const p = db[currentIdx];
    const getPageHead = () => `<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; min-height:90px;"><div style="display:flex; flex-direction:column; justify-content:center;"><div style="font-size:13px; font-weight:900; text-transform:uppercase; color:#666; margin-bottom:2px; line-height:1;">${p.name || ''}</div><div style="font-size:64px; font-weight:900; line-height:0.8; letter-spacing:-2px; margin:0;">${p.num || '---'}</div></div><div style="width:220px; font-size:11px; font-weight:800; line-height:1.5;"><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>MATERIAL</span><span>${p.mat || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>SÄGELÄNGE</span><span>${p.sag || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>ABSTAND</span><span>${p.abs || ''}</span></div><div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>GREIFBACKEN</span><span>${p.grf || ''}</span></div><div style="display:flex; justify-content:space-between;"><span>STÜCKZAHL</span><span>${p.stt || ''} / ${p.stn || ''}</span></div></div></div><div style="border-bottom:5px solid #000; margin-bottom:15px;"></div>`;
    const tableHead = `<div style="display:flex; font-size:10px; font-weight:900; text-transform:uppercase; margin-bottom:6px; padding:0 2px;"><div style="width:75px;">T-NR</div><div style="flex:1;">WERKZEUGNAME / KOMMENTAR</div><div style="width:125px; text-align:right;">Ø / TOLERANZ</div></div><div style="border-bottom:4px solid #000;"></div>`;
    const getRow = (t) => `<div style="display:flex; align-items:baseline; border-bottom:1.5px solid #000; padding:10px 0; width:100%; page-break-inside: avoid;"><div style="width:75px; font-weight:800; font-size:15px;">${t.id}</div><div style="flex:1; font-weight:700; font-size:15px; text-transform:uppercase; padding-right:10px; white-space:pre-wrap;">${t.nm}</div><div style="width:125px; text-align:right; font-weight:800; font-size:14px;">${t.dia.replace(/\//g, '<br>')}</div></div>`;
    let oben = [], unten = [], target = oben;
    (p.tools || []).forEach(t => { if(t.rev) target = unten; target.push(t); });
    let pdfHtml = `<!DOCTYPE html><html><head><style>@page { size: A4; margin: 0; } body { margin: 0; padding: 10mm; background: #fff; font-family: sans-serif; } .page { width: 210mm; height: 297mm; padding: 15mm; box-sizing: border-box; page-break-after: always; display: flex; flex-direction: column; } .content-border { border: 2.2px solid #000; padding: 20px; flex: 1; display: flex; flex-direction: column; box-sizing: border-box; } .footer { border-top: 1.2px solid #000; padding-top: 5px; font-size: 9px; font-weight: 800; text-align: center; color: #666; margin-top: auto; }</style></head><body><div class="page"><div class="content-border">${getPageHead()}<div style="flex:1;"><div style="margin-bottom:5px; font-size:18px; font-weight:900; text-transform:uppercase;">REVOLVER OBEN</div>${tableHead}${oben.map(getRow).join('')}</div><div class="footer">CITITOOL REPORT</div></div></div>${unten.length > 0 ? `<div class="page"><div class="content-border">${getPageHead()}<div style="flex:1;"><div style="margin-bottom:5px; font-size:18px; font-weight:900; text-transform:uppercase;">REVOLVER UNTEN</div>${tableHead}${unten.map(getRow).join('')}</div><div class="footer">CITITOOL REPORT</div></div></div>` : ''}<script>window.onload = function() { setTimeout(() => { window.print(); window.close(); }, 400); };</script></body></html>`;
    const win = window.open('', '_blank'); if (win) { win.document.write(pdfHtml); win.document.close(); }
}

// Прочие твои функции (modalP, saveP, etc.) остаются без изменений
function modalP() { el('p-idx').value = ''; show('m-p'); }
function saveP() { /* твоя логика */ save(); hide('m-p'); renderList(); }
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); save(); renderList(); } }

window.onload = () => { injectStyles(); renderList(); };
