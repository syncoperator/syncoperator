const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png'; 

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f2f5f8; 
            --accent: #007aff; 
            --neu-light: #ffffff;
            --neu-shadow: #cfd8e3;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, system-ui, sans-serif !important; 
            margin: 0; padding-bottom: 120px;
        }
        
        /* Верхняя панель управления */
        .top-nav {
            position: sticky; top: 0; z-index: 1000;
            background: rgba(242, 245, 248, 0.8); backdrop-filter: blur(15px);
            display: flex; justify-content: flex-end; padding: 15px 20px; gap: 12px;
        }

        /* Кнопки Premium */
        .btn-premium {
            background: var(--bg); border: none; border-radius: 12px;
            padding: 10px 18px; font-size: 11px; font-weight: 800; color: #666;
            box-shadow: 4px 4px 8px var(--neu-shadow), -4px -4px 8px var(--neu-light);
            cursor: pointer; text-transform: uppercase; transition: 0.2s;
        }
        .btn-premium.blue { background: var(--accent); color: white; box-shadow: 0 4px 12px rgba(0,122,255,0.3); }
        .btn-premium:active { transform: scale(0.95); }

        /* Центр: Логотип и Зеркало */
        .hero { display: flex; flex-direction: column; align-items: center; padding: 20px 0 40px; }
        .logo-huge { width: 200px; height: 200px; object-fit: contain; filter: drop-shadow(0 15px 30px rgba(0,0,0,0.1)); }
        
        .t-wrap { text-align: center; margin-top: 15px; }
        .t-main { font-size: 72px; font-weight: 900; letter-spacing: -4px; color: #1d1d1f; line-height: 0.8; }
        .t-refl { 
            font-size: 72px; font-weight: 900; letter-spacing: -4px; margin-top: -32px; 
            transform: scaleY(-1); opacity: 0.15;
            background: linear-gradient(to bottom, #000, transparent);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }

        /* Карточки Проектов */
        .project-card {
            background: var(--bg); border-radius: 30px; padding: 25px;
            margin: 0 20px 20px; box-shadow: 10px 10px 20px var(--neu-shadow), -10px -10px 20px var(--neu-light);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.7); cursor: pointer;
        }
        .p-title { font-size: 32px; font-weight: 900; color: #000; letter-spacing: -1px; }
        .p-name { font-size: 15px; font-weight: 700; color: #666; }

        /* Инструменты */
        .tool-card {
            background: #fff; border-radius: 20px; padding: 16px; margin: 0 16px 12px;
            display: flex; align-items: center; justify-content: space-between;
            box-shadow: 0 4px 15px rgba(0,0,0,0.03);
        }
        .t-id { font-size: 11px; font-weight: 800; color: var(--accent); }
        .t-nm { font-size: 19px; font-weight: 900; color: #000; }

        .hidden { display: none !important; }
        .fab { position: fixed; bottom: 30px; right: 20px; width: 64px; height: 64px; border-radius: 32px; background: var(--accent); color: white; display: flex; align-items: center; justify-content: center; font-size: 32px; box-shadow: 0 10px 25px rgba(0,122,255,0.3); z-index: 100; border: none; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// --- ЛОГИКА ---
function renderList() {
    const vHome = el('v-home');
    const vDet = el('v-det');
    if(!vHome) return;
    
    vHome.classList.remove('hidden');
    vDet.classList.add('hidden');

    vHome.innerHTML = `
        <div class="top-nav">
            <button class="btn-premium" onclick="exportJSON()">Export</button>
            <button class="btn-premium" onclick="importJSON()">Import</button>
            <button class="btn-premium blue" onclick="modalP()">+ NEU</button>
        </div>
        <div class="hero">
            <img src="${LOGO_URL}" class="logo-huge">
            <div class="t-wrap">
                <div class="t-main">CitiTool</div>
                <div class="t-refl">CitiTool</div>
            </div>
        </div>
        <div id="list-p"></div>
    `;

    el('list-p').innerHTML = db.map((p, i) => `
        <div class="project-card" onclick="openProject(${i})">
            <div>
                <div style="font-size:10px; font-weight:800; color:#adb5bd; letter-spacing:1px;">PROJEKT</div>
                <div class="p-title">${p.num || '---'}</div>
                <div class="p-name">${p.name || ''}</div>
            </div>
            <div style="color:#ff3b30; font-size:24px; font-weight:900;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:100px"></div>';
}

function openProject(i) {
    currentIdx = i;
    el('v-home').classList.add('hidden');
    el('v-det').classList.remove('hidden');
    renderTools();
}

function goHome() {
    currentIdx = null;
    renderList();
}

function renderTools() {
    const p = db[currentIdx];
    const vDet = el('v-det');
    vDet.innerHTML = `
        <div class="top-nav">
            <button class="btn-premium" onclick="goHome()">← BACK</button>
            <button class="btn-premium blue" onclick="makePDF()">PDF</button>
        </div>
        <div class="hero">
            <div class="t-wrap">
                <div class="t-main">${p.num}</div>
                <div style="font-weight:700; color:#666;">${p.name}</div>
            </div>
        </div>
        <div id="list-t"></div>
        <button class="fab" onclick="modalT()">+</button>
    `;

    const listT = el('list-t');
    const tools = p.tools || [];
    listT.innerHTML = tools.map((t, i) => `
        <div class="tool-card" onclick="modalT(${i})">
            <div>
                ${t.rev ? '<div style="background:#000; color:#fff; font-size:8px; padding:2px 6px; border-radius:4px; font-weight:900; width:fit-content; margin-bottom:4px;">REVOLVER UNTEN</div>' : ''}
                <div class="t-id">${t.id || 'T0000'}</div>
                <div class="t-nm">${t.nm || '---'}</div>
                <div style="color:var(--accent); font-weight:700; font-size:13px;">${t.dia || ''}</div>
            </div>
            <div style="font-weight:900; color:#ccc;">ID:${i}</div>
        </div>
    `).join('') + '<div style="height:150px"></div>';
}

// --- СИСТЕМНЫЕ ФУНКЦИИ (Твои оригинальные) ---
function modalP() {
    const num = prompt("Projekt Nummer:");
    const name = prompt("Name:");
    if(num) {
        db.push({num, name: name.toUpperCase(), tools: []});
        save(); renderList();
    }
}

function modalT(i = null) {
    const edit = i !== null;
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    
    const newId = prompt("T-NR:", t.id);
    const newNm = prompt("Name:", t.nm);
    const newDia = prompt("Ø / Toleranz:", t.dia);
    const isUnten = confirm("Revolver Unten? (Abbrechen = Oben)");

    if(newId) {
        const data = {id: newId.toUpperCase(), nm: newNm.toUpperCase(), dia: newDia, rev: isUnten};
        if(edit) db[currentIdx].tools[i] = data; else db[currentIdx].tools.push(data);
        save(); renderTools();
    }
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); save(); renderList(); } }

function exportJSON() {
    const data = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(db));
    const a = document.createElement('a'); a.href = data; a.download = "cititool_v8.json"; a.click();
}

function importJSON() {
    const inp = document.createElement('input'); inp.type = 'file';
    inp.onchange = e => {
        const r = new FileReader();
        r.onload = f => { db = JSON.parse(f.target.result); save(); renderList(); };
        r.readAsText(e.target.files[0]);
    };
    inp.click();
}

// PDF (Твой оригинальный код)
function makePDF() {
    const p = db[currentIdx];
    const getPageHead = () => `<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; min-height:90px;"><div style="display:flex; flex-direction:column; justify-content:center;"><div style="font-size:13px; font-weight:900; text-transform:uppercase; color:#666; margin-bottom:2px; line-height:1;">${p.name || ''}</div><div style="font-size:64px; font-weight:900; line-height:0.8; letter-spacing:-2px; margin:0;">${p.num || '---'}</div></div></div><div style="border-bottom:5px solid #000; margin-bottom:15px;"></div>`;
    const tableHead = `<div style="display:flex; font-size:10px; font-weight:900; text-transform:uppercase; margin-bottom:6px; padding:0 2px;"><div style="width:75px;">T-NR</div><div style="flex:1;">WERKZEUGNAME</div><div style="width:125px; text-align:right;">Ø / TOLERANZ</div></div><div style="border-bottom:4px solid #000;"></div>`;
    const getRow = (t) => `<div style="display:flex; align-items:baseline; border-bottom:1.5px solid #000; padding:10px 0; width:100%; page-break-inside: avoid;"><div style="width:75px; font-weight:800; font-size:15px;">${t.id}</div><div style="flex:1; font-weight:700; font-size:15px; text-transform:uppercase;">${t.nm}</div><div style="width:125px; text-align:right; font-weight:800; font-size:14px;">${t.dia}</div></div>`;
    
    let oben = [], unten = [];
    (p.tools || []).forEach(t => { if(t.rev) unten.push(t); else oben.push(t); });

    let pdfHtml = `<html><head><style>body { padding: 15mm; font-family: sans-serif; }</style></head><body>${getPageHead()}<h2>REVOLVER OBEN</h2>${tableHead}${oben.map(getRow).join('')}${unten.length > 0 ? `<h2>REVOLVER UNTEN</h2>${tableHead}${unten.map(getRow).join('')}` : ''}<script>window.onload=function(){window.print();}</script></body></html>`;
    const win = window.open('', '_blank'); win.document.write(pdfHtml); win.document.close();
}

window.onload = () => { injectStyles(); renderList(); };
