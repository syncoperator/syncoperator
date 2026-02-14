const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png'; 

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
            --neu-shadow: #cfd8e3;
            --neu-light: #ffffff;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important; 
            margin: 0; color: var(--text-main); padding-bottom: 120px;
        }
        
        /* Фиксированная шапка с кнопками */
        .top-bar {
            background: rgba(255,255,255,0.8); backdrop-filter: blur(15px);
            padding: 15px 20px; position: sticky; top: 0; z-index: 1000;
            display: flex; justify-content: flex-end; gap: 10px;
            border-bottom: 0.5px solid rgba(0,0,0,0.05);
        }

        .btn-premium {
            background: var(--bg); border: none; border-radius: 12px;
            padding: 10px 18px; font-size: 11px; font-weight: 800; color: #666;
            box-shadow: 4px 4px 8px var(--neu-shadow), -4px -4px 8px var(--neu-light);
            cursor: pointer; text-transform: uppercase;
        }
        .btn-premium.blue { background: var(--accent); color: white; box-shadow: 0 4px 12px rgba(0,122,255,0.3); }

        /* Центр: Лого и Зеркало */
        .hero-section { display: flex; flex-direction: column; align-items: center; padding: 20px 0 40px; }
        .logo-main { width: 200px; height: 200px; object-fit: contain; filter: drop-shadow(0 15px 30px rgba(0,0,0,0.1)); }
        
        .mirror-box { text-align: center; margin-top: 15px; }
        .t-high { font-size: 72px; font-weight: 900; letter-spacing: -4px; color: #1d1d1f; line-height: 0.8; }
        .t-refl { 
            font-size: 72px; font-weight: 900; letter-spacing: -4px; margin-top: -32px; 
            transform: scaleY(-1); opacity: 0.15;
            background: linear-gradient(to bottom, #000, transparent);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }

        /* Твои оригинальные карточки проектов */
        .project-card {
            background: var(--card-bg); border-radius: 24px;
            margin: 0 20px 16px 20px; padding: 20px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.04);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.5);
        }
        
        .p-info { display: flex; flex-direction: column; }
        .p-label { font-size: 11px; font-weight: 700; color: var(--text-sub); text-transform: uppercase; margin-bottom: 4px; }
        .p-title { font-size: 24px; font-weight: 900; color: #000; letter-spacing: -0.5px; }

        /* Инструменты */
        .tool-card {
            background: #fff; border-radius: 18px;
            padding: 14px 18px; margin: 0 16px 10px 16px;
            display: flex; align-items: center; justify-content: space-between;
            box-shadow: 0 2px 8px rgba(0,0,0,0.02);
        }

        .btn-delete { color: #ff3b30; font-weight: 800; padding: 10px; font-size: 18px; }
        .fab {
            position: fixed; bottom: 30px; right: 20px;
            background: var(--accent); color: #fff; width: 60px; height: 60px;
            border-radius: 30px; display: flex; align-items: center; justify-content: center;
            font-size: 30px; box-shadow: 0 8px 25px rgba(0,122,255,0.3); z-index: 1000; border: none;
        }
        .hidden { display: none !important; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// --- ГЛАВНАЯ СТРАНИЦА ---
function renderList() {
    const vHome = el('v-home');
    const vDet = el('v-det');
    if(!vHome) return;

    vHome.classList.add('active'); vHome.classList.remove('hidden');
    vDet.classList.remove('active'); vDet.classList.add('hidden');

    vHome.innerHTML = `
        <div class="top-bar">
            <button class="btn-premium" onclick="exportJSON()">Export</button>
            <button class="btn-premium" onclick="importJSON()">Import</button>
            <button class="btn-premium blue" onclick="modalP()">+ NEU</button>
        </div>
        <div class="hero-section">
            <img src="${LOGO_URL}" class="logo-main">
            <div class="mirror-box">
                <div class="t-high">CitiTool</div>
                <div class="t-refl">CitiTool</div>
            </div>
        </div>
        <div id="list-p"></div>
    `;

    const list = el('list-p');
    list.innerHTML = db.map((p, i) => `
        <div class="project-card" onclick="openProject(${i})">
            <div class="p-info">
                <div class="p-label">${p.name || 'UNBENANNT'}</div>
                <div class="p-title">${p.num || '---'}</div>
            </div>
            <div class="btn-delete" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:120px"></div>';
}

// --- ИНСТРУМЕНТЫ (Твоя логика без изменений) ---
function renderTools() {
    const vHome = el('v-home');
    const vDet = el('v-det');
    const p = db[currentIdx];

    vHome.classList.remove('active'); vHome.classList.add('hidden');
    vDet.classList.add('active'); vDet.classList.remove('hidden');

    vDet.innerHTML = `
        <div class="top-bar">
            <button class="btn-premium" onclick="goHome()">← BACK</button>
            <button class="btn-premium blue" onclick="makePDF()">PDF REPORT</button>
        </div>
        <div class="hero-section">
            <div class="mirror-box">
                <div class="t-high">${p.num}</div>
                <div style="font-weight:700; color:#666;">${p.name}</div>
            </div>
        </div>
        <div id="list-t"></div>
        <button class="fab" onclick="modalT()">+</button>
    `;

    const list = el('list-t');
    const tools = p.tools || [];
    list.innerHTML = tools.map((t, i) => `
        <div class="tool-card" onclick="modalT(${i})">
            <div style="flex:1;">
                ${t.rev ? `<div style="background:#000; color:#fff; font-size:8px; padding:2px 7px; border-radius:4px; margin-bottom:5px; font-weight:900; width:fit-content;">REVOLVER UNTEN</div>` : ''}
                <div style="font-size:10px; font-weight:700; color:var(--text-sub); text-transform:uppercase;">${t.id || 'T0000'}</div>
                <div style="font-size:18px; font-weight:800; color:#000;">${t.nm || '---'}</div>
                <div style="margin-top:4px; font-weight:700; color:var(--accent); font-size:13px;">${t.dia || ''}</div>
            </div>
        </div>`).join('') + '<div style="height:180px"></div>';
}

// --- ТВОИ ОРИГИНАЛЬНЫЕ ФУНКЦИИ УПРАВЛЕНИЯ ---
function openProject(i) { currentIdx = i; renderTools(); }
function goHome() { currentIdx = null; renderList(); }
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); save(); renderList(); } }

function exportJSON() { alert("JSON kopiert!"); console.log(JSON.stringify(db)); }
function importJSON() { const val = prompt("JSON сюда:"); if(val) { db = JSON.parse(val); save(); renderList(); } }

// Модалки (простые для стабильности)
function modalP() {
    const n = prompt("Projekt Nummer:");
    const m = prompt("Name:");
    if(n) { db.push({num:n, name:m, tools:[]}); save(); renderList(); }
}

function modalT(i = null) {
    const edit = i !== null;
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    const nid = prompt("T-NR:", t.id);
    const nnm = prompt("Name:", t.nm);
    const ndia = prompt("Dia:", t.dia);
    const nrev = confirm("Unten? (OK=Unten, Cancel=Oben)");
    if(nid) {
        const data = {id:nid, nm:nnm, dia:ndia, rev:nrev};
        if(edit) db[currentIdx].tools[i] = data; else db[currentIdx].tools.push(data);
        save(); renderTools();
    }
}

function makePDF() {
    const p = db[currentIdx];
    let pdfHtml = `<html><body style="padding:40px;font-family:sans-serif;"><h1>${p.num} - ${p.name}</h1><hr/>`;
    (p.tools || []).forEach(t => {
        pdfHtml += `<p><b>${t.id}</b>: ${t.nm} (${t.dia}) [${t.rev ? 'UNTEN' : 'OBEN'}]</p>`;
    });
    pdfHtml += `</body></html>`;
    const win = window.open('', '_blank'); win.document.write(pdfHtml); win.document.close();
}

window.onload = () => { injectStyles(); renderList(); };
