const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png'; 

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    if (document.getElementById('main-styles')) return;
    const style = document.createElement('style');
    style.id = 'main-styles';
    style.innerHTML = `
        :root { 
            --bg: #f2f5f8; 
            --neu-light: #ffffff;
            --neu-shadow: #cfd8e3;
            --accent: #007aff;
        }
        body { 
            background: var(--bg); 
            font-family: -apple-system, system-ui, sans-serif; 
            margin: 0; padding: 0;
            user-select: none; -webkit-tap-highlight-color: transparent;
        }

        /* Кнопки управления - всегда сверху и доступны */
        .nav-header {
            position: sticky; top: 0; width: 100%; height: 80px;
            display: flex; justify-content: flex-end; align-items: center;
            padding: 0 20px; box-sizing: border-box; z-index: 999;
            background: rgba(242, 245, 248, 0.8); backdrop-filter: blur(10px);
        }
        .btn-ui {
            background: var(--bg); border: none; border-radius: 12px;
            padding: 10px 18px; font-size: 11px; font-weight: 800; color: #666;
            box-shadow: 4px 4px 8px var(--neu-shadow), -4px -4px 8px var(--neu-light);
            cursor: pointer; text-transform: uppercase; margin-left: 10px;
        }
        .btn-ui.blue { background: var(--accent); color: white; box-shadow: 0 4px 12px rgba(0,122,255,0.3); }
        .btn-ui:active { transform: scale(0.95); box-shadow: inset 2px 2px 5px var(--neu-shadow); }

        /* Брендинг */
        .hero { display: flex; flex-direction: column; align-items: center; padding: 20px 0 40px 0; }
        .logo-huge {
            width: 200px; height: 200px; object-fit: contain;
            filter: drop-shadow(0 15px 30px rgba(0,0,0,0.1));
        }
        .title-group { text-align: center; margin-top: 15px; }
        .t-main { font-size: 72px; font-weight: 900; letter-spacing: -4px; color: #1d1d1f; line-height: 0.8; }
        .t-refl {
            font-size: 72px; font-weight: 900; letter-spacing: -4px;
            margin-top: -32px; transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.15), transparent 85%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            opacity: 0.2;
        }

        /* Список */
        .content { padding: 0 20px; max-width: 600px; margin: 0 auto; }
        .p-card {
            background: var(--bg); border-radius: 30px; padding: 25px;
            margin-bottom: 20px; box-shadow: 10px 10px 20px var(--neu-shadow), -10px -10px 20px var(--neu-light);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.7); cursor: pointer;
        }
        .p-num { font-size: 32px; font-weight: 900; color: #000; letter-spacing: -1px; }
        .p-name { font-size: 15px; font-weight: 700; color: #666; margin-top: 2px; }
        .btn-del { color: #ff3b30; font-size: 22px; font-weight: 900; padding: 5px; opacity: 0.5; }

        .hidden { display: none !important; }

        /* Revolver Styling */
        .tool-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-top: 20px; }
        .tool-card { 
            background: var(--bg); border-radius: 20px; padding: 15px;
            box-shadow: 6px 6px 12px var(--neu-shadow), -6px -6px 12px var(--neu-light);
        }
        .tool-t { font-size: 10px; font-weight: 800; color: var(--accent); }
        .tool-desc { font-size: 16px; font-weight: 900; color: #333; margin-top: 4px; }
    `;
    document.head.appendChild(style);
};

// --- CORE FUNCTIONS ---
window.save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

window.modalP = () => {
    const n = prompt("Projekt Nummer:");
    const m = prompt("Name:");
    if(n) {
        db.push({ num: n, name: m, tools: [] });
        window.save();
        window.renderList();
    }
};

window.openProject = (i) => {
    currentIdx = i;
    window.renderDetails();
};

window.goHome = () => {
    currentIdx = null;
    window.renderList();
};

window.exportJSON = () => {
    const data = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(db));
    const a = document.createElement('a');
    a.href = data; a.download = "cititool_save.json"; a.click();
};

window.importJSON = () => {
    const i = document.createElement('input'); i.type = 'file';
    i.onchange = e => {
        const r = new FileReader();
        r.onload = f => { db = JSON.parse(f.target.result); window.save(); window.renderList(); };
        r.readAsText(e.target.files[0]);
    };
    i.click();
};

// --- RENDERING ---
window.renderList = () => {
    const vHome = document.getElementById('v-home');
    const vDet = document.getElementById('v-det');
    if(!vHome || !vDet) return;

    vDet.classList.add('hidden');
    vHome.classList.remove('hidden');

    vHome.innerHTML = `
        <div class="nav-header">
            <button class="btn-ui" onclick="window.exportJSON()">Export</button>
            <button class="btn-ui" onclick="window.importJSON()">Import</button>
            <button class="btn-ui blue" onclick="window.modalP()">+ NEU</button>
        </div>
        <div class="hero">
            <img src="${LOGO_URL}" class="logo-huge">
            <div class="title-group">
                <div class="t-main">CitiTool</div>
                <div class="t-refl">CitiTool</div>
            </div>
        </div>
        <div class="content" id="list-target"></div>
    `;

    document.getElementById('list-target').innerHTML = db.map((p, i) => `
        <div class="p-card" onclick="window.openProject(${i})">
            <div>
                <div style="font-size:9px; font-weight:800; color:#adb5bd; letter-spacing:1px;">PROJEKT</div>
                <div class="p-num">${p.num || '---'}</div>
                <div class="p-name">${p.name || ''}</div>
            </div>
            <div class="btn-del" onclick="event.stopPropagation(); window.deleteP(${i})">✕</div>
        </div>
    `).join('') + '<div style="height:100px;"></div>';
};

window.renderDetails = () => {
    const vHome = document.getElementById('v-home');
    const vDet = document.getElementById('v-det');
    const p = db[currentIdx];

    vHome.classList.add('hidden');
    vDet.classList.remove('hidden');

    vDet.innerHTML = `
        <div class="nav-header">
            <button class="btn-ui" onclick="window.goHome()">← Home</button>
            <button class="btn-ui blue" onclick="window.addTool()">+ Tool</button>
        </div>
        <div class="hero">
            <div class="title-group">
                <div class="t-main">${p.num}</div>
                <div style="font-weight:700; color:#666; margin-top:5px;">${p.name}</div>
            </div>
        </div>
        <div class="content">
            <div class="tool-grid" id="tools-target"></div>
        </div>
    `;
    // Тут будет отрисовка инструментов Revolver Oben/Unten
};

window.deleteP = (i) => { if(confirm('Löschen?')) { db.splice(i,1); window.save(); window.renderList(); } };

// Инициализация
window.onload = () => {
    injectStyles();
    window.renderList();
};
