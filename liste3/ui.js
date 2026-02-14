const DB_KEY = 'QS_DATA_V8';
// Ссылка на твой логотип с GitHub
const LOGO_URL = 'https://raw.githubusercontent.com/USER/REPO/main/logo.png'; 

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
            --text: #1d1d1f;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, BlinkMacSystemFont, sans-serif !important; 
            margin: 0; padding-bottom: 150px; color: var(--text);
            -webkit-font-smoothing: antialiased;
        }

        /* Кнопки в верхнем правом углу */
        .top-right-actions {
            position: absolute; top: 20px; right: 20px;
            display: flex; gap: 10px; z-index: 1001;
        }
        .btn-action {
            background: var(--bg); border: none; border-radius: 10px;
            padding: 8px 14px; font-size: 11px; font-weight: 800; color: var(--accent);
            box-shadow: 4px 4px 8px var(--neu-shadow), -4px -4px 8px var(--neu-light);
            cursor: pointer; text-transform: uppercase;
        }
        .btn-action.main { background: var(--accent); color: white; box-shadow: 0 4px 10px rgba(0,122,255,0.3); }
        .btn-action:active { box-shadow: inset 2px 2px 5px var(--neu-shadow), inset -2px -2px 5px var(--neu-light); }

        /* Центрированная шапка */
        .premium-header {
            display: flex; flex-direction: column; align-items: center;
            padding: 60px 0 30px 0;
        }
        .main-logo {
            width: 110px; height: 110px; object-fit: contain;
            margin-bottom: 15px; filter: drop-shadow(0 10px 20px rgba(0,0,0,0.06));
        }
        .brand-name { text-align: center; margin-bottom: 30px; }
        .text-top { font-size: 58px; font-weight: 900; letter-spacing: -3px; line-height: 1; }
        .text-bottom {
            font-size: 58px; font-weight: 900; letter-spacing: -3px;
            margin-top: -25px; transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.15), transparent 90%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            opacity: 0.2; user-select: none;
        }

        /* Меню под названием */
        .nav-bar {
            display: flex; background: rgba(0,0,0,0.04); 
            padding: 4px; border-radius: 12px;
        }
        .nav-item {
            font-size: 10px; font-weight: 800; text-transform: uppercase;
            padding: 10px 16px; border-radius: 9px; color: #86868b;
            cursor: pointer; transition: 0.2s;
        }
        .nav-item.active { background: white; color: var(--accent); box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
        .nav-item.disabled { opacity: 0.4; pointer-events: none; }

        /* Карточки проектов */
        .project-card {
            background: var(--bg); border-radius: 28px; padding: 22px 25px;
            margin: 15px 20px; box-shadow: 10px 10px 20px var(--neu-shadow), -10px -10px 20px var(--neu-light);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.7);
            cursor: pointer; position: relative; z-index: 10;
        }
        .p-info .label { font-size: 9px; font-weight: 900; color: #adb5bd; letter-spacing: 1px; }
        .p-info .title { font-size: 26px; font-weight: 900; color: #000; margin: 2px 0; }
        .p-info .sub { font-size: 13px; font-weight: 700; color: #636366; }

        .btn-del { color: #ff3b30; font-size: 20px; font-weight: 900; padding: 10px; z-index: 11; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

function renderList() {
    const main = el('v-home'); if(!main) return;
    main.innerHTML = `
        <div class="top-right-actions">
            <button class="btn-action" onclick="exportJSON()">Export</button>
            <button class="btn-action" onclick="importJSON()">Import</button>
            <button class="btn-action main" onclick="modalP()">+ NEU</button>
        </div>

        <div class="premium-header">
            <img src="${LOGO_URL}" class="main-logo" alt="Logo">
            <div class="brand-name">
                <div class="text-top">CitiTool</div>
                <div class="text-bottom">CitiTool</div>
            </div>
            
            <div class="nav-bar">
                <div class="nav-item active" onclick="goHome()">Home</div>
                <div class="nav-item disabled">SyncOP</div>
                <div class="nav-item disabled">WKZListe</div>
                <div class="nav-item disabled">Stange</div>
                <div class="nav-item disabled">Coming Soon</div>
            </div>
        </div>
        
        <div id="list-p"></div>
    `;
    
    const list = el('list-p');
    list.innerHTML = db.map((p, i) => `
        <div class="project-card" onclick="openProject(${i})">
            <div class="p-info">
                <div class="label">PROJEKT</div>
                <div class="title">${p.num || '---'}</div>
                <div class="sub">${p.name || ''}</div>
            </div>
            <div class="btn-del" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:100px"></div>';
}

// Работа с JSON
function exportJSON() {
    console.log("Exporting...");
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(db));
    const dlAnchorElem = document.createElement('a');
    dlAnchorElem.setAttribute("href", dataStr);
    dlAnchorElem.setAttribute("download", "cititool_v8.json");
    dlAnchorElem.click();
}

function importJSON() {
    const input = document.createElement('input');
    input.type = 'file';
    input.onchange = e => {
        const reader = new FileReader();
        reader.onload = r => {
            db = JSON.parse(r.target.result);
            save(); renderList();
        };
        reader.readAsText(e.target.files[0]);
    };
    input.click();
}

// Навигация
function goHome() { 
    currentIdx = null; 
    const vHome = el('v-home');
    const vDet = el('v-det');
    if(vHome) vHome.style.display='block'; 
    if(vDet) vDet.style.display='none'; 
    renderList(); 
}

function deleteProject(i) {
    if(confirm('Projekt löschen?')) {
        db.splice(i, 1);
        save();
        renderList();
    }
}

// Базовый вызов
window.onload = () => {
    injectStyles();
    renderList();
};
