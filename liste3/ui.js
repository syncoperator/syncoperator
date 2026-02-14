const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png'; 

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f2f5f8; 
            --neu-light: #ffffff;
            --neu-shadow: #cfd8e3;
            --accent: #007aff;
            --text: #1d1d1f;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, system-ui, sans-serif !important; 
            margin: 0; padding-bottom: 150px; color: var(--text);
            overflow-x: hidden;
        }

        /* Кнопки управления - Теперь они гарантированно сверху */
        .top-nav-actions {
            position: fixed; top: 15px; right: 15px;
            display: flex; gap: 10px; z-index: 9999;
        }
        .btn-ctrl {
            background: var(--bg); border: none; border-radius: 12px;
            padding: 8px 16px; font-size: 11px; font-weight: 800; color: #666;
            box-shadow: 4px 4px 10px var(--neu-shadow), -4px -4px 10px var(--neu-light);
            cursor: pointer; text-transform: uppercase; transition: 0.2s ease;
            pointer-events: auto !important; /* Принудительная кликабельность */
        }
        .btn-ctrl.main-blue { color: var(--accent); border: 1px solid rgba(0,122,255,0.1); }
        .btn-ctrl:active { transform: scale(0.95); box-shadow: inset 2px 2px 5px var(--neu-shadow); }

        /* Центр: Логотип и Название */
        .hero-section {
            display: flex; flex-direction: column; align-items: center;
            padding: 80px 0 40px 0; pointer-events: none; /* Пропускает клики сквозь себя к кнопкам */
        }
        .logo-large {
            width: 150px; height: 150px; object-fit: contain;
            margin-bottom: 20px; filter: drop-shadow(0 15px 25px rgba(0,0,0,0.1));
            pointer-events: auto;
        }
        .brand-wrap { text-align: center; margin-bottom: 30px; pointer-events: auto; }
        .t-top { font-size: 64px; font-weight: 900; letter-spacing: -3px; line-height: 0.9; }
        .t-refl {
            font-size: 64px; font-weight: 900; letter-spacing: -3px;
            margin-top: -28px; transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.18), transparent 80%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            opacity: 0.25; user-select: none;
        }

        /* Навигация */
        .tabs-menu {
            display: flex; background: rgba(0,0,0,0.05); 
            padding: 5px; border-radius: 16px; pointer-events: auto;
        }
        .tab-btn {
            font-size: 11px; font-weight: 800; text-transform: uppercase;
            padding: 10px 20px; border-radius: 12px; color: #888;
            cursor: pointer; transition: 0.3s;
        }
        .tab-btn.active { background: white; color: var(--accent); box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
        .tab-btn.disabled { opacity: 0.3; }

        /* Список */
        #list-p { position: relative; z-index: 10; padding: 0 20px; }
        .card-project {
            background: var(--bg); border-radius: 35px; padding: 30px;
            margin-bottom: 20px; box-shadow: 12px 12px 24px var(--neu-shadow), -12px -12px 24px var(--neu-light);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.7);
            cursor: pointer;
        }
        .p-num { font-size: 32px; font-weight: 900; color: #000; letter-spacing: -1px; }
        .p-name { font-size: 15px; font-weight: 700; color: #666; margin-top: 4px; }
        .p-tag { font-size: 10px; font-weight: 800; color: #adb5bd; text-transform: uppercase; letter-spacing: 1px; }

        .x-del { color: #ff3b30; font-size: 24px; font-weight: 900; padding: 10px; opacity: 0.6; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

function renderList() {
    const main = el('v-home'); if(!main) return;
    
    // Кнопки вынесены в отдельный контейнер top-nav-actions
    main.innerHTML = `
        <div class="top-nav-actions">
            <button class="btn-ctrl" onclick="exportJSON()">Export</button>
            <button class="btn-ctrl" onclick="importJSON()">Import</button>
            <button class="btn-ctrl main-blue" onclick="modalP()">+ NEU</button>
        </div>

        <div class="hero-section">
            <img src="${LOGO_URL}" class="logo-large" alt="Logo">
            <div class="brand-wrap">
                <div class="t-top">CitiTool</div>
                <div class="t-refl">CitiTool</div>
            </div>
            
            <div class="tabs-menu">
                <div class="tab-btn active" onclick="goHome()">Home</div>
                <div class="tab-btn disabled">SyncOP</div>
                <div class="tab-btn disabled">WKZListe</div>
                <div class="tab-btn disabled">Stange</div>
            </div>
        </div>
        
        <div id="list-p"></div>
    `;
    
    const list = el('list-p');
    list.innerHTML = db.map((p, i) => `
        <div class="card-project" onclick="openProject(${i})">
            <div>
                <div class="p-tag">PROJEKT</div>
                <div class="p-num">${p.num || '---'}</div>
                <div class="p-name">${p.name || ''}</div>
            </div>
            <div class="x-del" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="padding-bottom:100px;"></div>';
}

// JSON Handlers
window.exportJSON = function() {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(db));
    const dlAnchorElem = document.createElement('a');
    dlAnchorElem.setAttribute("href", dataStr);
    dlAnchorElem.setAttribute("download", "cititool_v8.json");
    dlAnchorElem.click();
};

window.importJSON = function() {
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
};

function goHome() { 
    currentIdx = null; 
    if(el('v-home')) el('v-home').style.display='block'; 
    if(el('v-det')) el('v-det').style.display='none'; 
    renderList(); 
}

function deleteProject(i) {
    if(confirm('Löschen?')) { db.splice(i, 1); save(); renderList(); }
}

// Инициализация
window.onload = () => {
    injectStyles();
    renderList();
};
