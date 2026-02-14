const DB_KEY = 'QS_DATA_V8';
// Вставь сюда свою прямую ссылку на логотип (Raw link с GitHub)
const LOGO_URL = 'https://raw.githubusercontent.com/USER/REPO/main/logo.png'; 

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f5f7fa; 
            --neu-light: #ffffff;
            --neu-shadow: #d1d9e6;
            --accent: #007aff;
            --text: #1a1a1a;
            --glass: rgba(255, 255, 255, 0.7);
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important; 
            margin: 0; padding-top: 70px; padding-bottom: 120px; color: var(--text);
            -webkit-font-smoothing: antialiased;
        }

        /* Премиальное горизонтальное меню */
        .top-nav {
            position: fixed; top: 0; left: 0; right: 0; height: 65px;
            background: var(--glass); backdrop-filter: blur(15px);
            display: flex; align-items: center; justify-content: center;
            z-index: 1000; border-bottom: 1px solid rgba(255,255,255,0.4);
            box-shadow: 0 4px 20px rgba(0,0,0,0.03);
        }
        .nav-inner {
            display: flex; gap: 25px; font-size: 13px; font-weight: 700;
            text-transform: uppercase; letter-spacing: 0.5px;
        }
        .nav-item { 
            color: #8e8e93; cursor: pointer; transition: 0.3s; 
            text-decoration: none;
        }
        .nav-item:hover, .nav-item.active { color: var(--accent); }
        .nav-item.disabled { opacity: 0.4; cursor: not-allowed; }

        /* Логотип и Зеркало */
        .brand-header {
            padding: 40px 0 20px 0;
            display: flex; flex-direction: column; align-items: center;
        }
        .main-logo {
            width: 100px; height: 100px; object-fit: contain;
            margin-bottom: 12px; filter: drop-shadow(0 10px 15px rgba(0,0,0,0.05));
        }
        .mirror-title { position: relative; text-align: center; line-height: 1; }
        .text-top {
            font-size: 52px; font-weight: 900; letter-spacing: -2.5px;
            color: var(--text); position: relative; z-index: 2;
        }
        .text-bottom {
            font-size: 52px; font-weight: 900; letter-spacing: -2.5px;
            margin-top: -22px; transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.18) 0%, rgba(245,247,250,1) 90%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            opacity: 0.3; user-select: none;
        }

        /* Управление и кнопки */
        .action-bar { 
            display: flex; justify-content: flex-end; align-items: center;
            padding: 0 25px; margin: 20px 0; gap: 15px;
        }
        .btn-neu {
            background: var(--bg); border: none; border-radius: 14px;
            padding: 12px 24px; font-size: 13px; font-weight: 800; color: var(--accent);
            box-shadow: 5px 5px 12px var(--neu-shadow), -5px -5px 12px var(--neu-light);
            cursor: pointer;
        }
        .btn-neu:active { box-shadow: inset 3px 3px 6px var(--neu-shadow), inset -3px -3px 6px var(--neu-light); }
        .btn-data { font-size: 11px; padding: 8px 15px; color: #8e8e93; }

        /* Карточки проекта */
        .neu-card {
            background: var(--bg); border-radius: 28px; padding: 25px;
            margin: 15px 20px; box-shadow: 10px 10px 22px var(--neu-shadow), -10px -10px 22px var(--neu-light);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.7);
            cursor: pointer;
        }
        .p-label { font-size: 10px; font-weight: 800; color: #adb5bd; text-transform: uppercase; letter-spacing: 1.2px; }
        .p-title { font-size: 26px; font-weight: 900; color: #1c1c1e; margin: 4px 0; }
        .p-desc { font-size: 13px; font-weight: 600; color: #636366; }

        .btn-del { color: #ff3b30; font-size: 22px; font-weight: 900; padding: 10px; opacity: 0.8; }
        .btn-del:hover { opacity: 1; }

        /* Список инструментов и Drag-n-Drop */
        .tools-list { padding-bottom: 50px; }
        .dragging { opacity: 0.5; transform: scale(0.98); }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// Рендерим меню
const renderNav = (active = 'Home') => {
    return `
    <nav class="top-nav">
        <div class="nav-inner">
            <span class="nav-item ${active === 'Home' ? 'active' : ''}" onclick="goHome()">Home</span>
            <span class="nav-item disabled">SyncOP</span>
            <span class="nav-item disabled">WKZListe</span>
            <span class="nav-item disabled">Stange</span>
            <span class="nav-item disabled" style="font-size:10px">Coming Soon</span>
        </div>
    </nav>`;
};

function renderList() {
    const main = el('v-home'); if(!main) return;
    main.innerHTML = `
        ${renderNav('Home')}
        <div class="brand-header">
            <img src="${LOGO_URL}" class="main-logo" alt="Logo">
            <div class="mirror-title">
                <div class="text-top">CitiTool</div>
                <div class="text-bottom">CitiTool</div>
            </div>
        </div>
        <div class="action-bar">
            <button class="btn-neu btn-data" onclick="exportJSON()">Export</button>
            <button class="btn-neu btn-data" onclick="importJSON()">Import</button>
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
                <div class="p-desc">${p.name || ''}</div>
            </div>
            <div class="btn-del" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:100px"></div>';
}

// Функции управления данными
function exportJSON() {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(db));
    const dlAnchorElem = document.createElement('a');
    dlAnchorElem.setAttribute("href", dataStr);
    dlAnchorElem.setAttribute("download", "cititool_backup.json");
    dlAnchorElem.click();
}

function importJSON() {
    const input = document.createElement('input');
    input.type = 'file';
    input.onchange = e => {
        const file = e.target.files[0];
        const reader = new FileReader();
        reader.onload = readerEvent => {
            const content = readerEvent.target.result;
            db = JSON.parse(content);
            save(); renderList();
        };
        reader.readAsText(file);
    };
    input.click();
}

// Навигация и база
function goHome() { 
    currentIdx = null; 
    el('v-home').style.display = 'block'; 
    el('v-det').style.display = 'none'; 
    renderList(); 
}

function openProject(i) {
    currentIdx = i;
    el('v-home').style.display = 'none';
    el('v-det').style.display = 'block';
    el('h-num').innerText = db[i].num;
    el('h-nam').innerText = db[i].name;
    // Очистка и рендер инструментов будет тут
}

function deleteProject(i) {
    if(confirm('Projekt löschen?')) {
        db.splice(i, 1);
        save();
        renderList();
    }
}

// Инициализация
window.onload = () => {
    injectStyles();
    renderList();
};
