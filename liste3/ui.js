const DB_KEY = 'QS_DATA_V8';
// Ссылка на твой логотип с GitHub
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
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, system-ui, sans-serif !important; 
            margin: 0; padding-bottom: 120px; color: var(--text);
            -webkit-font-smoothing: antialiased;
        }

        /* Главный блок: Лого + Заголовок + Меню */
        .header-container {
            padding: 60px 0 30px 0;
            display: flex; flex-direction: column; align-items: center;
            background: linear-gradient(to bottom, rgba(255,255,255,0.5), transparent);
        }
        .main-logo {
            width: 120px; height: 120px; object-fit: contain;
            margin-bottom: 15px; filter: drop-shadow(0 10px 20px rgba(0,0,0,0.08));
        }
        .mirror-title { position: relative; text-align: center; margin-bottom: 35px; }
        .text-top {
            font-size: 62px; font-weight: 900; letter-spacing: -3px;
            color: var(--text); position: relative; z-index: 2; line-height: 1;
        }
        .text-bottom {
            font-size: 62px; font-weight: 900; letter-spacing: -3px;
            margin-top: -26px; transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.18) 0%, rgba(245,247,250,1) 90%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            opacity: 0.3; user-select: none;
        }

        /* Горизонтальное меню ПОД названием */
        .sub-nav {
            display: flex; gap: 15px; padding: 12px 20px;
            background: var(--bg); border-radius: 18px;
            box-shadow: 6px 6px 15px var(--neu-shadow), -6px -6px 15px var(--neu-light);
            border: 1px solid rgba(255,255,255,0.8);
        }
        .nav-item {
            font-size: 11px; font-weight: 800; text-transform: uppercase;
            color: #8e8e93; cursor: pointer; transition: 0.2s; letter-spacing: 0.5px;
            padding: 5px 10px; border-radius: 8px;
        }
        .nav-item.active { color: var(--accent); background: rgba(0,122,255,0.05); }
        .nav-item:hover:not(.disabled) { color: var(--text); }
        .nav-item.disabled { opacity: 0.3; cursor: not-allowed; }

        /* Кнопки данных и +NEU */
        .action-bar { 
            display: flex; justify-content: flex-end; padding: 0 25px; 
            margin: 20px 0; gap: 12px; align-items: center;
        }
        .btn-neu {
            background: var(--bg); border: none; border-radius: 12px;
            padding: 10px 20px; font-size: 13px; font-weight: 800; color: var(--accent);
            box-shadow: 4px 4px 10px var(--neu-shadow), -4px -4px 10px var(--neu-light);
            cursor: pointer;
        }
        .btn-neu:active { box-shadow: inset 2px 2px 5px var(--neu-shadow), inset -2px -2px 5px var(--neu-light); }
        .btn-small { font-size: 10px; padding: 6px 12px; color: #99a1ad; }

        /* Карточки */
        .neu-card {
            background: var(--bg); border-radius: 30px; padding: 25px;
            margin: 15px 20px; box-shadow: 10px 10px 20px var(--neu-shadow), -10px -10px 20px var(--neu-light);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.6);
        }
        .p-label { font-size: 10px; font-weight: 800; color: #adb5bd; text-transform: uppercase; }
        .p-title { font-size: 28px; font-weight: 900; color: #000; margin: 2px 0; }
        .p-desc { font-size: 13px; font-weight: 700; color: #636366; }

        .btn-del { color: #ff3b30; font-size: 22px; font-weight: 900; padding: 10px; cursor: pointer; }

        #list-p { padding-bottom: 80px; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

function renderList() {
    const main = el('v-home'); if(!main) return;
    main.innerHTML = `
        <div class="header-container">
            <img src="${LOGO_URL}" class="main-logo" onerror="this.style.opacity='0'">
            <div class="mirror-title">
                <div class="text-top">CitiTool</div>
                <div class="text-bottom">CitiTool</div>
            </div>
            
            <div class="sub-nav">
                <div class="nav-item active" onclick="goHome()">Home</div>
                <div class="nav-item disabled">SyncOP</div>
                <div class="nav-item disabled">WKZListe</div>
                <div class="nav-item disabled">Stange</div>
                <div class="nav-item disabled">Coming Soon</div>
            </div>
        </div>

        <div class="action-bar">
            <button class="btn-neu btn-small" onclick="exportJSON()">Export</button>
            <button class="btn-neu btn-small" onclick="importJSON()">Import</button>
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
        </div>`).join('');
}

// Управление JSON
function exportJSON() {
    const data = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(db));
    const link = document.createElement('a');
    link.setAttribute("href", data);
    link.setAttribute("download", "cititool_data.json");
    link.click();
}

function importJSON() {
    const file = document.createElement('input');
    file.type = 'file';
    file.onchange = e => {
        const reader = new FileReader();
        reader.onload = r => {
            db = JSON.parse(r.target.result);
            save(); renderList();
        };
        reader.readAsText(e.target.files[0]);
    };
    file.click();
}

function goHome() { currentIdx = null; el('v-home').style.display='block'; el('v-det').style.display='none'; renderList(); }
function openProject(i) { /* Логика открытия проекта */ }
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); save(); renderList(); } }

window.onload = () => { injectStyles(); renderList(); };
