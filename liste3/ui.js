const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png'; 

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f4f7f9; 
            --neu-light: #ffffff;
            --neu-shadow: #d1d9e6;
            --accent: #007aff;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, system-ui, sans-serif !important; 
            margin: 0; padding-bottom: 150px; 
        }

        /* Верхняя панель управления */
        .top-nav-bar {
            display: flex; justify-content: space-between; align-items: center;
            padding: 15px 25px; background: rgba(255,255,255,0.4);
            backdrop-filter: blur(10px); position: sticky; top: 0; z-index: 10000;
            border-bottom: 1px solid rgba(0,0,0,0.05);
        }
        .nav-left { font-size: 14px; font-weight: 800; color: var(--accent); cursor: pointer; }
        .nav-right { display: flex; gap: 10px; }

        .btn-ui {
            background: var(--bg); border: none; border-radius: 10px;
            padding: 8px 15px; font-size: 11px; font-weight: 800; color: #666;
            box-shadow: 4px 4px 8px var(--neu-shadow), -4px -4px 8px var(--neu-light);
            cursor: pointer; text-transform: uppercase;
        }
        .btn-ui.blue { background: var(--accent); color: white; box-shadow: 0 4px 10px rgba(0,122,255,0.3); }
        .btn-ui:active { transform: scale(0.95); box-shadow: inset 2px 2px 5px var(--neu-shadow); }

        /* Центральный брендинг */
        .brand-section {
            display: flex; flex-direction: column; align-items: center;
            padding: 40px 0;
        }
        .logo-main {
            width: 200px; height: 200px; object-fit: contain; /* ЛОГО ЕЩЕ БОЛЬШЕ */
            margin-bottom: 20px; filter: drop-shadow(0 15px 35px rgba(0,0,0,0.1));
        }
        .title-box { text-align: center; margin-bottom: 30px; }
        .t-high { font-size: 72px; font-weight: 900; letter-spacing: -4px; color: #1d1d1f; line-height: 0.8; }
        .t-refl {
            font-size: 72px; font-weight: 900; letter-spacing: -4px;
            margin-top: -32px; transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.2), transparent 80%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            opacity: 0.2; user-select: none;
        }

        /* Горизонтальное меню */
        .menu-bar {
            display: flex; gap: 8px; background: rgba(0,0,0,0.05); 
            padding: 5px; border-radius: 16px; margin-bottom: 40px;
        }
        .menu-item {
            font-size: 11px; font-weight: 800; text-transform: uppercase;
            padding: 10px 20px; border-radius: 12px; color: #888; cursor: pointer;
        }
        .menu-item.active { background: white; color: var(--accent); box-shadow: 0 4px 12px rgba(0,0,0,0.06); }
        .menu-item.disabled { opacity: 0.3; pointer-events: none; }

        /* Карточки */
        .list-p { padding: 0 20px; max-width: 900px; margin: 0 auto; }
        .p-card {
            background: var(--bg); border-radius: 35px; padding: 30px;
            margin-bottom: 25px; box-shadow: 12px 12px 25px var(--neu-shadow), -12px -12px 25px var(--neu-light);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.7); cursor: pointer;
        }
        .p-card:active { transform: scale(0.98); }
        .p-num { font-size: 36px; font-weight: 900; color: #000; letter-spacing: -1.5px; }
        .p-name { font-size: 16px; font-weight: 700; color: #666; }
        .p-label { font-size: 10px; font-weight: 800; color: #adb5bd; text-transform: uppercase; letter-spacing: 2px; }

        .btn-del { color: #ff3b30; font-size: 26px; font-weight: 900; padding: 10px; opacity: 0.5; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// Глобальные функции
window.exportJSON = () => {
    const blob = new Blob([JSON.stringify(db, null, 2)], {type: "application/json"});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = "cititool_v8.json";
    a.click();
};

window.importJSON = () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.onchange = e => {
        const reader = new FileReader();
        reader.onload = r => {
            db = JSON.parse(r.target.result);
            save(); window.renderList();
        };
        reader.readAsText(e.target.files[0]);
    };
    input.click();
};

window.renderList = () => {
    const main = el('v-home'); if(!main) return;
    
    main.innerHTML = `
        <div class="top-nav-bar">
            <div class="nav-left" onclick="window.renderList()">← v-home</div>
            <div class="nav-right">
                <button class="btn-ui" onclick="window.exportJSON()">Export</button>
                <button class="btn-ui" onclick="window.importJSON()">Import</button>
                <button class="btn-ui blue" onclick="window.modalP()">+ NEU</button>
            </div>
        </div>

        <div class="brand-section">
            <img src="${LOGO_URL}" class="logo-main" alt="Logo">
            <div class="title-box">
                <div class="t-high">CitiTool</div>
                <div class="t-refl">CitiTool</div>
            </div>
            
            <div class="menu-bar">
                <div class="menu-item active">Home</div>
                <div class="menu-item disabled">SyncOP</div>
                <div class="menu-item disabled">WKZListe</div>
                <div class="menu-item disabled">Stange</div>
            </div>
        </div>
        
        <div class="list-p" id="list-items"></div>
    `;
    
    const list = el('list-items');
    list.innerHTML = db.map((p, i) => `
        <div class="p-card" onclick="openProject(${i})">
            <div>
                <div class="p-label">PROJEKT</div>
                <div class="p-num">${p.num || '---'}</div>
                <div class="p-name">${p.name || ''}</div>
            </div>
            <div class="btn-del" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:100px;"></div>';
};

function deleteProject(i) {
    if(confirm('Löschen?')) { db.splice(i, 1); save(); window.renderList(); }
}

window.onload = () => {
    injectStyles();
    window.renderList();
};
