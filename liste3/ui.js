const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png'; 

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f5f7fa; 
            --neu-light: #ffffff;
            --neu-shadow: #d1d9e6;
            --accent: #007aff;
            --text: #1d1d1f;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, system-ui, sans-serif !important; 
            margin: 0; padding-top: 90px; padding-bottom: 120px;
        }

        /* Кнопки вынесены в отдельный слой с высшим приоритетом */
        .top-fixed-nav {
            position: fixed; top: 0; left: 0; right: 0; height: 80px;
            display: flex; justify-content: flex-end; align-items: center;
            padding: 0 25px; z-index: 10000; pointer-events: none;
        }
        .btn-wrap { display: flex; gap: 12px; pointer-events: auto; }

        .btn-ui {
            background: var(--bg); border: none; border-radius: 12px;
            padding: 10px 20px; font-size: 11px; font-weight: 800; color: #666;
            box-shadow: 4px 4px 8px var(--neu-shadow), -4px -4px 8px var(--neu-light);
            cursor: pointer; text-transform: uppercase; transition: 0.2s;
        }
        .btn-ui.blue { background: var(--accent); color: white; box-shadow: 0 4px 12px rgba(0,122,255,0.3); }
        .btn-ui:active { transform: scale(0.95); box-shadow: inset 2px 2px 5px var(--neu-shadow); }

        /* Главный блок: Логотип и Зеркало */
        .brand-hero {
            display: flex; flex-direction: column; align-items: center;
            padding: 40px 0; pointer-events: none;
        }
        .logo-img {
            width: 200px; height: 200px; object-fit: contain;
            margin-bottom: 25px; filter: drop-shadow(0 20px 40px rgba(0,0,0,0.1));
            pointer-events: auto;
        }
        .title-wrap { text-align: center; margin-bottom: 35px; pointer-events: auto; }
        .t-top { font-size: 72px; font-weight: 900; letter-spacing: -4px; color: #1d1d1f; line-height: 1; }
        .t-refl {
            font-size: 72px; font-weight: 900; letter-spacing: -4px;
            margin-top: -30px; transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.15), transparent 85%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            opacity: 0.2; user-select: none;
        }

        /* Меню */
        .nav-tabs {
            display: flex; background: rgba(0,0,0,0.05); 
            padding: 5px; border-radius: 18px; pointer-events: auto;
        }
        .tab {
            font-size: 11px; font-weight: 800; text-transform: uppercase;
            padding: 11px 22px; border-radius: 14px; color: #888; cursor: pointer;
        }
        .tab.active { background: white; color: var(--accent); box-shadow: 0 4px 12px rgba(0,0,0,0.06); }
        .tab.disabled { opacity: 0.3; pointer-events: none; }

        /* Карточки инструментов/проектов */
        .list-container { padding: 0 20px; max-width: 800px; margin: 0 auto; pointer-events: auto; }
        .card {
            background: var(--bg); border-radius: 30px; padding: 25px;
            margin-bottom: 20px; box-shadow: 10px 10px 20px var(--neu-shadow), -10px -10px 20px var(--neu-light);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.7); cursor: pointer;
        }
        .card-title { font-size: 28px; font-weight: 900; color: #000; }
        .card-label { font-size: 10px; font-weight: 800; color: #adb5bd; text-transform: uppercase; letter-spacing: 2px; }
        
        /* Кнопки управления данными внизу списка для удобства */
        .data-management {
            display: flex; justify-content: center; gap: 15px; margin-top: 40px; padding-bottom: 100px;
        }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// Глобальные обработчики
window.exportJSON = () => {
    const data = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(db));
    const a = document.createElement('a');
    a.href = data; a.download = "cititool_v8.json";
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
        <div class="top-fixed-nav">
            <div class="btn-wrap">
                <button class="btn-ui blue" onclick="window.modalP()">+ NEU</button>
            </div>
        </div>

        <div class="brand-hero">
            <img src="${LOGO_URL}" class="logo-img" alt="Logo">
            <div class="title-wrap">
                <div class="t-top">CitiTool</div>
                <div class="t-refl">CitiTool</div>
            </div>
            
            <div class="nav-tabs">
                <div class="tab active" onclick="window.renderList()">Home</div>
                <div class="tab disabled">SyncOP</div>
                <div class="tab disabled">WKZListe</div>
                <div class="tab disabled">Stange</div>
            </div>
        </div>
        
        <div class="list-container" id="project-items"></div>

        <div class="data-management">
            <button class="btn-ui" onclick="window.exportJSON()">Export JSON</button>
            <button class="btn-ui" onclick="window.importJSON()">Import JSON</button>
        </div>
    `;
    
    const list = el('project-items');
    list.innerHTML = db.map((p, i) => `
        <div class="card" onclick="openProject(${i})">
            <div>
                <div class="card-label">PROJEKT</div>
                <div class="card-title">${p.num || '---'}</div>
                <div style="font-size:14px; font-weight:700; color:#666;">${p.name || ''}</div>
            </div>
            <div style="color:#ff3b30; font-size:24px; font-weight:900;" onclick="event.stopPropagation(); deleteP(${i})">✕</div>
        </div>`).join('');
};

window.deleteP = (i) => {
    if(confirm('Löschen?')) { db.splice(i, 1); save(); window.renderList(); }
};

window.onload = () => {
    injectStyles();
    window.renderList();
};
