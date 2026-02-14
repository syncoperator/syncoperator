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
            -webkit-tap-highlight-color: transparent;
        }

        /* Кнопки управления - ГАРАНТИРОВАННЫЙ КЛИК */
        .top-nav-actions {
            position: fixed; top: 20px; right: 20px;
            display: flex; gap: 12px; z-index: 10000; /* Самый высокий слой */
            pointer-events: auto !important;
        }
        .btn-premium {
            background: var(--bg); border: none; border-radius: 12px;
            padding: 10px 18px; font-size: 11px; font-weight: 800; color: #666;
            box-shadow: 5px 5px 10px var(--neu-shadow), -5px -5px 10px var(--neu-light);
            cursor: pointer; text-transform: uppercase; transition: 0.2s;
        }
        .btn-premium.main { color: var(--accent); font-weight: 900; }
        .btn-premium:active { transform: scale(0.92); box-shadow: inset 2px 2px 5px var(--neu-shadow); }

        /* Шапка */
        .hero {
            display: flex; flex-direction: column; align-items: center;
            padding: 90px 0 40px 0; pointer-events: none; /* Пропускаем клики сквозь фон */
        }
        .logo-big {
            width: 180px; height: 180px; object-fit: contain;
            margin-bottom: 25px; filter: drop-shadow(0 15px 30px rgba(0,0,0,0.08));
            pointer-events: auto;
        }
        .brand-mirror { text-align: center; margin-bottom: 35px; pointer-events: auto; }
        .t1 { font-size: 68px; font-weight: 900; letter-spacing: -3.5px; line-height: 0.85; color: #1d1d1f; }
        .t2 {
            font-size: 68px; font-weight: 900; letter-spacing: -3.5px;
            margin-top: -30px; transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.2), transparent 75%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            opacity: 0.25;
        }

        /* Меню навигации */
        .nav-pills {
            display: flex; background: rgba(0,0,0,0.05); 
            padding: 5px; border-radius: 18px; pointer-events: auto;
        }
        .pill {
            font-size: 11px; font-weight: 800; text-transform: uppercase;
            padding: 12px 22px; border-radius: 14px; color: #888;
            cursor: pointer; transition: 0.3s;
        }
        .pill.active { background: white; color: var(--accent); box-shadow: 0 4px 15px rgba(0,0,0,0.06); }
        .pill.disabled { opacity: 0.3; cursor: default; }

        /* Список карточек */
        .list-container { padding: 0 20px; max-width: 800px; margin: 0 auto; pointer-events: auto; }
        .card {
            background: var(--bg); border-radius: 35px; padding: 30px;
            margin-bottom: 25px; box-shadow: 12px 12px 25px var(--neu-shadow), -12px -12px 25px var(--neu-light);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.7);
            cursor: pointer;
        }
        .card-num { font-size: 34px; font-weight: 900; color: #000; letter-spacing: -1.5px; }
        .card-name { font-size: 16px; font-weight: 700; color: #666; margin-top: 5px; }
        .card-label { font-size: 10px; font-weight: 800; color: #adb5bd; text-transform: uppercase; letter-spacing: 1.5px; }

        .btn-del { color: #ff3b30; font-size: 26px; font-weight: 900; padding: 10px; opacity: 0.5; }
        .btn-del:hover { opacity: 1; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// Глобальные обработчики для гарантии работы
window.exportJSON = function() {
    console.log("Export triggered");
    const blob = new Blob([JSON.stringify(db, null, 2)], {type: "application/json"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `cititool_v8_backup.json`;
    a.click();
};

window.importJSON = function() {
    console.log("Import triggered");
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

window.goHome = function() { 
    if(el('v-home')) el('v-home').style.display='block'; 
    if(el('v-det')) el('v-det').style.display='none'; 
    window.renderList(); 
};

window.renderList = function() {
    const main = el('v-home'); if(!main) return;
    
    main.innerHTML = `
        <div class="top-nav-actions">
            <button class="btn-premium" onclick="window.exportJSON()">Export</button>
            <button class="btn-premium" onclick="window.importJSON()">Import</button>
            <button class="btn-premium main" onclick="window.modalP()">+ NEU</button>
        </div>

        <div class="hero">
            <img src="${LOGO_URL}" class="logo-big" alt="Logo">
            <div class="brand-mirror">
                <div class="t1">CitiTool</div>
                <div class="t2">CitiTool</div>
            </div>
            
            <div class="nav-pills">
                <div class="pill active" onclick="window.goHome()">Home</div>
                <div class="pill disabled">SyncOP</div>
                <div class="pill disabled">WKZListe</div>
                <div class="pill disabled">Stange</div>
            </div>
        </div>
        
        <div class="list-container" id="list-p"></div>
    `;
    
    const list = el('list-p');
    list.innerHTML = db.map((p, i) => `
        <div class="card" onclick="openProject(${i})">
            <div>
                <div class="card-label">PROJEKT</div>
                <div class="card-num">${p.num || '---'}</div>
                <div class="card-name">${p.name || ''}</div>
            </div>
            <div class="btn-del" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:80px"></div>';
};

function deleteProject(i) {
    if(confirm('Löschen?')) { db.splice(i, 1); save(); window.renderList(); }
}

// Запуск приложения
window.onload = () => {
    injectStyles();
    window.renderList();
};
