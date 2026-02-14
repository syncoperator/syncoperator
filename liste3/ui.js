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
            margin: 0; padding-top: 80px; padding-bottom: 120px;
            color: var(--text);
        }

        /* Фиксированная верхняя панель (как на скриншотах) */
        .header-fixed {
            position: fixed; top: 0; left: 0; right: 0; height: 70px;
            background: rgba(255,255,255,0.7); backdrop-filter: blur(20px);
            display: flex; align-items: center; justify-content: space-between;
            padding: 0 25px; z-index: 10000; border-bottom: 1px solid rgba(0,0,0,0.05);
        }
        .nav-title { font-size: 14px; font-weight: 800; color: var(--accent); }
        .nav-btns { display: flex; gap: 12px; }

        /* Кнопки в стиле Premium */
        .btn-premium {
            background: var(--bg); border: none; border-radius: 12px;
            padding: 10px 18px; font-size: 11px; font-weight: 800; color: #666;
            box-shadow: 4px 4px 8px var(--neu-shadow), -4px -4px 8px var(--neu-light);
            cursor: pointer; text-transform: uppercase; transition: 0.2s;
        }
        .btn-premium.blue { background: var(--accent); color: white; box-shadow: 0 4px 12px rgba(0,122,255,0.3); }
        .btn-premium:active { transform: scale(0.95); box-shadow: inset 2px 2px 5px var(--neu-shadow); }

        /* Центр: Лого и Зеркало */
        .hero-block {
            display: flex; flex-direction: column; align-items: center;
            padding: 40px 0; margin-bottom: 20px;
        }
        .logo-img {
            width: 200px; height: 200px; object-fit: contain;
            margin-bottom: 25px; filter: drop-shadow(0 20px 40px rgba(0,0,0,0.1));
        }
        .mirror-wrap { text-align: center; margin-bottom: 40px; }
        .t-main { font-size: 72px; font-weight: 900; letter-spacing: -4px; line-height: 1; }
        .t-reflect {
            font-size: 72px; font-weight: 900; letter-spacing: -4px;
            margin-top: -30px; transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.2), transparent 80%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            opacity: 0.2; user-select: none;
        }

        /* Горизонтальное меню */
        .nav-tabs {
            display: flex; background: rgba(0,0,0,0.05); 
            padding: 5px; border-radius: 16px;
        }
        .tab-item {
            font-size: 11px; font-weight: 800; text-transform: uppercase;
            padding: 12px 22px; border-radius: 13px; color: #888; cursor: pointer;
        }
        .tab-item.active { background: white; color: var(--accent); box-shadow: 0 4px 12px rgba(0,0,0,0.06); }
        .tab-item.disabled { opacity: 0.3; pointer-events: none; }

        /* Список проектов */
        .container { padding: 0 20px; max-width: 800px; margin: 0 auto; }
        .card-p {
            background: var(--bg); border-radius: 35px; padding: 30px;
            margin-bottom: 25px; box-shadow: 12px 12px 25px var(--neu-shadow), -12px -12px 25px var(--neu-light);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.7); cursor: pointer;
        }
        .p-num { font-size: 36px; font-weight: 900; color: #000; letter-spacing: -1.5px; }
        .p-name { font-size: 16px; font-weight: 700; color: #666; margin-top: 4px; }
        .p-label { font-size: 10px; font-weight: 800; color: #adb5bd; text-transform: uppercase; letter-spacing: 2px; }

        .btn-del { color: #ff3b30; font-size: 26px; font-weight: 900; padding: 10px; opacity: 0.5; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// Глобальные методы
window.exportJSON = () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(db));
    const dlAnchorElem = document.createElement('a');
    dlAnchorElem.setAttribute("href", dataStr);
    dlAnchorElem.setAttribute("download", "cititool_v8.json");
    dlAnchorElem.click();
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
        <div class="header-fixed">
            <div class="nav-title" onclick="window.renderList()">CitiTool Home</div>
            <div class="nav-btns">
                <button class="btn-premium" onclick="window.exportJSON()">Export</button>
                <button class="btn-premium" onclick="window.importJSON()">Import</button>
                <button class="btn-premium blue" onclick="window.modalP()">+ NEU</button>
            </div>
        </div>

        <div class="hero-block">
            <img src="${LOGO_URL}" class="logo-img" alt="Logo">
            <div class="mirror-wrap">
                <div class="t-main">CitiTool</div>
                <div class="t-reflect">CitiTool</div>
            </div>
            
            <div class="nav-tabs">
                <div class="tab-item active">Home</div>
                <div class="tab-item disabled">SyncOP</div>
                <div class="tab-item disabled">WKZListe</div>
                <div class="tab-item disabled">Stange</div>
            </div>
        </div>
        
        <div class="container" id="list-target"></div>
    `;
    
    const list = el('list-target');
    list.innerHTML = db.map((p, i) => `
        <div class="card-p" onclick="openProject(${i})">
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

// Старт
window.onload = () => {
    injectStyles();
    window.renderList();
};
