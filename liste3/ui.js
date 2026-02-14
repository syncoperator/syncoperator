const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png'; 

// Загрузка данных
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];

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
            margin: 0; padding-top: 100px; padding-bottom: 150px;
            color: var(--text); overflow-x: hidden;
        }

        /* Кнопки в углу - Сверх-приоритет */
        .top-layer {
            position: fixed; top: 0; left: 0; right: 0; height: 90px;
            display: flex; justify-content: flex-end; align-items: center;
            padding: 0 25px; z-index: 10000; pointer-events: none;
        }
        .btn-group { display: flex; gap: 12px; pointer-events: auto; }

        .btn-premium {
            background: var(--bg); border: none; border-radius: 14px;
            padding: 10px 20px; font-size: 11px; font-weight: 800; color: #666;
            box-shadow: 5px 5px 10px var(--neu-shadow), -5px -5px 10px var(--neu-light);
            cursor: pointer; text-transform: uppercase; transition: 0.2s ease;
        }
        .btn-premium.blue { 
            background: var(--accent); color: white; 
            box-shadow: 0 5px 15px rgba(0,122,255,0.3); 
        }
        .btn-premium:active { transform: scale(0.94); box-shadow: inset 2px 2px 5px var(--neu-shadow); }

        /* Центральный брендинг */
        .hero {
            display: flex; flex-direction: column; align-items: center;
            padding-bottom: 40px; pointer-events: none;
        }
        .logo-huge {
            width: 200px; height: 200px; object-fit: contain;
            margin-bottom: 25px; filter: drop-shadow(0 20px 40px rgba(0,0,0,0.1));
            pointer-events: auto;
        }
        .title-mirror { text-align: center; margin-bottom: 40px; pointer-events: auto; }
        .t-top { font-size: 72px; font-weight: 900; letter-spacing: -4px; line-height: 0.8; }
        .t-refl {
            font-size: 72px; font-weight: 900; letter-spacing: -4px;
            margin-top: -32px; transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.15), transparent 85%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            opacity: 0.2; user-select: none;
        }

        /* Навигация под заголовком */
        .nav-pills {
            display: flex; background: rgba(0,0,0,0.05); 
            padding: 6px; border-radius: 20px; pointer-events: auto;
        }
        .pill {
            font-size: 11px; font-weight: 800; text-transform: uppercase;
            padding: 12px 24px; border-radius: 15px; color: #888; cursor: pointer;
        }
        .pill.active { background: white; color: var(--accent); box-shadow: 0 4px 15px rgba(0,0,0,0.06); }
        .pill.disabled { opacity: 0.3; pointer-events: none; }

        /* Список карточек */
        .main-container { padding: 0 25px; max-width: 850px; margin: 0 auto; }
        .p-card {
            background: var(--bg); border-radius: 35px; padding: 30px;
            margin-bottom: 25px; box-shadow: 12px 12px 25px var(--neu-shadow), -12px -12px 25px var(--neu-light);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.7); cursor: pointer;
        }
        .p-card:active { transform: scale(0.98); }
        .p-num { font-size: 34px; font-weight: 900; color: #000; letter-spacing: -1.5px; }
        .p-name { font-size: 16px; font-weight: 700; color: #666; margin-top: 4px; }
        .p-tag { font-size: 10px; font-weight: 800; color: #adb5bd; text-transform: uppercase; letter-spacing: 2px; }

        .btn-del { color: #ff3b30; font-size: 26px; font-weight: 900; padding: 10px; opacity: 0.6; }

        /* Модальное окно (NEU) */
        #modal-p {
            position: fixed; inset: 0; background: rgba(0,0,0,0.4); 
            backdrop-filter: blur(10px); z-index: 20000;
            display: none; align-items: center; justify-content: center;
        }
        .modal-content {
            background: var(--bg); padding: 40px; border-radius: 40px;
            width: 90%; max-width: 400px; box-shadow: 20px 20px 60px rgba(0,0,0,0.2);
            text-align: center;
        }
        input {
            width: 100%; padding: 15px; margin: 10px 0; border-radius: 15px;
            border: none; background: var(--bg); box-shadow: inset 3px 3px 6px var(--neu-shadow), inset -3px -3px 6px var(--neu-light);
            font-size: 16px; font-weight: 700; outline: none; box-sizing: border-box;
        }
    `;
    document.head.appendChild(style);
};

// --- Глобальные функции (привязаны к window для гарантии клика) ---

window.saveData = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

window.exportJSON = () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(db));
    const a = document.createElement('a');
    a.href = dataStr; a.download = "cititool_backup.json";
    a.click();
};

window.importJSON = () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.onchange = e => {
        const reader = new FileReader();
        reader.onload = r => {
            db = JSON.parse(r.target.result);
            window.saveData(); window.renderList();
        };
        reader.readAsText(e.target.files[0]);
    };
    input.click();
};

// Функция создания проекта (NEU)
window.modalP = () => {
    const num = prompt("Projekt Nummer:", "");
    const name = prompt("Projekt Name:", "");
    if(num) {
        db.push({ num, name, tools: [] });
        window.saveData();
        window.renderList();
    }
};

window.deleteP = (i) => {
    if(confirm('Projekt löschen?')) {
        db.splice(i, 1);
        window.saveData();
        window.renderList();
    }
};

window.renderList = () => {
    const main = document.getElementById('v-home');
    if(!main) return;
    
    main.innerHTML = `
        <div class="top-layer">
            <div class="btn-group">
                <button class="btn-premium" onclick="window.exportJSON()">Export</button>
                <button class="btn-premium" onclick="window.importJSON()">Import</button>
                <button class="btn-premium blue" onclick="window.modalP()">+ NEU</button>
            </div>
        </div>

        <div class="hero">
            <img src="${LOGO_URL}" class="logo-huge" alt="Logo">
            <div class="title-mirror">
                <div class="t-top">CitiTool</div>
                <div class="t-refl">CitiTool</div>
            </div>
            
            <div class="nav-pills">
                <div class="pill active" onclick="window.renderList()">Home</div>
                <div class="pill disabled">SyncOP</div>
                <div class="pill disabled">WKZListe</div>
                <div class="pill disabled">Stange</div>
            </div>
        </div>
        
        <div class="main-container" id="p-list"></div>
    `;
    
    const list = document.getElementById('p-list');
    list.innerHTML = db.map((p, i) => `
        <div class="p-card" onclick="openProject(${i})">
            <div>
                <div class="p-tag">Projekt</div>
                <div class="p-num">${p.num || '---'}</div>
                <div class="p-name">${p.name || ''}</div>
            </div>
            <div class="btn-del" onclick="event.stopPropagation(); window.deleteP(${i})">✕</div>
        </div>`).join('') + '<div style="height:100px;"></div>';
};

// Запуск
window.onload = () => {
    injectStyles();
    window.renderList();
};
