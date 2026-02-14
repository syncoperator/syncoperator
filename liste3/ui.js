const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png'; 

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f8f9fb; 
            --neu-light: #ffffff;
            --neu-shadow: #d1d9e6;
            --accent: #007aff;
            --text: #1d1d1f;
            --gray: #86868b;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, system-ui, sans-serif !important; 
            margin: 0; padding-bottom: 150px; color: var(--text);
            -webkit-font-smoothing: antialiased;
        }

        /* Верхний правый угол - компактные кнопки */
        .top-actions {
            position: absolute; top: 15px; right: 15px;
            display: flex; gap: 8px; z-index: 2000;
        }
        .btn-mini {
            background: var(--bg); border: none; border-radius: 10px;
            padding: 6px 12px; font-size: 10px; font-weight: 800; color: var(--gray);
            box-shadow: 3px 3px 6px var(--neu-shadow), -3px -3px 6px var(--neu-light);
            cursor: pointer; text-transform: uppercase; transition: 0.2s;
        }
        .btn-mini.blue { color: var(--accent); font-weight: 900; }
        .btn-mini:active { box-shadow: inset 2px 2px 4px var(--neu-shadow), inset -2px -2px 4px var(--neu-light); }

        /* Центр: Лого и Название */
        .brand-block {
            display: flex; flex-direction: column; align-items: center;
            padding: 50px 0 20px 0;
        }
        .logo-img {
            width: 90px; height: 90px; object-fit: contain;
            margin-bottom: 10px; filter: drop-shadow(0 10px 15px rgba(0,0,0,0.05));
        }
        .title-mirror { position: relative; text-align: center; margin-bottom: 25px; }
        .t-main { font-size: 52px; font-weight: 900; letter-spacing: -2.5px; line-height: 1; }
        .t-reflect {
            font-size: 52px; font-weight: 900; letter-spacing: -2.5px;
            margin-top: -24px; transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.15), transparent 85%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            opacity: 0.2; user-select: none;
        }

        /* Меню - Премиальный вид */
        .nav-container {
            display: flex; background: rgba(0,0,0,0.03); 
            padding: 4px; border-radius: 14px; margin-top: 5px;
            border: 1px solid rgba(255,255,255,0.5);
        }
        .nav-link {
            font-size: 10px; font-weight: 800; text-transform: uppercase;
            padding: 9px 15px; border-radius: 11px; color: var(--gray);
            cursor: pointer; transition: 0.2s; letter-spacing: 0.3px;
        }
        .nav-link.active { background: white; color: var(--accent); box-shadow: 0 4px 10px rgba(0,0,0,0.04); }
        .nav-link.disabled { opacity: 0.3; cursor: default; }

        /* Карточки проектов */
        .p-card {
            background: var(--bg); border-radius: 30px; padding: 25px;
            margin: 15px 20px; box-shadow: 10px 10px 20px var(--neu-shadow), -10px -10px 20px var(--neu-light);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.8);
            cursor: pointer; transition: transform 0.2s;
        }
        .p-card:active { transform: scale(0.98); }
        .p-meta .p-tag { font-size: 9px; font-weight: 900; color: #adb5bd; letter-spacing: 1px; text-transform: uppercase; }
        .p-meta .p-num { font-size: 28px; font-weight: 900; color: #000; margin: 2px 0; }
        .p-meta .p-name { font-size: 14px; font-weight: 700; color: var(--gray); }

        .btn-close { color: #ff3b30; font-size: 20px; font-weight: 900; padding: 10px; opacity: 0.7; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// Рендер главного экрана
function renderList() {
    const main = el('v-home'); if(!main) return;
    
    // Кнопки отдельно, чтобы не перекрывались заголовком
    main.innerHTML = `
        <div class="top-actions">
            <button class="btn-mini" onclick="exportJSON()">Export</button>
            <button class="btn-mini" onclick="importJSON()">Import</button>
            <button class="btn-mini blue" onclick="modalP()">+ NEU</button>
        </div>

        <div class="brand-block">
            <img src="${LOGO_URL}" class="logo-img" alt="CitiTool Logo">
            <div class="title-mirror">
                <div class="t-main">CitiTool</div>
                <div class="t-reflect">CitiTool</div>
            </div>
            
            <div class="nav-container">
                <div class="nav-link active" onclick="goHome()">Home</div>
                <div class="nav-link disabled">SyncOP</div>
                <div class="nav-link disabled">WKZListe</div>
                <div class="nav-item disabled" style="font-size:10px; color:#ccc; padding: 9px 10px;">...</div>
            </div>
        </div>
        
        <div id="list-p"></div>
    `;
    
    const list = el('list-p');
    list.innerHTML = db.map((p, i) => `
        <div class="p-card" onclick="openProject(${i})">
            <div class="p-meta">
                <div class="p-tag">Projekt</div>
                <div class="p-num">${p.num || '---'}</div>
                <div class="p-name">${p.name || ''}</div>
            </div>
            <div class="btn-close" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:80px"></div>';
}

// Работа с данными
function exportJSON() {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(db));
    const a = document.createElement('a');
    a.href = dataStr; a.download = "cititool_v8.json";
    a.click();
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

function goHome() { 
    currentIdx = null; 
    const vHome = el('v-home');
    const vDet = el('v-det');
    if(vHome) vHome.style.display='block'; 
    if(vDet) vDet.style.display='none'; 
    renderList(); 
}

function deleteProject(i) {
    if(confirm('Löschen?')) { db.splice(i, 1); save(); renderList(); }
}

// Инициализация при загрузке
window.onload = () => {
    injectStyles();
    renderList();
};
