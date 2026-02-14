const DB_KEY = 'QS_DATA_V8';
// Ссылка на твой логотип с GitHub
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png'; 

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f8f9fc; 
            --neu-light: #ffffff;
            --neu-shadow: #d1d9e6;
            --accent: #007aff;
            --text: #1d1d1f;
            --gray-text: #86868b;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", sans-serif !important; 
            margin: 0; padding-bottom: 120px; color: var(--text);
            -webkit-font-smoothing: antialiased;
        }

        /* Единый блок шапки */
        .premium-header {
            display: flex; flex-direction: column; align-items: center;
            padding: 40px 0 25px 0;
            background: linear-gradient(180deg, #ffffff 0%, var(--bg) 100%);
            border-bottom: 1px solid rgba(0,0,0,0.03);
        }
        .main-logo {
            width: 100px; height: 100px; object-fit: contain;
            margin-bottom: 15px; filter: drop-shadow(0 8px 15px rgba(0,0,0,0.08));
        }
        .brand-name {
            position: relative; text-align: center; margin-bottom: 25px;
        }
        .text-top {
            font-size: 54px; font-weight: 800; letter-spacing: -2.5px;
            color: var(--text); line-height: 1;
        }
        .text-bottom {
            font-size: 54px; font-weight: 800; letter-spacing: -2.5px;
            margin-top: -24px; transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.15) 0%, transparent 80%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            opacity: 0.2; user-select: none;
        }

        /* Навигация как в iOS/Mac */
        .nav-wrapper {
            display: flex; background: rgba(0,0,0,0.03); 
            padding: 4px; border-radius: 14px; margin-bottom: 30px;
            border: 1px solid rgba(0,0,0,0.02);
        }
        .nav-item {
            font-size: 11px; font-weight: 700; text-transform: uppercase;
            padding: 10px 18px; border-radius: 11px; color: var(--gray-text);
            cursor: pointer; transition: 0.2s; letter-spacing: 0.5px;
        }
        .nav-item.active { 
            background: white; color: var(--accent);
            box-shadow: 0 4px 10px rgba(0,0,0,0.05);
        }
        .nav-item.disabled { opacity: 0.4; font-size: 9px; }

        /* Панель управления (Action Bar) */
        .controls-row {
            display: flex; justify-content: space-between; align-items: center;
            padding: 0 20px; max-width: 600px; margin: 0 auto 20px auto;
            width: calc(100% - 40px);
        }
        .btn-group { display: flex; gap: 10px; }
        
        .btn-neu {
            background: var(--bg); border: none; border-radius: 14px;
            padding: 12px 24px; font-size: 14px; font-weight: 700; color: var(--accent);
            box-shadow: 6px 6px 14px var(--neu-shadow), -6px -6px 14px var(--neu-light);
            cursor: pointer; display: flex; align-items: center; gap: 8px;
        }
        .btn-neu:active { box-shadow: inset 3px 3px 6px var(--neu-shadow), inset -3px -3px 6px var(--neu-light); }
        .btn-sec { font-size: 11px; padding: 10px 15px; color: var(--gray-text); }

        /* Карточки проектов */
        .project-card {
            background: var(--bg); border-radius: 32px; padding: 25px;
            margin: 15px 20px; box-shadow: 10px 10px 20px var(--neu-shadow), -10px -10px 20px var(--neu-light);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.8);
            cursor: pointer;
        }
        .p-info .label { font-size: 10px; font-weight: 800; color: #adb5bd; letter-spacing: 1px; }
        .p-info .title { font-size: 28px; font-weight: 800; color: #000; margin-top: 2px; }
        .p-info .sub { font-size: 14px; font-weight: 600; color: var(--gray-text); }

        .btn-del { color: #ff3b30; font-size: 22px; font-weight: 900; padding: 10px; }

        #list-p { padding-bottom: 50px; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

function renderList() {
    const main = el('v-home'); if(!main) return;
    main.innerHTML = `
        <div class="premium-header">
            <img src="${LOGO_URL}" class="main-logo" onerror="this.style.opacity='0'">
            <div class="brand-name">
                <div class="text-top">CitiTool</div>
                <div class="text-bottom">CitiTool</div>
            </div>
            
            <div class="nav-wrapper">
                <div class="nav-item active">Home</div>
                <div class="nav-item disabled">SyncOP</div>
                <div class="nav-item disabled">WKZListe</div>
                <div class="nav-item disabled">Stange</div>
                <div class="nav-item disabled">Coming Soon</div>
            </div>
        </div>

        <div class="controls-row">
            <div class="btn-group">
                <button class="btn-neu btn-sec" onclick="exportJSON()">Export</button>
                <button class="btn-neu btn-sec" onclick="importJSON()">Import</button>
            </div>
            <button class="btn-neu" onclick="modalP()">
                <span>+</span> NEU
            </button>
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
        </div>`).join('');
}

// Функции данных (JSON)
function exportJSON() {
    const blob = new Blob([JSON.stringify(db, null, 2)], {type: "application/json"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `cititool_v8_${new Date().toLocaleDateString()}.json`;
    a.click();
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

function goHome() { currentIdx = null; el('v-home').style.display='block'; el('v-det').style.display='none'; renderList(); }
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); save(); renderList(); } }

window.onload = () => { injectStyles(); renderList(); };
