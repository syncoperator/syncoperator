const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png'; 

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
            --text: #1d1d1f;
        }
        body { background: var(--bg); font-family: -apple-system, system-ui, sans-serif; margin: 0; padding-top: 80px; padding-bottom: 150px; }

        /* Фикс панель */
        .header-fixed {
            position: fixed; top: 0; left: 0; right: 0; height: 70px;
            background: rgba(255,255,255,0.75); backdrop-filter: blur(20px);
            display: flex; align-items: center; justify-content: space-between;
            padding: 0 25px; z-index: 10000; border-bottom: 1px solid rgba(0,0,0,0.05);
        }
        .btn-premium {
            background: var(--bg); border: none; border-radius: 12px;
            padding: 10px 18px; font-size: 11px; font-weight: 800; color: #666;
            box-shadow: 4px 4px 8px var(--neu-shadow), -4px -4px 8px var(--neu-light);
            cursor: pointer; text-transform: uppercase;
        }
        .btn-premium.blue { background: var(--accent); color: white; }

        /* Центр */
        .hero { display: flex; flex-direction: column; align-items: center; padding: 40px 0; }
        .logo-img { width: 200px; height: 200px; object-fit: contain; filter: drop-shadow(0 20px 40px rgba(0,0,0,0.1)); }
        .t-main { font-size: 72px; font-weight: 900; letter-spacing: -4px; line-height: 1; margin-top: 20px;}
        .t-refl {
            font-size: 72px; font-weight: 900; letter-spacing: -4px;
            margin-top: -32px; transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.2), transparent 80%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent; opacity: 0.2;
        }

        /* Карточки Инструментов (Revolver) */
        .tool-card {
            background: var(--bg); border-radius: 20px; padding: 20px;
            margin: 15px 0; box-shadow: 8px 8px 16px var(--neu-shadow), -8px -8px 16px var(--neu-light);
            display: flex; justify-content: space-between; align-items: center;
        }
        .t-nr { font-size: 12px; font-weight: 800; color: var(--accent); text-transform: uppercase; }
        .t-desc { font-size: 20px; font-weight: 900; color: #000; display: block; margin-top: 4px; }
        
        .rev-label { 
            font-size: 10px; font-weight: 900; color: #ff9500; 
            padding: 4px 8px; border-radius: 6px; background: rgba(255,149,0,0.1);
        }

        .container { padding: 0 20px; max-width: 800px; margin: 0 auto; }
        .hidden { display: none; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);

window.renderList = () => {
    el('v-home').classList.remove('hidden');
    el('v-det').classList.add('hidden');
    
    el('v-home').innerHTML = `
        <div class="header-fixed">
            <div style="font-weight:900; color:var(--accent)">CITITOOL</div>
            <div class="nav-btns">
                <button class="btn-premium" onclick="exportJSON()">Export</button>
                <button class="btn-premium" onclick="importJSON()">Import</button>
                <button class="btn-premium blue" onclick="modalP()">+ NEU</button>
            </div>
        </div>

        <div class="hero">
            <img src="${LOGO_URL}" class="logo-img">
            <div class="t-main">CitiTool</div>
            <div class="t-reflect">CitiTool</div>
        </div>
        
        <div class="container" id="project-list"></div>
    `;
    renderProjectItems();
};

const renderProjectItems = () => {
    el('project-list').innerHTML = db.map((p, i) => `
        <div class="tool-card" onclick="openProject(${i})">
            <div>
                <div class="t-nr">Projekt</div>
                <div class="t-desc">${p.num || '---'}</div>
            </div>
            <div class="btn-premium" style="color:#ff3b30" onclick="event.stopPropagation(); deleteP(${i})">✕</div>
        </div>
    `).join('') + '<div style="height:100px"></div>';
};

window.openProject = (i) => {
    currentIdx = i;
    el('v-home').classList.add('hidden');
    el('v-det').classList.remove('hidden');
    renderDetails();
};

window.renderDetails = () => {
    const p = db[currentIdx];
    el('v-det').innerHTML = `
        <div class="header-fixed">
            <div onclick="renderList()" style="cursor:pointer; font-weight:800;">← BACK</div>
            <div style="font-weight:900;">${p.num}</div>
            <button class="btn-premium blue" onclick="addTool()">+ TOOL</button>
        </div>
        <div class="container" style="padding-top:40px">
            <h2 style="letter-spacing:-1px">Revolver Tools</h2>
            <div id="tool-list"></div>
        </div>
    `;
    renderTools();
};

const renderTools = () => {
    const tools = db[currentIdx].tools || [];
    el('tool-list').innerHTML = tools.map((t, i) => `
        <div class="tool-card">
            <div>
                <div class="t-nr">T-${t.nr || '00'}</div>
                <div class="t-desc">${t.desc || 'Beschreibung...'}</div>
                <div class="rev-label">${t.rev === 'oben' ? 'REVOLVER OBEN' : 'REVOLVER UNTEN'}</div>
            </div>
            <div class="btn-premium" onclick="deleteTool(${i})">✕</div>
        </div>
    `).join('');
};

// Системные функции
window.exportJSON = () => { /* логика экспорта */ };
window.importJSON = () => { /* логика импорта */ };
window.deleteP = (i) => { if(confirm('Löschen?')) { db.splice(i,1); save(); renderProjectItems(); } };

window.onload = () => {
    injectStyles();
    window.renderList();
};
