const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png'; 

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f2f5f8; 
            --neu-light: #ffffff;
            --neu-shadow: #cfd8e3;
            --accent: #007aff;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, system-ui, sans-serif !important; 
            margin: 0; padding-top: 100px; padding-bottom: 150px;
        }

        /* Кнопки в углу - Слой поверх всего */
        .top-nav {
            position: fixed; top: 0; left: 0; right: 0; height: 90px;
            display: flex; justify-content: flex-end; align-items: center;
            padding: 0 25px; z-index: 10000; pointer-events: none;
        }
        .btn-box { display: flex; gap: 12px; pointer-events: auto; }

        .btn-premium {
            background: var(--bg); border: none; border-radius: 14px;
            padding: 12px 20px; font-size: 11px; font-weight: 800; color: #666;
            box-shadow: 5px 5px 10px var(--neu-shadow), -5px -5px 10px var(--neu-light);
            cursor: pointer; text-transform: uppercase; transition: 0.2s;
        }
        .btn-premium.blue { background: var(--accent); color: white; box-shadow: 0 5px 15px rgba(0,122,255,0.3); }
        .btn-premium:active { transform: scale(0.94); box-shadow: inset 2px 2px 5px var(--neu-shadow); }

        /* Центр: Лого и Название */
        .hero { display: flex; flex-direction: column; align-items: center; padding-bottom: 50px; }
        .logo-huge {
            width: 200px; height: 200px; object-fit: contain;
            margin-bottom: 25px; filter: drop-shadow(0 20px 40px rgba(0,0,0,0.1));
        }
        .title-block { text-align: center; margin-bottom: 20px; }
        .t-top { font-size: 72px; font-weight: 900; letter-spacing: -4px; color: #1d1d1f; line-height: 0.8; }
        .t-refl {
            font-size: 72px; font-weight: 900; letter-spacing: -4px;
            margin-top: -32px; transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.15), transparent 85%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            opacity: 0.2; user-select: none;
        }

        /* Список карточек */
        .container { padding: 0 25px; max-width: 850px; margin: 0 auto; }
        .p-card {
            background: var(--bg); border-radius: 35px; padding: 30px;
            margin-bottom: 25px; box-shadow: 12px 12px 25px var(--neu-shadow), -12px -12px 25px var(--neu-light);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.7); cursor: pointer;
        }
        .p-num { font-size: 34px; font-weight: 900; color: #000; }
        .p-name { font-size: 16px; font-weight: 700; color: #666; margin-top: 4px; }
        .btn-del { color: #ff3b30; font-size: 26px; font-weight: 900; padding: 10px; opacity: 0.6; }

        .hidden { display: none !important; }
    `;
    document.head.appendChild(style);
};

// --- Глобальные функции ---
window.save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

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
            window.save(); window.renderList();
        };
        reader.readAsText(e.target.files[0]);
    };
    input.click();
};

window.modalP = () => {
    const num = prompt("Projekt Nummer:", "");
    const name = prompt("Projekt Name:", "");
    if(num) {
        db.push({ num, name, tools: [] });
        window.save();
        window.renderList();
    }
};

window.openProject = (i) => {
    currentIdx = i;
    document.getElementById('v-home').classList.add('hidden');
    document.getElementById('v-det').classList.remove('hidden');
    window.renderDetails();
};

window.goHome = () => {
    currentIdx = null;
    document.getElementById('v-home').classList.remove('hidden');
    document.getElementById('v-det').classList.add('hidden');
    window.renderList();
};

window.renderList = () => {
    const home = document.getElementById('v-home');
    if(!home) return;
    home.innerHTML = `
        <div class="top-nav">
            <div class="btn-box">
                <button class="btn-premium" onclick="window.exportJSON()">Export</button>
                <button class="btn-premium" onclick="window.importJSON()">Import</button>
                <button class="btn-premium blue" onclick="window.modalP()">+ NEU</button>
            </div>
        </div>

        <div class="hero">
            <img src="${LOGO_URL}" class="logo-huge">
            <div class="title-block">
                <div class="t-top">CitiTool</div>
                <div class="t-refl">CitiTool</div>
            </div>
        </div>
        
        <div class="container" id="p-items"></div>
    `;
    
    document.getElementById('p-items').innerHTML = db.map((p, i) => `
        <div class="p-card" onclick="window.openProject(${i})">
            <div>
                <div style="font-size:10px; font-weight:800; color:#adb5bd; letter-spacing:2px;">PROJEKT</div>
                <div class="p-num">${p.num || '---'}</div>
                <div class="p-name">${p.name || ''}</div>
            </div>
            <div class="btn-del" onclick="event.stopPropagation(); window.deleteP(${i})">✕</div>
        </div>`).join('') + '<div style="height:100px;"></div>';
};

window.renderDetails = () => {
    const det = document.getElementById('v-det');
    const p = db[currentIdx];
    det.innerHTML = `
        <div class="top-nav">
            <div class="btn-box">
                <button class="btn-premium" onclick="window.goHome()">← Home</button>
                <button class="btn-premium blue">+ TOOL</button>
            </div>
        </div>
        <div class="hero">
            <div class="title-block">
                <div class="t-top">${p.num}</div>
                <div style="font-weight:700; color:#666;">${p.name}</div>
            </div>
        </div>
        <div class="container">
            <p style="text-align:center; color:#adb5bd; font-weight:800;">REVOLVER TOOLS COMING SOON</p>
        </div>
    `;
};

window.deleteP = (i) => {
    if(confirm('Löschen?')) { db.splice(i, 1); window.save(); window.renderList(); }
};

window.onload = () => {
    injectStyles();
    window.renderList();
};
