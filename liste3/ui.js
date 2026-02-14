const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png'; 

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

// Стили - без изменений, только база
const injectStyles = () => {
    if (document.getElementById('main-style')) return;
    const style = document.createElement('style');
    style.id = 'main-style';
    style.innerHTML = `
        :root { --bg: #f2f5f8; --accent: #007aff; }
        body { background: var(--bg); font-family: sans-serif; margin: 0; padding-bottom: 100px; }
        
        .header { 
            position: sticky; top: 0; display: flex; justify-content: flex-end; 
            padding: 15px; gap: 10px; background: rgba(242,245,248,0.9); z-index: 999;
        }
        
        .btn { 
            padding: 10px 15px; border-radius: 10px; border: none; font-size: 11px; 
            font-weight: 800; cursor: pointer; text-transform: uppercase;
            box-shadow: 4px 4px 8px #cfd8e3, -4px -4px 8px #fff;
        }
        .btn-blue { background: var(--accent); color: white; }

        .hero { display: flex; flex-direction: column; align-items: center; padding: 20px 0; }
        .logo { width: 200px; height: 200px; object-fit: contain; }
        
        .t-wrap { text-align: center; margin-top: 10px; }
        .t-1 { font-size: 72px; font-weight: 900; letter-spacing: -4px; color: #1d1d1f; line-height: 0.8; }
        .t-2 { 
            font-size: 72px; font-weight: 900; letter-spacing: -4px; margin-top: -32px; 
            transform: scaleY(-1); opacity: 0.15;
            background: linear-gradient(to bottom, #000, transparent);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }

        .container { padding: 0 20px; max-width: 600px; margin: 0 auto; }
        .card { 
            background: var(--bg); border-radius: 25px; padding: 20px; margin-bottom: 20px;
            box-shadow: 8px 8px 16px #cfd8e3, -8px -8px 16px #fff;
            display: flex; justify-content: space-between; align-items: center; cursor: pointer;
        }
        .card-num { font-size: 30px; font-weight: 900; }
        
        .hidden { display: none !important; }
    `;
    document.head.appendChild(style);
};

// Функции - ЖЕСТКАЯ ГЛОБАЛИЗАЦИЯ
window.modalP = function() {
    const n = prompt("Projekt Nummer:");
    if (n) {
        db.push({ num: n, name: prompt("Name:"), tools: [] });
        localStorage.setItem(DB_KEY, JSON.stringify(db));
        window.renderList();
    }
};

window.openProject = function(i) {
    currentIdx = i;
    window.renderDetails();
};

window.goHome = function() {
    currentIdx = null;
    window.renderList();
};

window.exportJSON = function() {
    const a = document.createElement('a');
    a.href = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(db));
    a.download = "cititool.json"; a.click();
};

window.importJSON = function() {
    const inp = document.createElement('input'); inp.type = 'file';
    inp.onchange = e => {
        const r = new FileReader();
        r.onload = f => { db = JSON.parse(f.target.result); localStorage.setItem(DB_KEY, JSON.stringify(db)); window.renderList(); };
        r.readAsText(e.target.files[0]);
    };
    inp.click();
};

window.deleteP = function(e, i) {
    e.stopPropagation();
    if (confirm("Löschen?")) {
        db.splice(i, 1);
        localStorage.setItem(DB_KEY, JSON.stringify(db));
        window.renderList();
    }
};

// Рендеринг
window.renderList = function() {
    const vHome = document.getElementById('v-home');
    const vDet = document.getElementById('v-det');
    if (!vHome || !vDet) return;

    vHome.classList.remove('hidden');
    vDet.classList.add('hidden');

    vHome.innerHTML = `
        <div class="header">
            <button class="btn" onclick="window.exportJSON()">Export</button>
            <button class="btn" onclick="window.importJSON()">Import</button>
            <button class="btn btn-blue" onclick="window.modalP()">+ NEU</button>
        </div>
        <div class="hero">
            <img src="${LOGO_URL}" class="logo">
            <div class="t-wrap">
                <div class="t-1">CitiTool</div>
                <div class="t-2">CitiTool</div>
            </div>
        </div>
        <div class="container" id="list"></div>
    `;

    document.getElementById('list').innerHTML = db.map((p, i) => `
        <div class="card" onclick="window.openProject(${i})">
            <div>
                <div style="font-size:10px; color:#999; font-weight:800;">PROJEKT</div>
                <div class="card-num">${p.num}</div>
                <div style="color:#666; font-weight:700;">${p.name || ''}</div>
            </div>
            <div style="color:red; font-weight:900;" onclick="window.deleteP(event, ${i})">✕</div>
        </div>
    `).join('');
};

window.renderDetails = function() {
    const p = db[currentIdx];
    const vHome = document.getElementById('v-home');
    const vDet = document.getElementById('v-det');

    vHome.classList.add('hidden');
    vDet.classList.remove('hidden');

    vDet.innerHTML = `
        <div class="header">
            <button class="btn" onclick="window.goHome()">← Home</button>
            <button class="btn btn-blue">+ Tool</button>
        </div>
        <div class="hero">
            <div class="t-wrap">
                <div class="t-1">${p.num}</div>
                <div style="font-weight:700; color:#666;">${p.name}</div>
            </div>
        </div>
        <div class="container">
            <div style="text-align:center; padding:50px; color:#ccc; font-weight:800; border:2px dashed #ddd; border-radius:20px;">
                REVOLVER TOOLS
            </div>
        </div>
    `;
};

// Старт
window.onload = () => {
    injectStyles();
    window.renderList();
};
