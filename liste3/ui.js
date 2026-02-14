const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png';

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    if (document.getElementById('main-styles')) return;
    const style = document.createElement('style');
    style.id = 'main-styles';
    style.innerHTML = `
        :root { 
            --bg: #f8f9fb; 
            --accent: #007aff; 
            --text-main: #1c1c1e;
            --text-sub: #8e8e93; 
            --card-bg: #ffffff;
            --neu-shadow: #cfd8e3;
            --neu-light: #ffffff;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important; 
            margin: 0; color: var(--text-main);
            padding-bottom: 120px;
        }

        /* Контейнер для всего контента */
        .app-container { max-width: 600px; margin: 0 auto; }

        /* Шапка с логотипом */
        .brand-section {
            display: flex; flex-direction: column; align-items: center;
            padding: 40px 0 20px; background: var(--bg);
        }
        .logo-main { width: 200px; height: auto; margin-bottom: 10px; }
        .mirror-text { text-align: center; position: relative; }
        .t-real { font-size: 48px; font-weight: 900; letter-spacing: -2px; color: #000; line-height: 1; }
        .t-mirror { 
            font-size: 48px; font-weight: 900; letter-spacing: -2px; 
            margin-top: -20px; transform: scaleY(-1); opacity: 0.1;
            background: linear-gradient(to bottom, #000, transparent);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }

        /* Панель управления (кнопки) */
        .control-panel {
            display: flex; justify-content: center; gap: 12px;
            padding: 20px; sticky; top: 0; z-index: 50; background: rgba(248,249,251,0.8);
            backdrop-filter: blur(10px);
        }
        .btn-premium {
            background: var(--card-bg); border: none; border-radius: 14px;
            padding: 12px 20px; font-size: 12px; font-weight: 800; color: #666;
            box-shadow: 4px 4px 10px var(--neu-shadow), -4px -4px 10px var(--neu-light);
            text-transform: uppercase; cursor: pointer; transition: 0.2s;
        }
        .btn-premium.blue { background: var(--accent); color: white; box-shadow: 0 6px 15px rgba(0,122,255,0.3); }
        .btn-premium:active { transform: scale(0.96); }

        /* Карточки проектов */
        .project-card {
            background: var(--card-bg); border-radius: 24px;
            margin: 0 20px 16px; padding: 24px;
            display: flex; align-items: center; justify-content: space-between;
            box-shadow: 0 10px 30px rgba(0,0,0,0.03);
            border: 1px solid rgba(255,255,255,0.7);
        }
        .p-label { font-size: 11px; font-weight: 700; color: var(--text-sub); text-transform: uppercase; margin-bottom: 4px; }
        .p-title { font-size: 26px; font-weight: 900; color: #000; letter-spacing: -1px; }

        /* Инструменты */
        .tool-card {
            background: #fff; border-radius: 20px;
            padding: 16px 20px; margin: 0 16px 12px;
            display: flex; align-items: center; justify-content: space-between;
            box-shadow: 0 4px 15px rgba(0,0,0,0.02);
        }
        .rev-tag { 
            background: #000; color: #fff; font-size: 9px; font-weight: 900;
            padding: 3px 8px; border-radius: 6px; margin-bottom: 6px; width: fit-content;
        }

        .btn-delete { color: #ff3b30; font-size: 20px; padding: 10px; opacity: 0.3; }
        .btn-delete:hover { opacity: 1; }

        .fab {
            position: fixed; bottom: 30px; right: 25px;
            background: var(--accent); color: #fff; width: 65px; height: 65px;
            border-radius: 50%; display: flex; align-items: center; justify-content: center;
            font-size: 32px; box-shadow: 0 10px 25px rgba(0,122,255,0.4); border: none; z-index: 100;
        }
        
        /* Скрытые элементы */
        .view { display: none; }
        .view.active { display: block; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// --- РЕНДЕР ГЛАВНОЙ ---
function renderList() {
    const vHome = el('v-home');
    const vDet = el('v-det');
    vHome.classList.add('active');
    vDet.classList.remove('active');

    vHome.innerHTML = `
        <div class="brand-section">
            <img src="${LOGO_URL}" class="logo-main" onerror="this.style.display='none'">
            <div class="mirror-text">
                <div class="t-real">CitiTool</div>
                <div class="t-mirror">CitiTool</div>
            </div>
        </div>
        <div class="control-panel">
            <button class="btn-premium" onclick="exportJSON()">Export</button>
            <button class="btn-premium" onclick="show('m-imp')">Import</button>
            <button class="btn-premium blue" onclick="modalP()">+ NEU</button>
        </div>
        <div id="list-p-content"></div>
    `;

    const content = el('list-p-content');
    content.innerHTML = db.map((p, i) => `
        <div class="project-card" onclick="openProject(${i})">
            <div class="p-info">
                <div class="p-label">${p.name || 'UNBENANNT'}</div>
                <div class="p-title">${p.num || '---'}</div>
            </div>
            <div class="btn-delete" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>
    `).join('') + '<div style="height:50px"></div>';
}

// --- РЕНДЕР ИНСТРУМЕНТОВ ---
function renderTools() {
    const vHome = el('v-home');
    const vDet = el('v-det');
    const p = db[currentIdx];
    
    vHome.classList.remove('active');
    vDet.classList.add('active');

    vDet.innerHTML = `
        <div class="control-panel">
            <button class="btn-premium" onclick="goHome()">← Back</button>
            <button class="btn-premium blue" onclick="makePDF()">PDF Report</button>
        </div>
        <div class="brand-section">
            <div class="mirror-text">
                <div class="t-real">${p.num}</div>
                <div class="p-label" style="margin-top:10px">${p.name}</div>
            </div>
        </div>
        <div id="list-t-content"></div>
        <button class="fab" onclick="modalT()">+</button>
    `;

    const content = el('list-t-content');
    const tools = p.tools || [];
    content.innerHTML = tools.map((t, i) => `
        <div class="tool-card" onclick="modalT(${i})">
            <div style="flex:1">
                ${t.rev ? `<div class="rev-tag">REVOLVER UNTEN</div>` : ''}
                <div class="p-label">${t.id || 'T0000'}</div>
                <div style="font-size:18px; font-weight:800;">${t.nm || '---'}</div>
                <div style="color:var(--accent); font-weight:700; margin-top:4px;">${t.dia || ''}</div>
            </div>
            <div style="display:flex; flex-direction:column; gap:5px">
                <button class="btn-premium" style="padding:5px 10px" onclick="event.stopPropagation(); moveItem(${i}, -1)">↑</button>
                <button class="btn-premium" style="padding:5px 10px" onclick="event.stopPropagation(); moveItem(${i}, 1)">↓</button>
            </div>
        </div>
    `).join('') + '<div style="padding-bottom:100px"></div>';
}

// --- ФУНКЦИИ УПРАВЛЕНИЯ (Твой оригинал) ---
function openProject(i) { currentIdx = i; renderTools(); }
function goHome() { currentIdx = null; renderList(); }
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); save(); renderList(); } }

function moveItem(i, direction) {
    const tools = db[currentIdx].tools;
    const target = i + direction;
    if (target >= 0 && target < tools.length) {
        [tools[i], tools[target]] = [tools[target], tools[i]];
        save(); renderTools();
    }
}

// Привязываем инициализацию
window.onload = () => {
    injectStyles();
    renderList();
};

// Все модалки и экспорт (saveP, saveT, makePDF) используй из своего предыдущего рабочего кода, они не изменились.
