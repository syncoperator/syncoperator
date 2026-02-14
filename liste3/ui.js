const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png';

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f8f9fb; 
            --accent: #007aff; 
            --text-main: #1c1c1e;
            --text-sub: #8e8e93; 
            --card-bg: #ffffff;
            --shadow: 0 4px 15px rgba(0,0,0,0.05);
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important; 
            margin: 0; padding: 0; color: var(--text-main);
        }

        /* Центрированный контейнер приложения */
        .app-wrap { max-width: 500px; margin: 0 auto; padding-bottom: 120px; }

        /* Блок Брендинга: Логотип + Текст вплотную */
        .brand-header {
            display: flex; flex-direction: column; align-items: center;
            padding: 30px 20px 10px;
        }
        .logo-img { width: 120px; height: auto; margin-bottom: -5px; } /* Сблизили с текстом */
        .logo-text-wrap { text-align: center; }
        .logo-title { font-size: 50px; font-weight: 900; letter-spacing: -2.5px; line-height: 1; }
        .logo-mirror { 
            font-size: 50px; font-weight: 900; letter-spacing: -2.5px; 
            margin-top: -18px; transform: scaleY(-1); opacity: 0.1;
            background: linear-gradient(to bottom, #000, transparent);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }

        /* Панель кнопок: Ровный ряд */
        .action-bar {
            display: flex; justify-content: center; gap: 10px;
            padding: 15px 20px 25px;
        }
        .btn-p {
            background: #fff; border: none; border-radius: 12px;
            padding: 12px 0; width: 100px; font-size: 11px; font-weight: 800;
            color: #666; box-shadow: var(--shadow);
            text-transform: uppercase; cursor: pointer; transition: 0.2s;
        }
        .btn-p.blue { background: var(--accent); color: white; }
        .btn-p:active { transform: scale(0.95); }

        /* Карточки проектов */
        .card {
            background: var(--card-bg); border-radius: 20px;
            margin: 0 20px 12px; padding: 18px 22px;
            display: flex; align-items: center; justify-content: space-between;
            box-shadow: var(--shadow); border: 1px solid rgba(255,255,255,0.8);
        }
        .c-label { font-size: 10px; font-weight: 700; color: var(--text-sub); text-transform: uppercase; }
        .c-num { font-size: 24px; font-weight: 900; color: #000; letter-spacing: -0.5px; }
        .c-del { color: #ff3b30; font-size: 20px; padding: 5px; background: none; border: none; opacity: 0.3; }

        /* Плавающая кнопка */
        .fab {
            position: fixed; bottom: 30px; right: 25px;
            background: var(--accent); color: #fff; width: 60px; height: 60px;
            border-radius: 30px; display: flex; align-items: center; justify-content: center;
            font-size: 30px; box-shadow: 0 8px 25px rgba(0,122,255,0.3); border: none; z-index: 100;
        }

        /* Инструменты */
        .t-card {
            background: #fff; border-radius: 16px; margin: 0 16px 10px;
            padding: 15px 18px; display: flex; align-items: center; gap: 12px;
            box-shadow: var(--shadow);
        }
        .t-tag { background: #000; color: #fff; font-size: 8px; padding: 3px 6px; border-radius: 4px; font-weight: 900; margin-bottom: 4px; display: inline-block; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// --- РЕНДЕР ГЛАВНОЙ ---
function renderList() {
    const home = el('v-home');
    const det = el('v-det');
    home.style.display = 'block';
    det.style.display = 'none';

    home.innerHTML = `
        <div class="app-wrap">
            <div class="brand-header">
                <img src="${LOGO_URL}" class="logo-img">
                <div class="logo-text-wrap">
                    <div class="logo-title">CitiTool</div>
                    <div class="logo-mirror">CitiTool</div>
                </div>
            </div>
            
            <div class="action-bar">
                <button class="btn-p" onclick="exportJSON()">Export</button>
                <button class="btn-p" onclick="el('m-imp').classList.add('active')">Import</button>
                <button class="btn-p blue" onclick="modalP()">+ Neu</button>
            </div>

            <div id="projects-container"></div>
        </div>
    `;

    const container = el('projects-container');
    container.innerHTML = db.map((p, i) => `
        <div class="card" onclick="openProject(${i})">
            <div>
                <div class="c-label">${p.name || 'UNBENANNT'}</div>
                <div class="c-num">${p.num || '---'}</div>
            </div>
            <button class="c-del" onclick="event.stopPropagation(); deleteProject(${i})">✕</button>
        </div>
    `).join('');
}

// --- РЕНДЕР ИНСТРУМЕНТОВ ---
function renderTools() {
    const p = db[currentIdx];
    el('v-home').style.display = 'none';
    el('v-det').style.display = 'block';

    el('v-det').innerHTML = `
        <div class="app-wrap">
            <div class="action-bar">
                <button class="btn-p" onclick="goHome()">← Home</button>
                <div style="flex:1"></div>
                <button class="btn-p blue" onclick="makePDF()">PDF Report</button>
            </div>

            <div class="brand-header">
                <div class="logo-text-wrap">
                    <div class="logo-title">${p.num}</div>
                    <div class="c-label" style="margin-top:5px">${p.name}</div>
                </div>
            </div>

            <div id="tools-container"></div>
            <button class="fab" onclick="modalT()">+</button>
        </div>
    `;

    const container = el('tools-container');
    const tools = p.tools || [];
    container.innerHTML = tools.map((t, i) => `
        <div class="t-card" onclick="modalT(${i})">
            <div style="flex:1">
                ${t.rev ? `<span class="t-tag">REVOLVER UNTEN</span>` : ''}
                <div class="c-label">${t.id || 'T0000'}</div>
                <div style="font-size:17px; font-weight:800;">${t.nm || '---'}</div>
                <div style="color:var(--accent); font-weight:700; font-size:13px; margin-top:2px;">${t.dia || ''}</div>
            </div>
            <div style="display:flex; flex-direction:column; gap:4px">
                <button class="btn-p" style="width:34px; padding:5px 0;" onclick="event.stopPropagation(); moveItem(${i}, -1)">↑</button>
                <button class="btn-p" style="width:34px; padding:5px 0;" onclick="event.stopPropagation(); moveItem(${i}, 1)">↓</button>
            </div>
        </div>
    `).join('');
}

// --- ЛОГИКА (Функции открытия/сохранения остаются прежними) ---
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

// Функции modalP, saveP, modalT, saveT, exportJSON, importJSON, makePDF 
// должны быть в твоем коде ниже этого блока.

window.onload = () => {
    injectStyles();
    renderList();
};
