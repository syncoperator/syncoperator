const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png';

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f2f4f7; 
            --accent: #007aff; 
            --card-bg: #ffffff;
            --text-main: #1c1c1e;
            --text-sub: #8e8e93;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, sans-serif !important; 
            margin: 0; padding: 0; color: var(--text-main);
            padding-bottom: 120px;
        }

        /* Брендинг: Лого и Текст вплотную */
        .brand-section { display: flex; flex-direction: column; align-items: center; padding: 30px 0 10px; }
        .logo-img { width: 100px; height: auto; margin-bottom: -12px; z-index: 2; }
        .logo-title { font-size: 54px; font-weight: 900; letter-spacing: -3px; color: #000; line-height: 1; position: relative; }
        .logo-mirror { 
            font-size: 54px; font-weight: 900; letter-spacing: -3px; margin-top: -22px; 
            transform: scaleY(-1); opacity: 0.08;
            background: linear-gradient(to bottom, #000, transparent);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }

        /* Панель навигации */
        .nav-bar { display: flex; justify-content: center; gap: 10px; padding: 15px; position: sticky; top: 0; background: rgba(242,244,247,0.8); backdrop-filter: blur(15px); z-index: 100; }
        .btn-premium {
            background: #fff; border: none; border-radius: 14px;
            padding: 12px 18px; font-size: 11px; font-weight: 800; color: #555;
            box-shadow: 5px 5px 10px #d1d9e6, -5px -5px 10px #ffffff;
            text-transform: uppercase; cursor: pointer; transition: 0.2s;
        }
        .btn-premium.blue { background: var(--accent); color: #fff; box-shadow: 0 5px 15px rgba(0,122,255,0.3); }
        .btn-premium:active { transform: scale(0.96); box-shadow: inset 2px 2px 5px #d1d9e6; }

        /* Карточки */
        .card {
            background: #fff; border-radius: 24px; margin: 0 20px 15px; padding: 22px;
            display: flex; align-items: center; justify-content: space-between;
            box-shadow: 0 10px 25px rgba(0,0,0,0.03); border: 1px solid rgba(255,255,255,0.7);
        }
        .p-label { font-size: 10px; font-weight: 800; color: var(--text-sub); text-transform: uppercase; letter-spacing: 0.5px; }
        .p-num { font-size: 28px; font-weight: 900; color: #000; margin-top: 2px; }

        /* Селектор Револьвера в модалке */
        .rev-selector {
            width: 100%; padding: 15px; border-radius: 12px; border: 2px solid #eee;
            margin: 10px 0; font-weight: 800; text-align: center; cursor: pointer; transition: 0.3s;
        }
        .rev-selector.on { background: #000; color: #fff; border-color: #000; }

        .fab { position: fixed; bottom: 30px; right: 25px; background: var(--accent); color: #fff; width: 65px; height: 65px; border-radius: 35px; font-size: 35px; border: none; box-shadow: 0 10px 25px rgba(0,122,255,0.4); z-index: 1000; }
        
        /* Модалки */
        .modal { position: fixed; inset: 0; background: rgba(0,0,0,0.4); display: none; align-items: center; justify-content: center; z-index: 9999; backdrop-filter: blur(8px); }
        .modal.active { display: flex; }
        .modal-content { background: #fff; width: 90%; max-width: 400px; padding: 25px; border-radius: 30px; box-shadow: 0 20px 50px rgba(0,0,0,0.2); }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// --- РЕНДЕР ГЛАВНОЙ ---
function renderList() {
    el('v-home').style.display = 'block';
    el('v-det').style.display = 'none';
    
    el('v-home').innerHTML = `
        <div class="nav-bar">
            <button class="btn-premium" onclick="exportJSON()">Export</button>
            <button class="btn-premium" onclick="el('m-imp').classList.add('active')">Import</button>
            <button class="btn-premium blue" onclick="modalP()">+ NEU</button>
        </div>
        <div class="brand-section">
            <img src="${LOGO_URL}" class="logo-img">
            <div class="logo-title">CitiTool</div>
            <div class="logo-mirror">CitiTool</div>
        </div>
        <div id="list-p-box"></div>
    `;

    el('list-p-box').innerHTML = db.map((p, i) => `
        <div class="card" onclick="openProject(${i})">
            <div>
                <div class="p-label">${p.name || 'UNBENANNT'}</div>
                <div class="p-num">${p.num || '---'}</div>
            </div>
            <button style="border:none; background:none; color:#ff3b30; font-size:20px; opacity:0.3;" onclick="event.stopPropagation(); deleteProject(${i})">✕</button>
        </div>
    `).join('') + '<div style="height:50px"></div>';
}

// --- РЕНДЕР ИНСТРУМЕНТОВ ---
function renderTools() {
    const p = db[currentIdx];
    el('v-home').style.display = 'none';
    el('v-det').style.display = 'block';

    el('v-det').innerHTML = `
        <div class="nav-bar">
            <button class="btn-premium" onclick="goHome()">← Back</button>
            <button class="btn-premium blue" onclick="makePDF()">PDF Report</button>
        </div>
        <div class="brand-section">
            <div class="p-label">PROJECT NUMBER</div>
            <div class="logo-title">${p.num}</div>
            <div class="p-label" style="margin-top:5px; color:#000;">${p.name}</div>
        </div>
        <div id="list-t-box"></div>
        <button class="fab" onclick="modalT()">+</button>
    `;

    const tools = p.tools || [];
    el('list-t-box').innerHTML = tools.map((t, i) => `
        <div class="card" onclick="modalT(${i})">
            <div style="flex:1">
                ${t.rev ? `<div style="background:#000; color:#fff; font-size:9px; padding:3px 7px; border-radius:6px; width:fit-content; font-weight:900; margin-bottom:8px;">REVOLVER UNTEN</div>` : ''}
                <div class="p-label">${t.id}</div>
                <div style="font-size:20px; font-weight:900; color:#000;">${t.nm}</div>
                <div style="color:var(--accent); font-weight:800; font-size:14px; margin-top:4px;">${t.dia}</div>
            </div>
            <div style="display:flex; flex-direction:column; gap:6px;">
                <button class="btn-premium" style="padding:6px 12px;" onclick="event.stopPropagation(); moveItem(${i}, -1)">↑</button>
                <button class="btn-premium" style="padding:6px 12px;" onclick="event.stopPropagation(); moveItem(${i}, 1)">↓</button>
            </div>
        </div>
    `).join('') + '<div style="height:120px"></div>';
}

// --- ФУНКЦИИ УПРАВЛЕНИЯ ---
function openProject(i) { currentIdx = i; renderTools(); }
function goHome() { currentIdx = null; renderList(); }
function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); save(); renderList(); } }

function moveItem(i, dir) {
    const tools = db[currentIdx].tools;
    const target = i + dir;
    if (target >= 0 && target < tools.length) {
        [tools[i], tools[target]] = [tools[target], tools[i]];
        save(); renderTools();
    }
}

// Переключатель револьвера (ТО, ЧТО НЕ РАБОТАЛО)
function toggleRev() {
    const btn = el('btn-rev-toggle');
    btn.classList.toggle('on');
    btn.innerText = btn.classList.contains('on') ? 'REVOLVER UNTEN: START' : 'REVOLVER OBEN';
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    
    el('t-id').value = t.id; 
    el('t-nm').value = t.nm; 
    el('t-dia').value = t.dia;
    
    const btn = el('btn-rev-toggle');
    if(t.rev) {
        btn.classList.add('on');
        btn.innerText = 'REVOLVER UNTEN: START';
    } else {
        btn.classList.remove('on');
        btn.innerText = 'REVOLVER OBEN';
    }
    el('m-t').classList.add('active');
}

function saveT() {
    const i = el('t-idx').value;
    const isRev = el('btn-rev-toggle').classList.contains('on');
    const t = { 
        id: el('t-id').value.toUpperCase(), 
        nm: el('t-nm').value.toUpperCase(), 
        dia: el('t-dia').value, 
        rev: isRev 
    };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    save(); renderTools(); el('m-t').classList.remove('active');
}

// Остальные функции модалок проекта
function modalP() {
    el('p-idx').value = ''; el('p-num').value = ''; el('p-nam').value = '';
    el('m-p').classList.add('active');
}

function saveP() {
    const data = {
        num: el('p-num').value, name: el('p-nam').value.toUpperCase(),
        lzf:'', sag:'', stt:'', stn:'', abs:'', grf:'', mat:'', tools: []
    };
    db.push(data); save(); el('m-p').classList.remove('active'); renderList();
}

function exportJSON() { el('imp-area').value = JSON.stringify(db); el('imp-area').select(); document.execCommand('copy'); alert("Copied!"); }
function importJSON() { try { db = JSON.parse(el('imp-area').value); save(); renderList(); el('m-imp').classList.remove('active'); } catch(e){alert("Error");} }

window.onload = () => { injectStyles(); renderList(); };
