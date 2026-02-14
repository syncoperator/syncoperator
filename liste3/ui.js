const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png';

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f4f6f9; 
            --accent: #007aff; 
            --card: #ffffff;
            --text: #000000;
        }
        body { background: var(--bg); font-family: -apple-system, sans-serif; margin: 0; padding-bottom: 120px; }

        /* Брендинг: Огромное лого и текст вплотную */
        .brand-section { display: flex; flex-direction: column; align-items: center; padding: 40px 0 10px; }
        .logo-main { width: 220px; height: auto; margin-bottom: -25px; filter: contrast(1.1); }
        .logo-title { font-size: 60px; font-weight: 900; letter-spacing: -4px; margin: 0; line-height: 1; color: #000; }
        .logo-mirror { font-size: 60px; font-weight: 900; letter-spacing: -4px; margin-top: -26px; transform: scaleY(-1); opacity: 0.05; }

        /* Навигация */
        .nav-bar { display: flex; justify-content: center; gap: 12px; padding: 15px; position: sticky; top: 0; background: rgba(244,246,249,0.9); backdrop-filter: blur(15px); z-index: 100; }
        .btn-ui {
            background: #fff; border: none; border-radius: 14px; padding: 12px 20px;
            font-size: 11px; font-weight: 800; text-transform: uppercase;
            box-shadow: 0 4px 10px rgba(0,0,0,0.05); cursor: pointer;
        }
        .btn-ui.blue { background: var(--accent); color: #fff; }

        /* Карточки */
        .card {
            background: var(--card); border-radius: 24px; margin: 0 16px 15px; padding: 22px;
            display: flex; align-items: center; justify-content: space-between;
            box-shadow: 0 10px 30px rgba(0,0,0,0.04);
        }
        .p-label { font-size: 10px; font-weight: 800; color: #8e8e93; text-transform: uppercase; }
        .p-num { font-size: 30px; font-weight: 900; letter-spacing: -1px; }

        /* Модалки */
        .modal { position: fixed; inset: 0; background: rgba(0,0,0,0.4); display: none; align-items: center; justify-content: center; z-index: 2000; backdrop-filter: blur(10px); }
        .modal.active { display: flex; }
        .m-box { background: #fff; width: 90%; max-width: 400px; padding: 30px; border-radius: 35px; }

        /* РАБОЧАЯ КНОПКА РЕВОЛЬВЕРА */
        .rev-btn {
            width: 100%; padding: 18px; border-radius: 15px; margin: 15px 0;
            border: 2px solid #eee; font-weight: 900; font-size: 14px;
            text-align: center; cursor: pointer; transition: 0.3s;
            background: #f9f9f9;
        }
        .rev-btn.active-rev { background: #000; color: #fff; border-color: #000; }

        .fab { position: fixed; bottom: 30px; right: 25px; background: var(--accent); color: #fff; width: 65px; height: 65px; border-radius: 35px; border: none; font-size: 35px; box-shadow: 0 10px 25px rgba(0,122,255,0.4); }
        
        input { width: 100%; padding: 14px; border-radius: 12px; border: 1.5px solid #eee; margin-top: 5px; box-sizing: border-box; font-size: 16px; font-weight: 600; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const show = (id) => el(id).classList.add('active');
const hide = (id) => el(id).classList.remove('active');
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// --- РЕНДЕР ГЛАВНОЙ ---
function renderList() {
    el('v-home').style.display = 'block';
    el('v-det').style.display = 'none';
    el('v-home').innerHTML = `
        <div class="nav-bar">
            <button class="btn-ui" onclick="exportJSON()">Export</button>
            <button class="btn-ui" onclick="show('m-imp')">Import</button>
            <button class="btn-ui blue" onclick="modalP()">+ NEU</button>
        </div>
        <div class="brand-section">
            <img src="${LOGO_URL}" class="logo-main">
            <h1 class="logo-title">CitiTool</h1>
            <div class="logo-mirror">CitiTool</div>
        </div>
        <div id="p-container"></div>
    `;
    el('p-container').innerHTML = db.map((p, i) => `
        <div class="card" onclick="openProject(${i})">
            <div>
                <div class="p-label">${p.name || 'Project'}</div>
                <div class="p-num">${p.num || '---'}</div>
            </div>
            <button style="border:none; background:none; color:red; font-size:24px; opacity:0.2;" onclick="event.stopPropagation(); deleteProject(${i})">✕</button>
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
            <button class="btn-ui" onclick="goHome()">← Home</button>
            <button class="btn-ui blue" onclick="makePDF()">PDF REPORT</button>
        </div>
        <div class="brand-section">
            <div class="p-label">Zeichnungsnummer</div>
            <h1 class="logo-title">${p.num}</h1>
            <div class="p-label" style="margin-top:5px; color:#000">${p.name}</div>
        </div>
        <div id="t-container"></div>
        <button class="fab" onclick="modalT()">+</button>
    `;
    const tools = p.tools || [];
    el('t-container').innerHTML = tools.map((t, i) => `
        <div class="card" onclick="modalT(${i})">
            <div style="flex:1">
                ${t.rev ? `<div style="background:#000; color:#fff; font-size:8px; padding:3px 6px; border-radius:5px; width:fit-content; margin-bottom:8px; font-weight:900;">REVOLVER UNTEN</div>` : ''}
                <div class="p-label">${t.id}</div>
                <div style="font-size:21px; font-weight:900;">${t.nm}</div>
                <div style="color:var(--accent); font-weight:800; margin-top:3px;">${t.dia}</div>
            </div>
            <div style="display:flex; flex-direction:column; gap:5px;">
                <button class="btn-ui" style="padding:8px 14px;" onclick="event.stopPropagation(); moveItem(${i}, -1)">↑</button>
                <button class="btn-ui" style="padding:8px 14px;" onclick="event.stopPropagation(); moveItem(${i}, 1)">↓</button>
            </div>
        </div>
    `).join('') + '<div style="height:150px"></div>';
}

// --- ЛОГИКА ---
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

// ПЕРЕКЛЮЧАТЕЛЬ РЕВОЛЬВЕРА
function toggleRev() {
    const btn = el('btn-rev-toggle');
    btn.classList.toggle('active-rev');
    btn.innerText = btn.classList.contains('active-rev') ? 'REVOLVER UNTEN: START' : 'REVOLVER OBEN';
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    
    const btn = el('btn-rev-toggle');
    if(t.rev) {
        btn.classList.add('active-rev');
        btn.innerText = 'REVOLVER UNTEN: START';
    } else {
        btn.classList.remove('active-rev');
        btn.innerText = 'REVOLVER OBEN';
    }
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const t = {
        id: el('t-id').value.toUpperCase(),
        nm: el('t-nm').value.toUpperCase(),
        dia: el('t-dia').value,
        rev: el('btn-rev-toggle').classList.contains('active-rev')
    };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    save(); renderTools(); hide('m-t');
}

function modalP() {
    el('p-idx').value = ''; el('p-num').value = ''; el('p-nam').value = '';
    show('m-p');
}

function saveP() {
    const data = {
        num: el('p-num').value, name: el('p-nam').value.toUpperCase(),
        lzf:'', sag:'', stt:'', stn:'', abs:'', grf:'', mat:'', tools: []
    };
    db.push(data); save(); hide('m-p'); renderList();
}

// PDF И ЭКСПОРТ
function exportJSON() { el('imp-area').value = JSON.stringify(db); el('imp-area').select(); document.execCommand('copy'); alert("JSON Copied!"); }
function importJSON() { try { db = JSON.parse(el('imp-area').value); save(); renderList(); hide('m-imp'); } catch(e){ alert("Error"); } }

function makePDF() {
    // Вызывает системное окно печати текущего вида
    window.print();
}

window.onload = () => { injectStyles(); renderList(); };
