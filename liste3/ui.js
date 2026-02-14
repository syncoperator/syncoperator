const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png';

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f0f2f5; 
            --accent: #007aff; 
            --card: #ffffff;
            --text: #1c1c1e;
            --sub: #8e8e93;
            --neu-in: inset 2px 2px 5px #d1d9e6, inset -2px -2px 5px #ffffff;
            --neu-out: 5px 5px 10px #d1d9e6, -5px -5px 10px #ffffff;
        }
        body { background: var(--bg); font-family: -apple-system, sans-serif; margin: 0; padding-bottom: 100px; color: var(--text); }

        /* Логотип и Текст — Сближение */
        .brand-box { display: flex; flex-direction: column; align-items: center; padding: 20px 0 5px; }
        .logo-main { width: 90px; height: auto; margin-bottom: -15px; z-index: 10; filter: drop-shadow(0 2px 4px rgba(0,0,0,0.1)); }
        .logo-h1 { font-size: 56px; font-weight: 900; letter-spacing: -3px; margin: 0; line-height: 0.9; color: #000; }
        .logo-mirror { font-size: 56px; font-weight: 900; letter-spacing: -3px; margin-top: -24px; transform: scaleY(-1); opacity: 0.05; }

        /* Навигация */
        .top-nav { display: flex; justify-content: center; gap: 12px; padding: 15px; position: sticky; top: 0; background: rgba(240,242,245,0.8); backdrop-filter: blur(20px); z-index: 100; }
        .btn-p {
            background: var(--card); border: none; border-radius: 12px; padding: 12px 20px;
            font-size: 11px; font-weight: 800; text-transform: uppercase; color: #444;
            box-shadow: var(--neu-out); cursor: pointer; transition: 0.2s;
        }
        .btn-p.blue { background: var(--accent); color: #fff; box-shadow: 0 4px 15px rgba(0,122,255,0.3); }
        .btn-p:active { transform: scale(0.96); }

        /* Карточки */
        .item-card {
            background: var(--card); border-radius: 22px; margin: 0 16px 12px; padding: 20px;
            display: flex; align-items: center; justify-content: space-between;
            box-shadow: 0 8px 20px rgba(0,0,0,0.03); border: 1px solid #fff;
        }
        .tag { font-size: 9px; font-weight: 900; color: var(--sub); text-transform: uppercase; }
        .title { font-size: 26px; font-weight: 900; letter-spacing: -0.5px; }

        /* Модалки (Центрирование и Качество) */
        .modal { position: fixed; inset: 0; background: rgba(0,0,0,0.3); backdrop-filter: blur(10px); display: none; align-items: center; justify-content: center; z-index: 1000; }
        .modal.active { display: flex; }
        .m-content { background: #fff; width: 88%; max-width: 420px; padding: 25px; border-radius: 32px; box-shadow: 0 30px 60px rgba(0,0,0,0.15); }
        
        /* Кнопка-переключатель Револьвера */
        .rev-toggle {
            width: 100%; padding: 16px; border-radius: 14px; margin: 15px 0;
            text-align: center; font-weight: 900; font-size: 13px;
            background: #f8f9fa; border: 2px solid #eee; cursor: pointer; transition: 0.3s;
        }
        .rev-toggle.on { background: #000; color: #fff; border-color: #000; }

        .fab { position: fixed; bottom: 30px; right: 25px; background: var(--accent); color: #fff; width: 64px; height: 64px; border-radius: 32px; border: none; font-size: 32px; box-shadow: 0 10px 25px rgba(0,122,255,0.4); }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const hide = (id) => el(id).classList.remove('active');
const show = (id) => el(id).classList.add('active');
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

// --- РЕНДЕР ГЛАВНОЙ ---
function renderList() {
    el('v-home').style.display = 'block';
    el('v-det').style.display = 'none';
    el('v-home').innerHTML = `
        <div class="top-nav">
            <button class="btn-p" onclick="exportJSON()">Export</button>
            <button class="btn-p" onclick="show('m-imp')">Import</button>
            <button class="btn-p blue" onclick="modalP()">+ NEU</button>
        </div>
        <div class="brand-box">
            <img src="${LOGO_URL}" class="logo-main">
            <h1 class="logo-h1">CitiTool</h1>
            <div class="logo-mirror">CitiTool</div>
        </div>
        <div id="p-list"></div>
    `;
    el('p-list').innerHTML = db.map((p, i) => `
        <div class="item-card" onclick="openProject(${i})">
            <div>
                <div class="tag">${p.name || 'Projekt'}</div>
                <div class="title">${p.num || '---'}</div>
            </div>
            <button style="border:none; background:none; color:red; font-size:22px; opacity:0.2;" onclick="event.stopPropagation(); deleteProject(${i})">✕</button>
        </div>
    `).join('') + '<div style="height:50px"></div>';
}

// --- РЕНДЕР ИНСТРУМЕНТОВ ---
function renderTools() {
    const p = db[currentIdx];
    el('v-home').style.display = 'none';
    el('v-det').style.display = 'block';
    el('v-det').innerHTML = `
        <div class="top-nav">
            <button class="btn-p" onclick="goHome()">← Back</button>
            <button class="btn-p blue" onclick="makePDF()">PDF REPORT</button>
        </div>
        <div class="brand-box">
            <div class="tag">Drawing Number</div>
            <h1 class="logo-h1">${p.num}</h1>
            <div class="tag" style="margin-top:5px; color:#000">${p.name}</div>
        </div>
        <div id="t-list"></div>
        <button class="fab" onclick="modalT()">+</button>
    `;
    const tools = p.tools || [];
    el('t-list').innerHTML = tools.map((t, i) => `
        <div class="item-card" onclick="modalT(${i})">
            <div style="flex:1">
                ${t.rev ? `<div style="background:#000; color:#fff; font-size:8px; padding:3px 6px; border-radius:5px; width:fit-content; margin-bottom:6px; font-weight:900;">REVOLVER UNTEN</div>` : ''}
                <div class="tag">${t.id}</div>
                <div style="font-size:19px; font-weight:900;">${t.nm}</div>
                <div style="color:var(--accent); font-weight:800; font-size:13px; margin-top:2px;">${t.dia}</div>
            </div>
            <div style="display:flex; flex-direction:column; gap:5px;">
                <button class="btn-p" style="padding:6px 12px;" onclick="event.stopPropagation(); moveItem(${i}, -1)">↑</button>
                <button class="btn-p" style="padding:6px 12px;" onclick="event.stopPropagation(); moveItem(${i}, 1)">↓</button>
            </div>
        </div>
    `).join('') + '<div style="height:120px"></div>';
}

// --- ФУНКЦИИ ---
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

// РАБОЧИЙ ТРЕКЕР РЕВОЛЬВЕРА
function toggleRev() {
    const btn = el('btn-rev-toggle');
    const isOff = !btn.classList.contains('on');
    if(isOff) {
        btn.classList.add('on');
        btn.innerText = 'REVOLVER UNTEN: AKTIVIERT';
    } else {
        btn.classList.remove('on');
        btn.innerText = 'REVOLVER OBEN';
    }
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    
    const btn = el('btn-rev-toggle');
    if(t.rev) {
        btn.classList.add('on');
        btn.innerText = 'REVOLVER UNTEN: AKTIVIERT';
    } else {
        btn.classList.remove('on');
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
        rev: el('btn-rev-toggle').classList.contains('on')
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

// PDF и ЭКСПОРТ
function exportJSON() { el('imp-area').value = JSON.stringify(db); el('imp-area').select(); document.execCommand('copy'); alert("JSON Copied!"); }
function importJSON() { try { db = JSON.parse(el('imp-area').value); save(); renderList(); hide('m-imp'); } catch(e){ alert("Error"); } }

function makePDF() {
    const p = db[currentIdx];
    // Здесь твой код PDF из предыдущих версий (он рабочий)
    // ... логика генерации HTML для печати ...
    window.print();
}

window.onload = () => { injectStyles(); renderList(); };
