const DB_KEY = 'QS_DATA_V8';
const LOGO_URL = 'https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png';

let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    // Удаляем старые стили, если они были
    const oldStyle = document.getElementById('main-ui-styles');
    if(oldStyle) oldStyle.remove();

    const style = document.createElement('style');
    style.id = 'main-ui-styles';
    style.innerHTML = `
        :root { --bg: #f2f4f7; --accent: #007aff; --card: #ffffff; --text: #000; }
        body { background: var(--bg); font-family: -apple-system, BlinkMacSystemFont, sans-serif; margin: 0; padding-bottom: 150px; -webkit-font-smoothing: antialiased; }

        /* Брендинг: Максимальное сближение */
        .brand-section { display: flex; flex-direction: column; align-items: center; padding: 20px 0 0px; }
        .logo-main { width: 250px !important; height: auto; margin-bottom: -35px !important; z-index: 10; position: relative; }
        .logo-title { font-size: 68px; font-weight: 900; letter-spacing: -4px; margin: 0; line-height: 0.8; color: #000; z-index: 11; position: relative; }
        .logo-mirror { font-size: 68px; font-weight: 900; letter-spacing: -4px; margin-top: -32px; transform: scaleY(-1); opacity: 0.04; filter: blur(1px); }

        /* Кнопки Навигации */
        .nav-bar { display: flex; justify-content: center; gap: 12px; padding: 15px; position: sticky; top: 0; background: rgba(242,244,247,0.8); backdrop-filter: blur(20px); z-index: 100; }
        .btn-ui { background: #fff; border: none; border-radius: 14px; padding: 12px 20px; font-size: 11px; font-weight: 800; text-transform: uppercase; box-shadow: 0 4px 12px rgba(0,0,0,0.08); cursor: pointer; transition: 0.2s; }
        .btn-ui.blue { background: var(--accent); color: #fff; box-shadow: 0 6px 16px rgba(0,122,255,0.3); }

        /* Карточки */
        .card { background: var(--card); border-radius: 26px; margin: 0 16px 15px; padding: 22px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 10px 30px rgba(0,0,0,0.05); border: 1px solid rgba(255,255,255,0.8); }
        .p-label { font-size: 10px; font-weight: 800; color: #8e8e93; text-transform: uppercase; letter-spacing: 0.5px; }
        .p-num { font-size: 32px; font-weight: 900; letter-spacing: -1.5px; color: #000; line-height: 1.1; }

        /* МОДАЛКИ (Исправлено) */
        .modal { position: fixed; inset: 0; background: rgba(0,0,0,0.5); display: none; align-items: center; justify-content: center; z-index: 9999; backdrop-filter: blur(12px); }
        .modal.active { display: flex !important; }
        .m-box { background: #fff; width: 92%; max-width: 420px; padding: 30px; border-radius: 38px; box-sizing: border-box; box-shadow: 0 40px 80px rgba(0,0,0,0.2); }

        /* ТУМБЛЕР РЕВОЛЬВЕРА */
        .rev-container { display: flex; align-items: center; justify-content: space-between; background: #f8f9fa; padding: 15px; border-radius: 18px; margin: 20px 0; border: 1.5px solid #eee; }
        .rev-text { font-weight: 900; font-size: 13px; color: #333; }
        .rev-toggle-btn {
            background: #fff; border: 2px solid #ddd; border-radius: 12px;
            padding: 10px 18px; font-weight: 900; font-size: 11px; cursor: pointer;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .rev-toggle-btn.is-active { background: #000 !important; color: #fff !important; border-color: #000 !important; transform: scale(1.05); }

        .fab { position: fixed; bottom: 35px; right: 25px; background: var(--accent); color: #fff; width: 68px; height: 68px; border-radius: 34px; border: none; font-size: 38px; box-shadow: 0 12px 28px rgba(0,122,255,0.4); z-index: 500; }
        
        input { width: 100%; padding: 16px; border-radius: 15px; border: 2px solid #f0f0f0; margin-top: 6px; box-sizing: border-box; font-size: 16px; font-weight: 700; margin-bottom: 12px; transition: 0.3s; }
        input:focus { border-color: var(--accent); outline: none; background: #fff; }
        
        .m-btns { display: flex; gap: 12px; margin-top: 25px; }
        .m-btns button { flex: 1; padding: 16px; border-radius: 16px; font-weight: 900; text-transform: uppercase; font-size: 12px; cursor: pointer; border: none; }
        .btn-save { background: #34c759; color: #fff; }
        .btn-cancel { background: #f2f2f7; color: #8e8e93; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const show = (id) => { const m = el(id); if(m) m.classList.add('active'); };
const hide = (id) => { const m = el(id); if(m) m.classList.remove('active'); };
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
                <div class="p-label">${p.name || 'UNBENANNT'}</div>
                <div class="p-num">${p.num || '---'}</div>
            </div>
            <button style="border:none; background:none; color:#ff3b30; font-size:26px; opacity:0.3;" onclick="event.stopPropagation(); deleteProject(${i})">✕</button>
        </div>
    `).join('') + '<div style="height:80px"></div>';
}

// --- РЕНДЕР ИНСТРУМЕНТОВ ---
function renderTools() {
    const p = db[currentIdx];
    el('v-home').style.display = 'none';
    el('v-det').style.display = 'block';
    el('v-det').innerHTML = `
        <div class="nav-bar">
            <button class="btn-ui" onclick="goHome()">← Back</button>
            <button class="btn-ui blue" onclick="makePDF()">PDF Report</button>
        </div>
        <div class="brand-section">
            <div class="p-label">Projektnummer</div>
            <h1 class="logo-title">${p.num}</h1>
            <div class="p-label" style="margin-top:8px; color:#000; font-size:12px;">${p.name}</div>
        </div>
        <div id="t-container"></div>
        <button class="fab" onclick="modalT()">+</button>
    `;
    const tools = p.tools || [];
    el('t-container').innerHTML = tools.map((t, i) => `
        <div class="card" onclick="modalT(${i})">
            <div style="flex:1">
                ${t.rev ? `<div style="background:#000; color:#fff; font-size:9px; padding:4px 8px; border-radius:6px; width:fit-content; margin-bottom:10px; font-weight:900;">REVOLVER UNTEN</div>` : ''}
                <div class="p-label">${t.id}</div>
                <div style="font-size:22px; font-weight:900; color:#000;">${t.nm}</div>
                <div style="color:var(--accent); font-weight:800; margin-top:4px; font-size:15px;">${t.dia}</div>
            </div>
            <div style="display:flex; flex-direction:column; gap:6px;">
                <button class="btn-ui" style="padding:10px 16px;" onclick="event.stopPropagation(); moveItem(${i}, -1)">↑</button>
                <button class="btn-ui" style="padding:10px 16px;" onclick="event.stopPropagation(); moveItem(${i}, 1)">↓</button>
            </div>
        </div>
    `).join('') + '<div style="height:180px"></div>';
}

// --- УПРАВЛЕНИЕ ---
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

// ПЕРЕКЛЮЧАТЕЛЬ (ТОЛЬКО ОДНА ФУНКЦИЯ)
function toggleRev() {
    const btn = el('btn-rev-toggle');
    const isNowUnten = btn.classList.toggle('is-active');
    btn.innerText = isNowUnten ? 'UNTEN' : 'OBEN';
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
        btn.classList.add('is-active');
        btn.innerText = 'UNTEN';
    } else {
        btn.classList.remove('is-active');
        btn.innerText = 'OBEN';
    }
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const t = {
        id: el('t-id').value.toUpperCase(),
        nm: el('t-nm').value.toUpperCase(),
        dia: el('t-dia').value,
        rev: el('btn-rev-toggle').classList.contains('is-active')
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

function exportJSON() { el('imp-area').value = JSON.stringify(db); el('imp-area').select(); document.execCommand('copy'); alert("JSON Copied!"); }
function importJSON() { try { db = JSON.parse(el('imp-area').value); save(); renderList(); hide('m-imp'); } catch(e){ alert("Error"); } }
function makePDF() { window.print(); }

// Инициализация
window.onload = () => { injectStyles(); renderList(); };
