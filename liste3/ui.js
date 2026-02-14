const DB_KEY = 'QS_DATA_V8';
// Вставь сюда прямую ссылку на свой логотип с GitHub (Raw link)
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
            --text: #1a1a1a;
        }
        body { 
            background: var(--bg) !important; 
            font-family: -apple-system, system-ui, sans-serif !important; 
            margin: 0; padding-bottom: 120px; color: var(--text);
        }

        /* Шапка с лого и зеркалом */
        .brand-header {
            padding: 50px 0 20px 0;
            display: flex; flex-direction: column; align-items: center;
        }
        .main-logo {
            width: 300px; height: 300px; object-fit: contain;
            filter: drop-shadow(4px 6px 10px rgba(0,0,0,0.1));
            margin-bottom: 10px;
        }
        .mirror-box { position: relative; text-align: center; }
        .text-top {
            font-size: 58px; font-weight: 900; letter-spacing: -3px;
            color: var(--text); position: relative; z-index: 2; line-height: 1;
        }
        .text-bottom {
            font-size: 58px; font-weight: 900; letter-spacing: -3px;
            margin-top: -24px; transform: scaleY(-1);
            background: linear-gradient(to bottom, rgba(0,0,0,0.15) 0%, rgba(245,247,250,1) 85%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            opacity: 0.3; user-select: none;
        }

        /* Кнопка + NEU справа */
        .nav-bar { display: flex; justify-content: flex-end; padding: 0 25px; margin-bottom: 20px; }
        .btn-neu {
            background: var(--bg); border: none; border-radius: 14px;
            padding: 12px 24px; font-size: 14px; font-weight: 800; color: var(--accent);
            box-shadow: 6px 6px 12px var(--neu-shadow), -6px -6px 12px var(--neu-light);
            cursor: pointer; transition: 0.2s;
        }
        .btn-neu:active { box-shadow: inset 3px 3px 6px var(--neu-shadow), inset -3px -3px 6px var(--neu-light); }

        /* Премиальные карточки */
        .neu-card {
            background: var(--bg); border-radius: 30px; padding: 25px;
            margin: 15px 20px; box-shadow: 10px 10px 20px var(--neu-shadow), -10px -10px 20px var(--neu-light);
            display: flex; align-items: center; justify-content: space-between;
            border: 1px solid rgba(255,255,255,0.6);
        }
        .p-label { font-size: 10px; font-weight: 800; color: #99a1ad; text-transform: uppercase; letter-spacing: 1px; }
        .p-title { font-size: 26px; font-weight: 900; color: #000; }

        .btn-del { color: #ff3b30; font-size: 20px; font-weight: 900; padding: 10px; cursor: pointer; }

        /* Контейнер для инструментов со скроллом и отступом */
        #list-p, #list-t { padding-bottom: 100px; }
    `;
    document.head.appendChild(style);
};

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'block'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

function renderList() {
    const main = el('v-home'); if(!main) return;
    main.innerHTML = `
        <div class="brand-header">
            <img src="${LOGO_URL}" class="main-logo" onerror="this.style.display='none'">
            <div class="mirror-box">
                <div class="text-top">CitiTool</div>
                <div class="text-bottom">CitiTool</div>
            </div>
        </div>
        <div class="nav-bar">
            <button class="btn-neu" onclick="modalP()">+ NEU</button>
        </div>
        <div id="list-p"></div>
    `;
    const list = el('list-p');
    list.innerHTML = db.map((p, i) => `
        <div class="neu-card" onclick="openProject(${i})">
            <div>
                <div class="p-label">PROJEKT</div>
                <div class="p-title">${p.num || '---'}</div>
                <div style="font-size:13px; color:#666; font-weight:700; margin-top:5px;">${p.name || ''}</div>
            </div>
            <div class="btn-del" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('');
}

function openProject(i) { 
    currentIdx = i; 
    el('v-home').style.display='none'; 
    el('v-det').style.display='block'; 
    el('h-num').innerText = db[i].num; 
    el('h-nam').innerText = db[i].name; 
    renderTools(); 
}

function goHome() { 
    currentIdx = null; 
    el('v-home').style.display='block'; 
    el('v-det').style.display='none'; 
    renderList(); 
}

// Функции для работы с данными (сохранение проектов/инструментов) остаются без изменений
function saveP() {
    const idx = el('p-idx').value;
    const data = {
        num: el('p-num').value, name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value, sag: el('p-sag').value,
        stt: el('p-stt').value, stn: el('p-stn').value,
        abs: el('p-abs').value, grf: el('p-grf').value,
        mat: el('p-mat') ? el('p-mat').value.toUpperCase() : '',
        tools: (idx !== '' && db[idx]) ? (db[idx].tools || []) : []
    };
    if (idx === '') db.push(data); else db[idx] = data;
    save(); hide('m-p'); renderList();
}

function modalP(edit = false) {
    if (!edit) currentIdx = null;
    const p = (edit && db[currentIdx]) ? db[currentIdx] : {num:'', name:'', lzf:'', sag:'', stt:'', stn:'', abs:'', grf:'', mat:''};
    el('p-idx').value = edit ? currentIdx : '';
    el('p-num').value = p.num; el('p-nam').value = p.name;
    el('p-lzf').value = p.lzf; el('p-sag').value = p.sag;
    el('p-stt').value = p.stt; el('p-stn').value = p.stn;
    el('p-abs').value = p.abs; el('p-grf').value = p.grf;
    if(el('p-mat')) el('p-mat').value = p.mat;
    show('m-p');
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); save(); renderList(); } }

window.onload = () => { injectStyles(); renderList(); };
