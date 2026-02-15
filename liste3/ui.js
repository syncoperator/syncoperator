/* CITITOOL FINAL V9 - TERMINATOR EDITION
   Полная пересборка интерфейса. Удаляет старый HTML.
   Чинит PDF, Импорт и Дизайн.
*/

const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

// 1. СТИЛИ (CSS)
const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f5f5f7; 
            --card: #ffffff; 
            --text: #1d1d1f; 
            --blue: #007aff; 
            --red: #ff3b30; 
            --gray: #86868b;
        }
        
        body { 
            background-color: var(--bg); 
            font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, sans-serif; 
            margin: 0; 
            padding-top: 60px; /* Место под хедер */
            padding-bottom: 100px; 
            color: var(--text);
            -webkit-font-smoothing: antialiased;
        }

        /* НЕВИДИМАЯ ПАНЕЛЬ (Header) */
        #top-bar {
            position: fixed; top: 0; left: 0; right: 0;
            height: 60px;
            background: rgba(245, 245, 247, 0.8);
            backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            z-index: 9999;
            display: flex; align-items: center; justify-content: space-between;
            padding: 0 20px;
            border-bottom: 1px solid rgba(0,0,0,0.05);
        }

        .btn-head {
            background: #fff;
            color: var(--blue);
            border: 1px solid rgba(0,0,0,0.1);
            font-size: 11px; font-weight: 700;
            padding: 8px 16px;
            border-radius: 20px;
            cursor: pointer;
            box-shadow: 0 2px 6px rgba(0,0,0,0.03);
            transition: transform 0.1s;
        }
        .btn-head:active { transform: scale(0.96); background: #f0f0f0; }

        /* ЛОГОТИП И ЗАГОЛОВОК */
        .brand-container {
            display: flex; flex-direction: column; align-items: center;
            margin-top: 20px; margin-bottom: 30px;
        }
        .logo-img { width: 180px; height: auto; display: block; }
        .app-title {
            font-size: 42px; font-weight: 900; letter-spacing: -2px;
            margin-top: -30px; z-index: 10; color: #000;
        }

        /* СПИСОК ПРОЕКТОВ */
        .project-card {
            background: var(--card);
            margin: 0 20px 15px 20px;
            padding: 20px;
            border-radius: 22px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.03);
            display: flex; justify-content: space-between; align-items: center;
            transition: transform 0.2s;
            cursor: pointer;
        }
        .project-card:active { transform: scale(0.98); }
        .p-info h3 { margin: 0; font-size: 28px; font-weight: 800; letter-spacing: -1px; }
        .p-info span { font-size: 11px; font-weight: 700; color: var(--gray); text-transform: uppercase; }
        
        .btn-del {
            color: var(--red);
            font-weight: bold;
            padding: 10px;
            font-size: 18px;
        }

        /* ИНСТРУМЕНТЫ */
        .tool-card {
            background: var(--card);
            margin: 0 15px 10px 15px;
            padding: 15px 20px;
            border-radius: 18px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.02);
            display: flex; justify-content: space-between; align-items: center;
        }
        .tool-info { flex: 1; }
        .t-badge { 
            background: #000; color: #fff; 
            font-size: 9px; font-weight: 900; 
            padding: 3px 8px; border-radius: 6px; 
            display: inline-block; margin-bottom: 4px;
        }
        .t-id { font-size: 11px; font-weight: 800; color: var(--gray); }
        .t-name { font-size: 17px; font-weight: 800; color: #000; margin: 2px 0; }
        .t-dia { font-size: 14px; font-weight: 600; color: var(--blue); }

        .order-btns { display: flex; flex-direction: column; gap: 4px; margin-left: 10px; }
        .btn-arr {
            width: 30px; height: 30px;
            background: #f2f2f7; border: none; border-radius: 8px;
            color: var(--blue); font-weight: 900;
        }

        /* МОДАЛЬНЫЕ ОКНА */
        .modal-overlay {
            display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(0,0,0,0.4); z-index: 10000;
            align-items: flex-end; justify-content: center;
        }
        .modal-box {
            background: #fff; width: 100%; max-width: 600px;
            border-radius: 24px 24px 0 0; padding: 25px;
            box-sizing: border-box; max-height: 90vh; overflow-y: auto;
            animation: slideUp 0.3s ease-out;
        }
        @keyframes slideUp { from { transform: translateY(100%); } to { transform: translateY(0); } }

        .inp-group { margin-bottom: 15px; }
        .inp-label { font-size: 12px; font-weight: 700; color: var(--gray); margin-bottom: 6px; display: block; text-transform: uppercase; }
        input, textarea {
            width: 100%; padding: 14px; border: 1px solid #e5e5ea;
            background: #f9f9f9; border-radius: 12px;
            font-size: 16px; box-sizing: border-box; font-family: inherit;
        }
        input:focus { outline: none; border-color: var(--blue); background: #fff; }

        .btn-primary {
            background: var(--blue); color: #fff; width: 100%; padding: 16px;
            border: none; border-radius: 14px; font-size: 16px; font-weight: 700;
            margin-top: 10px;
        }
        .btn-rev {
            width: 100%; padding: 14px; border-radius: 12px;
            border: none; font-weight: 800; font-size: 13px;
            margin-bottom: 15px; cursor: pointer;
            transition: 0.2s;
        }
        .rev-off { background: #f2f2f7; color: var(--gray); }
        .rev-on { background: #1c1c1e; color: #fff; }

    `;
    document.head.appendChild(style);
};

// 2. СТРУКТУРА HTML (ГЕНЕРАЦИЯ)
function buildStructure() {
    // УБИВАЕМ СТАРЫЙ HTML (QS Central и прочее)
    document.body.innerHTML = ''; 

    // Строим новый
    document.body.innerHTML = `
        <div id="top-bar">
            <button class="btn-head" onclick="openImport()">JSON</button>
            <button class="btn-head" onclick="openProjectModal()">+ NEU</button>
        </div>

        <div id="view-home">
            <div class="brand-container">
                <img src="https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png" class="logo-img">
                <div class="app-title">CitiTool</div>
            </div>
            <div id="project-list"></div>
        </div>

        <div id="view-detail" style="display:none;">
            <div class="brand-container" style="transform: scale(0.8); margin-bottom: 0;">
                <h1 id="det-num" style="margin:0; font-size:40px;"></h1>
                <div id="det-name" style="font-weight:700; color:gray;"></div>
            </div>
            <div style="padding: 0 20px; margin-bottom: 10px; display:flex; gap:10px;">
                <button class="btn-head" style="flex:1" onclick="goHome()">← ZURÜCK</button>
                <button class="btn-head" style="flex:1" onclick="makePDF()">PDF REPORT</button>
                <button class="btn-head" style="flex:1" onclick="openProjectModal(true)">EDIT</button>
            </div>
            <div id="tool-list"></div>
            <button onclick="openToolModal()" style="position:fixed; bottom:30px; right:20px; width:60px; height:60px; border-radius:30px; background:var(--blue); color:#fff; border:none; font-size:30px; box-shadow:0 5px 20px rgba(0,122,255,0.4); z-index:900;">+</button>
        </div>

        <div id="modal-project" class="modal-overlay" onclick="if(event.target===this) closeModals()">
            <div class="modal-box">
                <h2 style="margin-top:0">Projekt</h2>
                <input type="hidden" id="p-idx">
                <div class="inp-group"><label class="inp-label">Projekt Nummer</label><input id="p-num" type="text" placeholder="z.B. 12345"></div>
                <div class="inp-group"><label class="inp-label">Name</label><input id="p-name" type="text"></div>
                
                <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px;">
                    <div class="inp-group"><label class="inp-label">Laufzeit</label><input id="p-lzf" type="text"></div>
                    <div class="inp-group"><label class="inp-label">Sägelänge</label><input id="p-sag" type="text"></div>
                    <div class="inp-group"><label class="inp-label">Stück Soll</label><input id="p-stt" type="text"></div>
                    <div class="inp-group"><label class="inp-label">Stück Ist</label><input id="p-stn" type="text"></div>
                </div>
                <div class="inp-group"><label class="inp-label">Abstand</label><input id="p-abs" type="text"></div>
                <div class="inp-group"><label class="inp-label">Greifbacken</label><input id="p-grf" type="text"></div>
                <div class="inp-group"><label class="inp-label">Material</label><input id="p-mat" type="text"></div>

                <button class="btn-primary" onclick="saveProject()">SPEICHERN</button>
                <button class="btn-primary" style="background:#f2f2f7; color:#000; margin-top:10px;" onclick="closeModals()">ABBRECHEN</button>
            </div>
        </div>

        <div id="modal-tool" class="modal-overlay" onclick="if(event.target===this) closeModals()">
            <div class="modal-box">
                <h2 style="margin-top:0">Werkzeug</h2>
                <input type="hidden" id="t-idx">
                
                <div class="inp-group"><label class="inp-label">T-NR</label><input id="t-id" type="text" placeholder="T0000"></div>
                <div class="inp-group"><label class="inp-label">Beschreibung</label><input id="t-nm" type="text"></div>
                <div class="inp-group"><label class="inp-label">Ø / Toleranz</label><input id="t-dia" type="text"></div>
                
                <label class="inp-label">Position</label>
                <button id="btn-rev-toggle" class="btn-rev rev-off" onclick="toggleRev()">REVOLVER OBEN</button>

                <button class="btn-primary" onclick="saveTool()">SPEICHERN</button>
                <button id="btn-del-tool" class="btn-primary" style="background:#ff3b30; margin-top:10px; display:none;" onclick="deleteTool()">LÖSCHEN</button>
            </div>
        </div>

        <div id="modal-import" class="modal-overlay" onclick="if(event.target===this) closeModals()">
            <div class="modal-box">
                <h2>Import / Export</h2>
                <p style="font-size:12px; color:gray;">Kopiere den Text zum Sichern oder füge ihn ein zum Wiederherstellen.</p>
                <textarea id="json-area" rows="8"></textarea>
                <div style="display:flex; gap:10px; margin-top:10px;">
                    <button class="btn-primary" onclick="copyJSON()">KOPIEREN (EXPORT)</button>
                    <button class="btn-primary" style="background:#34c759;" onclick="pasteJSON()">EINFÜGEN (IMPORT)</button>
                </div>
            </div>
        </div>
    `;
}

// 3. ЛОГИКА (JS)
const el = (id) => document.getElementById(id);
const saveDB = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

function renderProjects() {
    const list = el('project-list');
    list.innerHTML = db.map((p, i) => `
        <div class="project-card" onclick="openDetail(${i})">
            <div class="p-info">
                <span>${p.name || 'PROJEKT'}</span>
                <h3>${p.num || '---'}</h3>
            </div>
            <div class="btn-del" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>
    `).join('') + '<div style="height:100px;"></div>';
}

function openDetail(i) {
    currentIdx = i;
    el('view-home').style.display = 'none';
    el('view-detail').style.display = 'block';
    el('det-num').innerText = db[i].num;
    el('det-name').innerText = db[i].name;
    renderTools();
    // Скрываем +NEU в хедере когда внутри проекта, чтобы не мешал
    el('top-bar').style.display = 'none'; 
}

function goHome() {
    currentIdx = null;
    el('view-home').style.display = 'block';
    el('view-detail').style.display = 'none';
    el('top-bar').style.display = 'flex'; // Возвращаем хедер
    renderProjects();
}

function renderTools() {
    if (currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    el('tool-list').innerHTML = tools.map((t, i) => `
        <div class="tool-card" onclick="openToolModal(${i})">
            <div class="tool-info">
                ${t.rev ? '<span class="t-badge">REVOLVER UNTEN</span>' : ''}
                <div class="t-id">${t.id || 'T0000'}</div>
                <div class="t-name">${t.nm || '---'}</div>
                <div class="t-dia">${t.dia || ''}</div>
            </div>
            <div class="order-btns">
                <button class="btn-arr" onclick="event.stopPropagation(); moveTool(${i}, -1)">↑</button>
                <button class="btn-arr" onclick="event.stopPropagation(); moveTool(${i}, 1)">↓</button>
            </div>
        </div>
    `).join('') + '<div style="height:120px;"></div>';
}

// УПРАВЛЕНИЕ ДАННЫМИ
function openProjectModal(edit = false) {
    const p = (edit && currentIdx !== null) ? db[currentIdx] : {num:'',name:'',lzf:'',sag:'',stt:'',stn:'',abs:'',grf:'',mat:''};
    el('p-idx').value = edit ? currentIdx : '';
    el('p-num').value = p.num; el('p-name').value = p.name;
    el('p-lzf').value = p.lzf; el('p-sag').value = p.sag;
    el('p-stt').value = p.stt; el('p-stn').value = p.stn;
    el('p-abs').value = p.abs; el('p-grf').value = p.grf;
    el('p-mat').value = p.mat;
    el('modal-project').style.display = 'flex';
}

function saveProject() {
    const idx = el('p-idx').value;
    const newData = {
        num: el('p-num').value, name: el('p-name').value.toUpperCase(),
        lzf: el('p-lzf').value, sag: el('p-sag').value,
        stt: el('p-stt').value, stn: el('p-stn').value,
        abs: el('p-abs').value, grf: el('p-grf').value, mat: el('p-mat').value.toUpperCase(),
        tools: (idx !== '' && db[idx]) ? db[idx].tools : []
    };
    if (idx === '') db.push(newData); else db[idx] = newData;
    saveDB(); renderProjects(); 
    if(currentIdx !== null) { openDetail(currentIdx); } // Обновить шапку если редактировали внутри
    closeModals();
}

function deleteProject(i) {
    if(confirm('Wirklich löschen?')) {
        db.splice(i, 1); saveDB(); renderProjects();
    }
}

// УПРАВЛЕНИЕ ИНСТРУМЕНТАМИ
function openToolModal(i = null) {
    const edit = i !== null;
    const t = edit ? db[currentIdx].tools[i] : {id:'',nm:'',dia:'',rev:false};
    el('t-idx').value = edit ? i : '';
    el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    
    // Настройка кнопки револьвера
    const btn = el('btn-rev-toggle');
    btn.className = t.rev ? 'btn-rev rev-on' : 'btn-rev rev-off';
    btn.innerText = t.rev ? 'REVOLVER UNTEN' : 'REVOLVER OBEN';
    
    el('btn-del-tool').style.display = edit ? 'block' : 'none';
    el('modal-tool').style.display = 'flex';
}

function toggleRev() {
    const btn = el('btn-rev-toggle');
    const isUnten = btn.classList.contains('rev-on');
    if (isUnten) {
        btn.className = 'btn-rev rev-off'; btn.innerText = 'REVOLVER OBEN';
    } else {
        btn.className = 'btn-rev rev-on'; btn.innerText = 'REVOLVER UNTEN';
    }
}

function saveTool() {
    const i = el('t-idx').value;
    const isUnten = el('btn-rev-toggle').classList.contains('rev-on');
    const t = { 
        id: el('t-id').value.toUpperCase(), 
        nm: el('t-nm').value.toUpperCase(), 
        dia: el('t-dia').value, 
        rev: isUnten 
    };
    if (!db[currentIdx].tools) db[currentIdx].tools = [];
    if (i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    saveDB(); renderTools(); closeModals();
}

function deleteTool() {
    const i = el('t-idx').value;
    db[currentIdx].tools.splice(i, 1);
    saveDB(); renderTools(); closeModals();
}

function moveTool(i, dir) {
    const list = db[currentIdx].tools;
    if (i + dir >= 0 && i + dir < list.length) {
        [list[i], list[i+dir]] = [list[i+dir], list[i]];
        saveDB(); renderTools();
    }
}

// IMPORT / EXPORT
function openImport() {
    el('json-area').value = JSON.stringify(db, null, 2);
    el('modal-import').style.display = 'flex';
}
function copyJSON() {
    el('json-area').select(); document.execCommand('copy');
    alert('Kopiert!');
}
function pasteJSON() {
    try {
        const d = JSON.parse(el('json-area').value);
        if(Array.isArray(d)) { db = d; saveDB(); location.reload(); }
    } catch(e) { alert('Ungültiges JSON'); }
}

function closeModals() {
    document.querySelectorAll('.modal-overlay').forEach(m => m.style.display = 'none');
}

// 4. PDF ГЕНЕРАТОР (ИСПРАВЛЕННЫЙ)
function makePDF() {
    const p = db[currentIdx];
    
    // Стили для печати: break-inside: avoid ЗАПРЕЩАЕТ разрывать строку таблицы
    const styles = `
        <style>
            @page { size: A4; margin: 10mm; }
            body { font-family: sans-serif; -webkit-print-color-adjust: exact; }
            .header { border-bottom: 5px solid #000; margin-bottom: 20px; padding-bottom: 10px; }
            .meta-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 5px; font-size: 11px; font-weight: 800; margin-left: auto; width: 250px; }
            .meta-row { display: flex; justify-content: space-between; border-bottom: 1px solid #eee; padding: 2px 0; }
            
            .section-title { font-size: 18px; font-weight: 900; margin-top: 20px; border-bottom: 2px solid #000; padding-bottom: 5px; }
            
            /* ТАБЛИЦА */
            .t-row { 
                display: flex; align-items: baseline; 
                border-bottom: 1px solid #000; 
                padding: 8px 0; 
                page-break-inside: avoid; /* ГЛАВНОЕ: НЕ РАЗРЫВАТЬ СТРОКУ */
            }
            .c-id { width: 80px; font-weight: 800; font-size: 14px; }
            .c-nm { flex: 1; font-weight: 700; font-size: 14px; }
            .c-dia { width: 120px; text-align: right; font-weight: 800; font-size: 13px; }
        </style>
    `;

    const meta = `
        <div class="header" style="display:flex; justify-content:space-between;">
            <div>
                <div style="font-size:12px; font-weight:900; color:#666;">${p.name || ''}</div>
                <div style="font-size:50px; font-weight:900; line-height:1;">${p.num || '---'}</div>
            </div>
            <div class="meta-grid">
                <div class="meta-row"><span>LZF</span><span>${p.lzf}</span></div>
                <div class="meta-row"><span>MAT</span><span>${p.mat}</span></div>
                <div class="meta-row"><span>SÄGE</span><span>${p.sag}</span></div>
                <div class="meta-row"><span>ABS</span><span>${p.abs}</span></div>
                <div class="meta-row"><span>BACKEN</span><span>${p.grf}</span></div>
                <div class="meta-row"><span>STK</span><span>${p.stt} / ${p.stn}</span></div>
            </div>
        </div>
    `;

    const getTable = (list, title) => {
        if (!list || list.length === 0) return '';
        return `
            <div class="section-title">${title}</div>
            ${list.map(t => `
                <div class="t-row">
                    <div class="c-id">${t.id}</div>
                    <div class="c-nm">${t.nm}</div>
                    <div class="c-dia">${t.dia}</div>
                </div>
            `).join('')}
        `;
    };

    let oben = [], unten = [];
    (p.tools || []).forEach(t => t.rev ? unten.push(t) : oben.push(t));

    const html = `
        <!DOCTYPE html>
        <html>
        <head>${styles}</head>
        <body>
            ${meta}
            ${getTable(oben, 'REVOLVER OBEN')}
            ${getTable(unten, 'REVOLVER UNTEN')}
            <script>window.onload = () => setTimeout(() => window.print(), 500);<\/script>
        </body>
        </html>
    `;

    const win = window.open('', '_blank');
    win.document.write(html);
    win.document.close();
}

// ЗАПУСК
window.onload = () => {
    injectStyles();
    buildStructure();
    renderProjects();
};
