/* CITITOOL PREMUM V11 - FINAL STABLE 
   - Авто-очистка старого HTML (QS Central больше не мешает)
   - Исправлен PDF (умные разрывы, всё в одном окне)
   - Drag & Drop логика и правильные отступы снизу
*/

const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

// --- 1. СТИЛИ ПРИЛОЖЕНИЯ ---
const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { 
            --bg: #f2f2f7; --accent: #007aff; --card: #ffffff; 
            --text: #1c1c1e; --sub: #8e8e93; 
        }
        body { 
            background: var(--bg); font-family: -apple-system, sans-serif; 
            margin: 0; color: var(--text); padding-bottom: 120px;
            overscroll-behavior: none;
        }
        
        /* Glass Header */
        header {
            background: rgba(255,255,255,0.8); backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            padding: 15px 20px; position: sticky; top: 0; z-index: 1000;
            display: flex; justify-content: space-between; align-items: center;
            border-bottom: 0.5px solid rgba(0,0,0,0.1);
        }
        .h-title { font-weight: 800; font-size: 22px; letter-spacing: -0.5px; }

        /* Projects */
        .project-card {
            background: var(--card); border-radius: 20px;
            margin: 12px 16px; padding: 20px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.03);
            display: flex; align-items: center; justify-content: space-between;
        }
        .p-num-big { font-size: 28px; font-weight: 900; line-height: 1; }
        .p-name-sub { font-size: 11px; font-weight: 700; color: var(--sub); text-transform: uppercase; }

        /* Tools */
        .tool-card {
            background: var(--card); border-radius: 16px;
            margin: 8px 16px; padding: 16px;
            display: flex; align-items: center; justify-content: space-between;
            box-shadow: 0 2px 8px rgba(0,0,0,0.02);
        }
        .t-nr-label { font-size: 10px; font-weight: 700; color: var(--sub); }
        .t-name-main { font-size: 18px; font-weight: 800; margin: 2px 0; }
        .t-dia-val { color: var(--accent); font-weight: 700; font-size: 14px; }
        .rev-badge { 
            background: #000; color: #fff; font-size: 8px; font-weight: 900; 
            padding: 2px 6px; border-radius: 4px; margin-bottom: 5px; width: fit-content;
        }

        /* Controls */
        .btn-icon { background: #f2f2f7; border: none; border-radius: 10px; width: 36px; height: 36px; color: var(--accent); font-weight: 900; }
        .fab {
            position: fixed; bottom: 30px; right: 20px;
            background: var(--accent); color: white; width: 60px; height: 60px;
            border-radius: 30px; display: flex; align-items: center; justify-content: center;
            font-size: 30px; box-shadow: 0 10px 25px rgba(0,122,255,0.3); z-index: 2000;
        }

        /* Modals */
        .modal { 
            display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.4); 
            z-index: 3000; align-items: flex-end; 
        }
        .modal-content { 
            background: white; width: 100%; border-radius: 25px 25px 0 0; 
            padding: 25px; box-sizing: border-box; max-height: 90vh; overflow-y: auto;
        }
        input { 
            width: 100%; padding: 14px; margin: 8px 0; border: 1px solid #efeff4; 
            border-radius: 12px; background: #f9f9fb; font-size: 16px; box-sizing: border-box;
        }
        .btn-save { background: var(--accent); color: white; width: 100%; padding: 16px; border: none; border-radius: 14px; font-weight: 700; margin-top: 10px; }
    `;
    document.head.appendChild(style);
};

// --- 2. ГЕНЕРАЦИЯ ЧИСТОЙ СТРУКТУРЫ ---
function setupApp() {
    document.body.innerHTML = `
        <header id="main-header">
            <div class="h-title" id="h-text">CitiTool</div>
            <button class="btn-icon" onclick="openImport()" style="font-size:10px">JSON</button>
        </header>

        <div id="v-home">
            <div id="list-p"></div>
            <div class="fab" onclick="modalP()">+</div>
        </div>

        <div id="v-det" style="display:none">
            <div style="padding: 20px 20px 10px 20px">
                <div id="det-name" class="p-name-sub"></div>
                <div id="det-num" style="font-size:40px; font-weight:900; letter-spacing:-1px"></div>
                <div style="display:flex; gap:10px; margin-top:15px">
                    <button class="btn-save" style="margin:0; flex:1" onclick="goHome()">Назад</button>
                    <button class="btn-save" style="margin:0; flex:1; background:#000" onclick="makePDF()">PDF Report</button>
                </div>
            </div>
            <div id="list-t"></div>
            <div class="fab" onclick="modalT()">+</div>
        </div>

        <div id="m-p" class="modal" onclick="if(event.target==this)this.style.display='none'">
            <div class="modal-content">
                <input type="hidden" id="p-idx">
                <input id="p-num" placeholder="PROJEKT NUMMER">
                <input id="p-nam" placeholder="NAME / KUNDE">
                <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px">
                    <input id="p-lzf" placeholder="Laufzeit">
                    <input id="p-mat" placeholder="Material">
                </div>
                <input id="p-sag" placeholder="Sägelänge">
                <input id="p-abs" placeholder="Abstand">
                <input id="p-grf" placeholder="Greifbacken">
                <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px">
                    <input id="p-stt" placeholder="Stück Soll">
                    <input id="p-stn" placeholder="Stück Ist">
                </div>
                <button class="btn-save" onclick="saveP()">SPEICHERN</button>
            </div>
        </div>

        <div id="m-t" class="modal" onclick="if(event.target==this)this.style.display='none'">
            <div class="modal-content">
                <input type="hidden" id="t-idx">
                <input id="t-id" placeholder="T-NR (z.B. T01)">
                <input id="t-nm" placeholder="WERKZEUGNAME">
                <input id="t-dia" placeholder="Ø / TOLERANZ">
                <button id="t-rev-btn" class="btn-save" style="background:#f2f2f7; color:#000" onclick="this.classList.toggle('active'); this.innerText=this.classList.contains('active')?'REVOLVER UNTEN':'REVOLVER OBEN'">REVOLVER OBEN</button>
                <button class="btn-save" onclick="saveT()">SPEICHERN</button>
                <button id="btn-del-t" class="btn-save" style="background:none; color:red" onclick="delT()">Löschen</button>
            </div>
        </div>

        <div id="m-imp" class="modal" onclick="if(event.target==this)this.style.display='none'">
            <div class="modal-content">
                <textarea id="imp-area" style="width:100%; height:200px; border-radius:12px; border:1px solid #ddd; padding:10px"></textarea>
                <button class="btn-save" onclick="importJSON()">IMPORT</button>
            </div>
        </div>
    `;
}

// --- 3. ФУНКЦИОНАЛ ---
const save = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

function renderList() {
    const list = document.getElementById('list-p');
    list.innerHTML = db.map((p, i) => `
        <div class="project-card" onclick="openProject(${i})">
            <div>
                <div class="p-name-sub">${p.name || 'UNBENANNT'}</div>
                <div class="p-num-big">${p.num || '---'}</div>
            </div>
            <div style="color:red; font-weight:800; padding:10px" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>
    `).join('') + '<div style="height:100px"></div>';
}

function openProject(i) {
    currentIdx = i;
    document.getElementById('v-home').style.display = 'none';
    document.getElementById('v-det').style.display = 'block';
    document.getElementById('det-name').innerText = db[i].name;
    document.getElementById('det-num').innerText = db[i].num;
    renderTools();
}

function goHome() {
    currentIdx = null;
    document.getElementById('v-home').style.display = 'block';
    document.getElementById('v-det').style.display = 'none';
    renderList();
}

function renderTools() {
    const list = document.getElementById('list-t');
    const tools = db[currentIdx].tools || [];
    list.innerHTML = tools.map((t, i) => `
        <div class="tool-card" onclick="modalT(${i})">
            <div style="flex:1">
                ${t.rev ? '<div class="rev-badge">REVOLVER UNTEN</div>' : ''}
                <div class="t-nr-label">${t.id}</div>
                <div class="t-name-main">${t.nm}</div>
                <div class="t-dia-val">${t.dia}</div>
            </div>
            <div style="display:flex; flex-direction:column; gap:5px">
                <button class="btn-icon" onclick="event.stopPropagation(); moveItem(${i},-1)">↑</button>
                <button class="btn-icon" onclick="event.stopPropagation(); moveItem(${i},1)">↓</button>
            </div>
        </div>
    `).join('') + '<div style="height:150px"></div>';
}

function modalP() {
    document.getElementById('p-idx').value = '';
    document.getElementById('p-num').value = '';
    document.getElementById('p-nam').value = '';
    document.getElementById('m-p').style.display = 'flex';
}

function saveP() {
    const data = {
        num: document.getElementById('p-num').value,
        name: document.getElementById('p-nam').value.toUpperCase(),
        lzf: document.getElementById('p-lzf').value,
        mat: document.getElementById('p-mat').value,
        sag: document.getElementById('p-sag').value,
        abs: document.getElementById('p-abs').value,
        grf: document.getElementById('p-grf').value,
        stt: document.getElementById('p-stt').value,
        stn: document.getElementById('p-stn').value,
        tools: []
    };
    db.push(data); save(); goHome(); document.getElementById('m-p').style.display = 'none';
}

function modalT(i = null) {
    const edit = i !== null;
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', rev:false};
    document.getElementById('t-idx').value = edit ? i : '';
    document.getElementById('t-id').value = t.id;
    document.getElementById('t-nm').value = t.nm;
    document.getElementById('t-dia').value = t.dia;
    const btn = document.getElementById('t-rev-btn');
    if(t.rev) { btn.classList.add('active'); btn.innerText = 'REVOLVER UNTEN'; }
    else { btn.classList.remove('active'); btn.innerText = 'REVOLVER OBEN'; }
    document.getElementById('btn-del-t').style.display = edit ? 'block' : 'none';
    document.getElementById('m-t').style.display = 'flex';
}

function saveT() {
    const i = document.getElementById('t-idx').value;
    const isUnten = document.getElementById('t-rev-btn').classList.contains('active');
    const t = { 
        id: document.getElementById('t-id').value.toUpperCase(), 
        nm: document.getElementById('t-nm').value.toUpperCase(), 
        dia: document.getElementById('t-dia').value, 
        rev: isUnten 
    };
    if(i === '') db[currentIdx].tools.push(t);
    else db[currentIdx].tools[i] = t;
    save(); renderTools(); document.getElementById('m-t').style.display = 'none';
}

function moveItem(i, dir) {
    const tools = db[currentIdx].tools;
    if(i+dir >= 0 && i+dir < tools.length) {
        [tools[i], tools[i+dir]] = [tools[i+dir], tools[i]];
        save(); renderTools();
    }
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i,1); save(); renderList(); } }
function delT() { const i = document.getElementById('t-idx').value; db[currentIdx].tools.splice(i,1); save(); renderTools(); document.getElementById('m-t').style.display='none'; }
function openImport() { document.getElementById('imp-area').value = JSON.stringify(db); document.getElementById('m-imp').style.display='flex'; }
function importJSON() { try { db = JSON.parse(document.getElementById('imp-area').value); save(); location.reload(); } catch(e){ alert('Error'); } }

// --- 4. PDF ENGINE (STABLE) ---
function makePDF() {
    const p = db[currentIdx];
    const oben = (p.tools || []).filter(t => !t.rev);
    const unten = (p.tools || []).filter(t => t.rev);

    const getTable = (title, list) => {
        if(list.length === 0) return '';
        return `
            <div style="background:#000; color:#fff; padding:6px 10px; font-weight:900; margin-top:20px">${title}</div>
            <table style="width:100%; border-collapse:collapse">
                <thead>
                    <tr style="font-size:10px; border-bottom:2px solid #000">
                        <th style="text-align:left; padding:8px 0">T-NR</th>
                        <th style="text-align:left">WERKZEUGNAME</th>
                        <th style="text-align:right">Ø / TOLERANZ</th>
                    </tr>
                </thead>
                <tbody>
                    ${list.map(t => `
                        <tr style="border-bottom:1px solid #000; page-break-inside:avoid">
                            <td style="padding:12px 0; font-weight:800; font-size:15px; width:70px">${t.id}</td>
                            <td style="font-weight:700; text-transform:uppercase">${t.nm}</td>
                            <td style="text-align:right; font-weight:800">${t.dia}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        `;
    };

    const pdfHtml = `
    <html>
    <head>
        <style>
            @page { size: A4; margin: 10mm; }
            body { font-family: sans-serif; margin: 0; }
            .h-grid { display: flex; justify-content: space-between; border-bottom: 5px solid #000; padding-bottom: 10px; }
            .meta { width: 220px; font-size: 11px; font-weight: 800; }
            .meta div { display: flex; justify-content: space-between; border-bottom: 1px solid #eee; padding: 2px 0; }
        </style>
    </head>
    <body>
        <div class="h-grid">
            <div>
                <div style="font-size:12px; font-weight:800; color:#666">${p.name || ''}</div>
                <div style="font-size:54px; font-weight:900; line-height:0.9">${p.num || '---'}</div>
            </div>
            <div class="meta">
                <div><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div>
                <div><span>MATERIAL</span><span>${p.mat || ''}</span></div>
                <div><span>SÄGELÄНGE</span><span>${p.sag || ''}</span></div>
                <div><span>ABSTAND</span><span>${p.abs || ''}</span></div>
                <div><span>GREIFBACKEN</span><span>${p.grf || ''}</span></div>
                <div><span>STÜCKZAHL</span><span>${p.stt || ''} / ${p.stn || ''}</span></div>
            </div>
        </div>
        ${getTable('REVOLVER OBEN', oben)}
        ${getTable('REVOLVER UNTEN', unten)}
        <script>window.onload=()=>{setTimeout(()=>{window.print();window.close()},500)}<\/script>
    </body>
    </html>`;

    const win = window.open('', '_blank');
    win.document.write(pdfHtml);
    win.document.close();
}

// INIT
window.onload = () => {
    injectStyles();
    setupApp();
    renderList();
};
