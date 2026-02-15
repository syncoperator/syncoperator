/* CITITOOL FINAL FIXED PDF V10
   Исправлена верстка PDF, убраны разрывы страниц и мусор.
*/

const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const injectStyles = () => {
    const style = document.createElement('style');
    style.innerHTML = `
        :root { --bg: #f5f5f7; --card: #ffffff; --text: #1d1d1f; --blue: #007aff; --red: #ff3b30; --gray: #86868b; }
        body { background-color: var(--bg); font-family: -apple-system, sans-serif; margin: 0; padding-top: 70px; padding-bottom: 100px; color: var(--text); }
        #top-bar { position: fixed; top: 0; left: 0; right: 0; height: 60px; background: rgba(245,245,247,0.8); backdrop-filter: blur(20px); -webkit-backdrop-filter: blur(20px); z-index: 9999; display: flex; align-items: center; justify-content: space-between; padding: 0 20px; border-bottom: 1px solid rgba(0,0,0,0.05); }
        .btn-head { background: #fff; color: var(--blue); border: 1px solid rgba(0,0,0,0.1); font-size: 11px; font-weight: 700; padding: 8px 16px; border-radius: 20px; cursor: pointer; }
        .brand-container { display: flex; flex-direction: column; align-items: center; margin-bottom: 20px; }
        .logo-img { width: 150px; height: auto; }
        .app-title { font-size: 32px; font-weight: 900; letter-spacing: -1.5px; margin-top: -20px; color: #000; }
        .project-card { background: var(--card); margin: 0 20px 12px 20px; padding: 18px; border-radius: 20px; box-shadow: 0 4px 15px rgba(0,0,0,0.03); display: flex; justify-content: space-between; align-items: center; }
        .p-info h3 { margin: 0; font-size: 24px; font-weight: 800; }
        .p-info span { font-size: 10px; font-weight: 700; color: var(--gray); text-transform: uppercase; }
        .tool-card { background: var(--card); margin: 0 15px 10px 15px; padding: 15px; border-radius: 15px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 8px rgba(0,0,0,0.02); }
        .t-badge { background: #000; color: #fff; font-size: 9px; font-weight: 900; padding: 3px 7px; border-radius: 5px; display: inline-block; margin-bottom: 4px; }
        .t-name { font-size: 16px; font-weight: 800; color: #000; }
        .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.4); z-index: 10000; align-items: flex-end; justify-content: center; }
        .modal-box { background: #fff; width: 100%; max-width: 500px; border-radius: 24px 24px 0 0; padding: 20px; box-sizing: border-box; }
        input { width: 100%; padding: 12px; border: 1px solid #e5e5ea; background: #f9f9f9; border-radius: 10px; font-size: 16px; margin-bottom: 10px; box-sizing: border-box; }
        .btn-primary { background: var(--blue); color: #fff; width: 100%; padding: 14px; border: none; border-radius: 12px; font-size: 16px; font-weight: 700; cursor: pointer; }
    `;
    document.head.appendChild(style);
};

function buildStructure() {
    document.body.innerHTML = `
        <div id="top-bar"><button class="btn-head" onclick="openImport()">JSON</button><button class="btn-head" onclick="openProjectModal()">+ NEU</button></div>
        <div id="view-home"><div class="brand-container"><img src="https://raw.githubusercontent.com/syncoperator/syncoperator/refs/heads/main/IMG_2810.png" class="logo-img"><div class="app-title">CitiTool</div></div><div id="project-list"></div></div>
        <div id="view-detail" style="display:none;"><div class="brand-container"><h1 id="det-num" style="margin:0;"></h1><div id="det-name" style="color:gray;font-weight:700;"></div></div><div style="padding:0 20px;margin-bottom:15px;display:flex;gap:8px;"><button class="btn-head" style="flex:1" onclick="goHome()">← BACK</button><button class="btn-head" style="flex:1" onclick="makePDF()">PDF</button><button class="btn-head" style="flex:1" onclick="openProjectModal(true)">EDIT</button></div><div id="tool-list"></div><button onclick="openToolModal()" style="position:fixed;bottom:30px;right:20px;width:60px;height:60px;border-radius:30px;background:var(--blue);color:#fff;border:none;font-size:30px;z-index:900;">+</button></div>
        <div id="modal-project" class="modal-overlay" onclick="if(event.target===this)closeModals()"><div class="modal-box"><h2>Projekt</h2><input type="hidden" id="p-idx"><input id="p-num" placeholder="Nummer"><input id="p-name" placeholder="Name"><input id="p-lzf" placeholder="Laufzeit"><input id="p-mat" placeholder="Material"><input id="p-sag" placeholder="Sägelänge"><input id="p-abs" placeholder="Abstand"><input id="p-grf" placeholder="Greifbacken"><input id="p-stt" placeholder="Soll"><input id="p-stn" placeholder="Ist"><button class="btn-primary" onclick="saveProject()">OK</button></div></div>
        <div id="modal-tool" class="modal-overlay" onclick="if(event.target===this)closeModals()"><div class="modal-box"><h2>Tool</h2><input type="hidden" id="t-idx"><input id="t-id" placeholder="T-NR"><input id="t-nm" placeholder="Name"><input id="t-dia" placeholder="Ø"><button id="btn-rev" class="btn-primary" style="background:#eee;color:#000;margin-bottom:10px;" onclick="this.classList.toggle('active')">REV UNTEN</button><button class="btn-primary" onclick="saveTool()">OK</button></div></div>
        <div id="modal-import" class="modal-overlay" onclick="if(event.target===this)closeModals()"><div class="modal-box"><h2>JSON</h2><textarea id="json-area" style="width:100%;height:150px;"></textarea><button class="btn-primary" onclick="pasteJSON()">IMPORT</button></div></div>
    `;
}

// ЛОГИКА (БАЗОВАЯ)
const el = (id) => document.getElementById(id);
const saveDB = () => localStorage.setItem(DB_KEY, JSON.stringify(db));

function renderProjects() {
    el('project-list').innerHTML = db.map((p, i) => `<div class="project-card" onclick="openDetail(${i})"><div class="p-info"><span>${p.name}</span><h3>${p.num}</h3></div><div style="color:red;font-weight:bold;padding:10px;" onclick="event.stopPropagation();deleteProject(${i})">✕</div></div>`).join('');
}
function openDetail(i) { currentIdx = i; el('view-home').style.display = 'none'; el('view-detail').style.display = 'block'; el('det-num').innerText = db[i].num; el('det-name').innerText = db[i].name; renderTools(); el('top-bar').style.display = 'none'; }
function goHome() { currentIdx = null; el('view-home').style.display = 'block'; el('view-detail').style.display = 'none'; el('top-bar').style.display = 'flex'; renderProjects(); }
function renderTools() { const tools = db[currentIdx].tools || []; el('tool-list').innerHTML = tools.map((t, i) => `<div class="tool-card" onclick="openToolModal(${i})"><div>${t.rev ? '<span class="t-badge">UNTEN</span>' : ''}<div style="color:gray;font-size:10px;font-weight:700;">${t.id}</div><div class="t-name">${t.nm}</div><div style="color:var(--blue);font-weight:700;">${t.dia}</div></div></div>`).join(''); }

// МОДАЛКИ
function openProjectModal(edit = false) {
    const p = (edit) ? db[currentIdx] : {num:'',name:'',lzf:'',mat:'',sag:'',abs:'',grf:'',stt:'',stn:''};
    el('p-idx').value = edit ? currentIdx : ''; el('p-num').value = p.num; el('p-name').value = p.name;
    el('p-lzf').value = p.lzf; el('p-mat').value = p.mat; el('p-sag').value = p.sag; el('p-abs').value = p.abs; el('p-grf').value = p.grf; el('p-stt').value = p.stt; el('p-stn').value = p.stn;
    el('modal-project').style.display = 'flex';
}
function saveProject() {
    const i = el('p-idx').value;
    const d = { num: el('p-num').value, name: el('p-name').value.toUpperCase(), lzf: el('p-lzf').value, mat: el('p-mat').value.toUpperCase(), sag: el('p-sag').value, abs: el('p-abs').value, grf: el('p-grf').value, stt: el('p-stt').value, stn: el('p-stn').value, tools: (i !== '' ? db[i].tools : []) };
    if(i === '') db.push(d); else db[i] = d;
    saveDB(); closeModals(); renderProjects(); if(currentIdx !== null) openDetail(currentIdx);
}
function openToolModal(i = null) {
    const edit = i !== null; const t = edit ? db[currentIdx].tools[i] : {id:'',nm:'',dia:'',rev:false};
    el('t-idx').value = edit ? i : ''; el('t-id').value = t.id; el('t-nm').value = t.nm; el('t-dia').value = t.dia;
    if(t.rev) el('btn-rev').classList.add('active'); else el('btn-rev').classList.remove('active');
    el('modal-tool').style.display = 'flex';
}
function saveTool() {
    const i = el('t-idx').value; const t = { id: el('t-id').value.toUpperCase(), nm: el('t-nm').value.toUpperCase(), dia: el('t-dia').value, rev: el('btn-rev').classList.contains('active') };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    saveDB(); closeModals(); renderTools();
}
function deleteProject(i) { if(confirm('Delete?')) { db.splice(i,1); saveDB(); renderProjects(); } }
function openImport() { el('json-area').value = JSON.stringify(db); el('modal-import').style.display = 'flex'; }
function pasteJSON() { try { db = JSON.parse(el('json-area').value); saveDB(); location.reload(); } catch(e) { alert('Error'); } }
function closeModals() { document.querySelectorAll('.modal-overlay').forEach(m => m.style.display = 'none'); }

// ===================== PDF GENERATOR (FIXED) =====================
function makePDF() {
    const p = db[currentIdx];
    const oben = (p.tools || []).filter(t => !t.rev);
    const unten = (p.tools || []).filter(t => t.rev);

    const html = `
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            @page { size: A4; margin: 10mm; }
            body { font-family: -apple-system, sans-serif; margin: 0; padding: 0; color: #000; }
            .page { width: 190mm; margin: 0 auto; }
            .header { border-bottom: 4px solid #000; display: flex; justify-content: space-between; padding-bottom: 10px; margin-bottom: 20px; }
            .title-block { display: flex; flex-direction: column; }
            .p-num { font-size: 48px; font-weight: 900; line-height: 1; margin: 0; }
            .p-name { font-size: 14px; font-weight: 800; color: #555; text-transform: uppercase; }
            
            .meta-table { width: 240px; font-size: 11px; font-weight: 800; border-collapse: collapse; }
            .meta-table td { border-bottom: 1px solid #eee; padding: 3px 0; }
            .meta-val { text-align: right; }

            .section-header { background: #000; color: #fff; padding: 6px 10px; font-size: 16px; font-weight: 900; margin-top: 20px; }
            
            .tool-table { width: 100%; border-collapse: collapse; margin-top: 5px; }
            .tool-table th { text-align: left; font-size: 10px; padding: 5px; border-bottom: 2px solid #000; }
            .tool-table td { padding: 10px 5px; border-bottom: 1.5px solid #000; font-size: 14px; font-weight: 700; page-break-inside: avoid; }
            .c-id { width: 70px; font-size: 15px; }
            .c-nm { text-transform: uppercase; }
            .c-dia { width: 100px; text-align: right; font-size: 14px; }
            
            @media print { body { -webkit-print-color-adjust: exact; } }
        </style>
    </head>
    <body>
        <div class="page">
            <div class="header">
                <div class="title-block">
                    <div class="p-name">${p.name || ''}</div>
                    <div class="p-num">${p.num || '---'}</div>
                </div>
                <table class="meta-table">
                    <tr><td>LAUFZEIT</td><td class="meta-val">${p.lzf || ''}</td></tr>
                    <tr><td>MATERIAL</td><td class="meta-val">${p.mat || ''}</td></tr>
                    <tr><td>SÄGELÄNGE</td><td class="meta-val">${p.sag || ''}</td></tr>
                    <tr><td>ABSTAND</td><td class="meta-val">${p.abs || ''}</td></tr>
                    <tr><td>GREIFBACKEN</td><td class="meta-val">${p.grf || ''}</td></tr>
                    <tr><td>STÜCKZAHL</td><td class="meta-val">${p.stt || ''} / ${p.stn || ''}</td></tr>
                </table>
            </div>

            ${oben.length > 0 ? `
                <div class="section-header">REVOLVER OBEN</div>
                <table class="tool-table">
                    <thead><tr><th>T-NR</th><th>WERKZEUGNAME</th><th style="text-align:right">Ø / TOLERANZ</th></tr></thead>
                    <tbody>
                        ${oben.map(t => `<tr><td class="c-id">${t.id}</td><td class="c-nm">${t.nm}</td><td class="c-dia">${t.dia}</td></tr>`).join('')}
                    </tbody>
                </table>
            ` : ''}

            ${unten.length > 0 ? `
                <div class="section-header">REVOLVER UNTEN</div>
                <table class="tool-table">
                    <thead><tr><th>T-NR</th><th>WERKZEUGNAME</th><th style="text-align:right">Ø / TOLERANZ</th></tr></thead>
                    <tbody>
                        ${unten.map(t => `<tr><td class="c-id">${t.id}</td><td class="c-nm">${t.nm}</td><td class="c-dia">${t.dia}</td></tr>`).join('')}
                    </tbody>
                </table>
            ` : ''}
        </div>
        <script>window.onload = () => { setTimeout(() => { window.print(); window.close(); }, 500); };<\/script>
    </body>
    </html>`;

    const win = window.open('', '_blank');
    win.document.write(html);
    win.document.close();
}

window.onload = () => { injectStyles(); buildStructure(); renderProjects(); };
