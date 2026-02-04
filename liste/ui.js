const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'flex'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

// --- ПРОЕКТЫ ---
function renderList() {
    const list = el('list-p'); if(!list) return;
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})">
            <div><small>${p.name || '---'}</small><b>${p.num || '---'}</b></div>
            <div style="color:var(--danger); font-weight:900; padding:15px; z-index:20;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:100px"></div>';
}

function openProject(i) {
    if (db[i]) {
        currentIdx = i;
        el('v-home').classList.remove('active');
        el('v-det').classList.add('active');
        el('h-num').innerText = db[i].num || '---';
        el('h-nam').innerText = db[i].name || '---';
        renderTools();
    }
}

function goHome() { currentIdx = null; el('v-home').classList.add('active'); el('v-det').classList.remove('active'); renderList(); }

// --- ИНСТРУМЕНТЫ (APP) ---
function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'list-item';
        item.draggable = true;
        // Выравнивание строго по верхнему краю
        item.style.alignItems = 'flex-start';
        item.style.padding = '15px';

        item.innerHTML = `
            <div style="flex:1; padding-right:10px; display:flex; flex-direction:column; min-width:0;" onclick="modalT(${i})">
                <small style="color:#8e8e93; font-weight:700; font-size:11px; margin-bottom:2px;">${t.id || 'T0000'}</small>
                <b style="font-size:20px; font-weight:900; display:block; line-height:1.2; word-wrap:break-word; white-space:pre-wrap;">${t.nm || '---'}</b>
            </div>
            <div style="font-weight:900; color:var(--accent); font-size:18px; text-align:right; white-space:pre-line; min-width:110px; line-height:1.2; padding-top:14px;">${t.dia}</div>
        `;
        
        item.ondragstart = (e) => { e.dataTransfer.setData('text/plain', i); item.style.opacity = '0.4'; };
        item.ondragend = () => { item.style.opacity = '1'; renderTools(); };
        item.ondragover = (e) => e.preventDefault();
        item.ondrop = (e) => {
            e.preventDefault();
            const from = e.dataTransfer.getData('text/plain');
            moveTool(parseInt(from), i);
        };
        list.appendChild(item);
    });
    list.innerHTML += '<div style="height:120px; pointer-events:none;"></div>';
}

function moveTool(from, to) {
    const tools = db[currentIdx].tools;
    const item = tools.splice(from, 1)[0];
    tools.splice(to, 0, item);
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools();
}

// --- PDF (ВОЗВРАТ К ИСХОДНОМУ ВИДУ) ---
function makePDF() {
    const p = db[currentIdx];
    const rows = (p.tools || []).map(t => `
        <div style="display:flex; align-items:flex-start; border-bottom:0.5px solid #ccc; padding:10px 0; width:100%; min-height:35px;">
            <div style="width:70px; font-weight:800; font-size:13px; font-family:sans-serif; padding-top:2px;">${t.id}</div>
            <div style="flex:1; font-weight:700; font-size:13px; text-transform:uppercase; font-family:sans-serif; padding-right:15px; white-space:pre-wrap; line-height:1.3; padding-top:2px;">${t.nm}</div>
            <div style="width:120px; text-align:right; font-weight:800; font-size:13px; font-family:sans-serif; white-space:pre-line; line-height:1.3;">${t.dia}</div>
        </div>`).join('');

    const html = `
    <div style="width:210mm; padding:12mm; box-sizing:border-box; background:#fff; font-family:sans-serif; color:#000;">
        <div style="border:1px solid #000; padding:20px; min-height:265mm; display:flex; flex-direction:column; box-sizing:border-box;">
            
            <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:15px;">
                <div style="display:flex; flex-direction:column;">
                    <div style="font-size:11px; font-weight:900; text-transform:uppercase; color:#000;">${p.name || ''}</div>
                    <div style="font-size:55px; font-weight:900; line-height:0.9; letter-spacing:-1px;">${p.num || '---'}</div>
                </div>

                <div style="width:200px; font-size:10px; font-weight:800; line-height:1.4;">
                    <div style="display:flex; justify-content:space-between; border-bottom:0.5px solid #eee;"><span>ABSTAND</span><span>${p.abs || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:0.5px solid #eee;"><span>GREIFBACKEN</span><span>${p.grf || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:0.5px solid #eee;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:0.5px solid #eee;"><span>SÄGELÄNGE</span><span>${p.sag || ''}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:0.5px solid #eee;"><span>STÜCK T</span><span>${p.stt || ''}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>STÜCK N</span><span>${p.stn || ''}</span></div>
                </div>
            </div>

            <div style="border-bottom:4px solid #000; margin-bottom:10px;"></div>

            <div style="display:flex; font-size:9px; font-weight:900; text-transform:uppercase; margin-bottom:5px; padding:0 2px;">
                <div style="width:70px;">T-NR</div>
                <div style="flex:1;">WERKZEUGNAME / KOMMENTAR</div>
                <div style="width:120px; text-align:right;">Ø / TOLERANZ</div>
            </div>

            <div style="border-bottom:2px solid #000; margin-bottom:0px;"></div>
            <div style="flex:1;">${rows}</div>
        </div>
    </div>`;

    el('print-container').innerHTML = html;
    setTimeout(() => { window.print(); }, 150);
}

// (Остальные функции: saveP, modalP, saveT, modalT, delT, deleteProject, exportJSON, importJSON, runImp остаются без изменений)
