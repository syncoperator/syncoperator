const DB_KEY = 'SMARTBOX_PRO_V1';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const hide = (id) => el(id).style.display = 'none';

// --- PROJECTS ---
function renderList() {
    const list = el('list-p');
    if(!list) return;
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})">
            <div style="flex:1">
                <div class="t-id-label">${p.name || 'PROJEKT'}</div>
                <div class="t-name-label">${p.num || '---'}</div>
            </div>
            <div style="font-size:18px; padding:10px; color:#86868b;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:100px"></div>';
}

function openProject(i) {
    currentIdx = i;
    el('v-home').classList.remove('active');
    el('v-det').classList.add('active');
    el('h-num').innerText = db[i].num;
    el('h-nam').innerText = db[i].name;
    renderTools();
}

function goHome() {
    currentIdx = null;
    el('v-home').classList.add('active');
    el('v-det').classList.remove('active');
    renderList();
}

// --- TOOLS ---
function renderTools() {
    const list = el('list-t');
    if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'list-item';
        item.setAttribute('data-idx', i);
        
        const locClass = t.isStandart ? 'standart' : 'sonder';
        const locText = t.isStandart ? `STANDART | ${t.loc || 'FACH'}` : 'SONDER | BLAUKISTE';

        item.innerHTML = `
            <div class="handle">☰</div>
            <div style="flex:1" onclick="modalT(${i})">
                <div class="t-id-label">${t.id}</div>
                <div class="t-name-label">${t.nm}</div>
                <div class="loc-tag ${locClass}">${locText}</div>
                ${t.dia ? `<br><div class="t-dia-badge">${t.dia}</div>` : ''}
            </div>`;
        
        // Drag & Drop Logic
        const handle = item.querySelector('.handle');
        handle.ontouchstart = () => { startIdx = i; };
        handle.ontouchmove = (e) => {
            const touch = e.touches[0];
            const target = document.elementFromPoint(touch.clientX, touch.clientY)?.closest('.list-item');
            if (target) {
                const overIdx = parseInt(target.getAttribute('data-idx'));
                if (overIdx !== startIdx) {
                    const currentTools = db[currentIdx].tools;
                    const itemToMove = currentTools.splice(startIdx, 1)[0];
                    currentTools.splice(overIdx, 0, itemToMove);
                    startIdx = overIdx;
                    localStorage.setItem(DB_KEY, JSON.stringify(db));
                    renderTools();
                }
            }
        };
        list.appendChild(item);
    });
    list.innerHTML += '<div style="height:150px"></div>';
}

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:'', isStandart:false, loc:''};
    el('t-id').value = t.id; 
    el('t-nm').value = t.nm; 
    el('t-dia').value = t.dia;
    el('t-loc').value = t.loc || '';
    
    const btn = el('btn-storage-toggle');
    if(t.isStandart) btn.classList.add('on'); else btn.classList.remove('on');
    btn.onclick = () => btn.classList.toggle('on');
    
    el('btn-del-t').style.display = edit ? 'block' : 'none';
    el('m-t').style.display = 'flex';
}

function saveT() {
    const i = el('t-idx').value;
    const t = { 
        id: el('t-id').value.toUpperCase(), 
        nm: el('t-nm').value.toUpperCase(), 
        dia: el('t-dia').value, 
        isStandart: el('btn-storage-toggle').classList.contains('on'),
        loc: el('t-loc').value.toUpperCase()
    };
    if(!db[currentIdx].tools) db[currentIdx].tools = [];
    if(i === '') db[currentIdx].tools.push(t); else db[currentIdx].tools[i] = t;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools(); 
    hide('m-t');
}

function modalP() { 
    el('p-idx').value = ''; 
    el('p-num').value = '';
    el('p-nam').value = '';
    el('m-p').style.display = 'flex'; 
}

function saveP() {
    const p = { 
        num: el('p-num').value, 
        name: el('p-nam').value.toUpperCase(), 
        tools: [] 
    };
    db.push(p); 
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderList(); 
    hide('m-p');
}

function deleteProject(i) { 
    if(confirm('Löschen?')) { 
        db.splice(i,1); 
        localStorage.setItem(DB_KEY, JSON.stringify(db)); 
        renderList(); 
    } 
}

function delT() { 
    db[currentIdx].tools.splice(el('t-idx').value, 1); 
    localStorage.setItem(DB_KEY, JSON.stringify(db)); 
    renderTools(); 
    hide('m-t'); 
}

// --- PDF ГЕНЕРАТОР (БЕЗ РАЗРЫВОВ) ---
function makePDF() {
    const p = db[currentIdx];
    const sonder = (p.tools || []).filter(t => !t.isStandart);
    const standart = (p.tools || []).filter(t => t.isStandart);

    const getRow = (t) => `
        <div style="display:flex; border-bottom:1.5px solid #000; padding:8px 0; align-items:center; page-break-inside:avoid;">
            <div style="width:75px; font-weight:800; font-size:15px;">${t.id}</div>
            <div style="flex:1;">
                <div style="font-weight:700; font-size:15px; text-transform:uppercase;">${t.nm}</div>
                <div style="font-size:10px; font-weight:800; color:#666;">
                    ${t.isStandart ? 'PLATZ: ' + (t.loc || '---') : 'IN BLAUKISTE'}
                </div>
            </div>
            <div style="width:100px; text-align:right; font-weight:800; font-size:14px;">${t.dia}</div>
        </div>`;

    const html = `<html>
    <head>
        <style>
            @page { size: A4; margin: 10mm; }
            body { margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, sans-serif; background: #fff; }
            .container { 
                border: 2.5px solid #000; 
                padding: 25px; 
                width: 100%; 
                box-sizing: border-box; 
                page-break-inside: avoid;
                display: block;
            }
            .section-header { 
                background: #000; color: #fff; 
                padding: 6px 12px; margin-top: 25px; 
                font-weight: 900; font-size: 13px; 
            }
            .footer { 
                margin-top: 30px; text-align: center; 
                font-size: 9px; font-weight: 800; 
                border-top: 1px solid #000; padding-top: 10px; 
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                <div>
                    <div style="font-size:14px; font-weight:900; color:#666; text-transform:uppercase;">${p.name || ''}</div>
                    <div style="font-size:60px; font-weight:900; line-height:0.8; letter-spacing:-1.5px;">${p.num || '---'}</div>
                </div>
            </div>
            
            <div style="margin-top:25px; border-top:5px solid #000;"></div>
            
            ${sonder.length > 0 ? `
                <div class="section-header">IN BLAUKISTE (SONDER)</div>
                ${sonder.map(getRow).join('')}
            ` : ''}
            
            ${standart.length > 0 ? `
                <div class="section-header">IN SCHUBLADEN (STANDART)</div>
                ${standart.map(getRow).join('')}
            ` : ''}
            
            <div class="footer">CITITOOL SMARTBOX LOGISTIC</div>
        </div>
        <script>
            window.onload = () => {
                setTimeout(() => { 
                    window.print(); 
                    window.close(); 
                }, 300);
            };
        <\/script>
    </body>
    </html>`;
    
    const win = window.open('', '_blank');
    if (win) {
        win.document.write(html);
        win.document.close();
    }
}

window.onload = () => renderList();
