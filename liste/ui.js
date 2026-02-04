const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

// --- CORE FUNCTIONS ---
const el = (id) => document.getElementById(id);
const show = (id) => el(id).style.display = 'flex';
const hide = (id) => el(id).style.display = 'none';

function renderList() {
    const list = el('list-p');
    if(db.length === 0) {
        list.innerHTML = '<div style="padding:40px; text-align:center; color:#555;">Keine Projekte</div>';
        return;
    }
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})">
            <div>
                <b>${p.num}</b>
                <small>${p.name}</small>
            </div>
            <div style="color: #FF453A; padding: 10px;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>
    `).join('');
}

function renderTools() {
    const tools = db[currentIdx].tools;
    el('list-t').innerHTML = tools.map((t, i) => `
        <div class="list-item" onclick="modalT(${i})">
            <div>
                <b>${t.id}</b>
                <small>${t.nm}</small>
            </div>
            <div class="meta">${t.dia}</div>
        </div>
    `).join('');
}

// --- ACTIONS ---

function goHome() {
    currentIdx = null;
    el('v-det').classList.remove('active');
    el('v-home').classList.add('active');
    renderList();
}

function openProject(i) {
    currentIdx = i;
    el('v-home').classList.remove('active');
    el('v-det').classList.add('active');
    el('h-num').innerText = db[i].num;
    el('h-nam').innerText = db[i].name;
    renderTools();
}

function modalP(edit = false) {
    const p = edit ? db[currentIdx] : {num:'', name:'', lzf:'', sag:'', stt:'', stn:'', abs:'', grf:''};
    el('p-idx').value = edit ? currentIdx : '';
    el('p-num').value = p.num;
    el('p-nam').value = p.name;
    el('p-lzf').value = p.lzf;
    el('p-sag').value = p.sag;
    el('p-stt').value = p.stt;
    el('p-stn').value = p.stn;
    el('p-abs').value = p.abs;
    el('p-grf').value = p.grf;
    show('m-p');
}

function saveP() {
    const idx = el('p-idx').value;
    const newP = {
        num: el('p-num').value,
        name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value,
        sag: el('p-sag').value,
        stt: el('p-stt').value,
        stn: el('p-stn').value,
        abs: el('p-abs').value,
        grf: el('p-grf').value,
        tools: (idx !== '' && db[idx]) ? db[idx].tools : []
    };

    if (idx === '') db.push(newP);
    else db[idx] = newP;
    
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    hide('m-p');
    if(idx === '') goHome();
    else openProject(currentIdx);
}

function deleteProject(i) {
    if(confirm('Projekt löschen?')) {
        db.splice(i, 1);
        localStorage.setItem(DB_KEY, JSON.stringify(db));
        renderList();
    }
}

// --- TOOLS & IMPORT ---

function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : '';
    const t = edit ? db[currentIdx].tools[i] : {id:'', nm:'', dia:''};
    el('t-id').value = t.id;
    el('t-nm').value = t.nm;
    el('t-dia').value = t.dia;
    el('btn-del-t').style.display = edit ? 'block' : 'none';
    show('m-t');
}

function saveT() {
    const i = el('t-idx').value;
    const t = {
        id: el('t-id').value.toUpperCase(),
        nm: el('t-nm').value.toUpperCase(),
        dia: el('t-dia').value
    };
    if(i === '') db[currentIdx].tools.push(t);
    else db[currentIdx].tools[i] = t;
    
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools();
    hide('m-t');
}

function delT() {
    const i = el('t-idx').value;
    db[currentIdx].tools.splice(i, 1);
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools();
    hide('m-t');
}

function runImp() {
    const text = el('imp-area').value;
    const lines = text.split('\n');
    lines.forEach(line => {
        const clean = line.trim();
        if(clean) {
            db[currentIdx].tools.push({ id:'T?', nm: clean.toUpperCase(), dia:'' });
        }
    });
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    el('imp-area').value = '';
    renderTools();
    hide('m-imp');
}

// --- PDF GENERATION (PIXEL PERFECT REPLICA) ---
function makePDF() {
    const p = db[currentIdx];
    
    // Генерируем строки таблицы
    const rows = p.tools.map(t => `
        <div style="display: flex; align-items: flex-start; border-bottom: 1px solid #000; padding: 10px 0;">
            <div style="width: 60px; font-weight: 700; font-size: 14px;">${t.id}</div>
            <div style="flex: 1; font-weight: 500; font-size: 14px; text-transform: uppercase;">${t.nm}</div>
            <div style="width: 80px; text-align: right; font-weight: 700; font-size: 14px;">${t.dia}</div>
        </div>
    `).join('');

    // HTML структура точь-в-точь как на фото 19:54
    const html = `
    <div style="width: 210mm; min-height: 297mm; padding: 15mm; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; box-sizing: border-box; position: relative;">
        
        <div style="border: 2px solid #000; height: 100%; min-height: 250mm; padding: 25px; box-sizing: border-box; display: flex; flex-direction: column;">
            
            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 20px;">
                
                <div>
                    <div style="font-size: 12px; font-weight: 700; color: #888; text-transform: uppercase; margin-bottom: 5px;">${p.name}</div>
                    <div style="font-size: 64px; font-weight: 900; line-height: 0.8; letter-spacing: -2px;">${p.num}</div>
                </div>

                <div style="width: 220px; font-size: 11px; font-weight: 700; line-height: 1.8;">
                    <div style="display:flex; justify-content:space-between; border-bottom: 1px solid #ccc; padding-bottom: 2px; margin-bottom: 2px;">
                        <span>ABSTAND:</span> <span>${p.abs}</span>
                    </div>
                    <div style="display:flex; justify-content:space-between; border-bottom: 1px solid #ccc; padding-bottom: 2px; margin-bottom: 2px;">
                        <span>GREIFBACKEN:</span> <span>${p.grf}</span>
                    </div>
                    <div style="display:flex; justify-content:space-between; border-bottom: 1px solid #ccc; padding-bottom: 2px; margin-bottom: 2px;">
                        <span>LAUFZEIT:</span> <span>${p.lzf}</span>
                    </div>
                    <div style="display:flex; justify-content:space-between; border-bottom: 1px solid #ccc; padding-bottom: 2px; margin-bottom: 2px;">
                        <span>SÄGELÄNGE:</span> <span>${p.sag}</span>
                    </div>
                    <div style="display:flex; justify-content:space-between; border-bottom: 1px solid #ccc; padding-bottom: 2px; margin-bottom: 2px;">
                        <span>STÜCK T:</span> <span>${p.stt}</span>
                    </div>
                     <div style="display:flex; justify-content:space-between; border-bottom: 1px solid #ccc; padding-bottom: 2px;">
                        <span>STÜCK N:</span> <span>${p.stn}</span>
                    </div>
                </div>
            </div>

            <div style="border-bottom: 4px solid #000; margin-bottom: 15px;"></div>

            <div style="display: flex; font-size: 10px; font-weight: 700; color: #888; text-transform: uppercase; margin-bottom: 5px;">
                <div style="width: 60px;">T-NR</div>
                <div style="flex: 1;">WERKZEUGNAME / KOMMENTAR</div>
                <div style="width: 80px; text-align: right;">Ø / TOLERANZ</div>
            </div>

             <div style="border-bottom: 2px solid #000; margin-bottom: 0px;"></div>

            <div style="flex: 1;">
                ${rows}
            </div>

        </div>
    </div>
    `;

    el('print-container').innerHTML = html;
    window.print();
}

// Init
renderList();
