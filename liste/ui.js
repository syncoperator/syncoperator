const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;

const el = (id) => document.getElementById(id);
const show = (id) => { 
    const target = el(id);
    if(target) target.style.display = 'flex';
};
const hide = (id) => {
    const target = el(id);
    if(target) target.style.display = 'none';
};

// Исправленная функция для кнопки +NEU
function modalP(edit = false) {
    const p = edit && currentIdx !== null ? db[currentIdx] : {num:'', name:'', lzf:'', sag:'', stt:'', stn:'', abs:'', grf:''};
    
    el('p-idx').value = edit ? currentIdx : '';
    el('p-num').value = p.num || '';
    el('p-nam').value = p.name || '';
    el('p-lzf').value = p.lzf || '';
    el('p-sag').value = p.sag || '';
    el('p-stt').value = p.stt || '';
    el('p-stn').value = p.stn || '';
    el('p-abs').value = p.abs || '';
    el('p-grf').value = p.grf || '';
    
    show('m-p');
}

function renderList() {
    const list = el('list-p');
    if(!list) return;
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
            <div style="color: #FF453A; padding: 10px; font-weight:bold;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>
    `).join('');
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
    el('v-det').classList.remove('active');
    el('v-home').classList.add('active');
    renderList();
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
    renderList();
    if(idx !== '') openProject(idx);
    else goHome();
}

function deleteProject(i) {
    if(confirm('Projekt löschen?')) {
        db.splice(i, 1);
        localStorage.setItem(DB_KEY, JSON.stringify(db));
        renderList();
    }
}

function renderTools() {
    const listT = el('list-t');
    if(!listT || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    listT.innerHTML = tools.map((t, i) => `
        <div class="list-item" onclick="modalT(${i})">
            <div>
                <b>${t.id}</b>
                <small>${t.nm}</small>
            </div>
            <div class="meta">${t.dia}</div>
        </div>
    `).join('');
}

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

function makePDF() {
    const p = db[currentIdx];
    const rows = p.tools.map(t => `
        <div style="display: flex; align-items: flex-start; border-bottom: 1px solid #f2f2f2; padding: 10px 0;">
            <div style="width: 65px; font-weight: 800; font-size: 13px; color: #000;">${t.id}</div>
            <div style="flex: 1; font-weight: 700; font-size: 13px; text-transform: uppercase; color: #000;">${t.nm}</div>
            <div style="width: 85px; text-align: right; font-weight: 800; font-size: 13px; color: #000;">${t.dia}</div>
        </div>
    `).join('');

    const html = `
    <div style="width: 210mm; padding: 15mm; font-family: 'Helvetica', Arial, sans-serif; color: #000;">
        <div style="border: 1.5px solid #000; padding: 20px; min-height: 260mm; display: flex; flex-direction: column;">
            <div style="display: flex; justify-content: space-between; margin-bottom: 25px;">
                <div>
                    <div style="font-size: 13px; font-weight: 900; text-transform: uppercase; margin-bottom: 3px;">${p.name}</div>
                    <div style="font-size: 60px; font-weight: 900; line-height: 0.9; letter-spacing: -1.5px;">${p.num}</div>
                </div>
                <div style="width: 180px; font-size: 11px; font-weight: 800; line-height: 1.6;">
                    <div style="display:flex; justify-content:space-between;"><span>ABSTAND</span><span>${p.abs}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>GREIFBACKEN</span><span>${p.grf}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>LAUFZEIT</span><span>${p.lzf}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>SÄGELÄNGE</span><span>${p.sag}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>STÜCK T</span><span>${p.stt}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>STÜCK N</span><span>${p.stn}</span></div>
                </div>
            </div>
            <div style="border-bottom: 3.5px solid #000; margin-bottom: 12px;"></div>
            <div style="display: flex; font-size: 10px; font-weight: 900; text-transform: uppercase; margin-bottom: 5px;">
                <div style="width: 65px;">T-NR</div>
                <div style="flex: 1;">WERKZEUGNAME / KOMMENTAR</div>
                <div style="width: 85px; text-align: right;">Ø / TOLERANZ</div>
            </div>
            <div style="border-bottom: 1.5px solid #f2f2f2; margin-bottom: 0px;"></div>
            <div style="flex: 1;">${rows}</div>
        </div>
    </div>`;

    const container = el('print-container');
    if(container) {
        container.innerHTML = html;
        window.print();
    }
}

// Запуск при загрузке
window.onload = renderList;
