// БАЗА ДАННЫХ
let db = JSON.parse(localStorage.getItem('QS_FINAL_V1')) || [];
let cur = null; // Индекс текущего проекта

// ВСПОМОГАТЕЛЬНЫЕ
const el = (id) => document.getElementById(id);
const show = (id) => el(id).style.display = 'flex';
const hide = (id) => el(id).style.display = 'none';

// НАВИГАЦИЯ
function goHome() {
    cur = null;
    el('v-det').classList.remove('active');
    el('v-home').classList.add('active');
    renderHome();
}

function openP(i) {
    cur = i;
    el('v-home').classList.remove('active');
    el('v-det').classList.add('active');
    el('h-num').innerText = db[i].num;
    el('h-nam').innerText = db[i].name;
    renderTools();
}

// РЕНДЕР СПИСКОВ
function renderHome() {
    el('list-p').innerHTML = db.map((p, i) => `
        <div class="item" onclick="openP(${i})">
            <div class="item-info"><b>${p.num}</b><div>${p.name}</div></div>
            <div class="btn-del-mini" onclick="event.stopPropagation(); delP(${i})">✕</div>
        </div>
    `).join('');
}

function renderTools() {
    el('list-t').innerHTML = db[cur].tools.map((t, i) => `
        <div class="item" onclick="modalT(${i})">
            <div class="item-info"><b>${t.id}</b><div>${t.nm}</div></div>
            <div class="item-val">${t.dia}</div>
        </div>
    `).join('');
}

// МОДАЛКА ПРОЕКТА
function modalP(edit = false) {
    const p = edit ? db[cur] : {num:'',name:'',lzf:'',sag:'',stt:'',stn:'',abs:'',grf:''};
    el('p-idx').value = edit ? cur : "";
    
    // Заполняем поля
    ['num','nam','lzf','sag','stt','stn','abs','grf'].forEach(k => {
        el('p-'+k).value = p[k] || "";
    });
    show('m-p');
}

function saveP() {
    const idx = el('p-idx').value;
    const data = {
        num: el('p-num').value,
        name: el('p-nam').value.toUpperCase(),
        lzf: el('p-lzf').value,
        sag: el('p-sag').value,
        stt: el('p-stt').value,
        stn: el('p-stn').value,
        abs: el('p-abs').value,
        grf: el('p-grf').value,
        tools: (idx !== "" && db[idx]) ? db[idx].tools : []
    };

    if(idx === "") db.push(data);
    else db[idx] = data;

    localStorage.setItem('QS_FINAL_V1', JSON.stringify(db));
    hide('m-p');
    
    if(idx === "") goHome();
    else { openP(cur); } // Обновить заголовок если редактировали
}

// МОДАЛКА ИНСТРУМЕНТА
function modalT(i = null) {
    const edit = i !== null;
    el('t-idx').value = edit ? i : "";
    const t = edit ? db[cur].tools[i] : {id:'',nm:'',dia:''};
    
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

    if(i === "") db[cur].tools.push(t);
    else db[cur].tools[i] = t;

    localStorage.setItem('QS_FINAL_V1', JSON.stringify(db));
    renderTools();
    hide('m-t');
}

// УДАЛЕНИЕ
function delP(i) {
    if(confirm('Удалить проект?')) {
        db.splice(i, 1);
        localStorage.setItem('QS_FINAL_V1', JSON.stringify(db));
        renderHome();
    }
}
function delT() {
    const i = el('t-idx').value;
    db[cur].tools.splice(i, 1);
    localStorage.setItem('QS_FINAL_V1', JSON.stringify(db));
    renderTools();
    hide('m-t');
}

// ИМПОРТ
function runImp() {
    const lines = el('imp-area').value.split('\n');
    lines.forEach(line => {
        if(line.trim()) {
            db[cur].tools.push({ id:'T?', nm: line.trim().toUpperCase(), dia:'' });
        }
    });
    localStorage.setItem('QS_FINAL_V1', JSON.stringify(db));
    el('imp-area').value = "";
    renderTools();
    hide('m-imp');
}

// PDF ГЕНЕРАЦИЯ (ЖЕСТКАЯ ВЕРСТКА)
function makePDF() {
    const p = db[cur];
    
    // Генерация строк таблицы с жесткими линиями
    const rows = p.tools.map(t => `
        <tr style="border-bottom: 1px solid black;">
            <td style="padding:10px 0; font-weight:900; font-size:11pt; width:60px;">${t.id}</td>
            <td style="padding:10px 10px; font-weight:500; font-size:11pt; text-transform:uppercase;">${t.nm}</td>
            <td style="padding:10px 0; font-weight:900; font-size:12pt; text-align:right;">${t.dia}</td>
        </tr>
    `).join('');

    // HTML для печати
    el('pdf-box').innerHTML = `
        <div style="width:210mm; padding:10mm; background:white; font-family:sans-serif; color:black;">
            <div style="border:3px solid black; padding:30px; min-height:270mm;">
                
                <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                    <div>
                        <div style="font-size:10pt; font-weight:bold; color:#555;">${p.name}</div>
                        <div style="font-size:60pt; font-weight:900; line-height:0.8; letter-spacing:-3px; margin-top:5px;">${p.num}</div>
                    </div>
                    
                    <div style="width:260px; font-size:10pt; font-weight:800; line-height:2.0;">
                        <div style="display:flex; justify-content:space-between; border-bottom:1px solid #000;">
                            <span>LAUFZEIT:</span><span>${p.lzf}</span>
                        </div>
                        <div style="display:flex; justify-content:space-between; border-bottom:1px solid #000;">
                            <span>SÄGELÄNGE:</span><span>${p.sag}</span>
                        </div>
                        <div style="display:flex; justify-content:space-between; border-bottom:1px solid #000;">
                            <span>STÜCK T:</span><span>${p.stt}</span>
                        </div>
                        <div style="display:flex; justify-content:space-between; border-bottom:1px solid #000;">
                            <span>STÜCK N:</span><span>${p.stn}</span>
                        </div>
                        <div style="display:flex; justify-content:space-between; border-bottom:1px solid #000;">
                            <span>ABSTAND:</span><span>${p.abs}</span>
                        </div>
                        <div style="display:flex; justify-content:space-between; border-bottom:1px solid #000;">
                            <span>GREIFBACKEN:</span><span>${p.grf}</span>
                        </div>
                    </div>
                </div>

                <div style="height:6px; background:black; margin:30px 0 10px 0;"></div>

                <table style="width:100%; border-collapse:collapse;">
                    <thead>
                        <tr style="border-bottom:3px solid black;">
                            <th align="left" style="font-size:8pt; padding-bottom:5px;">T-NR</th>
                            <th align="left" style="font-size:8pt; padding-bottom:5px; padding-left:10px;">WERKZEUGNAME</th>
                            <th align="right" style="font-size:8pt; padding-bottom:5px;">Ø / TOLERANZ</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${rows}
                    </tbody>
                </table>

            </div>
        </div>
    `;
    
    window.print();
}

// Запуск
renderHome();
