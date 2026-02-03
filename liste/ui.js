let db = JSON.parse(localStorage.getItem('QS_DB_V3')) || [];
let curP = null;

const showM = (id) => document.getElementById(id).style.display = 'flex';
const hideM = (id) => document.getElementById(id).style.display = 'none';

function goHome() { 
    curP = null; 
    document.getElementById('v-det').classList.remove('active');
    document.getElementById('v-home').classList.add('active');
    renderHome();
}

function renderHome() {
    const c = document.getElementById('list-p');
    c.innerHTML = db.map((p, i) => `
        <div class="item-card" onclick="openP(${i})">
            <div class="item-main"><b>${p.num}</b><div>${p.name}</div></div>
            <button class="btn-danger" style="width:auto" onclick="event.stopPropagation();delP(${i})">✕</button>
        </div>`).join('');
}

function openP(i) {
    curP = i;
    document.getElementById('v-home').classList.remove('active');
    document.getElementById('v-det').classList.add('active');
    document.getElementById('h-num').innerText = db[i].num;
    document.getElementById('h-nam').innerText = db[i].name;
    renderT();
}

function renderT() {
    const c = document.getElementById('list-t');
    c.innerHTML = db[curP].tools.map((t, i) => `
        <div class="item-card" data-idx="${i}">
            <div class="item-main" onclick="modalT(${i})"><b>${t.id}</b><div>${t.nm}</div></div>
            <div class="item-dia">${t.dia}</div>
        </div>`).join('');
}

function modalP(edit = false) {
    const p = edit ? db[curP] : {num:'',name:'',abs:'',grf:'',lzf:'',sag:'',sta:'',stb:''};
    document.getElementById('p-idx').value = edit ? curP : "";
    ['num','nam','abs','grf','lzf','sag','sta','stb'].forEach(k => document.getElementById('p-'+k).value = p[k] || p.name || "");
    showM('m-p');
}

function saveP() {
    const i = document.getElementById('p-idx').value;
    const data = {
        num: document.getElementById('p-num').value,
        name: document.getElementById('p-nam').value.toUpperCase(),
        abs: document.getElementById('p-abs').value,
        grf: document.getElementById('p-grf').value,
        lzf: document.getElementById('p-lzf').value,
        sag: document.getElementById('p-sag').value,
        sta: document.getElementById('p-sta').value,
        stb: document.getElementById('p-stb').value,
        tools: i === "" ? [] : db[i].tools
    };
    if(i === "") db.push(data); else db[i] = data;
    localStorage.setItem('QS_DB_V3', JSON.stringify(db));
    hideM('m-p'); goHome();
}

function runImp() {
    const lines = document.getElementById('imp-area').value.split('\n').filter(l => l.trim());
    lines.forEach(l => db[curP].tools.push({ id:'T?', nm:l.trim().toUpperCase(), dia:'' }));
    localStorage.setItem('QS_DB_V3', JSON.stringify(db));
    renderT(); hideM('m-imp');
}

function modalT(i = null) {
    const edit = i !== null;
    document.getElementById('t-idx').value = edit ? i : "";
    const t = edit ? db[curP].tools[i] : {id:'',nm:'',dia:''};
    document.getElementById('t-id').value = t.id;
    document.getElementById('t-nm').value = t.nm;
    document.getElementById('t-dia').value = t.dia;
    document.getElementById('btn-del-t').style.display = edit ? 'block' : 'none';
    showM('m-t');
}

function saveT() {
    const i = document.getElementById('t-idx').value;
    const t = { id: document.getElementById('t-id').value.toUpperCase(), nm: document.getElementById('t-nm').value.toUpperCase(), dia: document.getElementById('t-dia').value };
    if(i === "") db[curP].tools.push(t); else db[curP].tools[i] = t;
    localStorage.setItem('QS_DB_V3', JSON.stringify(db));
    renderT(); hideM('m-t');
}

function delP(i) { if(confirm('Löschen?')) { db.splice(i,1); localStorage.setItem('QS_DB_V3', JSON.stringify(db)); renderHome(); } }
function delT() { db[curP].tools.splice(document.getElementById('t-idx').value, 1); localStorage.setItem('QS_DB_V3', JSON.stringify(db)); renderT(); hideM('m-t'); }

function makePDF() {
    const p = db[curP];
    const rows = p.tools.map(t => `
        <tr style="border-bottom: 0.5pt solid black;">
            <td style="padding: 12px 0; font-weight: 900; font-size: 10pt; width: 60px;">${t.id}</td>
            <td style="padding: 12px 10px; font-weight: 500; font-size: 10pt; text-transform: uppercase;">${t.nm}</td>
            <td style="padding: 12px 0; font-weight: 900; font-size: 11pt; text-align: right; width: 100px;">${t.dia}</td>
        </tr>`).join('');

    document.getElementById('pdf-render').innerHTML = `
    <div style="width: 210mm; min-height: 297mm; padding: 12mm; background: white; color: black; box-sizing: border-box;">
        <div style="border: 1.5pt solid black; height: 100%; min-height: 270mm; padding: 30px; display: flex; flex-direction: column; box-sizing: border-box;">
            
            <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                <div>
                    <div style="font-size: 10pt; font-weight: 700; color: #666;">${p.name}</div>
                    <div style="font-size: 56pt; font-weight: 900; line-height: 0.8; letter-spacing: -3px; margin-top: 5px;">${p.num}</div>
                </div>
                <div style="text-align: right; font-size: 9pt; font-weight: 800; line-height: 1.6; min-width: 220px;">
                    <div style="display: flex; justify-content: space-between; border-bottom: 0.5pt solid #ddd;"><span>ABSTAND:</span> <span>${p.abs}</span></div>
                    <div style="display: flex; justify-content: space-between; border-bottom: 0.5pt solid #ddd;"><span>GREIFBACKEN:</span> <span>${p.grf}</span></div>
                    <div style="display: flex; justify-content: space-between; border-bottom: 0.5pt solid #ddd;"><span>LAUFZEIT:</span> <span>${p.lzf}</span></div>
                    <div style="display: flex; justify-content: space-between; border-bottom: 0.5pt solid #ddd;"><span>SÄGELÄNGE:</span> <span>${p.sag}</span></div>
                    <div style="display: flex; justify-content: space-between;"><span>STÜCK A: ${p.sta}</span> <span>B: ${p.stb}</span></div>
                </div>
            </div>

            <div style="height: 4pt; background: black; margin: 25px 0 10px 0;"></div>

            <table style="width: 100%; border-collapse: collapse; table-layout: fixed;">
                <thead>
                    <tr style="border-bottom: 2pt solid black;">
                        <th align="left" style="width: 60px; font-size: 7.5pt; font-weight: 900; padding-bottom: 5px;">T-NR</th>
                        <th align="left" style="font-size: 7.5pt; font-weight: 900; padding-bottom: 5px; padding-left: 10px;">WERKZEUGNAME</th>
                        <th align="right" style="width: 100px; font-size: 7.5pt; font-weight: 900; padding-bottom: 5px;">Ø / TOLERANZ</th>
                    </tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>
        </div>
    </div>`;
    window.print();
}

renderHome();
