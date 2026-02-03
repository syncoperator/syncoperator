let db = JSON.parse(localStorage.getItem('QS_PRO_V4')) || [];
let cur = null;

const show = (id) => document.getElementById(id).style.display = 'flex';
const hide = (id) => document.getElementById(id).style.display = 'none';

function goHome() {
    cur = null;
    document.getElementById('v-home').classList.add('active');
    document.getElementById('v-det').classList.remove('active');
    renderHome();
}

function renderHome() {
    const list = document.getElementById('list-p');
    list.innerHTML = db.map((p, i) => `
        <div class="card" onclick="openP(${i})">
            <div class="card-main"><b>${p.num}</b><div>${p.name}</div></div>
            <div style="color:#FF453A" onclick="event.stopPropagation();delP(${i})">✕</div>
        </div>`).join('');
}

function openP(i) {
    cur = i;
    document.getElementById('v-home').classList.remove('active');
    document.getElementById('v-det').classList.add('active');
    document.getElementById('h-num').innerText = db[i].num;
    document.getElementById('h-nam').innerText = db[i].name;
    renderT();
}

function renderT() {
    const list = document.getElementById('list-t');
    list.innerHTML = db[cur].tools.map((t, i) => `
        <div class="card" onclick="modalT(${i})">
            <div class="card-main"><b>${t.id}</b><div>${t.nm}</div></div>
            <div class="card-val">${t.dia}</div>
        </div>`).join('');
}

function modalP(edit = false) {
    const p = edit ? db[cur] : {num:'',name:'',abs:'',grf:'',lzf:'',sag:'',sta:'',stb:''};
    document.getElementById('p-idx').value = edit ? cur : "";
    ['num','nam','abs','grf','lzf','sag','sta','stb'].forEach(k => document.getElementById('p-'+k).value = p[k] || "");
    show('m-p');
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
    if(i==="") db.push(data); else db[i] = data;
    localStorage.setItem('QS_PRO_V4', JSON.stringify(db));
    hide('m-p'); goHome();
}

function runImp() {
    const lines = document.getElementById('imp-area').value.split('\n').filter(l => l.trim());
    lines.forEach(l => db[cur].tools.push({ id:'T?', nm:l.trim().toUpperCase(), dia:'' }));
    localStorage.setItem('QS_PRO_V4', JSON.stringify(db));
    renderT(); hide('m-imp');
}

function modalT(i=null) {
    const edit = i !== null;
    document.getElementById('t-idx').value = edit ? i : "";
    const t = edit ? db[cur].tools[i] : {id:'',nm:'',dia:''};
    document.getElementById('t-id').value = t.id;
    document.getElementById('t-nm').value = t.nm;
    document.getElementById('t-dia').value = t.dia;
    show('m-t');
}

function saveT() {
    const i = document.getElementById('t-idx').value;
    const t = { id: document.getElementById('t-id').value.toUpperCase(), nm: document.getElementById('t-nm').value.toUpperCase(), dia: document.getElementById('t-dia').value };
    if(i==="") db[cur].tools.push(t); else db[cur].tools[i] = t;
    localStorage.setItem('QS_PRO_V4', JSON.stringify(db));
    renderT(); hide('m-t');
}

function delP(i) { if(confirm('Löschen?')) { db.splice(i,1); localStorage.setItem('QS_PRO_V4', JSON.stringify(db)); renderHome(); } }
function delT() { db[cur].tools.splice(document.getElementById('t-idx').value, 1); localStorage.setItem('QS_PRO_V4', JSON.stringify(db)); renderT(); hide('m-t'); }

function makePDF() {
    const p = db[cur];
    const rows = p.tools.map(t => `
        <tr style="border-bottom: 0.5pt solid #000;">
            <td style="padding: 12px 0; font-weight: 900; font-size: 10pt; width: 60px;">${t.id}</td>
            <td style="padding: 12px 10px; font-weight: 500; font-size: 10pt; text-transform: uppercase;">${t.nm}</td>
            <td style="padding: 12px 0; font-weight: 900; font-size: 11pt; text-align: right; width: 100px;">${t.dia}</td>
        </tr>`).join('');

    document.getElementById('pdf-box').innerHTML = `
    <div style="width: 210mm; padding: 15mm; background: white; color: black; font-family: sans-serif;">
        <div style="border: 2pt solid black; padding: 30px; min-height: 260mm; position: relative;">
            <div style="display: flex; justify-content: space-between;">
                <div>
                    <div style="font-size: 10pt; font-weight: bold; color: #555;">${p.name}</div>
                    <div style="font-size: 56pt; font-weight: 900; line-height: 0.8; letter-spacing: -3px;">${p.num}</div>
                </div>
                <div style="width: 230px; font-size: 9pt; font-weight: 800; line-height: 1.8;">
                    <div style="display:flex; justify-content:space-between; border-bottom:0.5pt solid #ccc"><span>ABSTAND:</span><span>${p.abs}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:0.5pt solid #ccc"><span>GREIFBACKEN:</span><span>${p.grf}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:0.5pt solid #ccc"><span>LAUFZEIT:</span><span>${p.lzf}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:0.5pt solid #ccc"><span>SÄGELÄНGE:</span><span>${p.sag}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>STÜCK A: ${p.sta}</span><span>B: ${p.stb}</span></div>
                </div>
            </div>
            <div style="height: 5pt; background: black; margin: 30px 0 10px 0;"></div>
            <table style="width: 100%; border-collapse: collapse;">
                <thead>
                    <tr style="border-bottom: 2pt solid black;">
                        <th align="left" style="width: 60px; font-size: 8pt; padding-bottom: 5px;">T-NR</th>
                        <th align="left" style="font-size: 8pt; padding-bottom: 5px; padding-left: 10px;">WERKZEUGNAME</th>
                        <th align="right" style="width: 100px; font-size: 8pt; padding-bottom: 5px;">Ø / TOLERANZ</th>
                    </tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>
        </div>
    </div>`;
    window.print();
}

renderHome();
