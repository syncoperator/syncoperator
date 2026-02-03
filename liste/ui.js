let db = JSON.parse(localStorage.getItem('QS_PRO_V5')) || [];
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
            <div style="color:#FF453A; padding: 10px;" onclick="event.stopPropagation();delP(${i})">✕</div>
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
            <div style="font-weight:900; font-size:18px; color:var(--accent)">${t.dia}</div>
        </div>`).join('');
}

function modalP(edit = false) {
    const p = edit ? db[cur] : {num:'',name:'',lzf:'',sag:'',stt:'',stn:'',abs:'',grf:''};
    document.getElementById('p-idx').value = edit ? cur : "";
    ['num','nam','lzf','sag','stt','stn','abs','grf'].forEach(k => document.getElementById('p-'+k).value = p[k] || "");
    show('m-p');
}

function saveP() {
    const i = document.getElementById('p-idx').value;
    const data = {
        num: document.getElementById('p-num').value,
        name: document.getElementById('p-nam').value.toUpperCase(),
        lzf: document.getElementById('p-lzf').value,
        sag: document.getElementById('p-sag').value,
        stt: document.getElementById('p-stt').value,
        stn: document.getElementById('p-stn').value,
        abs: document.getElementById('p-abs').value,
        grf: document.getElementById('p-grf').value,
        tools: (i !== "" && db[i]) ? db[i].tools : []
    };
    if(i==="") db.push(data); else db[i] = data;
    localStorage.setItem('QS_PRO_V5', JSON.stringify(db));
    hide('m-p'); goHome();
}

function modalT(i=null) {
    const edit = i !== null;
    document.getElementById('t-idx').value = edit ? i : "";
    const t = edit ? db[cur].tools[i] : {id:'',nm:'',dia:''};
    document.getElementById('t-id').value = t.id;
    document.getElementById('t-nm').value = t.nm;
    document.getElementById('t-dia').value = t.dia;
    document.getElementById('btn-del-t').style.display = edit ? 'block' : 'none';
    show('m-t');
}

function saveT() {
    const i = document.getElementById('t-idx').value;
    const t = { id: document.getElementById('t-id').value.toUpperCase(), nm: document.getElementById('t-nm').value.toUpperCase(), dia: document.getElementById('t-dia').value };
    if(i==="") db[cur].tools.push(t); else db[cur].tools[i] = t;
    localStorage.setItem('QS_PRO_V5', JSON.stringify(db));
    renderT(); hide('m-t');
}

function runImp() {
    const lines = document.getElementById('imp-area').value.split('\n').filter(l => l.trim());
    lines.forEach(l => db[cur].tools.push({ id:'T?', nm:l.trim().toUpperCase(), dia:'' }));
    localStorage.setItem('QS_PRO_V5', JSON.stringify(db));
    document.getElementById('imp-area').value = "";
    renderT(); hide('m-imp');
}

function delP(i) { if(confirm('Löschen?')) { db.splice(i,1); localStorage.setItem('QS_PRO_V5', JSON.stringify(db)); renderHome(); } }
function delT() { db[cur].tools.splice(document.getElementById('t-idx').value, 1); localStorage.setItem('QS_PRO_V5', JSON.stringify(db)); renderT(); hide('m-t'); }

function makePDF() {
    const p = db[cur];
    const rows = p.tools.map(t => `
        <tr style="border-bottom: 1px solid black !important;">
            <td style="padding: 15px 0; font-weight: 900; font-size: 10pt; width: 60px;">${t.id}</td>
            <td style="padding: 15px 10px; font-weight: 500; font-size: 10pt; text-transform: uppercase;">${t.nm}</td>
            <td style="padding: 15px 0; font-weight: 900; font-size: 11pt; text-align: right; width: 120px;">${t.dia}</td>
        </tr>`).join('');

    document.getElementById('pdf-box').innerHTML = `
    <div style="width: 210mm; padding: 15mm; background: white; color: black; font-family: sans-serif;">
        <div style="border: 2.5pt solid black; padding: 35px; min-height: 260mm; box-sizing: border-box;">
            <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                <div>
                    <div style="font-size: 10pt; font-weight: 700; color: #555;">${p.name}</div>
                    <div style="font-size: 58pt; font-weight: 900; line-height: 0.8; letter-spacing: -3px; margin-top: 10px;">${p.num}</div>
                </div>
                <div style="width: 260px; font-size: 10pt; font-weight: 800; line-height: 2.1;">
                    <div style="display:flex; justify-content:space-between; border-bottom:0.8pt solid black"><span>LAUFZEIT:</span><span>${p.lzf}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:0.8pt solid black"><span>SÄGELÄNGE:</span><span>${p.sag}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:0.8pt solid black"><span>STÜCK T:</span><span>${p.stt}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:0.8pt solid black"><span>STÜCK N:</span><span>${p.stn}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:0.8pt solid black"><span>ABSTAND:</span><span>${p.abs}</span></div>
                    <div style="display:flex; justify-content:space-between; border-bottom:0.8pt solid black"><span>GREIFBACKEN:</span><span>${p.grf}</span></div>
                </div>
            </div>
            <div style="height: 6pt; background: black; margin: 35px 0 10px 0;"></div>
            <table style="width: 100%; border-collapse: collapse; table-layout: fixed;">
                <thead>
                    <tr style="border-bottom: 2.5pt solid black;">
                        <th align="left" style="width: 70px; font-size: 8pt; padding-bottom: 10px;">T-NR</th>
                        <th align="left" style="font-size: 8pt; padding-bottom: 10px; padding-left: 10px;">WERKZEUGNAME / KOMMENTAR</th>
                        <th align="right" style="width: 120px; font-size: 8pt; padding-bottom: 10px;">Ø / TOLERANZ</th>
                    </tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>
        </div>
    </div>`;
    window.print();
}

renderHome();
