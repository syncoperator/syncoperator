let db = JSON.parse(localStorage.getItem('qs_v22')) || [];
let cur = null;

const showM = id => document.getElementById(id).style.display = 'flex';
const hideM = id => document.getElementById(id).style.display = 'none';
const save = () => localStorage.setItem('qs_v22', JSON.stringify(db));

function renderP() {
    const l = document.getElementById('list-p'); l.innerHTML = "";
    db.forEach((p, i) => {
        l.innerHTML += `<div class="card" onclick="openP(${i})">
            <div style="flex-grow:1"><b>${p.num}</b><br><small style="color:#888">${p.name}</small></div>
            <button class="c-del" onclick="event.stopPropagation();delP(${i})">✕</button>
        </div>`;
    });
}

function openP(i) {
    cur = i;
    document.getElementById('v-home').classList.remove('active');
    document.getElementById('v-det').classList.add('active');
    document.getElementById('d-num').innerText = db[i].num;
    document.getElementById('d-nam').innerText = db[i].name;
    renderT();
}

function renderT() {
    const l = document.getElementById('list-t'); l.innerHTML = "";
    db[cur].tools.forEach((t, i) => {
        l.innerHTML += `<div class="card" onclick="editT(${i})">
            <div class="c-drag">☰</div>
            <div class="c-id">${t.id}</div>
            <div class="c-name">${t.nm}</div>
            <div class="c-diam">${t.dia || '-'}</div>
            <button class="c-del" onclick="event.stopPropagation();delT(${i})">✕</button>
        </div>`;
    });
    new Sortable(l, { handle:'.c-drag', animation:150, onEnd: () => {
        const items = Array.from(l.querySelectorAll('.card'));
        db[cur].tools = items.map(el => {
            const tid = el.querySelector('.c-id').innerText;
            return db[cur].tools.find(x => x.id === tid);
        });
        save();
    }});
}

// ГЕНЕРАЦИЯ ПЕЧАТИ (СТИЛЬ СКРИНШОТА)
function printPDF() {
    const p = db[cur];
    const rows = p.tools.map(t => `
        <tr>
            <td width="15%">${t.id}</td>
            <td width="75%">${t.nm}</td>
            <td width="10%" align="right">${t.dia||'-'}</td>
        </tr>`).join('');

    document.getElementById('print-area').innerHTML = `
        <div class="print-frame">
            <p class="p-label">${p.name}</p>
            <h1 class="p-title">${p.num}</h1>
            <div class="p-hr"></div>
            <table>
                <thead>
                    <tr><th>T-NR</th><th>WERKZEUGNAME / KOMMENTAR</th><th align="right">Ø</th></tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>
        </div>`;

    window.print();
}

function addP() {
    const num = document.getElementById('p-num').value;
    const nam = document.getElementById('p-nam').value;
    if(!num) return;
    db.push({num, name:nam, tools:[]}); save(); renderP(); hideM('m-p');
    document.getElementById('p-num').value=""; document.getElementById('p-nam').value="";
}

function addT() {
    const idx = document.getElementById('t-idx').value;
    const t = { id: document.getElementById('t-id').value, nm: document.getElementById('t-nm').value, dia: document.getElementById('t-dia').value };
    if(!t.id) return;
    if(idx==="") db[cur].tools.push(t); else db[cur].tools[idx]=t;
    save(); renderT(); hideM('m-t');
}

function editT(i) {
    const t = db[cur].tools[i];
    document.getElementById('t-idx').value = i;
    document.getElementById('t-id').value = t.id;
    document.getElementById('t-nm').value = t.nm;
    document.getElementById('t-dia').value = t.dia;
    showM('m-t');
}

function delP(i) { if(confirm('Löschen?')) {db.splice(i,1); save(); renderP(); }}
function delT(i) { db[cur].tools.splice(i,1); save(); renderT(); }

function doImp() {
    const lines = document.getElementById('i-txt').value.split('\n');
    let cid = null, cnm = [];
    lines.forEach(l => {
        l = l.trim(); if(!l) return;
        if(/^[T][0-9]+/i.test(l)) {
            if(cid) db[cur].tools.push({id:cid, nm:cnm.join(' ').toUpperCase(), dia:''});
            cid = l.toUpperCase(); cnm = [];
        } else if(cid) cnm.push(l);
    });
    if(cid) db[cur].tools.push({id:cid, nm:cnm.join(' ').toUpperCase(), dia:''});
    save(); renderT(); hideM('m-i');
}

function goHome() {
    document.getElementById('v-det').classList.remove('active');
    document.getElementById('v-home').classList.add('active');
    renderP();
}
renderP();
