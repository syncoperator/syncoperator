let db = JSON.parse(localStorage.getItem('qs_v27')) || [];
let cur = null;

const showM = id => document.getElementById(id).style.display = 'flex';
const hideM = id => document.getElementById(id).style.display = 'none';
const save = () => localStorage.setItem('qs_v27', JSON.stringify(db));

// Демо-данные
function loadDemo() {
    const demoTools = [];
    for(let i=1; i<=15; i++) {
        demoTools.push({
            id: `T01${String(i).padStart(2,'0')}`,
            nm: `WERKZEUGNAME NR ${i} KOMMENTAR G54`,
            dia: i === 1 ? "29\n-0.1" : ""
        });
    }
    db.push({num: "233562", name: "BUCHSE", tools: demoTools});
    save(); renderP();
}

function renderP() {
    const l = document.getElementById('list-p'); l.innerHTML = "";
    db.forEach((p, i) => {
        l.innerHTML += `<div class="card" onclick="openP(${i})">
            <div style="flex-grow:1"><b>${p.num}</b><br><small style="color:#888">${p.name}</small></div>
            <button onclick="event.stopPropagation();delP(${i})" style="border:none;background:none;color:red;font-size:18px">✕</button>
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
            <div class="c-id">${t.id}</div>
            <div class="c-name">${t.nm}</div>
            <div class="c-diam">${t.dia || ''}</div>
        </div>`;
    });
}

function printPDF() {
    const p = db[cur];
    const rows = p.tools.map(t => `
        <tr>
            <td class="td-num">${t.id}</td>
            <td class="td-name">${t.nm}</td>
            <td class="td-dia">${t.dia || ''}</td>
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

// Стандартные функции
function addP() {
    const num = document.getElementById('p-num').value;
    const nam = document.getElementById('p-nam').value;
    if(!num) return;
    db.push({num, name:nam.toUpperCase(), tools:[]}); save(); renderP(); hideM('m-p');
}

function addT() {
    const idx = document.getElementById('t-idx').value;
    const t = { 
        id: document.getElementById('t-id').value.toUpperCase(), 
        nm: document.getElementById('t-nm').value.toUpperCase(), 
        dia: document.getElementById('t-dia').value 
    };
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
function goHome() {
    document.getElementById('v-det').classList.remove('active');
    document.getElementById('v-home').classList.add('active');
    renderP();
}
renderP();
