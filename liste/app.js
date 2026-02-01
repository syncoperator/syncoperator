let db = JSON.parse(localStorage.getItem('qs_v21')) || [];
let cur = null;

const showM = id => document.getElementById(id).style.display = 'flex';
const hideM = id => document.getElementById(id).style.display = 'none';
const save = () => localStorage.setItem('qs_v21', JSON.stringify(db));

function renderP() {
    const l = document.getElementById('list-p'); l.innerHTML = "";
    db.forEach((p, i) => {
        l.innerHTML += `<div class="card" onclick="openP(${i})">
            <div style="flex-grow:1"><b>${p.num}</b><br><small>${p.name}</small></div>
            <button onclick="event.stopPropagation();delP(${i})" style="border:none;background:none;color:red">✕</button>
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
            <div class="c-diam">${t.dia || '-'}</div>
        </div>`;
    });
}

// --- СИСТЕМНАЯ ПЕЧАТЬ ---
function printPDF() {
    const p = db[cur];
    const rows = p.tools.map(t => `
        <tr>
            <td width="20%">${t.id}</td>
            <td width="70%">${t.nm}</td>
            <td width="10%" align="right">${t.dia||'-'}</td>
        </tr>`).join('');

    document.getElementById('print-area').innerHTML = `
        <div class="print-border">
            <p class="pdf-proj">${p.name}</p>
            <h1 class="pdf-title">${p.num}</h1>
            <div class="pdf-hr"></div>
            <table>
                <thead><tr><th>T-NR</th><th>BEZEICHNUNG</th><th align="right">Ø</th></tr></thead>
                <tbody>${rows}</tbody>
            </table>
        </div>`;

    window.print(); // Вызов нативного окна iOS
}

// (остальные функции addP, addT, doImp остаются без изменений)
function addP() {
    const num = document.getElementById('p-num').value;
    const nam = document.getElementById('p-nam').value;
    if(!num) return;
    db.push({num, name:nam, tools:[]}); save(); renderP(); hideM('m-p');
}

function addT() {
    const idx = document.getElementById('t-idx').value;
    const t = { id: document.getElementById('t-id').value, nm: document.getElementById('t-nm').value, dia: document.getElementById('t-dia').value };
    if(!t.id) return;
    if(idx==="") db[cur].tools.push(t); else db[cur].tools[idx]=t;
    save(); renderT(); hideM('m-t');
}

function delP(i) { if(confirm('Löschen?')) {db.splice(i,1); save(); renderP(); }}
function goHome() {
    document.getElementById('v-det').classList.remove('active');
    document.getElementById('v-home').classList.add('active');
    renderP();
}
renderP();
