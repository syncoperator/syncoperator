let db = JSON.parse(localStorage.getItem('qs_v16')) || [];
let cur = null;

const showM = id => document.getElementById(id).style.display = 'flex';
const hideM = id => document.getElementById(id).style.display = 'none';

function save() { localStorage.setItem('qs_v16', JSON.stringify(db)); }

function goHome() {
    document.getElementById('v-det').classList.remove('active');
    document.getElementById('v-home').classList.add('active');
    renderP();
}

function renderP() {
    const l = document.getElementById('list-p'); l.innerHTML = "";
    db.forEach((p, i) => {
        l.innerHTML += `<div class="card" onclick="openP(${i})">
            <div style="flex-grow:1"><b>${p.num}</b><br><small>${p.name}</small></div>
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
            <div class="c-diam">${t.dia || ''}</div>
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

function addP() {
    const num = document.getElementById('p-num').value;
    const nam = document.getElementById('p-nam').value;
    if(!num) return;
    db.push({num, name:nam, tools:[]}); save(); renderP(); hideM('m-p');
}

function openTAdd() { document.getElementById('t-idx').value=""; document.getElementById('t-id').value=""; document.getElementById('t-nm').value=""; document.getElementById('t-dia').value=""; showM('m-t'); }

function addT() {
    const idx = document.getElementById('t-idx').value;
    const t = { id: document.getElementById('t-id').value, nm: document.getElementById('t-nm').value, dia: document.getElementById('t-dia').value };
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
function clearT() { if(confirm('Leeren?')) {db[cur].tools=[]; save(); renderT(); }}

function doImp() {
    const lines = document.getElementById('i-txt').value.split('\n');
    let cid = null, cnm = [];
    lines.forEach(l => {
        l = l.trim(); if(!l) return;
        if(/^[T][0-9]+/i.test(l)) {
            if(cid) db[cur].tools.push({id:cid, nm:cnm.join(' '), dia:''});
            cid = l.toUpperCase(); cnm = [];
        } else if(cid) cnm.push(l);
    });
    if(cid) db[cur].tools.push({id:cid, nm:cnm.join(' '), dia:''});
    save(); renderT(); hideM('m-i');
}

function getPDF() {
    const p = db[cur];
    const rows = p.tools.map(t => `<tr><td>${t.id}</td><td>${t.nm}</td><td style="text-align:right">${t.dia||''}</td></tr>`).join('');
    return `<div class="pdf-box"><div class="pdf-h"><div class="pdf-title">${p.name}</div><div class="pdf-num">${p.num}</div></div>
    <table class="pdf-table"><thead><tr><th width="15%">T-NR</th><th width="70%">BEZEICHNUNG</th><th width="15%" style="text-align:right">Ø</th></tr></thead>
    <tbody>${rows}</tbody></table></div>`;
}

function openPre() { document.getElementById('a4').innerHTML = getPDF(); showM('m-pre'); }

async function makePDF() {
    const b = document.getElementById('d-btn'); b.innerText = "WAIT...";
    const el = document.createElement('div');
    el.style.width = "210mm"; el.style.background = "white";
    el.innerHTML = `<div style="padding:15mm">${getPDF()}</div>`;
    document.body.appendChild(el);
    
    const opt = { margin:0, filename:`${db[cur].num}.pdf`, image:{type:'jpeg',quality:1}, html2canvas:{scale:2}, jsPDF:{unit:'mm',format:'a4'} };
    
    try {
        await html2pdf().set(opt).from(el).save();
    } catch(e) { alert('Err'); }
    finally { document.body.removeChild(el); b.innerText = "PDF SPEICHERN"; }
}

renderP();
