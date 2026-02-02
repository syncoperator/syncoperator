let db = JSON.parse(localStorage.getItem('qs_central_v3')) || [];
let cur = null;

const showM = id => document.getElementById(id).style.display = 'flex';
const hideM = id => document.getElementById(id).style.display = 'none';
const saveDB = () => localStorage.setItem('qs_central_v3', JSON.stringify(db));

// --- ПРОЕКТЫ ---
function saveProject() {
    const idx = document.getElementById('p-idx').value;
    const num = document.getElementById('p-num').value;
    const nam = document.getElementById('p-nam').value.toUpperCase();
    if(!num) return;
    if(idx === "") db.push({num, name: nam, tools: []});
    else { db[idx].num = num; db[idx].name = nam; }
    saveDB(); renderP(); hideM('m-p');
    if(cur !== null) openP(cur); 
}

function renderP() {
    const l = document.getElementById('list-p'); l.innerHTML = "";
    db.forEach((p, i) => {
        l.innerHTML += `<div class="card" onclick="openP(${i})">
            <div style="flex-grow:1"><b>${p.num}</b><br><small style="color:#888">${p.name}</small></div>
            <button onclick="event.stopPropagation();delP(${i})" style="border:none;background:none;color:#FF3B30;font-size:20px">✕</button>
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

// --- ИНСТРУМЕНТЫ ---
function smartParse(val) {
    let res = {id: "", nm: "", dia: ""};
    const tMatch = val.match(/T\d+/i);
    if(tMatch) { res.id = tMatch[0].toUpperCase(); val = val.replace(tMatch[0], '').trim(); }
    const diaMatch = val.match(/(\d+[\.,]?\d*\s*[\+\-]\s*\d+[\.,]?\d*|\d+[\.,]?\d*)$/);
    if(diaMatch) { res.dia = diaMatch[0].replace(' ', '\n'); val = val.replace(diaMatch[0], '').trim(); }
    res.nm = val.toUpperCase();
    return res;
}

function doSmart() {
    const p = smartParse(document.getElementById('t-smart').value);
    document.getElementById('t-id').value = p.id;
    document.getElementById('t-dia').value = p.dia;
    document.getElementById('t-nm').value = p.nm;
}

function processMassImport() {
    const lines = document.getElementById('imp-area').value.split('\n');
    lines.forEach(line => {
        if(line.trim()) {
            const p = smartParse(line);
            if(p.id) db[cur].tools.push({id: p.id, nm: p.nm, dia: p.dia});
        }
    });
    saveDB(); renderT(); hideM('m-imp');
    document.getElementById('imp-area').value = "";
}

function renderT() {
    const l = document.getElementById('list-t'); l.innerHTML = "";
    db[cur].tools.forEach((t, i) => {
        const d = document.createElement('div');
        d.className = 'card'; d.dataset.id = i;
        d.onclick = () => editT(i);
        d.innerHTML = `<div class="c-id">${t.id}</div><div class="c-name">${t.nm}</div><div class="c-diam">${t.dia || ''}</div>`;
        l.appendChild(d);
    });
    new Sortable(l, { animation: 150, delay: 150, delayOnTouchOnly: true, ghostClass: 'sortable-ghost', onEnd: function() {
        const updated = [];
        l.querySelectorAll('.card').forEach(el => updated.push(db[cur].tools[el.dataset.id]));
        db[cur].tools = updated; saveDB(); renderT();
    }});
}

function saveTool() {
    const idx = document.getElementById('t-idx').value;
    const t = { id: document.getElementById('t-id').value.toUpperCase(), nm: document.getElementById('t-nm').value.toUpperCase(), dia: document.getElementById('t-dia').value };
    if(!t.id) return;
    if(idx === "") db[cur].tools.push(t); else db[cur].tools[idx] = t;
    saveDB(); renderT(); hideM('m-t');
}

// Стандартные навигационные функции
function openNewProject() { document.getElementById('p-idx').value = ""; document.getElementById('p-num').value = ""; document.getElementById('p-nam').value = ""; showM('m-p'); }
function editCurrentProject() { document.getElementById('p-idx').value = cur; document.getElementById('p-num').value = db[cur].num; document.getElementById('p-nam').value = db[cur].name; showM('m-p'); }
function openNewTool() { ['t-idx','t-smart','t-id','t-nm','t-dia'].forEach(id => document.getElementById(id).value = ""); document.getElementById('btn-del-t').style.display = 'none'; showM('m-t'); }
function editT(i) { const t = db[cur].tools[i]; document.getElementById('t-idx').value = i; document.getElementById('t-id').value = t.id; document.getElementById('t-nm').value = t.nm; document.getElementById('t-dia').value = t.dia; document.getElementById('btn-del-t').style.display = 'block'; showM('m-t'); }
function delTool() { const idx = document.getElementById('t-idx').value; if(confirm('Löschen?')) { db[cur].tools.splice(idx, 1); saveDB(); renderT(); hideM('m-t'); }}
function delP(i) { if(confirm('Projekt löschen?')) { db.splice(i,1); saveDB(); renderP(); }}
function goHome() { document.getElementById('v-det').classList.remove('active'); document.getElementById('v-home').classList.add('active'); }
function loadDemo() {
    const d = []; for(let i=1; i<=15; i++) d.push({id:`T${i}`, nm:`WERKZEUG NR ${i}`, dia: i==1 ? "29\n-0.1" : ""});
    db.push({num:"233562", name:"BUCHSE", tools: d}); saveDB(); renderP();
}
function printPDF() {
    const p = db[cur];
    const rows = p.tools.map(t => `<tr><td class="td-num">${t.id}</td><td class="td-name">${t.nm}</td><td class="td-dia">${t.dia || ''}</td></tr>`).join('');
    document.getElementById('print-area').innerHTML = `<div class="print-frame"><p class="p-label">${p.name}</p><h1 class="p-title">${p.num}</h1><div class="p-hr"></div><table><thead><tr><th>T-NR</th><th>WERKZEUGNAME</th><th align="right">Ø</th></tr></thead><tbody>${rows}</tbody></table></div>`;
    window.print();
}
renderP();
