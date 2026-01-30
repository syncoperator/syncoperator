let db = JSON.parse(localStorage.getItem('qs_v15_fixed')) || [];
let curP = null;

// UI FUNCTIONS
function goHome() {
    document.getElementById('view-detail').classList.remove('active-view');
    document.getElementById('view-home').classList.add('active-view');
    renderHome();
}
function openModal(id) { document.getElementById(id).style.display = 'flex'; }
function closeModal(id) { document.getElementById(id).style.display = 'none'; }
function save() { localStorage.setItem('qs_v15_fixed', JSON.stringify(db)); }

// RENDER HOME
function renderHome() {
    const list = document.getElementById('project-list');
    list.innerHTML = "";
    db.forEach((p, i) => {
        list.innerHTML += `
        <div class="tool-card" onclick="openDetail(${i})">
            <div style="flex-grow:1;">
                <div style="font-weight:900; font-size:18px;">${p.num}</div>
                <div style="color:#777; font-size:14px;">${p.name}</div>
            </div>
            <button onclick="event.stopPropagation(); deleteProject(${i})" class="col-del" style="text-align:center;">✕</button>
        </div>`;
    });
}

// RENDER TOOLS
function openDetail(i) {
    curP = i;
    document.getElementById('view-home').classList.remove('active-view');
    document.getElementById('view-detail').classList.add('active-view');
    document.getElementById('det-num').innerText = db[i].num;
    document.getElementById('det-name').innerText = db[i].name;
    renderTools();
}

function renderTools() {
    const list = document.getElementById('tool-list');
    list.innerHTML = "";
    db[curP].tools.forEach((t, i) => {
        // Мы используем классы col-..., которые прописаны в CSS Flexbox
        list.innerHTML += `
        <div class="tool-card" onclick="editTool(${i})">
            <div class="col-drag">☰</div>
            <div class="col-id">${t.id}</div>
            <div class="col-name">${t.nm}</div>
            <div class="col-diam">${t.diam || ''}</div>
            <button onclick="event.stopPropagation(); deleteTool(${i})" class="col-del">✕</button>
        </div>`;
    });

    // Сортировка
    new Sortable(list, {
        handle: '.col-drag',
        animation: 150,
        onEnd: () => {
            const rows = Array.from(list.querySelectorAll('.tool-card'));
            const newTools = rows.map(row => {
                const id = row.querySelector('.col-id').innerText;
                // Находим оригинал по ID (упрощенно)
                return db[curP].tools.find(t => t.id === id); 
            });
            db[curP].tools = newTools;
            save();
        }
    });
}

// LOGIC
function saveProject() {
    const num = document.getElementById('p-in-num').value;
    const name = document.getElementById('p-in-name').value;
    if(num) {
        db.push({ num, name, tools: [] });
        save(); renderHome(); closeModal('modal-p');
        document.getElementById('p-in-num').value = "";
        document.getElementById('p-in-name').value = "";
    }
}

function saveTool() {
    const idx = document.getElementById('t-idx').value;
    const id = document.getElementById('t-in-id').value;
    const nm = document.getElementById('t-in-nm').value;
    const diam = document.getElementById('t-in-diam').value;
    const obj = { id, nm, diam };
    
    if(idx === "") db[curP].tools.push(obj);
    else db[curP].tools[idx] = obj;
    
    save(); renderTools(); closeModal('modal-t');
}

function editTool(i) {
    const t = db[curP].tools[i];
    document.getElementById('t-idx').value = i;
    document.getElementById('t-in-id').value = t.id;
    document.getElementById('t-in-nm').value = t.nm;
    document.getElementById('t-in-diam').value = t.diam;
    openModal('modal-t');
}

function openTModal() {
    document.getElementById('t-idx').value = "";
    document.getElementById('t-in-id').value = "";
    document.getElementById('t-in-nm').value = "";
    document.getElementById('t-in-diam').value = "";
    openModal('modal-t');
}

function runImport() {
    const text = document.getElementById('import-area').value;
    const lines = text.split('\n');
    let cid = null, cnm = [];
    lines.forEach(l => {
        l = l.trim();
        if(!l) return;
        if(/^[T][0-9]+/i.test(l)) {
            if(cid) db[curP].tools.push({ id: cid, nm: cnm.join(' '), diam: '' });
            cid = l.toUpperCase(); cnm = [];
        } else if(cid) { cnm.push(l); }
    });
    if(cid) db[curP].tools.push({ id: cid, nm: cnm.join(' '), diam: '' });
    save(); renderTools(); closeModal('modal-import');
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); save(); renderHome(); } }
function deleteTool(i) { db[curP].tools.splice(i, 1); save(); renderTools(); }
function clearTools() { if(confirm('Leeren?')) { db[curP].tools = []; save(); renderTools(); } }

// PDF ENGINE (С методом Overlay от белого листа)
function getPDFHTML() {
    const p = db[curP];
    let rows = p.tools.map(t => `<tr><th>${t.id}</th><td>${t.nm}</td><td style="text-align:right">${t.diam||''}</td></tr>`).join('');
    return `
    <div class="pdf-border">
        <div class="pdf-head">
            <div class="pdf-title">${p.name}</div>
            <div class="pdf-num">${p.num}</div>
        </div>
        <table class="pdf-table">
            <thead><tr><th width="15%">T-NR</th><th width="70%">WERKZEUGNAME</th><th width="15%" align="right">Ø</th></tr></thead>
            <tbody>${rows}</tbody>
        </table>
    </div>`;
}

function openPreview() {
    document.getElementById('preview-stage').innerHTML = getPDFHTML();
    openModal('modal-prev');
}

async function downloadPDF() {
    const btn = document.getElementById('download-btn');
    btn.innerText = "Processing...";
    
    // Создаем временный слой поверх всего, чтобы браузер его "видел"
    const layer = document.createElement('div');
    layer.style.position = 'fixed'; 
    layer.style.top = '0'; layer.style.left = '0'; 
    layer.style.zIndex = '9999'; layer.style.background = 'white';
    // Вставляем контент шириной A4 (210мм)
    layer.innerHTML = `<div style="width:210mm; padding:0;">${getPDFHTML()}</div>`;
    document.body.appendChild(layer);

    const opt = {
        margin: 0,
        filename: `Setup_${db[curP].num}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true, logging: false },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    try {
        await new Promise(r => setTimeout(r, 100)); // Даем браузеру отрисовать
        await html2pdf().set(opt).from(layer.firstChild).save();
    } catch(e) { alert("Error"); }
    finally {
        document.body.removeChild(layer);
        btn.innerText = "PDF SPEICHERN";
    }
}

renderHome();
