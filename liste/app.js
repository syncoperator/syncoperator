// --- STATE MANAGEMENT ---
let db = JSON.parse(localStorage.getItem('qs_tool_v15')) || [];
let curP = null;

// --- NAVIGATION & UI ---
function goHome() {
    document.getElementById('view-detail').classList.remove('active-view');
    document.getElementById('view-home').classList.add('active-view');
    renderHome();
}

function openModal(id) { document.getElementById(id).style.display = 'flex'; }
function closeModal(id) { document.getElementById(id).style.display = 'none'; }

// --- RENDERERS ---
function renderHome() {
    const list = document.getElementById('project-list');
    list.innerHTML = "";
    db.forEach((p, i) => {
        list.innerHTML += `
        <div class="tool-card" style="grid-template-columns: 1fr 40px;" onclick="openDetail(${i})">
            <div>
                <div style="font-size:18px; font-weight:900;">${p.num}</div>
                <div style="color:#888; font-size:13px;">${p.name}</div>
            </div>
            <button onclick="event.stopPropagation(); deleteProject(${i})" class="btn-del">✕</button>
        </div>`;
    });
}

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
        list.innerHTML += `
        <div class="tool-card" onclick="editTool(${i})">
            <div class="drag-handle">☰</div>
            <div class="t-id">${t.id}</div>
            <div class="t-name">${t.nm}</div>
            <div class="t-diam">${t.diam || ''}</div>
            <button onclick="event.stopPropagation(); deleteTool(${i})" class="btn-del">✕</button>
        </div>`;
    });

    // Initialize Sortable
    new Sortable(list, {
        handle: '.drag-handle',
        animation: 150,
        onEnd: function() {
            // Reorder array based on DOM
            const rows = Array.from(list.querySelectorAll('.tool-card'));
            const newTools = rows.map(row => {
                const id = row.querySelector('.t-id').innerText;
                // Find original object (simple match)
                return db[curP].tools.find(t => t.id === id);
            });
            db[curP].tools = newTools;
            save();
        }
    });
}

// --- DATA HANDLING ---
function save() { localStorage.setItem('qs_tool_v15', JSON.stringify(db)); }

function saveProject() {
    const num = document.getElementById('p-in-num').value;
    const name = document.getElementById('p-in-name').value;
    if(num) {
        db.push({ num, name, tools: [] });
        save();
        renderHome();
        closeModal('modal-p');
        // Clear inputs
        document.getElementById('p-in-num').value = "";
        document.getElementById('p-in-name').value = "";
    }
}

function saveTool() {
    const idx = document.getElementById('t-idx').value;
    const id = document.getElementById('t-in-id').value;
    const nm = document.getElementById('t-in-nm').value;
    const diam = document.getElementById('t-in-diam').value;

    const toolObj = { id, nm, diam };

    if(idx === "") {
        db[curP].tools.push(toolObj);
    } else {
        db[curP].tools[idx] = toolObj;
    }
    save();
    renderTools();
    closeModal('modal-t');
}

function editTool(i) {
    document.getElementById('t-idx').value = i;
    document.getElementById('t-in-id').value = db[curP].tools[i].id;
    document.getElementById('t-in-nm').value = db[curP].tools[i].nm;
    document.getElementById('t-in-diam').value = db[curP].tools[i].diam;
    openModal('modal-t');
}

function openTModal() {
    document.getElementById('t-idx').value = "";
    document.getElementById('t-in-id').value = "";
    document.getElementById('t-in-nm').value = "";
    document.getElementById('t-in-diam').value = "";
    openModal('modal-t');
}

function deleteProject(i) { if(confirm('Löschen?')) { db.splice(i, 1); save(); renderHome(); } }
function deleteTool(i) { db[curP].tools.splice(i, 1); save(); renderTools(); }
function clearTools() { if(confirm('Alle Tools löschen?')) { db[curP].tools = []; save(); renderTools(); } }

function runImport() {
    const text = document.getElementById('import-area').value;
    const lines = text.split('\n');
    let currentId = null;
    let currentName = [];

    lines.forEach(line => {
        const l = line.trim();
        if(!l) return;
        
        // Check for T-Number (e.g. T0101, T12, etc)
        if(/^[T][0-9]+/i.test(l)) {
            // Save previous if exists
            if(currentId) {
                db[curP].tools.push({ id: currentId, nm: currentName.join(' '), diam: '' });
            }
            // Start new
            currentId = l.toUpperCase();
            currentName = [];
        } else {
            // Append to name
            if(currentId) currentName.push(l);
        }
    });
    // Push last one
    if(currentId) {
        db[curP].tools.push({ id: currentId, nm: currentName.join(' '), diam: '' });
    }
    
    save();
    renderTools();
    closeModal('modal-import');
    document.getElementById('import-area').value = "";
}

// --- PDF GENERATION ENGINE ---

function getPDFHTML() {
    const p = db[curP];
    let rows = p.tools.map(t => `
        <tr>
            <th>${t.id}</th>
            <td>${t.nm}</td>
            <td class="center">${t.diam || ''}</td>
        </tr>
    `).join('');

    return `
    <div class="pdf-border">
        <div class="pdf-head">
            <div class="pdf-title">${p.name}</div>
            <div class="pdf-num">${p.num}</div>
        </div>
        <table class="pdf-table">
            <thead>
                <tr>
                    <th style="width:15%">T-NR</th>
                    <th style="width:70%">WERKZEUG</th>
                    <th style="width:15%; text-align:right">Ø</th>
                </tr>
            </thead>
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
    
    // 1. Create a temporary Overlay for printing
    // This is crucial for mobile browsers to render correctly
    const printLayer = document.createElement('div');
    printLayer.id = 'print-layer';
    printLayer.style.position = 'fixed';
    printLayer.style.top = '0';
    printLayer.style.left = '0';
    printLayer.style.zIndex = '9999';
    printLayer.style.background = 'white';
    
    // 2. Insert the PDF content (High Quality HTML)
    // We wrap it in a container 210mm wide
    printLayer.innerHTML = `
        <div style="width: 210mm; padding: 0; background: white;">
            ${getPDFHTML()}
        </div>
    `;
    
    document.body.appendChild(printLayer);
    
    // 3. Configure HTML2PDF
    const opt = {
        margin: 0,
        filename: `Setup_${db[curP].num}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true, logging: false },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    try {
        // Wait 100ms for DOM to paint
        await new Promise(resolve => setTimeout(resolve, 100));
        
        // Generate from the temporary layer
        await html2pdf().set(opt).from(printLayer.firstElementChild).save();
    } catch (e) {
        alert('Error generating PDF: ' + e.message);
    } finally {
        // Cleanup
        document.body.removeChild(printLayer);
        btn.innerText = "PDF SPEICHERN";
    }
}

// Initial Render
renderHome();
