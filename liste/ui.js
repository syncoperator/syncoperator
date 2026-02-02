let currentProjectIdx = null;

function showM(id) { document.getElementById(id).style.display = 'flex'; }
function hideM(id) { document.getElementById(id).style.display = 'none'; }

function goHome() {
    currentProjectIdx = null;
    document.getElementById('v-det').classList.remove('active');
    document.getElementById('v-home').classList.add('active');
    renderProjects();
}

function renderProjects() {
    const cont = document.getElementById('list-projects');
    cont.innerHTML = "";
    db.forEach((p, i) => {
        const d = document.createElement('div');
        d.className = 'project-card';
        d.onclick = () => openProject(i);
        d.innerHTML = `<div style="flex:1"><b>${p.num}</b><br><small>${p.name}</small></div>
        <button class="btn-danger" style="width:auto" onclick="event.stopPropagation();deleteProject(${i})">LÖSCHEN</button>`;
        cont.appendChild(d);
    });
}

function openProject(i) {
    currentProjectIdx = i;
    document.getElementById('v-home').classList.remove('active');
    document.getElementById('v-det').classList.add('active');
    document.getElementById('h-num').innerText = db[i].num;
    document.getElementById('h-nam').innerText = db[i].name;
    renderTools();
}

function renderTools() {
    const cont = document.getElementById('list-tools');
    cont.innerHTML = "";
    db[currentProjectIdx].tools.forEach((t, i) => {
        const d = document.createElement('div');
        d.className = 'tool-card';
        d.dataset.id = i;
        d.innerHTML = `<div class="drag-handle">☰</div>
        <div class="tool-main" onclick="modalTool(${i})"><b>${t.id}</b><br><small>${t.nm}</small></div>
        <div class="tool-dia">${t.dia || ''}</div>`;
        cont.appendChild(d);
    });
    new Sortable(cont, { handle: '.drag-handle', animation: 150, onEnd: () => {
        let n = [];
        cont.querySelectorAll('.tool-card').forEach(el => n.push(db[currentProjectIdx].tools[el.dataset.id]));
        db[currentProjectIdx].tools = n; saveDB();
    }});
}

function modalProject(edit = false) {
    const p = edit ? db[currentProjectIdx] : {num:'',name:'',abs:'',grf:'',lzf:'',sta:'',stb:''};
    document.getElementById('p-idx').value = edit ? currentProjectIdx : "";
    document.getElementById('p-num').value = p.num;
    document.getElementById('p-nam').value = p.name;
    document.getElementById('p-abs').value = p.abs || '';
    document.getElementById('p-grf').value = p.grf || '';
    document.getElementById('p-lzf').value = p.lzf || '';
    document.getElementById('p-sta').value = p.sta || '';
    document.getElementById('p-stb').value = p.stb || '';
    showM('m-p');
}

function saveProject() {
    const i = document.getElementById('p-idx').value;
    const data = {
        num: document.getElementById('p-num').value,
        name: document.getElementById('p-nam').value.toUpperCase(),
        abs: document.getElementById('p-abs').value,
        grf: document.getElementById('p-grf').value,
        lzf: document.getElementById('p-lzf').value,
        sta: document.getElementById('p-sta').value,
        stb: document.getElementById('p-stb').value,
        tools: i === "" ? [] : db[i].tools
    };
    if(i==="") db.push(data); else db[i] = data;
    saveDB(); hideM('m-p'); goHome();
}

function modalTool(i=null) {
    const edit = i !== null;
    document.getElementById('t-idx').value = edit ? i : "";
    document.getElementById('t-id').value = edit ? db[currentProjectIdx].tools[i].id : "";
    document.getElementById('t-nm').value = edit ? db[currentProjectIdx].tools[i].nm : "";
    document.getElementById('t-dia').value = edit ? db[currentProjectIdx].tools[i].dia : "";
    showM('m-t');
}

function saveTool() {
    const i = document.getElementById('t-idx').value;
    const t = { id: document.getElementById('t-id').value.toUpperCase(), nm: document.getElementById('t-nm').value.toUpperCase(), dia: document.getElementById('t-dia').value };
    if(i==="") db[currentProjectIdx].tools.push(t); else db[currentProjectIdx].tools[i] = t;
    saveDB(); hideM('m-t'); renderTools();
}

function deleteProject(i) { db.splice(i,1); saveDB(); renderProjects(); }
function deleteTool() { db[currentProjectIdx].tools.splice(document.getElementById('t-idx').value, 1); saveDB(); hideM('m-t'); renderTools(); }
function runMassImport() { logicMassImport(currentProjectIdx, document.getElementById('imp-area').value); hideM('m-imp'); renderTools(); }

function makePDF() {
    const p = db[currentProjectIdx];
    const rows = p.tools.map(t => `
        <tr style="border-bottom:1.5pt solid black;">
            <td style="padding:12px 5px; font-weight:900; width:60px;">${t.id}</td>
            <td style="padding:12px 30px; font-weight:700;">${t.nm}</td>
            <td style="padding:12px 5px; font-weight:900; text-align:right; width:100px;">${t.dia||''}</td>
        </tr>`).join('');

    document.getElementById('pdf-render').innerHTML = `
    <div style="width:210mm; height:297mm; padding:12mm; background:white; color:black; font-family:sans-serif; box-sizing:border-box;">
        <div style="border:1.5pt solid black; height:100%; padding:30px; display:flex; flex-direction:column; box-sizing:border-box;">
            <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                <div>
                    <div style="font-size:10pt; font-weight:bold; color:#666;">${p.name}</div>
                    <div style="font-size:50pt; font-weight:900; line-height:1;">${p.num}</div>
                </div>
                <div style="text-align:right; font-size:10pt; font-weight:800; line-height:1.8;">
                    <div>ABSTAND: <span style="border-bottom:1pt solid black; min-width:80px; display:inline-block">${p.abs||''}</span></div>
                    <div>GREIFBACKEN: <span style="border-bottom:1pt solid black; min-width:80px; display:inline-block">${p.grf||''}</span></div>
                    <div>LAUFZEIT: <span style="border-bottom:1pt solid black; min-width:80px; display:inline-block">${p.lzf||''}</span></div>
                    <div>STÜCK A: <span style="border-bottom:1pt solid black; min-width:40px; display:inline-block">${p.sta||''}</span> | B: <span style="border-bottom:1pt solid black; min-width:40px; display:inline-block">${p.stb||''}</span></div>
                </div>
            </div>
            <div style="height:4pt; background:black; margin:20px 0;"></div>
            <table style="width:100%; border-collapse:collapse; table-layout:fixed;">
                <thead>
                    <tr style="border-bottom:2.5pt solid black; font-size:8pt; font-weight:900;">
                        <th align="left" style="width:60px; padding-bottom:5px;">T-NR</th>
                        <th align="left" style="padding-bottom:5px; padding-left:30px;">WERKZEUGNAME / KOMMENTAR</th>
                        <th align="right" style="width:100px; padding-bottom:5px;">Ø / TOLERANZ</th>
                    </tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>
        </div>
    </div>`;
    window.print();
}

renderProjects();
