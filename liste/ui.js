/**
 * UI CONTROLLER
 * Управление экранами, отрисовкой и событиями
 */

let currentProjectIdx = null;

// Навигация
function goHome() {
    currentProjectIdx = null;
    document.getElementById('v-det').classList.remove('active');
    document.getElementById('v-home').classList.add('active');
    renderProjects();
}

function showM(id) { document.getElementById(id).style.display = 'flex'; }
function hideM(id) { document.getElementById(id).style.display = 'none'; }

/**
 * Рендер списка проектов
 */
function renderProjects() {
    const container = document.getElementById('list-projects');
    container.innerHTML = "";
    
    db.forEach((project, index) => {
        const div = document.createElement('div');
        div.className = 'project-card';
        div.onclick = () => openProject(index);
        div.innerHTML = `
            <div style="flex:1">
                <b>${project.num}</b>
                <small style="color:#8E8E93">${project.name}</small>
            </div>
            <button class="btn-danger" style="padding:0" onclick="event.stopPropagation(); deleteProject(${index})">LÖSCHEN</button>
        `;
        container.appendChild(div);
    });
}

function openProject(index) {
    currentProjectIdx = index;
    document.getElementById('v-home').classList.remove('active');
    document.getElementById('v-det').classList.add('active');
    document.getElementById('h-num').innerText = db[index].num;
    document.getElementById('h-nam').innerText = db[index].name;
    renderTools();
}

/**
 * Рендер инструментов с поддержкой Sortable
 */
function renderTools() {
    const container = document.getElementById('list-tools');
    container.innerHTML = "";
    
    db[currentProjectIdx].tools.forEach((tool, index) => {
        const div = document.createElement('div');
        div.className = 'tool-card';
        div.dataset.id = index;
        div.innerHTML = `
            <div class="drag-handle">☰</div>
            <div class="tool-main" onclick="modalTool(${index})">
                <div class="tool-id">${tool.id}</div>
                <div class="tool-name">${tool.nm}</div>
            </div>
            <div class="tool-dia">${tool.dia || ''}</div>
        `;
        container.appendChild(div);
    });

    // Инициализация Drag-and-Drop
    new Sortable(container, {
        handle: '.drag-handle',
        animation: 200,
        ghostClass: 'sortable-ghost',
        onEnd: function() {
            let newOrder = [];
            container.querySelectorAll('.tool-card').forEach(el => {
                newOrder.push(db[currentProjectIdx].tools[el.dataset.id]);
            });
            db[currentProjectIdx].tools = newOrder;
            saveDB();
            renderTools();
        }
    });
}

/**
 * Работа с модальными окнами
 */
function modalProject(edit = false) {
    document.getElementById('p-idx').value = edit ? currentProjectIdx : "";
    document.getElementById('p-num').value = edit ? db[currentProjectIdx].num : "";
    document.getElementById('p-nam').value = edit ? db[currentProjectIdx].name : "";
    showM('m-p');
}

function saveProject() {
    const idx = document.getElementById('p-idx').value;
    const num = document.getElementById('p-num').value.trim();
    const nam = document.getElementById('p-nam').value.trim().toUpperCase();
    
    if (!num) return;
    
    if (idx === "") {
        db.push({ num, name: nam, tools: [] });
    } else {
        db[idx].num = num;
        db[idx].name = nam;
        document.getElementById('h-num').innerText = num;
        document.getElementById('h-nam').innerText = nam;
    }
    
    saveDB();
    renderProjects();
    hideM('m-p');
}

function modalTool(index = null) {
    const edit = index !== null;
    document.getElementById('t-idx').value = edit ? index : "";
    document.getElementById('t-smart').value = "";
    document.getElementById('t-id').value = edit ? db[currentProjectIdx].tools[index].id : "";
    document.getElementById('t-nm').value = edit ? db[currentProjectIdx].tools[index].nm : "";
    document.getElementById('t-dia').value = edit ? db[currentProjectIdx].tools[index].dia : "";
    document.getElementById('btn-del-t').style.display = edit ? 'block' : 'none';
    showM('m-t');
}

function runSmart() {
    const p = parseSmartString(document.getElementById('t-smart').value);
    if (p.id) document.getElementById('t-id').value = p.id;
    if (p.nm) document.getElementById('t-nm').value = p.nm;
    if (p.dia) document.getElementById('t-dia').value = p.dia;
}

function saveTool() {
    const idx = document.getElementById('t-idx').value;
    const tool = {
        id: document.getElementById('t-id').value.trim().toUpperCase(),
        nm: document.getElementById('t-nm').value.trim().toUpperCase(),
        dia: document.getElementById('t-dia').value.trim()
    };
    
    if (!tool.id) return;
    
    if (idx === "") db[currentProjectIdx].tools.push(tool);
    else db[currentProjectIdx].tools[idx] = tool;
    
    cleanProjectSlots(currentProjectIdx);
    renderTools();
    hideM('m-t');
}

function runMassImport() {
    const text = document.getElementById('imp-area').value;
    logicMassImport(currentProjectIdx, text);
    document.getElementById('imp-area').value = "";
    renderTools();
    hideM('m-imp');
}

function deleteTool() {
    const idx = document.getElementById('t-idx').value;
    db[currentProjectIdx].tools.splice(idx, 1);
    saveDB();
    renderTools();
    hideM('m-t');
}

function deleteProject(index) {
    db.splice(index, 1);
    saveDB();
    renderProjects();
}

/**
 * Генерация PDF
 */
function makePDF() {
    const p = db[currentProjectIdx];
    const rows = p.tools.map(t => `
        <tr>
            <td class="td-id">${t.id}</td>
            <td class="td-nm">${t.nm}</td>
            <td class="td-dia">${t.dia || ''}</td>
        </tr>
    `).join('');

    document.getElementById('pdf-render').innerHTML = `
        <div class="pdf-frame">
            <div class="pdf-meta">${p.name}</div>
            <div class="pdf-number">${p.num}</div>
            <div class="pdf-line"></div>
            <table>
                <thead>
                    <tr>
                        <th>T-NR</th>
                        <th>WERKZEUGNAME / KOMMENTAR</th>
                        <th align="right">Ø / TOLERANZ</th>
                    </tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>
        </div>
    `;
    window.print();
}

// Demo data
document.getElementById('btn-demo').onclick = () => {
    let demo = [];
    for(let i=1; i<=15; i++) demo.push({id:`T${i}`, nm:`BEISPIEL WERKZEUG ${i}`, dia: i==1?"10 -0.02":""});
    db.push({num:"233562", name:"MUSTERPOT", tools: demo});
    saveDB();
    renderProjects();
};

// Start
renderProjects();
