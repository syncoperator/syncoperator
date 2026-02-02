/**
 * UI CONTROLLER
 * Управление экранами и генерация PDF "1 в 1"
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

    new Sortable(container, {
        handle: '.drag-handle',
        animation: 200,
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
 * ФУНКЦИЯ ПЕЧАТИ: СОЗДАЕТ ОТДЕЛЬНЫЙ СТИЛЬ ДЛЯ PDF
 */
function makePDF() {
    const p = db[currentProjectIdx];
    const printZone = document.getElementById('pdf-render');
    
    // Генерируем строки таблицы
    const rows = p.tools.map(t => `
        <tr>
            <td style="width:15%; border-bottom:1.5pt solid black; padding:12px 5px;">${t.id}</td>
            <td style="width:65%; border-bottom:1.5pt solid black; padding:12px 5px; text-align:left; padding-left:30px;">${t.nm}</td>
            <td style="width:20%; border-bottom:1.5pt solid black; padding:12px 5px; text-align:right;">${t.dia || ''}</td>
        </tr>
    `).join('');

    // Вставляем контент с инлайновыми стилями (игнорируя внешний CSS)
    printZone.innerHTML = `
        <div style="
            width: 210mm; height: 297mm; padding: 15mm; 
            background: white; font-family: sans-serif; box-sizing: border-box;
        ">
            <div style="
                border: 2pt solid black; padding: 40px; height: 100%; 
                display: flex; flex-direction: column; box-sizing: border-box;
            ">
                <div style="font-size: 11pt; color: #777; text-transform: uppercase; font-weight: bold; margin-bottom: 5px;">
                    ${p.name}
                </div>
                <div style="font-size: 50pt; font-weight: 900; margin: 0; line-height: 1;">
                    ${p.num}
                </div>
                <div style="height: 4pt; background: black; margin: 20px 0;"></div>
                
                <table style="width: 100%; border-collapse: collapse; table-layout: fixed;">
                    <thead>
                        <tr>
                            <th style="text-align:left; font-size:8pt; padding:10px 5px; border-bottom:2.5pt solid black;">T-NR</th>
                            <th style="text-align:left; font-size:8pt; padding:10px 5px 10px 30px; border-bottom:2.5pt solid black;">WERKZEUGNAME / KOMMENTAR</th>
                            <th style="text-align:right; font-size:8pt; padding:10px 5px; border-bottom:2.5pt solid black;">Ø / TOLERANZ</th>
                        </tr>
                    </thead>
                    <tbody style="font-size: 11pt; font-weight: 800;">
                        ${rows}
                    </tbody>
                </table>
            </div>
        </div>
    `;

    window.print();
}

// Утилиты модалок (Project/Tool)
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
    if (idx === "") db.push({ num, name: nam, tools: [] });
    else { db[idx].num = num; db[idx].name = nam; }
    saveDB(); renderProjects(); hideM('m-p');
    if(currentProjectIdx !== null) openProject(currentProjectIdx);
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
    renderTools(); hideM('m-t');
}

function runMassImport() {
    logicMassImport(currentProjectIdx, document.getElementById('imp-area').value);
    document.getElementById('imp-area').value = "";
    renderTools(); hideM('m-imp');
}

function deleteTool() {
    db[currentProjectIdx].tools.splice(document.getElementById('t-idx').value, 1);
    saveDB(); renderTools(); hideM('m-t');
}

function deleteProject(index) {
    db.splice(index, 1);
    saveDB(); renderProjects();
}

// Запуск
renderProjects();
