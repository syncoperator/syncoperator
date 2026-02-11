const DB_KEY = 'QS_DATA_V8';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];
let currentIdx = null;
let startIdx = null;

const el = (id) => document.getElementById(id);
const show = (id) => { if(el(id)) el(id).style.display = 'block'; };
const hide = (id) => { if(el(id)) el(id).style.display = 'none'; };

// Главная страница
function renderList() {
    const list = el('list-p'); 
    if(!list) return;
    list.innerHTML = db.map((p, i) => `
        <div class="list-item" onclick="openProject(${i})">
            <div class="t-label">PROJEKT</div>
            <div class="t-name">${p.num || '---'}</div>
            <div style="font-size:12px; opacity:0.6;">${p.name || ''}</div>
            <div style="position:absolute; right:15px; top:15px; color:red;" onclick="event.stopPropagation(); deleteProject(${i})">✕</div>
        </div>`).join('') + '<div style="height:100px"></div>';
}

// Рендер инструментов (Карточки T-01 / T-02 со скрина)
function renderTools() {
    const list = el('list-t'); 
    if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'tool-card';
        item.setAttribute('data-idx', i);
        
        const dotColor = t.rev ? '#007aff' : '#62cc71'; // Синий или зеленый как на фото

        item.innerHTML = `
            <div class="dot" style="background:${dotColor}; box-shadow: 0 0 8px ${dotColor}"></div>
            <div onclick="modalT(${i})">
                <div class="t-label">${t.id || 'T-NR'}</div>
                <div class="t-name">${t.nm || '---'}</div>
                <div class="t-val">${t.dia || '0.00'}</div>
            </div>
        `;

        // Drag & Drop
        item.draggable = true;
        item.ondragstart = () => { startIdx = i; item.style.opacity = '0.4'; };
        item.ondragover = (e) => e.preventDefault();
        item.ondrop = () => { if(startIdx !== i) moveTool(startIdx, i); };
        item.ondragend = () => renderTools();
        
        list.appendChild(item);
    });
}

function moveTool(from, to) {
    const tools = db[currentIdx].tools;
    const item = tools.splice(from, 1)[0];
    tools.splice(to, 0, item);
    save(); renderTools();
}

function openProject(i) {
    currentIdx = i;
    el('v-home').style.display = 'none';
    el('v-det').style.display = 'block';
    el('h-num').innerText = db[i].num;
    renderTools();
}

function goHome() { 
    currentIdx = null; 
    el('v-home').style.display = 'block'; 
    el('v-det').style.display = 'none'; 
    renderList(); 
}

function save() { localStorage.setItem(DB_KEY, JSON.stringify(db)); }

// Функции кнопок (теперь точно нажимаются)
function modalP(edit = false) { 
    if(!edit) currentIdx = null;
    show('m-p'); 
}
function modalT(i = null) { 
    el('t-idx').value = i !== null ? i : ''; 
    show('m-t'); 
}

// Инициализация
window.onload = renderList;
