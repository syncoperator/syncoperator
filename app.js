/**
 * SyncOperator CORE MODULE
 */

// --- STATE MANAGEMENT ---
let state = {
    operations: {}, // ID -> Object
    slots: {
        "1": [null, null, null, null, null],
        "2": [null, null, null, null, null]
    }
};

const STORAGE_KEY = 'sync_operator_data';

function loadState() {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
        state = JSON.parse(saved);
    }
}

function saveState() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    renderAll();
}

// --- LOGIC: L-CODE GENERATION ---
function calculateLCode(channel, index) {
    const xx = String(index + 1).padStart(2, '0');
    return `L${channel}1${xx}`;
}

// --- LOGIC: SETTING UP DATA ---
function getSetupSheet() {
    const tools = {};

    const processChannel = (chanId, key) => {
        state.slots[chanId].forEach(opId => {
            if (opId && state.operations[opId]) {
                const op = state.operations[opId];
                if (!tools[op.toolNo]) {
                    tools[op.toolNo] = { toolNo: op.toolNo, oben: "", unten: "" };
                }
                tools[op.toolNo][key] = op.title;
            }
        });
    };

    processChannel("1", "oben");
    processChannel("2", "unten");
    return Object.values(tools);
}

// --- UI RENDERING ---

function renderAll() {
    renderLibrary();
    renderSlots("1");
    renderSlots("2");
    renderSetupSheet();
}

function renderLibrary() {
    const list = document.getElementById('libraryList');
    const catFilter = document.getElementById('filterCategory').value;
    const spinFilter = document.getElementById('filterSpindle').value;
    
    list.innerHTML = '';
    
    Object.values(state.operations).forEach(op => {
        if (catFilter && op.category !== catFilter) return;
        if (spinFilter && op.spindle !== spinFilter) return;

        const el = document.createElement('div');
        el.className = 'op-card';
        el.draggable = true;
        el.textContent = `${op.title} (${op.toolNo})`;
        el.dataset.id = op.id;
        
        el.addEventListener('dragstart', (e) => {
            e.dataTransfer.setData('operationId', op.id);
            e.dataTransfer.setData('source', 'library');
        });

        el.onclick = () => openModal(op.id);
        list.appendChild(el);
    });
}

function renderSlots(channelId) {
    const container = document.getElementById(`slots-${channelId}`);
    container.innerHTML = '';

    state.slots[channelId].forEach((opId, index) => {
        const slotEl = document.createElement('div');
        slotEl.className = 'slot';
        slotEl.dataset.index = index;
        slotEl.dataset.channel = channelId;

        if (opId && state.operations[opId]) {
            const op = state.operations[opId];
            const lCode = calculateLCode(channelId, index);
            slotEl.innerHTML = `
                <div style="display:flex; justify-content:space-between">
                    <strong>${op.title}</strong>
                    <span>${lCode}</span>
                </div>
                <small>${op.toolNo}</small>
            `;
            slotEl.onclick = (e) => {
                e.stopPropagation();
                openModal(opId);
            };
            slotEl.draggable = true;
            slotEl.addEventListener('dragstart', (e) => {
                e.dataTransfer.setData('sourceIndex', index);
                e.dataTransfer.setData('sourceChannel', channelId);
                e.dataTransfer.setData('operationId', opId);
                e.dataTransfer.setData('source', 'slot');
            });
        } else {
            slotEl.classList.add('empty');
            slotEl.textContent = 'Leer';
            slotEl.onclick = () => openModal(null, channelId, index);
        }

        // Drop Logic
        slotEl.addEventListener('dragover', (e) => e.preventDefault());
        slotEl.addEventListener('drop', (e) => {
            e.preventDefault();
            const source = e.dataTransfer.getData('source');
            const opId = e.dataTransfer.getData('operationId');

            if (source === 'library') {
                state.slots[channelId][index] = opId;
            } else {
                const sIdx = e.dataTransfer.getData('sourceIndex');
                const sChan = e.dataTransfer.getData('sourceChannel');
                if (sChan === channelId) {
                    // Move in same channel
                    const temp = state.slots[channelId][index];
                    state.slots[channelId][index] = state.slots[channelId][sIdx];
                    state.slots[channelId][sIdx] = temp;
                } else {
                    // From other channel
                    state.slots[channelId][index] = opId;
                }
            }
            saveState();
        });

        container.appendChild(slotEl);
    });
}

function renderSetupSheet() {
    const body = document.getElementById('setupBody');
    body.innerHTML = '';
    const data = getSetupSheet();
    
    data.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${row.toolNo}</td><td>${row.oben}</td><td>${row.unten}</td>`;
        body.appendChild(tr);
    });
}

// --- MODAL LOGIC ---
const modal = document.getElementById('opModal');
let currentContext = null; // { channel, index }

function openModal(opId = null, channel = null, index = null) {
    currentContext = { channel, index };
    const form = document.getElementById('opForm');
    form.reset();
    
    if (opId) {
        const op = state.operations[opId];
        document.getElementById('field-id').value = op.id;
        document.getElementById('field-title').value = op.title;
        document.getElementById('field-code').value = op.code;
        document.getElementById('field-spindle').value = op.spindle;
        document.getElementById('field-category').value = op.category;
        document.getElementById('field-toolNo').value = op.toolNo;
        document.getElementById('field-toolName').value = op.toolName;
        document.getElementById('field-doppelhalter').checked = op.doppelhalter;
        document.getElementById('deleteOpBtn').style.display = 'block';
    } else {
        document.getElementById('field-id').value = '';
        document.getElementById('deleteOpBtn').style.display = 'none';
    }
    
    modal.style.display = 'flex';
}

document.getElementById('opForm').onsubmit = (e) => {
    e.preventDefault();
    const id = document.getElementById('field-id').value || crypto.randomUUID();
    const op = {
        id,
        title: document.getElementById('field-title').value,
        code: document.getElementById('field-code').value,
        spindle: document.getElementById('field-spindle').value,
        category: document.getElementById('field-category').value,
        toolNo: document.getElementById('field-toolNo').value,
        toolName: document.getElementById('field-toolName').value,
        doppelhalter: document.getElementById('field-doppelhalter').checked
    };

    state.operations[id] = op;
    
    if (currentContext.channel !== null) {
        state.slots[currentContext.channel][currentContext.index] = id;
    }
    
    modal.style.display = 'none';
    saveState();
};

document.getElementById('deleteOpBtn').onclick = () => {
    const id = document.getElementById('field-id').value;
    if (!id) return;
    
    delete state.operations[id];
    // Remove from slots
    ["1", "2"].forEach(c => {
        state.slots[c] = state.slots[c].map(slotId => slotId === id ? null : slotId);
    });
    
    modal.style.display = 'none';
    saveState();
};

document.getElementById('closeModalBtn').onclick = () => modal.style.display = 'none';

// --- SLOTS MANAGEMENT ---
document.querySelectorAll('.add-slot-btn').forEach(btn => {
    btn.onclick = () => {
        const c = btn.dataset.channel;
        state.slots[c].push(null);
        saveState();
    };
});

// --- IMPORT / EXPORT ---
document.getElementById('exportBtn').onclick = () => {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `syncoperator_backup_${new Date().getTime()}.json`;
    a.click();
};

document.getElementById('importBtn').onclick = () => document.getElementById('importInput').click();
document.getElementById('importInput').onchange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const imported = JSON.parse(e.target.result);
            if (imported.operations && imported.slots) {
                state = imported;
                saveState();
            }
        } catch (err) { alert("Ungültige Datei"); }
    };
    reader.readAsText(file);
};

// --- FILTERS ---
document.getElementById('filterCategory').onchange = renderLibrary;
document.getElementById('filterSpindle').onchange = renderLibrary;
document.getElementById('addNewOpBtn').onclick = () => openModal();

// --- INIT ---
loadState();
renderAll();
