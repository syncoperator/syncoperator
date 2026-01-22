/**
 * SyncOperator PRO - Core Logic & UI Handler
 */

let state = {
    operations: {},
    slots: { "1": [null, null, null, null, null], "2": [null, null, null, null, null] }
};

const STORAGE_KEY = 'sync_pro_data';

function init() {
    loadState();
    setupEventListeners();
    renderAll();
}

function loadState() {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) state = JSON.parse(saved);
}

function saveState() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    renderAll();
}

// --- RENDERING ---

function renderAll() {
    renderSlots("1");
    renderSlots("2");
    renderLibrary();
    renderSetupSheet();
}

function renderSlots(channel) {
    const container = document.getElementById(`slots-${channel}`);
    container.innerHTML = '';

    state.slots[channel].forEach((opId, index) => {
        const op = state.operations[opId];
        const slotEl = document.createElement('div');
        slotEl.className = `slot ${!op ? 'empty' : ''}`;
        
        const lCode = `L${channel}1${String(index + 1).padStart(2, '0')}`;

        if (op) {
            slotEl.innerHTML = `
                <div class="slot-header">
                    <strong>${op.title}</strong>
                    <span class="l-code">${lCode}</span>
                </div>
                <div class="slot-footer">
                    <small style="color:var(--text-dim)">${op.toolNo} • ${op.category}</small>
                </div>
            `;
            slotEl.onclick = () => openModal(opId);
        } else {
            slotEl.innerHTML = `<span>+ Leer</span>`;
            slotEl.onclick = () => openModal(null, channel, index);
        }
        
        // Drag and Drop (Simple implementation for Pro feel)
        slotEl.draggable = !!op;
        slotEl.onsdragstart = (e) => {
            e.dataTransfer.setData('opId', opId);
            e.dataTransfer.setData('fromChan', channel);
            e.dataTransfer.setData('fromIdx', index);
        };

        container.appendChild(slotEl);
    });
}

function renderLibrary() {
    const list = document.getElementById('libraryList');
    const catF = document.getElementById('filterCategory').value;
    const spinF = document.getElementById('filterSpindle').value;
    list.innerHTML = '';

    Object.values(state.operations).forEach(op => {
        if (catF && op.category !== catF) return;
        if (spinF && op.spindle !== spinF) return;

        const el = document.createElement('div');
        el.className = 'op-card';
        el.innerHTML = `<strong>${op.title}</strong><br><small>${op.toolNo} | ${op.spindle}</small>`;
        el.onclick = () => openModal(op.id);
        list.appendChild(el);
    });
}

function renderSetupSheet() {
    const container = document.getElementById('setupBody');
    container.innerHTML = '';
    
    // Group by toolNo
    const tools = {};
    [1, 2].forEach(c => {
        state.slots[c].forEach(id => {
            if (id && state.operations[id]) {
                const op = state.operations[id];
                if (!tools[op.toolNo]) tools[op.toolNo] = { id: op.toolNo, name: op.toolName, ch1: [], ch2: [] };
                tools[op.toolNo][`ch${c}`].push(op.title);
            }
        });
    });

    Object.values(tools).forEach(t => {
        const card = document.createElement('div');
        card.className = 'setup-card';
        card.innerHTML = `
            <div style="display:flex; justify-content:space-between">
                <strong>${t.id}</strong>
                <small>${t.name || ''}</small>
            </div>
            <div class="setup-grid">
                <div class="setup-box"><strong>CH1:</strong> ${t.ch1.join(', ') || '-'}</div>
                <div class="setup-box"><strong>CH2:</strong> ${t.ch2.join(', ') || '-'}</div>
            </div>
        `;
        container.appendChild(card);
    });
}

// --- UI CONTROLS ---

function setupEventListeners() {
    // Navigation
    document.querySelectorAll('.nav-item').forEach(btn => {
        btn.onclick = () => {
            document.querySelectorAll('.nav-item, .view-section').forEach(el => el.classList.remove('active'));
            btn.classList.add('active');
            document.getElementById(btn.dataset.target).classList.add('active');
        };
    });

    // Modal
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
        if (activeContext.channel) state.slots[activeContext.channel][activeContext.index] = id;
        closeModal();
        saveState();
    };

    document.getElementById('closeModalBtn').onclick = closeModal;
    document.getElementById('addNewOpBtn').onclick = () => openModal();
}

let activeContext = { channel: null, index: null };

function openModal(opId = null, channel = null, index = null) {
    activeContext = { channel, index };
    const form = document.getElementById('opForm');
    form.reset();
    document.getElementById('deleteOpBtn').style.display = opId ? 'block' : 'none';
    
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
    }
    document.getElementById('opModal').style.display = 'flex';
}

function closeModal() { document.getElementById('opModal').style.display = 'none'; }

// Init call
init();
