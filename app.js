const App = {
    state: JSON.parse(localStorage.getItem('sync_v21')) || {
        slots: { "1": Array(10).fill(null), "2": Array(10).fill(null) },
        currentChannel: "1"
    },
    currentIdx: null,

    init() {
        this.renderUI();
    },

    switchChannel(ch) {
        this.state.currentChannel = ch;
        this.renderUI();
    },

    renderUI() {
        const ch = this.state.currentChannel;
        const root = document.getElementById('ui-root');
        
        // Обновляем визуальное состояние вкладок
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active', 'border-b-4', 'border-blue-600', 'bg-white'));
        document.getElementById('tab-' + ch).classList.add('bg-white', 'border-b-4', 'border-blue-600');

        root.innerHTML = this.state.slots[ch].map((s, i) => `
            <div class="slot-card ${!s ? 'empty' : ''}" onclick="App.openModal(${i})">
                <div class="cell-content">
                    <span style="font-size:10px; color:#94a3b8; font-weight:bold;">${ch === '1' ? 'L11' : 'L21'}${i + 1}</span>
                    <span class="cell-title" style="font-size:14px;">${s ? s.title : '--- LEER ---'}</span>
                    ${s ? `<span class="cell-sub">${s.tool} | ${s.spindle}</span>` : ''}
                </div>
                ${s ? `<button onclick="event.stopPropagation(); App.deleteSlot(${i})" style="color:#ef4444; font-weight:bold; padding:10px;">✕</button>` : ''}
            </div>
        `).join('');
    },

    openModal(i) {
        this.currentIdx = i;
        const s = this.state.slots[this.state.currentChannel][i];
        document.getElementById('in-title').value = s ? s.title : "";
        document.getElementById('in-tool').value = s ? s.tool : "";
        document.getElementById('in-tool-name').value = s ? s.toolName : "";
        document.getElementById('modal').classList.remove('hidden');
    },

    saveSlot() {
        this.state.slots[this.state.currentChannel][this.currentIdx] = {
            title: document.getElementById('in-title').value || "Operation",
            tool: document.getElementById('in-tool').value || "T0",
            spindle: document.getElementById('in-spindle').value,
            toolName: document.getElementById('in-tool-name').value || "Standard Tool"
        };
        this.save();
        this.closeModal();
    },

    deleteSlot(i) { this.state.slots[this.state.currentChannel][i] = null; this.save(); },
    save() { localStorage.setItem('sync_v21', JSON.stringify(this.state)); this.renderUI(); },
    closeModal() { document.getElementById('modal').classList.add('hidden'); },

    exportPDF(type) {
        const buf = document.getElementById('pdf-buffer');
        buf.innerHTML = `<div class="pdf-page">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; border-bottom:3px solid black; padding-bottom:5px;">
                <h1 style="margin:0; text-transform:uppercase; font-size:18px;">${type === 'plan' ? 'Operation Plan' : 'Werkzeugliste'}</h1>
                <span style="font-weight:bold; font-size:12px;">SyncOp X | DMG MORI</span>
            </div>
            ${type === 'plan' ? TableRenderer.renderOperationPlan(this.state.slots["1"], this.state.slots["2"]) : TableRenderer.renderToolList(this.state.slots["1"], this.state.slots["2"])}
        </div>`;

        html2pdf().from(buf).set({
            margin: 0,
            filename: `SyncOp_${type}.pdf`,
            html2canvas: { scale: 2 },
            jsPDF: { orientation: 'landscape', unit: 'mm', format: 'a4' }
        }).save();
    }
};

App.init();
