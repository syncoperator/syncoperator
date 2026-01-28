const App = {
    state: JSON.parse(localStorage.getItem('sync_v20')) || {
        slots: { "1": Array(8).fill(null), "2": Array(8).fill(null) },
        currentChannel: "1"
    },
    currentIdx: null,

    init() { this.renderUI(); },

    renderUI() {
        const ch = this.state.currentChannel;
        const root = document.getElementById('ui-root');
        root.innerHTML = this.state.slots[ch].map((s, i) => `
            <div class="card" onclick="App.openModal(${i})">
                <div>
                    <div style="font-size:9px; color:#999; font-family:monospace">${ch==='1'?'L11':'L21'}${i+1}</div>
                    <div style="font-weight:900; text-transform:uppercase">${s ? s.title : '---'}</div>
                    ${s ? `<div style="font-size:10px; color:var(--dmg-blue)">${s.tool} | ${s.spindle}</div>` : ''}
                </div>
                <button onclick="event.stopPropagation(); App.deleteSlot(${i})" style="opacity:0.3 hover:opacity:1">✖</button>
            </div>
        `).join('');
    },

    openModal(i) {
        this.currentIdx = i;
        const s = this.state.slots[this.state.currentChannel][i];
        document.getElementById('in-title').value = s ? s.title : "";
        document.getElementById('in-tool').value = s ? s.tool : "";
        document.getElementById('modal').classList.remove('hidden');
    },

    saveSlot() {
        this.state.slots[this.state.currentChannel][this.currentIdx] = {
            title: document.getElementById('in-title').value,
            tool: document.getElementById('in-tool').value,
            spindle: document.getElementById('in-spindle').value,
            toolName: "Standard Tool"
        };
        this.save();
        this.closeModal();
    },

    deleteSlot(i) { this.state.slots[this.state.currentChannel][i] = null; this.save(); },
    save() { localStorage.setItem('sync_v20', JSON.stringify(this.state)); this.renderUI(); },
    closeModal() { document.getElementById('modal').classList.add('hidden'); },
    switchChannel(ch) { this.state.currentChannel = ch; this.renderUI(); },

    exportPDF(type) {
        const buf = document.getElementById('pdf-buffer');
        buf.innerHTML = `<div class="pdf-page">
            <h2 style="text-transform:uppercase; margin-bottom:15px; border-bottom:2px solid #000">${type==='plan'?'Operation Plan':'Werkzeugliste'}</h2>
            ${type==='plan' ? TableRenderer.renderOperationPlan(this.state.slots["1"], this.state.slots["2"]) : TableRenderer.renderToolList(this.state.slots["1"], this.state.slots["2"])}
        </div>`;
        html2pdf().from(buf).set({ margin: 5, filename: `SyncOp_${type}.pdf`, jsPDF: { orientation: 'landscape', format: 'a4' } }).save();
    }
};
App.init();
