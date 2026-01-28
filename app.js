const App = {
    state: JSON.parse(localStorage.getItem('sync_v18')) || {
        slots: { "1": Array(10).fill(null), "2": Array(10).fill(null) },
        currentChannel: "1"
    },

    init() {
        this.renderUI();
    },

    switchChannel(ch) {
        this.state.currentChannel = ch;
        this.renderUI();
    },

    renderUI() {
        const root = document.getElementById('ui-root');
        const ch = this.state.currentChannel;
        const slots = this.state.slots[ch];

        root.innerHTML = `
            <div style="display:flex; margin-bottom:20px; gap:10px;">
                <button onclick="App.switchChannel('1')" class="btn ${ch==='1'?'btn-accent':'btn-primary'}">Kanal 1</button>
                <button onclick="App.switchChannel('2')" class="btn ${ch==='2'?'btn-accent':'btn-primary'}">Kanal 2</button>
            </div>
            ${slots.map((s, i) => `
                <div class="slot-card ${!s?'empty':''}" onclick="App.editSlot(${i})">
                    <div>
                        <small style="color:#999">${ch==='1'?'L11':'L21'}${i+1}</small>
                        <div style="font-weight:bold">${s ? s.title : '--- LEER ---'}</div>
                        ${s ? `<small>${s.tool} | ${s.spindle}</small>` : ''}
                    </div>
                    ${s ? `<button onclick="event.stopPropagation(); App.deleteSlot(${i})" style="color:red; background:none; border:none; cursor:pointer;">✖</button>` : ''}
                </div>
            `).join('')}
        `;
    },

    editSlot(i) {
        const title = prompt("Operation Name:", this.state.slots[this.state.currentChannel][i]?.title || "");
        if (title === null) return;
        const tool = prompt("T-Nummer (z.B. T1):", "T1");
        const sp = prompt("Spindel (SP4 или SP3):", "SP4");
        
        this.state.slots[this.state.currentChannel][i] = { title, tool, spindle: sp, toolName: "Spezifikation..." };
        this.save();
    },

    deleteSlot(i) {
        this.state.slots[this.state.currentChannel][i] = null;
        this.save();
    },

    save() {
        localStorage.setItem('sync_v18', JSON.stringify(this.state));
        this.renderUI();
    },

    exportPDF(type) {
        const buf = document.getElementById('pdf-buffer');
        const html = (type === 'plan') 
            ? TableRenderer.renderOperationPlan(this.state.slots["1"], this.state.slots["2"])
            : TableRenderer.renderToolList(this.state.slots["1"], this.state.slots["2"]);

        buf.innerHTML = `<div class="pdf-page">
            <h2 style="text-transform:uppercase; margin-bottom:20px;">${type==='plan'?'Operation Plan':'Werkzeugliste'}</h2>
            ${html}
        </div>`;

        html2pdf().from(buf).set({
            margin: 5, filename: `SyncOp_${type}.pdf`,
            jsPDF: { orientation: 'landscape', unit: 'mm', format: 'a4' }
        }).save();
    }
};

App.init();
