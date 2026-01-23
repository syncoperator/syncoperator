// script.js
const App = {
    state: {
        currentKanal: 1,
        kanals: { 1: [], 2: [] },
        lib: [
            { name: "Bohrer D12", length: 145.5 },
            { name: "Fräser D25", length: 98.2 },
            { name: "Gewinde M10", length: 112.0 }
        ]
    },

    init() {
        this.renderLib();
        this.renderSlots();
        document.getElementById('report-date').innerText = new Date().toLocaleDateString();
    },

    switchKanal(n) {
        this.state.currentKanal = n;
        document.querySelectorAll('.k-btn').forEach((b, i) => b.classList.toggle('active', i+1 === n));
        this.renderSlots();
    },

    addSlot() {
        this.state.kanals[this.state.currentKanal].push({ ...this.state.lib[0] });
        this.renderSlots();
    },

    renderSlots() {
        const container = document.getElementById('slot-list');
        container.innerHTML = this.state.kanals[this.state.currentKanal].map((slot, i) => `
            <div class="slot-item">
                <div style="font-weight:800; color:var(--accent)">${i+1}</div>
                <div>
                    <b>${slot.name}</b><br>
                    <small>L-XXXX: <input type="number" value="${slot.length}" onchange="App.updateLen(${i}, this.value)" style="width:60px; border:none; background:#f0f0f0; border-radius:4px;"> mm</small>
                </div>
                <button onclick="App.removeSlot(${i})" style="border:none; background:none; color:red">✕</button>
            </div>
        `).join('');
        this.updateReport();
    },

    updateLen(index, val) {
        this.state.kanals[this.state.currentKanal][index].length = parseFloat(val);
        this.updateReport();
    },

    updateReport() {
        const body = document.getElementById('report-body');
        const allData = [...this.state.kanals[1], ...this.state.kanals[2]];
        body.innerHTML = allData.map((s, i) => `
            <tr>
                <td>${i+1}</td>
                <td><b>${s.name}</b></td>
                <td>H6 / BMT60</td>
                <td style="font-family:monospace; font-weight:700">${s.length.toFixed(3)}</td>
                <td><span style="color:green">● READY</span></td>
            </tr>
        `).join('');
    },

    toggleView() {
        const config = document.getElementById('config-view');
        const report = document.getElementById('report-view');
        const btn = document.getElementById('view-toggle');
        
        const isReport = config.classList.toggle('hidden');
        report.classList.toggle('hidden');
        btn.innerText = isReport ? "BACK TO CONFIG" : "EINRICHTEBLATT";
    },

    renderLib() {
        document.getElementById('tool-lib').innerHTML = this.state.lib.map(t => `
            <div class="lib-card" onclick="App.quickAdd('${t.name}', ${t.length})">
                <small>TOOL</small><br><b>${t.name}</b>
            </div>
        `).join('');
    },

    quickAdd(name, len) {
        this.state.kanals[this.state.currentKanal].push({ name, length: len });
        this.renderSlots();
    },

    exportPDF() { window.print(); }
};

App.init();
