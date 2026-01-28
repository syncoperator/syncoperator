const SyncOp = {
    state: JSON.parse(localStorage.getItem('sync_v15_state')) || {
        slots: { "1": Array(8).fill(null), "2": Array(8).fill(null) },
        currentChannel: "1"
    },

    save() {
        localStorage.setItem('sync_v15_state', JSON.stringify(this.state));
        this.renderUI();
    },

    // Рендер таблицы для PDF (Универсальный метод)
    generateTableHTML(type) {
        const slots1 = this.state.slots["1"];
        const slots2 = this.state.slots["2"];
        
        if (type === 'plan') {
            const max = Math.max(slots1.length, slots2.length);
            let rows = "";
            for (let i = 0; i < max; i++) {
                rows += `<tr>
                    <td style="width:30px; text-align:center; font-weight:bold; background:#eee;">${i+1}</td>
                    ${this.renderOpCell(slots1[i], 'SP4', 'L11'+(i+1))}
                    ${this.renderOpCell(slots1[i], 'SP3', 'L11'+(i+1))}
                    ${this.renderOpCell(slots2[i], 'SP3', 'L21'+(i+1))}
                    ${this.renderOpCell(slots2[i], 'SP4', 'L21'+(i+1))}
                </tr>`;
            }
            return `<table class="dmg-grid-table">
                <thead>
                    <tr><th style="width:30px">#</th><th colspan="2">Kanal 1</th><th colspan="2">Kanal 2</th></tr>
                    <tr><th></th><th>SP 4</th><th>SP 3</th><th>SP 3</th><th>SP 4</th></tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>`;
        } 
        
        if (type === 'tools') {
            const allTools = [...slots1, ...slots2].filter(s => s !== null);
            const sp4 = allTools.filter(t => t.spindle === 'SP4');
            const sp3 = allTools.filter(t => t.spindle === 'SP3');
            const max = Math.max(sp4.length, sp3.length);

            let rows = "";
            for (let i = 0; i < max; i++) {
                rows += `<tr>
                    <td>${this.renderToolContent(sp4[i])}</td>
                    <td>${this.renderToolContent(sp3[i])}</td>
                </tr>`;
            }
            return `<table class="dmg-grid-table">
                <thead>
                    <tr><th>Werkzeuge Spindel 4</th><th>Werkzeuge Spindel 3</th></tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>`;
        }
    },

    renderOpCell(op, sp, label) {
        const isMatch = op && op.spindle === sp;
        return `<td class="${!isMatch ? 'empty-cell' : ''}">
            ${isMatch ? `<div class="cell-content">
                <span class="cell-title">${op.title}</span>
                <span class="cell-sub">${label} | ${op.tool}</span>
            </div>` : ''}
        </td>`;
    },

    renderToolContent(tool) {
        if (!tool) return "";
        return `<div class="cell-content">
            <span class="cell-title">${tool.tool}</span>
            <span class="cell-sub">${tool.toolName}</span>
        </div>`;
    },

    exportPDF(type) {
        const element = document.getElementById('pdf-buffer');
        element.innerHTML = `<div class="pdf-page">
            <h2 style="text-transform:uppercase; font-size:14px; margin-bottom:10px;">${type === 'plan' ? 'Operation Plan' : 'Werkzeugliste'}</h2>
            ${this.generateTableHTML(type)}
        </div>`;
        
        html2pdf().from(element).set({
            margin: 0,
            filename: `SyncOp_${type}.pdf`,
            jsPDF: { orientation: 'landscape', unit: 'mm', format: 'a4' }
        }).save();
    }
};
