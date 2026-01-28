const TableRenderer = {
    // План операций (4 колонки + номер)
    renderOperationPlan(slots1, slots2) {
        const max = Math.max(slots1.length, slots2.length);
        let rows = "";
        for (let i = 0; i < max; i++) {
            rows += `
                <tr>
                    <td style="width:30px; text-align:center; background:#eee; font-weight:bold;">${i+1}</td>
                    ${this._opCell(slots1[i], 'SP4', 'L11'+(i+1))}
                    ${this._opCell(slots1[i], 'SP3', 'L11'+(i+1))}
                    ${this._opCell(slots2[i], 'SP3', 'L21'+(i+1))}
                    ${this._opCell(slots2[i], 'SP4', 'L21'+(i+1))}
                </tr>`;
        }
        return `
            <table class="dmg-grid-table">
                <thead>
                    <tr><th style="width:30px">#</th><th colspan="2">Kanal 1</th><th colspan="2">Kanal 2</th></tr>
                    <tr><th></th><th>SP 4</th><th>SP 3</th><th>SP 3</th><th>SP 4</th></tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>`;
    },

    // Список инструментов (2 колонки, БЕЗ #, БЕЗ Kanal)
    renderToolList(slots1, slots2) {
        const all = [...slots1, ...slots2].filter(x => x);
        const sp4 = all.filter(t => t.spindle === 'SP4');
        const sp3 = all.filter(t => t.spindle === 'SP3');
        const max = Math.max(sp4.length, sp3.length);

        let rows = "";
        for (let i = 0; i < max; i++) {
            rows += `
                <tr>
                    <td>${this._toolContent(sp4[i])}</td>
                    <td>${this._toolContent(sp3[i])}</td>
                </tr>`;
        }
        return `
            <table class="dmg-grid-table">
                <thead>
                    <tr><th>Werkzeuge Spindel 4</th><th>Werkzeuge Spindel 3</th></tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>`;
    },

    _opCell(op, sp, label) {
        const match = op && op.spindle === sp;
        return `<td class="${!match ? 'empty-cell' : ''}">
            ${match ? `<div class="cell-content"><span class="cell-title">${op.title}</span><span class="cell-sub">${label} | ${op.tool}</span></div>` : ''}
        </td>`;
    },

    _toolContent(t) {
        return t ? `<div class="cell-content"><span class="cell-title">${t.tool}</span><span class="cell-sub">${t.toolName}</span></div>` : '';
    }
};
