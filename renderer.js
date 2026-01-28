const TableRenderer = {
    renderOperationPlan(slots1, slots2) {
        const max = Math.max(slots1.length, slots2.length);
        let rows = "";
        for (let i = 0; i < max; i++) {
            rows += `<tr>
                <td style="width:25px; text-align:center; font-weight:bold; background:#eee;">${i+1}</td>
                ${this._cell(slots1[i], 'SP4', 'L11'+(i+1))} ${this._cell(slots1[i], 'SP3', 'L11'+(i+1))}
                ${this._cell(slots2[i], 'SP3', 'L21'+(i+1))} ${this._cell(slots2[i], 'SP4', 'L21'+(i+1))}
            </tr>`;
        }
        return `
            <table class="dmg-grid-table">
                <thead>
                    <tr><th style="width:25px">#</th><th colspan="2">Kanal 1</th><th colspan="2">Kanal 2</th></tr>
                    <tr><th></th><th style="color:var(--dmg-sp4)">SP 4</th><th style="color:var(--dmg-sp3)">SP 3</th>
                    <th style="color:var(--dmg-sp3)">SP 3</th><th style="color:var(--dmg-sp4)">SP 4</th></tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>`;
    },

    renderToolList(slots1, slots2) {
        const all = [...slots1, ...slots2].filter(x => x);
        const sp4 = all.filter(t => t.spindle === 'SP4');
        const sp3 = all.filter(t => t.spindle === 'SP3');
        let rows = "";
        for (let i = 0; i < Math.max(sp4.length, sp3.length); i++) {
            rows += `<tr><td>${this._tool(sp4[i])}</td><td>${this._tool(sp3[i])}</td></tr>`;
        }
        return `
            <table class="dmg-grid-table">
                <thead><tr><th>Werkzeuge Spindel 4</th><th>Werkzeuge Spindel 3</th></tr></thead>
                <tbody>${rows}</tbody>
            </table>`;
    },

    _cell(op, sp, l) {
        const ok = op && op.spindle === sp;
        return `<td style="${!ok ? 'background:#fafafa' : ''}">
            ${ok ? `<span class="cell-title">${op.title}</span><span class="cell-sub">${l} | ${op.tool}</span>` : ''}
        </td>`;
    },

    _tool(t) {
        return t ? `<span class="cell-title">${t.tool}</span><span class="cell-sub">${t.toolName}</span>` : '';
    }
};
