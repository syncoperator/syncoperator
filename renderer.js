const TableRenderer = {
    renderOperationPlan(s1, s2) {
        const max = Math.max(s1.length, s2.length);
        let rows = "";
        for (let i = 0; i < max; i++) {
            rows += `<tr>
                <td style="width:35px; text-align:center; font-weight:900; background:#f1f5f9; border:1px solid #000;">${i+1}</td>
                ${this._cell(s1[i], 'SP4', 'L11'+(i+1))}
                ${this._cell(s1[i], 'SP3', 'L11'+(i+1))}
                ${this._cell(s2[i], 'SP3', 'L21'+(i+1))}
                ${this._cell(s2[i], 'SP4', 'L21'+(i+1))}
            </tr>`;
        }
        return `
            <table class="dmg-grid-table">
                <thead>
                    <tr><th style="width:35px">#</th><th colspan="2">Kanal 1</th><th colspan="2">Kanal 2</th></tr>
                    <tr><th></th><th style="color:var(--dmg-sp4)">Spindel 4</th><th style="color:var(--dmg-sp3)">Spindel 3</th>
                    <th style="color:var(--dmg-sp3)">Spindel 3</th><th style="color:var(--dmg-sp4)">Spindel 4</th></tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>`;
    },

    renderToolList(s1, s2) {
        const tools = [...s1, ...s2].filter(x => x);
        const sp4 = tools.filter(t => t.spindle === 'SP4');
        const sp3 = tools.filter(t => t.spindle === 'SP3');
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
        const active = op && op.spindle === sp;
        return `<td class="${!active ? 'empty-cell' : ''}">
            ${active ? `<div class="cell-content"><span class="cell-title">${op.title}</span><span class="cell-sub">${l} | ${op.tool}</span></div>` : ''}
        </td>`;
    },

    _tool(t) {
        return t ? `<div class="cell-content"><span class="cell-title">${t.tool}</span><span class="cell-sub">${t.toolName}</span></div>` : '';
    }
};
