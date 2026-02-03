const PDFEngine = {
    generate() {
        const data = App.state;
        const win = window.open('', '_blank');
        
        const toolRows = data.tools.map((t, i) => `
            <tr style="border-bottom: 1px solid black;">
                <td style="padding: 10pt; border-right: 1px solid black;">${i+1}</td>
                <td style="padding: 10pt; font-weight: 900; text-transform: uppercase;">${t.nm}</td>
                <td style="padding: 10pt; text-align: center;">${t.id}</td>
                <td style="padding: 10pt; text-align: right;">${t.dia}</td>
            </tr>
        `).join('');

        win.document.write(`
            <html>
            <body style="font-family: monospace; padding: 0; margin: 0;">
                <div style="border: 2.5pt solid black; margin: 10mm; padding: 20pt; min-height: 275mm;">
                    <div style="display: flex; justify-content: space-between;">
                        <div style="font-size: 58pt; font-weight: 900;">QS-PRO</div>
                        <div style="width: 200pt;">
                            ${Object.entries(data.fields).map(([k, v]) => `
                                <div style="display: flex; justify-content: space-between; border-bottom: 0.8pt solid black; padding: 3pt 0;">
                                    <b style="font-size: 9pt; text-transform: uppercase;">${k}</b>
                                    <span>${v || '-'}</span>
                                </div>
                            `).join('')}
                        </div>
                    </div>
                    <div style="height: 6pt; background: black; margin: 20pt 0;"></div>
                    <table style="width: 100%; border-collapse: collapse; border: 1px solid black;">
                        <thead style="background: #EEE;">
                            <tr>
                                <th style="border: 1px solid black; padding: 10pt;">POS</th>
                                <th style="border: 1px solid black; padding: 10pt; text-align: left;">DESCRIPTION</th>
                                <th style="border: 1px solid black; padding: 10pt;">CODE</th>
                                <th style="border: 1px solid black; padding: 10pt; text-align: right;">DIA</th>
                            </tr>
                        </thead>
                        <tbody>${toolRows}</tbody>
                    </table>
                </div>
            </body>
            </html>
        `);
        win.document.close();
        win.print();
    }
};
