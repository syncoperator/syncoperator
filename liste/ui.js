function makePDF() {
    const p = db[currentProjectIdx];
    
    // Генерируем строки: теперь используем только проценты для стабильности
    const rows = p.tools.map(t => `
        <tr style="border-bottom: 1.2pt solid black;">
            <td style="width: 12%; padding: 10px 5px; font-weight: 900; font-size: 10pt; white-space: nowrap;">${t.id}</td>
            <td style="width: 68%; padding: 10px 10px; font-weight: 700; font-size: 10pt; text-align: left;">${t.nm}</td>
            <td style="width: 20%; padding: 10px 5px; font-weight: 900; font-size: 10pt; text-align: right; white-space: nowrap;">${t.dia || ''}</td>
        </tr>
    `).join('');

    document.getElementById('pdf-render').innerHTML = `
    <div style="width: 210mm; height: 297mm; padding: 10mm; background: white; box-sizing: border-box; color: black; font-family: sans-serif;">
        <div style="border: 1.5pt solid black; height: 100%; padding: 30px; display: flex; flex-direction: column; box-sizing: border-box;">
            
            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px;">
                <div style="flex: 1;">
                    <div style="font-size: 10pt; font-weight: 600; color: #555; text-transform: uppercase;">${p.name}</div>
                    <div style="font-size: 52pt; font-weight: 900; line-height: 0.9; letter-spacing: -2px;">${p.num}</div>
                </div>
                
                <div style="text-align: right; font-size: 9.5pt; font-weight: 800; line-height: 1.7; min-width: 200px;">
                    <div style="display: flex; justify-content: flex-end; gap: 5px;">ABSTAND: <span style="border-bottom: 1pt solid black; min-width: 60px; text-align: center;">${p.abs || ''}</span></div>
                    <div style="display: flex; justify-content: flex-end; gap: 5px;">GREIFBACKEN: <span style="border-bottom: 1pt solid black; min-width: 60px; text-align: center;">${p.grf || ''}</span></div>
                    <div style="display: flex; justify-content: flex-end; gap: 5px;">LAUFZEIT: <span style="border-bottom: 1pt solid black; min-width: 60px; text-align: center;">${p.lzf || ''}</span></div>
                    <div style="display: flex; justify-content: flex-end; gap: 5px;">STÜCK A: <span style="border-bottom: 1pt solid black; min-width: 30px; text-align: center;">${p.sta || ''}</span> | B: <span style="border-bottom: 1pt solid black; min-width: 30px; text-align: center;">${p.stb || ''}</span></div>
                </div>
            </div>

            <div style="height: 4pt; background: black; margin: 15px 0;"></div>

            <table style="width: 100%; border-collapse: collapse; table-layout: fixed;">
                <thead>
                    <tr style="border-bottom: 2.2pt solid black;">
                        <th style="width: 12%; text-align: left; font-size: 7.5pt; font-weight: 900; padding-bottom: 5px;">T-NR</th>
                        <th style="width: 68%; text-align: left; font-size: 7.5pt; font-weight: 900; padding-bottom: 5px; padding-left: 10px;">WERKZEUGNAME / KOMMENTAR</th>
                        <th style="width: 20%; text-align: right; font-size: 7.5pt; font-weight: 900; padding-bottom: 5px;">Ø / TOLERANZ</th>
                    </tr>
                </thead>
                <tbody>
                    ${rows}
                </tbody>
            </table>
        </div>
    </div>`;

    window.print();
}
