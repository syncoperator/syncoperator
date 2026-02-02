function makePDF() {
    const p = db[currentProjectIdx];
    
    // Тонкие линии для элегантного вида
    const rows = p.tools.map(t => `
        <tr style="border-bottom: 0.5pt solid #333;">
            <td style="width: 15%; padding: 12px 0; font-weight: 900; font-size: 10pt;">${t.id}</td>
            <td style="width: 65%; padding: 12px 0; font-weight: 500; font-size: 10pt; text-transform: uppercase;">${t.nm}</td>
            <td style="width: 20%; padding: 12px 0; font-weight: 900; font-size: 11pt; text-align: right;">${t.dia || ''}</td>
        </tr>
    `).join('');

    document.getElementById('pdf-render').innerHTML = `
    <div style="width: 210mm; height: 297mm; padding: 12mm; background: white; box-sizing: border-box; color: black; font-family: 'Helvetica', sans-serif;">
        <div style="border: 1.2pt solid black; height: 100%; padding: 35px; display: flex; flex-direction: column; box-sizing: border-box; position: relative;">
            
            <div style="display: flex; justify-content: space-between;">
                <div>
                    <div style="font-size: 10pt; font-weight: 600; color: #888; text-transform: uppercase; margin-bottom: 2px;">${p.name}</div>
                    <div style="font-size: 48pt; font-weight: 900; line-height: 0.8; letter-spacing: -2px;">${p.num}</div>
                </div>

                <div style="text-align: right; font-size: 9pt; font-weight: 700; line-height: 1.8;">
                    <div style="border-bottom: 0.5pt solid #DDD; margin-bottom: 4px;">ABSTAND: <span style="display:inline-block; min-width:60px; text-align:right">${p.abs||'—'}</span></div>
                    <div style="border-bottom: 0.5pt solid #DDD; margin-bottom: 4px;">GREIFBACKEN: <span style="display:inline-block; min-width:60px; text-align:right">${p.grf||'—'}</span></div>
                    <div style="border-bottom: 0.5pt solid #DDD; margin-bottom: 4px;">LAUFZEIT: <span style="display:inline-block; min-width:60px; text-align:right">${p.lzf||'—'}</span></div>
                    <div style="border-bottom: 0.5pt solid #DDD;">STÜCK A: ${p.sta||'—'} | B: ${p.stb||'—'}</div>
                </div>
            </div>

            <div style="height: 3pt; background: black; margin: 25px 0 15px 0;"></div>

            <table style="width: 100%; border-collapse: collapse; table-layout: fixed;">
                <thead>
                    <tr style="border-bottom: 2pt solid black;">
                        <th style="width: 15%; text-align: left; font-size: 7pt; font-weight: 900; padding-bottom: 8px;">T-NR</th>
                        <th style="width: 65%; text-align: left; font-size: 7pt; font-weight: 900; padding-bottom: 8px;">WERKZEUGNAME / KOMMENTAR</th>
                        <th style="width: 20%; text-align: right; font-size: 7pt; font-weight: 900; padding-bottom: 8px;">Ø / TOLERANZ</th>
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

// ПРОСТОЙ ИМПОРТ БЕЗ "УМНЫХ" ГЛЮКОВ
function runMassImport() {
    const text = document.getElementById('imp-area').value;
    if (!text.trim()) return;
    
    const lines = text.split('\n');
    lines.forEach(line => {
        if(line.trim() !== "") {
            db[currentProjectIdx].tools.push({
                id: "T??", // Ты сам впишешь номер
                nm: line.trim().toUpperCase(),
                dia: ""
            });
        }
    });
    saveDB(); renderTools(); hideM('m-imp');
    document.getElementById('imp-area').value = "";
}
