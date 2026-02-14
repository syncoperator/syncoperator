function makePDF() {
    const p = db[currentIdx];
    if (!p) return;

    const getPageHead = () => `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; min-height:90px;">
            <div style="display:flex; flex-direction:column;">
                <div style="font-size:13px; font-weight:900; text-transform:uppercase; color:#666;">${p.name || ''}</div>
                <div style="font-size:64px; font-weight:900; line-height:0.8; letter-spacing:-2px; margin:0;">${p.num || '---'}</div>
            </div>
            <div style="width:220px; font-size:11px; font-weight:800; line-height:1.5;">
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>MATERIAL</span><span>${p.mat || ''}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>SÄGELÄNGE</span><span>${p.sag || ''}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>ABSTAND</span><span>${p.abs || ''}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee;"><span>GREIFBACKEN</span><span>${p.grf || ''}</span></div>
                <div style="display:flex; justify-content:space-between;"><span>STÜCKZAHL</span><span>${p.stt || ''} / ${p.stn || ''}</span></div>
            </div>
        </div>
        <div style="border-bottom:5px solid #000; margin-bottom:15px;"></div>`;

    const tableHead = (title) => `
        <div style="margin: 20px 0 5px 0; font-size:18px; font-weight:900; text-transform:uppercase;">${title}</div>
        <div style="display:flex; font-size:10px; font-weight:900; text-transform:uppercase; margin-bottom:6px;">
            <div style="width:75px;">T-NR</div>
            <div style="flex:1;">WERKZEUGNAME</div>
            <div style="width:125px; text-align:right;">Ø / TOLERANZ</div>
        </div>
        <div style="border-bottom:4px solid #000;"></div>`;

    const getRow = (t) => `
        <div style="display:flex; align-items:baseline; border-bottom:1.5px solid #000; padding:10px 0; page-break-inside: avoid;">
            <div style="width:75px; font-weight:800; font-size:15px;">${t.id}</div>
            <div style="flex:1; font-weight:700; font-size:15px; text-transform:uppercase;">${t.nm}</div>
            <div style="width:125px; text-align:right; font-weight:800; font-size:14px;">${t.dia.replace(/\//g, '<br>')}</div>
        </div>`;

    let oben = [], unten = [];
    (p.tools || []).forEach(t => { if(t.rev) unten.push(t); else oben.push(t); });

    // Логика переноса: если инструментов больше 10, переносим Unten на новую страницу
    const totalItems = (p.tools || []).length;
    const shouldBreak = totalItems > 10; 

    let pdfHtml = `
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            @page { size: A4; margin: 0; }
            body { margin: 0; padding: 10mm; font-family: sans-serif; -webkit-print-color-adjust: exact; }
            .content-border { border: 2.2px solid #000; padding: 20px; min-height: 260mm; display: flex; flex-direction: column; box-sizing: border-box; }
            .section-unten { ${shouldBreak ? 'page-break-before: always; margin-top: 20px;' : ''} }
            .footer { border-top: 1.2px solid #000; padding-top: 5px; font-size: 9px; font-weight: 800; text-align: center; margin-top: auto; }
            @media print { .page-break { display: block; page-break-before: always; } }
        </style>
    </head>
    <body>
        <div class="content-border">
            ${getPageHead()}
            
            <div class="section-oben">
                ${tableHead('REVOLVER OBEN')}
                ${oben.map(getRow).join('')}
            </div>

            ${unten.length > 0 ? `
                <div class="section-unten">
                    ${shouldBreak ? getPageHead() : ''} 
                    ${tableHead('REVOLVER UNTEN')}
                    ${unten.map(getRow).join('')}
                </div>
            ` : ''}

            <div class="footer">CITITOOL REPORT</div>
        </div>
        <script>window.onload = () => { setTimeout(() => { window.print(); window.close(); }, 500); };</script>
    </body>
    </html>`;

    const win = window.open('', '_blank');
    if (win) {
        win.document.write(pdfHtml);
        win.document.close();
    }
}
