// --- PDF GENERATION ---
function makePDF() {
    const p = db[currentIdx];
    
    // Строки таблицы: линии максимально светлые (#eee)
    const rows = p.tools.map(t => `
        <div style="display: flex; align-items: flex-start; border-bottom: 1px solid #eee; padding: 10px 0;">
            <div style="width: 60px; font-weight: 700; font-size: 14px; color: #000;">${t.id}</div>
            <div style="flex: 1; font-weight: 700; font-size: 14px; text-transform: uppercase; color: #000;">${t.nm}</div>
            <div style="width: 80px; text-align: right; font-weight: 700; font-size: 14px; color: #000;">${t.dia}</div>
        </div>
    `).join('');

    const html = `
    <div style="width: 210mm; min-height: 297mm; padding: 15mm; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; box-sizing: border-box; position: relative;">
        
        <div style="border: 2px solid #000; height: 100%; min-height: 250mm; padding: 25px; box-sizing: border-box; display: flex; flex-direction: column;">
            
            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 20px;">
                
                <div>
                    <div style="font-size: 12px; font-weight: 900; color: #000; text-transform: uppercase; margin-bottom: 5px;">${p.name}</div>
                    <div style="font-size: 64px; font-weight: 900; line-height: 0.8; letter-spacing: -2px; color: #000;">${p.num}</div>
                </div>

                <div style="width: 200px; font-size: 11px; font-weight: 700; line-height: 1.6; color: #000;">
                    <div style="display:flex; justify-content:space-between; margin-bottom: 2px;">
                        <span>ABSTAND</span> <span>${p.abs}</span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom: 2px;">
                        <span>GREIFBACKEN</span> <span>${p.grf}</span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom: 2px;">
                        <span>LAUFZEIT</span> <span>${p.lzf}</span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom: 2px;">
                        <span>SÄGELÄNGE</span> <span>${p.sag}</span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom: 2px;">
                        <span>STÜCK T</span> <span>${p.stt}</span>
                    </div>
                    <div style="display:flex; justify-content:space-between;">
                        <span>STÜCK N</span> <span>${p.stn}</span>
                    </div>
                </div>
            </div>

            <div style="border-bottom: 4px solid #000; margin-bottom: 15px;"></div>

            <div style="display: flex; font-size: 10px; font-weight: 900; color: #000; text-transform: uppercase; margin-bottom: 5px;">
                <div style="width: 60px;">T-NR</div>
                <div style="flex: 1;">WERKZEUGNAME / KOMMENTAR</div>
                <div style="width: 80px; text-align: right;">Ø / TOLERANZ</div>
            </div>

             <div style="border-bottom: 2px solid #eee; margin-bottom: 0px;"></div>

            <div style="flex: 1;">
                ${rows}
            </div>

        </div>
    </div>
    `;

    el('print-container').innerHTML = html;
    window.print();
}
