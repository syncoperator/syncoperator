function makePDF() {
    const p = db[currentIdx];
    
    // Список инструментов: Шрифт крупнее (15px), больше отступов для стиля
    const rows = p.tools.map(t => `
        <div style="display: flex; align-items: baseline; border-bottom: 1px solid #f0f0f0; padding: 14px 0;">
            <div style="width: 70px; font-weight: 800; font-size: 15px; color: #000; letter-spacing: -0.5px;">${t.id}</div>
            <div style="flex: 1; font-weight: 700; font-size: 15px; text-transform: uppercase; color: #000; padding-right: 10px;">${t.nm}</div>
            <div style="width: 100px; text-align: right; font-weight: 800; font-size: 15px; color: #000;">${t.dia}</div>
        </div>
    `).join('');

    const html = `
    <div style="width: 210mm; padding: 12mm; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif; color: #000; box-sizing: border-box;">
        
        <div style="border: 2px solid #000; padding: 30px; min-height: 265mm; display: flex; flex-direction: column; box-sizing: border-box;">
            
            <div style="display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 8px;">
                
                <div style="margin-bottom: -4px;">
                    <div style="font-size: 14px; font-weight: 900; text-transform: uppercase; color: #000; margin-bottom: 2px; letter-spacing: 0.5px;">${p.name}</div>
                    <div style="font-size: 72px; font-weight: 900; line-height: 0.85; letter-spacing: -3px; color: #000;">${p.num}</div>
                </div>

                <div style="width: 200px; font-size: 11px; font-weight: 800; line-height: 1.7; color: #000; margin-bottom: 2px;">
                    <div style="display:flex; justify-content:space-between;"><span>ABSTAND</span> <span>${p.abs}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>GREIFBACKEN</span> <span>${p.grf}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>LAUFZEIT</span> <span>${p.lzf}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>SÄGELÄNGE</span> <span>${p.sag}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>STÜCK T</span> <span>${p.stt}</span></div>
                    <div style="display:flex; justify-content:space-between;"><span>STÜCK N</span> <span>${p.stn}</span></div>
                </div>
            </div>

            <div style="border-bottom: 5px solid #000; margin-bottom: 20px;"></div>

            <div style="display: flex; font-size: 10px; font-weight: 900; text-transform: uppercase; color: #000; letter-spacing: 1px; margin-bottom: 8px; padding: 0 2px;">
                <div style="width: 70px;">T-NR</div>
                <div style="flex: 1;">WERKZEUGNAME / KOMMENTAR</div>
                <div style="width: 100px; text-align: right;">Ø / TOLERANZ</div>
            </div>

            <div style="border-bottom: 2px solid #f0f0f0;"></div>

            <div style="flex: 1;">
                ${rows}
            </div>

        </div>
    </div>
    `;

    const container = el('print-container');
    if(container) {
        container.innerHTML = html;
        // Даем браузеру долю секунды на рендеринг шрифтов перед печатью
        setTimeout(() => { window.print(); }, 50);
    }
}
