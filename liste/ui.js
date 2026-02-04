function runImp() {
    const text = el('imp-area').value;
    if (!text.trim()) return;

    // Регулярное выражение ищет T или TO и цифры (T0101, TO202 и т.д.)
    const regex = /(T[0O]\d{2,4})/gi;
    const parts = text.split(regex);
    
    // parts[0] — это текст до первого T, игнорируем его если пустой
    for (let i = 1; i < parts.length; i += 2) {
        let id = parts[i].trim().toUpperCase().replace('O', '0'); // Исправляем TO на T0
        let name = (parts[i + 1] || '').trim()
            .replace(/[\r\n]+/g, ' ') // Убираем переносы строк внутри названия
            .replace(/\s\s+/g, ' ');  // Убираем двойные пробелы
        
        db[currentIdx].tools.push({
            id: id,
            nm: name.toUpperCase(),
            dia: ''
        });
    }

    localStorage.setItem(DB_KEY, JSON.stringify(db));
    el('imp-area').value = '';
    renderTools();
    hide('m-imp');
}

function makePDF() {
    const p = db[currentIdx];
    const rows = p.tools.map(t => `
        <div style="display: flex; align-items: baseline; border-bottom: 1px solid #f0f0f0; padding: 12px 0;">
            <div style="width: 75px; font-weight: 800; font-size: 15px; color: #000; letter-spacing: -0.5px;">${t.id}</div>
            <div style="flex: 1; font-weight: 700; font-size: 15px; text-transform: uppercase; color: #000; padding-right: 10px;">${t.nm}</div>
            <div style="width: 95px; text-align: right; font-weight: 800; font-size: 15px; color: #000;">${t.dia}</div>
        </div>
    `).join('');

    const html = `
    <div style="width: 210mm; padding: 12mm; font-family: -apple-system, system-ui, sans-serif; color: #000; box-sizing: border-box;">
        <div style="border: 2px solid #000; padding: 25px; min-height: 265mm; display: flex; flex-direction: column; box-sizing: border-box;">
            
            <div style="display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 6px;">
                <div style="margin-bottom: -5px;">
                    <div style="font-size: 14px; font-weight: 900; text-transform: uppercase; color: #000; margin-bottom: 1px;">${p.name}</div>
                    <div style="font-size: 70px; font-weight: 900; line-height: 0.82; letter-spacing: -3px; color: #000;">${p.num}</div>
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

            <div style="border-bottom: 5px solid #000; margin-bottom: 15px;"></div>

            <div style="display: flex; font-size: 10px; font-weight: 900; text-transform: uppercase; color: #000; letter-spacing: 0.5px; margin-bottom: 5px;">
                <div style="width: 75px;">T-NR</div>
                <div style="flex: 1;">WERKZEUGNAME / KOMMENTAR</div>
                <div style="width: 95px; text-align: right;">Ø / TOLERANZ</div>
            </div>

            <div style="border-bottom: 2px solid #f0f0f0;"></div>

            <div style="flex: 1;">
                ${rows}
            </div>
        </div>
    </div>`;

    const container = el('print-container');
    if(container) {
        container.innerHTML = html;
        setTimeout(() => { window.print(); }, 100);
    }
}
