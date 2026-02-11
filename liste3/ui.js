// --- РЕАЛЬНЫЙ DRAG & DROP (ФИКСИРОВАННЫЙ) ---
let dragEl = null;

function renderTools() {
    const list = el('list-t'); if(!list || currentIdx === null) return;
    const tools = db[currentIdx].tools || [];
    list.innerHTML = '';
    
    tools.forEach((t, i) => {
        const item = document.createElement('div');
        item.className = 'list-item drag-item';
        item.setAttribute('data-idx', i);
        
        const revMark = t.rev ? `<div style="background:#000; color:#fff; font-size:9px; padding:2px 6px; border-radius:4px; margin-bottom:5px; font-weight:900; width:fit-content;">UNTEN START ↓</div>` : '';

        item.innerHTML = `
            <div class="handle" style="cursor:grab; color:#ccc; font-size:24px; padding:10px; touch-action:none; user-select:none;">☰</div>
            <div style="flex:1; min-width:0;" onclick="modalT(${i})">
                ${revMark}
                <small style="color:#8e8e93; font-weight:700; font-size:11px;">${t.id || 'T0000'}</small>
                <b style="font-size:20px; font-weight:900; display:block; line-height:1.2;">${t.nm || '---'}</b>
            </div>
            <div style="font-weight:800; color:var(--accent);">${t.dia || ''}</div>
        `;
        
        const handle = item.querySelector('.handle');

        handle.addEventListener('touchstart', (e) => {
            dragEl = item;
            item.style.opacity = '0.5';
            item.style.background = '#f0f0f0';
        }, {passive: false});

        handle.addEventListener('touchmove', (e) => {
            e.preventDefault();
            const touch = e.touches[0];
            const target = document.elementFromPoint(touch.clientX, touch.clientY);
            const afterElement = getDragAfterElement(list, touch.clientY);
            if (afterElement == null) {
                list.appendChild(item);
            } else {
                list.insertBefore(item, afterElement);
            }
        }, {passive: false});

        handle.addEventListener('touchend', () => {
            item.style.opacity = '1';
            item.style.background = '';
            saveOrder();
            dragEl = null;
        });

        list.appendChild(item);
    });
    // Тот самый отступ в конце списка инструментов
    list.innerHTML += '<div style="height:180px; pointer-events:none;"></div>'; 
}

function getDragAfterElement(container, y) {
    const draggableElements = [...container.querySelectorAll('.drag-item:not(.dragging)')];
    return draggableElements.reduce((closest, child) => {
        const box = child.getBoundingClientRect();
        const offset = y - box.top - box.height / 2;
        if (offset < 0 && offset > closest.offset) {
            return { offset: offset, element: child };
        } else {
            return closest;
        }
    }, { offset: Number.NEGATIVE_INFINITY }).element;
}

function saveOrder() {
    const list = el('list-t');
    const items = [...list.querySelectorAll('.drag-item')];
    const newTools = items.map(item => {
        const idx = item.getAttribute('data-idx');
        return db[currentIdx].tools[idx];
    });
    db[currentIdx].tools = newTools;
    localStorage.setItem(DB_KEY, JSON.stringify(db));
    renderTools();
}

// --- PDF С ПОЛОСКОЙ НА КАЖДОЙ СТРАНИЦЕ ---
function makePDF() {
    const p = db[currentIdx];
    
    const getHeader = () => `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; min-height:90px;">
            <div>
                <div style="font-size:13px; font-weight:900; text-transform:uppercase; color:#666;">${p.name || ''}</div>
                <div style="font-size:64px; font-weight:900; line-height:0.8; letter-spacing:-2px;">${p.num || '---'}</div>
            </div>
            <div style="width:220px; font-size:11px; font-weight:800; line-height:1.5;">
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>LAUFZEIT</span><span>${p.lzf || ''}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>MATERIAL</span><span>${p.mat || ''}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>SÄGELÄNGE</span><span>${p.sag || ''}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>ABSTAND</span><span>${p.abs || ''}</span></div>
                <div style="display:flex; justify-content:space-between; border-bottom:1px solid #f0f0f0;"><span>GREIFBACKEN</span><span>${p.grf || ''}</span></div>
                <div style="display:flex; justify-content:space-between;"><span>STÜCKZAHL</span><span>${p.stt || ''}/${p.stn || ''}</span></div>
            </div>
        </div>
        <div style="border-bottom:5px solid #000; margin-bottom:15px;"></div>`;

    const subHeader = `
        <div style="display:flex; font-size:10px; font-weight:900; text-transform:uppercase; margin-bottom:6px;">
            <div style="width:75px;">T-NR</div><div style="flex:1;">WERKZEUGNAME / KOMMENTAR</div><div style="width:125px; text-align:right;">Ø / TOLERANZ</div>
        </div><div style="border-bottom:4px solid #000; margin-bottom:0px;"></div>`;

    const getRow = (t) => `
        <div style="display:flex; align-items:baseline; border-bottom:1.5px solid #000; padding:10px 0;">
            <div style="width:75px; font-weight:800; font-size:15px;">${t.id}</div>
            <div style="flex:1; font-weight:700; font-size:15px; text-transform:uppercase;">${t.nm}</div>
            <div style="width:125px; text-align:right; font-weight:800; font-size:14px;">${t.dia}</div>
        </div>`;

    const getFooter = () => `
        <div style="margin-top: auto; padding-top: 20px;">
            <div style="border-top: 1.5px solid #000; width: 100%; margin-bottom: 5px;"></div>
            <div style="text-align: center; font-size: 9px; font-weight: 800; color: #000; text-transform: uppercase; letter-spacing: 1px;">
                QS CENTRAL PREMIUM REPORT
            </div>
        </div>`;

    let oben = [], unten = [], target = oben;
    (p.tools || []).forEach(t => { if(t.rev) target = unten; target.push(t); });

    // Условие разделения (если инструментов больше 14 в сумме или Oben длинный)
    const forceSplit = oben.length > 12 || (oben.length + unten.length > 15);

    let html = `<div style="width:210mm; background:#fff; font-family:sans-serif; color:#000;">
        <div style="border:2px solid #000; padding:25px; min-height:282mm; display:flex; flex-direction:column; box-sizing:border-box; page-break-after: always;">
            ${getHeader()}
            <div style="font-size:18px; font-weight:900; margin-bottom:5px;">REVOLVER OBEN</div>
            ${subHeader}
            ${oben.map(getRow).join('')}
            ${(!forceSplit && unten.length > 0) ? `
                <div style="margin-top:30px; font-size:18px; font-weight:900;">REVOLVER UNTEN</div>
                ${subHeader}${unten.map(getRow).join('')}` : ''}
            ${getFooter()}
        </div>`;

    // СТРАНИЦА 2 (если нужно)
    if (forceSplit && unten.length > 0) {
        html += `
        <div style="border:2px solid #000; padding:25px; min-height:282mm; display:flex; flex-direction:column; margin-top:10px; box-sizing:border-box;">
            ${getHeader()}
            <div style="font-size:18px; font-weight:900; margin-bottom:5px;">REVOLVER UNTEN</div>
            ${subHeader}
            ${unten.map(getRow).join('')}
            ${getFooter()}
        </div>`;
    }
    
    html += `</div>`;
    el('print-container').innerHTML = html;
    setTimeout(() => { window.print(); }, 250);
}
