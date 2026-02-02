// Функция "Смарт-решала" для ручного ввода
function smartDistribute() {
    let input = document.getElementById('t-nm').value;
    
    // 1. Вытаскиваем T-номер (например, T05 или T12)
    const tMatch = input.match(/T\d+/i);
    if (tMatch) {
        document.getElementById('t-id').value = tMatch[0].toUpperCase();
        input = input.replace(tMatch[0], '').trim();
    }

    // 2. Ищем диаметр и допуски в конце строки
    // Паттерн ищет числа, знаки + - и символы диаметра в конце
    const diaMatch = input.match(/(\d+[\.,]?\d*\s*[\+\-]\s*\d+[\.,]?\d*|\d+[\.,]?\d*)$/);
    if (diaMatch) {
        // Если нашли что-то похожее на размер, кладем в Ø
        document.getElementById('t-dia').value = diaMatch[0].replace(' ', '\n');
        input = input.replace(diaMatch[0], '').trim();
    }

    // 3. Остаток текста — это название
    document.getElementById('t-nm').value = input.toUpperCase();
}

// При сохранении инструмента
function addT() {
    const idx = document.getElementById('t-idx').value;
    const t = { 
        id: document.getElementById('t-id').value.toUpperCase(), 
        nm: document.getElementById('t-nm').value.toUpperCase(), 
        dia: document.getElementById('t-dia').value 
    };
    
    if(!t.id) return;
    
    if(idx==="") db[cur].tools.push(t); 
    else db[cur].tools[idx] = t;
    
    save(); renderT(); hideM('m-t');
}
