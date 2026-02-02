/**
 * DATA LOGIC
 * Парсинг текста и массовые операции
 */

/**
 * Универсальный парсер строки
 * Разбирает "T01 FRAESER 10 +0.02" на ID, Имя и Диаметр
 */
function parseSmartString(input) {
    let result = { id: "", nm: "", dia: "" };
    let text = input.trim();

    // 1. Ищем T-номер (например, T01, T12)
    const tMatch = text.match(/T\d+/i);
    if (tMatch) {
        result.id = tMatch[0].toUpperCase();
        text = text.replace(tMatch[0], "").trim();
    }

    // 2. Ищем диаметр и допуск в конце строки (например, 10.5 -0.02 или 12)
    const diaMatch = text.match(/(\d+[\.,]?\d*\s*[\+\-]?\s*\d*[\.,]?\d*)$/);
    if (diaMatch) {
        result.dia = diaMatch[0].trim();
        text = text.replace(diaMatch[0], "").trim();
    }

    // 3. Все, что осталось посередине — это название
    result.nm = text.toUpperCase();
    
    return result;
}

/**
 * Логика массового импорта
 */
function logicMassImport(projectIndex, rawText) {
    const lines = rawText.split('\n');
    lines.forEach(line => {
        if (line.trim()) {
            const parsed = parseSmartString(line);
            if (parsed.id) {
                db[projectIndex].tools.push(parsed);
            }
        }
    });
    cleanProjectSlots(projectIndex);
}
