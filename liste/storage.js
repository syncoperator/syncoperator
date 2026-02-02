/**
 * STORAGE ENGINE
 * Управляет сохранением, очисткой и JSON-файлами
 */

const DB_KEY = 'QS_CENTRAL_DATABASE_V1';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];

// Сохранение всей базы на диск
function saveDB() {
    localStorage.setItem(DB_KEY, JSON.stringify(db));
}

/**
 * Удаление пустых инструментов из проекта
 * Вызывается автоматически, чтобы в PDF и списках не было "дырок"
 */
function cleanProjectSlots(projectIndex) {
    if (db[projectIndex] && db[projectIndex].tools) {
        db[projectIndex].tools = db[projectIndex].tools.filter(t => 
            t.id && t.id.trim() !== ""
        );
    }
    saveDB();
}

/**
 * EXPORT: Создание физического JSON файла
 */
function exportJSON() {
    if (db.length === 0) return;
    const blob = new Blob([JSON.stringify(db, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `QS_Backup_${new Date().toISOString().slice(0,10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
}

/**
 * IMPORT: Обработка вставленного JSON кода
 */
function importJSON() {
    const area = document.getElementById('json-area');
    try {
        const data = JSON.parse(area.value);
        if (Array.isArray(data)) {
            db = data;
            saveDB();
            area.value = "";
            hideM('m-json');
            location.reload(); // Перезагрузка для чистого обновления всех UI модулей
        }
    } catch (e) {
        // Ошибка формата - игнорируем согласно правилу "без лишних окон"
        console.error("Invalid JSON");
    }
}
