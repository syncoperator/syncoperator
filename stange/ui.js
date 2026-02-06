// --- Логика Drag and Drop ---
const toolList = document.querySelector('.tool-list');

toolList.addEventListener('dragstart', (e) => {
    e.target.classList.add('selected');
});

toolList.addEventListener('dragend', (e) => {
    e.target.classList.remove('selected');
});

toolList.addEventListener('dragover', (e) => {
    e.preventDefault();
    const activeElement = toolList.querySelector('.selected');
    const currentElement = e.target.closest('.tool-card');
    
    if (!currentElement || activeElement === currentElement) return;

    const nextElement = (currentElement === activeElement.nextElementSibling) ?
        currentElement.nextElementSibling : currentElement;

    toolList.insertBefore(activeElement, nextElement);
});

// --- Data Management (JSON) ---
function exportData(id) {
    const data = {
        t_nr: "T-03",
        beschreibung: "Оптимальная длина: 979 мм",
        params: { material: 912, zange: 50, ausspann: 55, verlust: 54 }
    };
    const blob = new Blob([JSON.stringify(data, null, 2)], {type: 'application/json'});
    alert("JSON сформирован для: " + data.beschreibung);
    console.log(data);
}

function importData() {
    alert("Функция импорта активна. Выберите файл в системе.");
}
