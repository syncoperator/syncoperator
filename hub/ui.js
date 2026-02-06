// Инициализация Drag and Drop для реордеринга
const el = document.getElementById('draggable-zone');
const sortable = Sortable.create(el, {
    animation: 250,
    forceFallback: true, // для плавности на мобилках
    onEnd: function() {
        console.log('Order saved to memory');
    }
});

function exportData() {
    console.log("Exporting JSON...");
    // Логика экспорта
}

function importData() {
    console.log("Importing JSON...");
    // Логика импорта
}
