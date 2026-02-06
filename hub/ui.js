document.addEventListener('DOMContentLoaded', () => {
    const grid = document.getElementById('drag-zone');

    // Drag and Drop инициализация
    new Sortable(grid, {
        animation: 300,
        ghostClass: "sortable-ghost",
        easing: "cubic-bezier(1, 0, 0, 1)",
        onEnd: () => {
            console.log("Configuration Updated");
        }
    });
});

function exportData() {
    alert("Экспорт конфигурации JSON...");
}

function importData() {
    alert("Импорт конфигурации JSON...");
}
