document.addEventListener('DOMContentLoaded', () => {
    const grid = document.getElementById('draggable-zone');

    // Клик по карточкам
    grid.addEventListener('click', (e) => {
        const tile = e.target.closest('.tool-tile');
        if (tile && tile.dataset.url) {
            window.location.href = tile.dataset.url;
        }
    });

    // Drag and Drop логика
    new Sortable(grid, {
        animation: 350,
        easing: "cubic-bezier(1, 0, 0, 1)",
        ghostClass: "sortable-ghost",
        onEnd: () => {
            const order = Array.from(document.querySelectorAll('.tool-tile'))
                               .map(t => t.querySelector('.t-name').innerText);
            localStorage.setItem('sync_order', JSON.stringify(order));
            console.log("Order saved:", order);
        }
    });
});

function exportData() {
    alert("Экспорт конфигурации JSON...");
}

function importData() {
    alert("Загрузка JSON...");
}
