document.addEventListener('DOMContentLoaded', () => {
    const el = document.getElementById('draggable-list');

    // Drag and Drop: Мягкое перетаскивание карточек
    new Sortable(el, {
        animation: 350,
        ghostClass: 'sortable-ghost',
        onEnd: () => {
            console.log('Order updated and saved to system.');
        }
    });
});

function exportData() { console.log('Exporting JSON...'); }
function importData() { console.log('Importing JSON...'); }
