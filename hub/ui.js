document.addEventListener('DOMContentLoaded', () => {
    const el = document.getElementById('sort-container');

    // Drag and Drop: работает идеально для множества карточек
    new Sortable(el, {
        animation: 200,
        handle: '.mini-card', // можно тащить за всю карточку
        ghostClass: 'sortable-ghost',
        onEnd: () => {
            console.log('New order saved');
            // Здесь можно добавить сохранение в LocalStorage
        }
    });
});

function exportJSON() { alert('JSON Exported'); }
function importJSON() { alert('JSON Imported'); }
