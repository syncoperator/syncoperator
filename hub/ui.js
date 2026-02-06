document.addEventListener('DOMContentLoaded', () => {
    const el = document.getElementById('draggable-list');

    new Sortable(el, {
        animation: 250,
        ghostClass: 'sortable-ghost',
        onEnd: () => console.log('Order saved.')
    });
});

function exportData() { alert('JSON Export...'); }
function importData() { alert('JSON Import...'); }
