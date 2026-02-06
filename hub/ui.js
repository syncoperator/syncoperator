document.addEventListener('DOMContentLoaded', () => {
    const grid = document.getElementById('sortable-grid');

    new Sortable(grid, {
        animation: 300,
        ghostClass: 'sortable-ghost',
        onEnd: () => {
            console.log("Configuration saved");
        }
    });
});

function exportData() { alert('Exporting...'); }
function importData() { alert('Importing...'); }
