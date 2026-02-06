document.addEventListener('DOMContentLoaded', () => {
    const grid = document.getElementById('drag-zone');

    new Sortable(grid, {
        animation: 400,
        ghostClass: 'sortable-ghost',
        dragClass: 'dragging-card',
        onEnd: () => {
            console.log("System layout updated.");
        }
    });
});
