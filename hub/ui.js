document.addEventListener('DOMContentLoaded', () => {
    const el = document.getElementById('drag-zone');

    // Drag and Drop реордеринг
    new Sortable(el, {
        animation: 250,
        ghostClass: 'sortable-ghost',
        onEnd: () => {
            console.log('Order Updated');
        }
    });
});
