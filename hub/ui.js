document.addEventListener('DOMContentLoaded', () => {
    const list = document.getElementById('sort-list');

    new Sortable(list, {
        animation: 300,
        ghostClass: 'sortable-ghost',
        onEnd: () => {
            console.log("Configuration saved.");
        }
    });
});
