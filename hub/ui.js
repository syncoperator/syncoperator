document.addEventListener('DOMContentLoaded', () => {
    const el = document.getElementById('sortable-list');

    // Инициализация Drag and Drop
    new Sortable(el, {
        animation: 150,
        ghostClass: 'sortable-ghost',
        onEnd: () => {
            console.log("Order updated");
            // Здесь можно сохранить порядок в localStorage
        }
    });

    // Экспорт данных
    document.getElementById('exportBtn').addEventListener('click', () => {
        const data = { version: "1.0", timestamp: new Date() };
        const blob = new Blob([JSON.stringify(data)], {type: 'application/json'});
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'config.json';
        a.click();
    });

    // Импорт (заглушка под логику)
    document.getElementById('importBtn').addEventListener('click', () => {
        const input = document.createElement('input');
        input.type = 'file';
        input.onchange = e => { 
            const file = e.target.files[0];
            alert('Файл ' + file.name + ' выбран');
        };
        input.click();
    });
});
