function logicMassImport(pIdx, text) {
    if (!text.trim()) return;
    const lines = text.split('\n');
    lines.forEach(line => {
        const parts = line.trim().split(/\s+/);
        if (parts.length >= 2) {
            const id = parts[0].toUpperCase();
            const dia = parts[parts.length - 1];
            const nm = parts.slice(1, -1).join(' ').toUpperCase() || "WERKZEUG";
            db[pIdx].tools.push({ id, nm, dia });
        }
    });
    saveDB();
}
