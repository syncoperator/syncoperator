const DB_KEY = 'QS_CENTRAL_DB_V2';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];

function saveDB() {
    localStorage.setItem(DB_KEY, JSON.stringify(db));
}

function cleanProjectSlots(pIdx) {
    if (db[pIdx] && db[pIdx].tools) {
        db[pIdx].tools = db[pIdx].tools.filter(t => t.id && t.id.trim() !== "");
    }
    saveDB();
}
