const DB_KEY = 'QS_CENTRAL_FINAL_V1';
let db = JSON.parse(localStorage.getItem(DB_KEY)) || [];

function saveDB() {
    localStorage.setItem(DB_KEY, JSON.stringify(db));
}
