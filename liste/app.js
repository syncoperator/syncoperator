"use strict";

/* DEMO DATA */
const project = {
    num: "PRJ-001",
    name: "Demo Werkzeugliste",
    tools: Array.from({ length: 15 }).map((_, i) => ({
        id: "T" + String(i + 1).padStart(3, "0"),
        name: "Werkzeug Nummer " + (i + 1),
        dia: (10 + i) + " mm"
    }))
};

function $(id) {
    return document.getElementById(id);
}

/* RENDER */
function render() {
    $("p-title").textContent = `${project.num} – ${project.name}`;
    const body = $("tool-body");
    body.innerHTML = "";
    project.tools.forEach(t => {
        body.innerHTML += `
            <tr>
                <td>${t.id}</td>
                <td>${t.name}</td>
                <td>${t.dia}</td>
            </tr>
        `;
    });
}

/* PRINT */
function printPDF() {
    window.print();
}

render();
