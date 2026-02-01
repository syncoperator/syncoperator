"use strict";

/* ===== DEMO DATA ===== */
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

/* ===== UI ===== */
function openProject() {
    $("p-num").textContent = project.num;
    $("p-name").textContent = project.name;
    $("view-home").classList.remove("active");
    $("view-project").classList.add("active");
    renderTools();
}

function goHome() {
    $("view-project").classList.remove("active");
    $("view-home").classList.add("active");
}

function renderTools() {
    const list = $("tool-list");
    list.innerHTML = "";
    project.tools.forEach(t => {
        list.innerHTML += `
            <div class="card">
                <div class="c-id">${t.id}</div>
                <div class="c-name">${t.name}</div>
                <div class="c-dia">${t.dia}</div>
            </div>
        `;
    });
}

/* ===== PDF ===== */
function exportPDF() {
    const root = $("pdf-root");

    root.innerHTML = `
        <div class="pdf-title">Werkzeugliste</div>
        <div class="pdf-sub">${project.num} – ${project.name}</div>

        <table class="pdf-table">
            <thead>
                <tr>
                    <th>T-NR</th>
                    <th>Werkzeug</th>
                    <th>Ø</th>
                </tr>
            </thead>
            <tbody>
                ${project.tools.map(t => `
                    <tr>
                        <td>${t.id}</td>
                        <td>${t.name}</td>
                        <td>${t.dia}</td>
                    </tr>
                `).join("")}
            </tbody>
        </table>
    `;

    /* ВАЖНО: без таймаутов, элемент уже в DOM */
    html2pdf()
        .from(root)
        .set({
            filename: `Werkzeugliste_${project.num}.pdf`,
            html2canvas: {
                scale: 2,
                scrollY: 0
            },
            jsPDF: {
                unit: "mm",
                format: "a4",
                orientation: "portrait"
            }
        })
        .save();
}
