"use strict";

let db = JSON.parse(localStorage.getItem("core_final")) || [];
let cur = null;

const $ = id => document.getElementById(id);
const save = () => localStorage.setItem("core_final", JSON.stringify(db));

function renderProjects() {
    const l = $("project-list");
    l.innerHTML = "";
    db.forEach((p,i)=>{
        l.innerHTML += `
            <div class="card" onclick="openProject(${i})">
                <b>${p.num}</b><br><small>${p.name}</small>
            </div>`;
    });
}

function openProjectModal() {
    $("modal-project").classList.add("active");
}

function saveProject() {
    const num = $("mp-num").value.trim();
    if(!num) return;
    db.push({ num, name: $("mp-name").value, tools: [] });
    save();
    $("mp-num").value = "";
    $("mp-name").value = "";
    $("modal-project").classList.remove("active");
    renderProjects();
}

function openProject(i) {
    cur = i;
    $("p-num").textContent = db[i].num;
    $("p-name").textContent = db[i].name;
    $("view-home").classList.remove("active");
    $("view-project").classList.add("active");
    renderTools();
}

function goHome() {
    $("view-project").classList.remove("active");
    $("view-home").classList.add("active");
    renderProjects();
}

function renderTools() {
    const l = $("tool-list");
    l.innerHTML = "";
    db[cur].tools.forEach(t=>{
        l.innerHTML += `
            <div class="card">
                <div class="c-id">${t.id}</div>
                <div class="c-name">${t.name}</div>
                <div class="c-dia">${t.dia||"-"}</div>
            </div>`;
    });
}

function openToolModal() {
    $("mt-old").value = "";
    $("mt-id").value = "";
    $("mt-name").value = "";
    $("mt-dia").value = "";
    $("modal-tool").classList.add("active");
}

function saveTool() {
    const id = $("mt-id").value.trim();
    if(!id) return;
    db[cur].tools.push({
        id,
        name: $("mt-name").value,
        dia: $("mt-dia").value
    });
    save();
    $("modal-tool").classList.remove("active");
    renderTools();
}

function exportWerkzeugPDF() {
    const p = db[cur];
    const tpl = $("pdf-template");

    tpl.innerHTML = `
        <div class="pdf-title">Werkzeugliste</div>
        <div class="pdf-sub">${p.num} – ${p.name}</div>

        <table class="pdf-table">
            <thead>
                <tr>
                    <th>T-NR</th>
                    <th>Werkzeug</th>
                    <th>Ø</th>
                </tr>
            </thead>
            <tbody>
                ${p.tools.map(t=>`
                    <tr>
                        <td>${t.id}</td>
                        <td>${t.name}</td>
                        <td>${t.dia||""}</td>
                    </tr>`).join("")}
            </tbody>
        </table>
    `;

    html2pdf().from(tpl).set({
        filename: `Werkzeugliste_${p.num}.pdf`,
        html2canvas: { scale: 2 },
        jsPDF: { unit: "mm", format: "a4" }
    }).save();
}

renderProjects();
