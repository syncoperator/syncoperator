"use strict";

let db = JSON.parse(localStorage.getItem("qs_v20")) || [];
let cur = null;
let sortable = null;

const $ = id => document.getElementById(id);
const save = () => localStorage.setItem("qs_v20", JSON.stringify(db));

const showM = id => $(id).classList.add("active");
const hideM = id => $(id).classList.remove("active");

/* PROJECTS */
function renderP() {
    const list = $("list-p");
    list.innerHTML = "";

    db.forEach((p, i) => {
        list.insertAdjacentHTML("beforeend", `
            <div class="card" onclick="openP(${i})">
                <div style="flex:1">
                    <b>${p.num}</b><br>
                    <small>${p.name}</small>
                </div>
                <button class="c-del" onclick="event.stopPropagation();delP(${i})">✕</button>
            </div>
        `);
    });
}

function addP() {
    const num = $("p-num").value.trim();
    const name = $("p-nam").value.trim();

    if (!num) {
        alert("Nummer fehlt");
        return;
    }

    db.push({ num, name, tools: [] });

    $("p-num").value = "";
    $("p-nam").value = "";

    save();
    renderP();
    hideM("m-p");
}

function delP(i) {
    if (confirm("Löschen?")) {
        db.splice(i, 1);
        save();
        renderP();
    }
}

function openP(i) {
    cur = i;
    $("v-home").classList.remove("active");
    $("v-det").classList.add("active");
    $("d-num").textContent = db[i].num;
    $("d-nam").textContent = db[i].name;
    renderT();
}

function goHome() {
    $("v-det").classList.remove("active");
    $("v-home").classList.add("active");
    renderP();
}

/* TOOLS */
function renderT() {
    const list = $("list-t");
    list.innerHTML = "";

    db[cur].tools.forEach(t => {
        list.insertAdjacentHTML("beforeend", `
            <div class="card" data-id="${t.id}">
                <div class="c-drag">☰</div>
                <div class="c-id">${t.id}</div>
                <div class="c-name" onclick="editT('${t.id}')">${t.nm}</div>
                <div class="c-diam">${t.dia || "-"}</div>
                <button class="c-del" onclick="delT('${t.id}')">✕</button>
            </div>
        `);
    });

    if (sortable) sortable.destroy();

    sortable = new Sortable(list, {
        handle: ".c-drag",
        animation: 150,
        onEnd() {
            const ids = [...list.children].map(el => el.dataset.id);
            db[cur].tools = ids.map(id => db[cur].tools.find(t => t.id === id));
            save();
        }
    });
}

function newTool() {
    $("t-old-id").value = "";
    $("t-id").value = "";
    $("t-nm").value = "";
    $("t-dia").value = "";
    showM("m-t");
}

function editT(id) {
    const t = db[cur].tools.find(x => x.id === id);
    $("t-old-id").value = t.id;
    $("t-id").value = t.id;
    $("t-nm").value = t.nm;
    $("t-dia").value = t.dia;
    showM("m-t");
}

function saveT() {
    const id = $("t-id").value.trim();
    if (!id) return;

    const tool = {
        id,
        nm: $("t-nm").value.trim(),
        dia: $("t-dia").value.trim()
    };

    const oldId = $("t-old-id").value;

    if (oldId) {
        const i = db[cur].tools.findIndex(t => t.id === oldId);
        db[cur].tools[i] = tool;
    } else {
        db[cur].tools.push(tool);
    }

    save();
    renderT();
    hideM("m-t");
}

function delT(id) {
    db[cur].tools = db[cur].tools.filter(t => t.id !== id);
    save();
    renderT();
}

/* IMPORT */
function doImp() {
    const lines = $("i-txt").value.split("\n");
    let cid = null, buf = [];

    lines.forEach(l => {
        l = l.trim();
        if (!l) return;

        if (/^T\d+/i.test(l)) {
            if (cid) db[cur].tools.push({ id: cid, nm: buf.join(" "), dia: "" });
            cid = l.match(/^T\d+/i)[0].toUpperCase();
            buf = [];
        } else if (cid) {
            buf.push(l);
        }
    });

    if (cid) db[cur].tools.push({ id: cid, nm: buf.join(" "), dia: "" });

    save();
    renderT();
    hideM("m-i");
}

/* PDF */
function downloadPDF() {
    const p = db[cur];

    $("pdf-template").innerHTML = `
        <div class="pdf-border">
            <p class="pdf-proj">${p.name}</p>
            <h1 class="pdf-title">${p.num}</h1>
            <div class="pdf-hr"></div>
            <table class="pdf-table">
                <thead>
                    <tr><th>T</th><th>NAME</th><th>Ø</th></tr>
                </thead>
                <tbody>
                    ${p.tools.map(t => `
                        <tr>
                            <td>${t.id}</td>
                            <td>${t.nm}</td>
                            <td align="right">${t.dia || "-"}</td>
                        </tr>
                    `).join("")}
                </tbody>
            </table>
        </div>
    `;

    html2pdf().from($("pdf-template")).set({
        filename: `Setup_${p.num}.pdf`
    }).save();
}

renderP();
