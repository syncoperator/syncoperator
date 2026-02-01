"use strict";

let db = JSON.parse(localStorage.getItem("qs_final")) || [];
let cur = null;

const $ = id => document.getElementById(id);
const save = () => localStorage.setItem("qs_final", JSON.stringify(db));

const showM = id => $(id).classList.add("active");
const hideM = id => $(id).classList.remove("active");

/* PROJECTS */
function renderP() {
    const l = $("list-p");
    l.innerHTML = "";
    db.forEach((p,i)=>{
        l.innerHTML += `
            <div class="card" onclick="openP(${i})">
                <b>${p.num}</b><br><small>${p.name}</small>
            </div>`;
    });
}

function addP() {
    const num = $("p-num").value.trim();
    if(!num) return;
    db.push({ num, name: $("p-nam").value, tools: [] });
    save();
    hideM("m-p");
    renderP();
}

function openP(i) {
    cur = i;
    $("d-num").textContent = db[i].num;
    $("d-nam").textContent = db[i].name;
    $("v-home").classList.remove("active");
    $("v-det").classList.add("active");
    renderT();
}

function goHome() {
    $("v-det").classList.remove("active");
    $("v-home").classList.add("active");
}

/* TOOLS */
function renderT() {
    const l = $("list-t");
    l.innerHTML = "";
    db[cur].tools.forEach(t=>{
        l.innerHTML += `
            <tr>
                <td>${t.id}</td>
                <td>${t.nm}</td>
                <td>${t.dia || ""}</td>
            </tr>`;
    });
}

function addT() {
    const t = {
        id: $("t-id").value.trim(),
        nm: $("t-nm").value.trim(),
        dia: $("t-dia").value.trim()
    };
    if(!t.id) return;
    db[cur].tools.push(t);
    save();
    hideM("m-t");
    renderT();
}

/* IMPORT */
function doImp() {
    const lines = $("i-txt").value.split("\n");
    let id=null, buf=[];
    lines.forEach(l=>{
        l=l.trim();
        if(/^T\d+/i.test(l)){
            if(id) db[cur].tools.push({id,nm:buf.join(" "),dia:""});
            id=l.match(/^T\d+/i)[0];
            buf=[];
        } else if(id) buf.push(l);
    });
    if(id) db[cur].tools.push({id,nm:buf.join(" "),dia:""});
    save();
    hideM("m-i");
    renderT();
}

function printPDF() {
    window.print();
}

renderP();
