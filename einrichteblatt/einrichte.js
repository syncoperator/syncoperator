// einrichte.js
// Einrichteblatt – Datenquelle: SyncOperator localStorage

const STORAGE_KEY = "CitiTool_SyncOperator_v1";

/* helpers */
function $(sel, root = document) {
  return root.querySelector(sel);
}

function readSyncOperatorData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed.data || null;
  } catch {
    return null;
  }
}

function buildToolMap(state, kanal) {
  const result = [];
  const seen = new Set();

  const slots = state.slots?.[kanal] || [];
  const library = state.library || [];

  for (const opId of slots) {
    if (!opId) continue;

    const op = library.find(o => o.id === opId);
    if (!op) continue;

    const toolNo = (op.toolNo || "").trim();
    if (!toolNo) continue;
    if (seen.has(toolNo)) continue;

    seen.add(toolNo);

    result.push({
      toolNo,
      name: op.toolName?.trim() || op.title || "",
      print: true
    });
  }

  return result;
}

function renderTable(section, tools) {
  const tbody = $("tbody", section);
  tbody.innerHTML = "";

  if (!tools.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="3" class="empty">Keine Werkzeuge</td>`;
    tbody.appendChild(tr);
    return;
  }

  tools.forEach(tool => {
    const tr = document.createElement("tr");

    const tdPrint = document.createElement("td");
    tdPrint.className = "col-print";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = tool.print;
    cb.className = "print-toggle";
    cb.addEventListener("change", () => {
      tr.classList.toggle("no-print", !cb.checked);
    });
    tdPrint.appendChild(cb);

    const tdT = document.createElement("td");
    tdT.className = "col-t";
    tdT.textContent = tool.toolNo;

    const tdName = document.createElement("td");
    tdName.contentEditable = "true";
    tdName.textContent = tool.name;

    tr.append(tdPrint, tdT, tdName);
    tbody.appendChild(tr);
  });
}

function initEinrichteblatt() {
  const data = readSyncOperatorData();

  // fallback demo
  const demo = {
    "1": [
      { toolNo: "T0101", name: "Planen / Vordrehen", print: true },
      { toolNo: "T0202", name: "Außen Schlichten", print: true }
    ],
    "2": [
      { toolNo: "T1101", name: "Planen / Vordrehen", print: true }
    ]
  };

  let kanal1 = demo["1"];
  let kanal2 = demo["2"];

  if (data) {
    kanal1 = buildToolMap(data, "1");
    kanal2 = buildToolMap(data, "2");
  }

  renderTable($(".kanal-1"), kanal1);
  renderTable($(".kanal-2"), kanal2);
}

/* PRINT CLEANUP */
function preparePrint() {
  document.querySelectorAll("tr.no-print").forEach(tr => {
    tr.style.display = "none";
  });
}

function restoreAfterPrint() {
  document.querySelectorAll("tr.no-print").forEach(tr => {
    tr.style.display = "";
  });
}

window.addEventListener("beforeprint", preparePrint);
window.addEventListener("afterprint", restoreAfterPrint);

document.addEventListener("DOMContentLoaded", initEinrichteblatt);
