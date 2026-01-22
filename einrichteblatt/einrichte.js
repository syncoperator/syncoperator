const STORAGE_KEY = "CitiTool_SyncOperator_v1";

const $ = id => document.getElementById(id);

function today() {
  return new Date().toLocaleDateString("de-DE");
}

function loadData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    return JSON.parse(raw).data;
  } catch {
    return null;
  }
}

function buildTools(state, kanal) {
  const seen = new Set();
  const list = [];

  const slots = state.slots[kanal] || [];
  const lib = state.library || [];

  slots.forEach(id => {
    const op = lib.find(o => o.id === id);
    if (!op) return;

    const t = (op.toolNo || "").trim();
    if (!t || seen.has(t)) return;

    seen.add(t);

    list.push({
      t,
      name: op.toolName || op.title || "",
      print: true
    });
  });

  return list;
}

function render(bodyId, tools) {
  const body = $(bodyId);
  body.innerHTML = "";

  if (!tools.length) {
    body.innerHTML = `<tr><td colspan="3" class="empty">Keine Werkzeuge</td></tr>`;
    return;
  }

  tools.forEach(tool => {
    const tr = document.createElement("tr");

    const tdP = document.createElement("td");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = tool.print;
    cb.addEventListener("change", () => {
      tr.classList.toggle("no-print", !cb.checked);
    });
    tdP.appendChild(cb);

    const tdT = document.createElement("td");
    tdT.textContent = tool.t;

    const tdN = document.createElement("td");
    tdN.contentEditable = true;
    tdN.textContent = tool.name;

    tr.append(tdP, tdT, tdN);
    body.appendChild(tr);
  });
}

function init() {
  $("dateField").value = today();

  const data = loadData();

  let k1 = [], k2 = [];

  if (data) {
    k1 = buildTools(data, "1");
    k2 = buildTools(data, "2");
  } else {
    k1 = [{ t: "T0101", name: "Planen / Vordrehen", print: true }];
    k2 = [{ t: "T1101", name: "Planen / Vordrehen", print: true }];
  }

  render("kanal1Body", k1);
  render("kanal2Body", k2);

  $("printAll").onclick = () => {
    document.querySelectorAll("tbody input[type=checkbox]").forEach(cb => {
      cb.checked = true;
      cb.dispatchEvent(new Event("change"));
    });
  };

  $("printNone").onclick = () => {
    document.querySelectorAll("tbody input[type=checkbox]").forEach(cb => {
      cb.checked = false;
      cb.dispatchEvent(new Event("change"));
    });
  };

  $("printBtn").onclick = () => window.print();
}

window.addEventListener("beforeprint", () => {
  document.querySelectorAll("tr.no-print").forEach(tr => tr.style.display = "none");
});

window.addEventListener("afterprint", () => {
  document.querySelectorAll("tr.no-print").forEach(tr => tr.style.display = "");
});

document.addEventListener("DOMContentLoaded", init);
