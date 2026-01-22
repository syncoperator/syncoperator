const STORAGE_KEY = "CitiTool_SyncOperator_v1";

function loadPlanData() {
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
  const map = {};
  const slots = state.slots[kanal] || [];

  slots.forEach((opId) => {
    if (!opId) return;
    const op = state.library.find(o => o.id === opId);
    if (!op) return;
    if (!op.toolNo) return;

    if (!map[op.toolNo]) {
      map[op.toolNo] = {
        toolNo: op.toolNo,
        name: op.toolName || op.title || "",
        print: true
      };
    }
  });

  return Object.values(map);
}

function renderKanal(tableBody, tools) {
  tableBody.innerHTML = "";

  tools.forEach(t => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>${t.toolNo}</td>
      <td contenteditable="true">${t.name}</td>
      <td><input type="checkbox" checked /></td>
    `;

    tableBody.appendChild(tr);
  });
}

function init() {
  const data = loadPlanData();
  if (!data) return;

  const tools1 = buildToolMap(data, "1");
  const tools2 = buildToolMap(data, "2");

  renderKanal(document.getElementById("kanal1Body"), tools1);
  renderKanal(document.getElementById("kanal2Body"), tools2);

  document.getElementById("printBtn").addEventListener("click", () => {
    window.print();
  });
}

document.addEventListener("DOMContentLoaded", init);
