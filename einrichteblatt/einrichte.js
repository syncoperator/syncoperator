/* ===============================
   DATA SOURCE
================================ */

// versucht aus SyncOperator zu lesen
function loadFromSyncOperator() {
  try {
    const raw = localStorage.getItem("CitiTool_SyncOperator_v1");
    if (!raw) return null;

    const parsed = JSON.parse(raw);
    if (!parsed?.data?.slots || !parsed?.data?.library) return null;

    const lib = parsed.data.library;
    const slots = parsed.data.slots;

    function extract(kanal) {
      const result = {};
      slots[kanal].forEach(id => {
        if (!id) return;
        const op = lib.find(o => o.id === id);
        if (!op || !op.toolNo) return;

        if (!result[op.toolNo]) {
          result[op.toolNo] = {
            toolNo: op.toolNo,
            name: op.toolName || op.title || ""
          };
        }
      });
      return Object.values(result);
    }

    return {
      kanal1: extract("1"),
      kanal2: extract("2")
    };
  } catch {
    return null;
  }
}

/* ===============================
   DEMO FALLBACK
================================ */

const DEMO_DATA = {
  kanal1: [
    { toolNo: "T0101", name: "Planhalter Ø80" },
    { toolNo: "T0202", name: "Bohrer Ø12.5" },
    { toolNo: "T0303", name: "Innenstechstahl 2mm" },
    { toolNo: "T0404", name: "Gewindeschneider M26" }
  ],
  kanal2: [
    { toolNo: "T1101", name: "Abstechstahl 3mm" },
    { toolNo: "T1202", name: "Radialbohrer Ø5" },
    { toolNo: "T1303", name: "Entgrater Lochkreis" }
  ]
};

const data = loadFromSyncOperator() || DEMO_DATA;

/* ===============================
   RENDER
================================ */

function renderKanal(bodyId, tools) {
  const body = document.getElementById(bodyId);
  body.innerHTML = "";

  tools
    .sort((a, b) => a.toolNo.localeCompare(b.toolNo))
    .forEach(t => {
      const tr = document.createElement("tr");

      tr.innerHTML = `
        <td><input type="checkbox" checked></td>
        <td class="tool-no">${t.toolNo}</td>
        <td>
          <input class="tool-name" value="${t.name}">
        </td>
      `;

      body.appendChild(tr);
    });
}

renderKanal("kanal1Body", data.kanal1);
renderKanal("kanal2Body", data.kanal2);

/* ===============================
   PRINT
================================ */

document.getElementById("printBtn").addEventListener("click", () => {
  window.print();
});
