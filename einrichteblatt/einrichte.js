const demoK1 = [
  { t: "T0101", name: "Planen / Vordrehen", print: true },
  { t: "T0202", name: "Außen Schlichten", print: true },
  { t: "T0303", name: "Bohren Ø12.5", print: false },
];

const demoK2 = [
  { t: "T1101", name: "Planen / Vordrehen", print: true },
  { t: "T1202", name: "Innen Schlichten", print: true },
  { t: "T1303", name: "Abstechen", print: false },
];

function renderTable(targetId, data) {
  const tbody = document.getElementById(targetId);
  tbody.innerHTML = "";

  data.forEach(row => {
    const tr = document.createElement("tr");

    const tdPrint = document.createElement("td");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = row.print;
    checkbox.addEventListener("change", () => {
      tr.classList.toggle("no-print", !checkbox.checked);
    });
    tdPrint.appendChild(checkbox);

    const tdT = document.createElement("td");
    tdT.textContent = row.t;

    const tdName = document.createElement("td");
    tdName.textContent = row.name;

    tr.append(tdPrint, tdT, tdName);
    if (!row.print) tr.classList.add("no-print");

    tbody.appendChild(tr);
  });
}

document.getElementById("printBtn").addEventListener("click", () => {
  window.print();
});

renderTable("kanal1", demoK1);
renderTable("kanal2", demoK2);
