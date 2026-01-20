const tools = [
  {
    id: "T0101",
    name: "Planmeißel",
    k1: "Planen / Vordrehen",
    k2: "",
    print: true
  },
  {
    id: "T0203",
    name: "Bohrer Ø12.5",
    k1: "",
    k2: "Bohren Ø12.5",
    print: true
  }
];

const table = document.getElementById("toolTable");

function render() {
  table.innerHTML = "";
  tools.forEach((t, i) => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>
        <input type="checkbox" ${t.print ? "checked" : ""}>
        <div class="controls">
          <button>↑</button>
          <button>↓</button>
        </div>
      </td>
      <td>
        <div class="tool-id">${t.id}</div>
        <div class="tool-name">${t.name}</div>
      </td>
      <td>${t.k1 || "—"}</td>
      <td>${t.k2 || "—"}</td>
    `;

    const checkbox = tr.querySelector("input");
    checkbox.onchange = () => {
      t.print = checkbox.checked;
    };

    const [up, down] = tr.querySelectorAll("button");

    up.onclick = () => {
      if (i > 0) {
        [tools[i - 1], tools[i]] = [tools[i], tools[i - 1]];
        render();
      }
    };

    down.onclick = () => {
      if (i < tools.length - 1) {
        [tools[i + 1], tools[i]] = [tools[i], tools[i + 1]];
        render();
      }
    };

    table.appendChild(tr);
  });
}

render();
