const App = {
    state: Storage.load(),

    init() {
        this.renderFields();
        this.renderTools();
        console.log("QS CENTRAL PREMIUM LOADED");
    },

    showScreen(name) {
        document.querySelectorAll('.screen').forEach(s => s.style.display = 'none');
        document.getElementById(`screen-${name}`).style.display = 'block';
    },

    renderFields() {
        const container = document.getElementById('project-fields');
        container.innerHTML = Object.keys(this.state.fields).map(key => `
            <div class="field-box">
                <label>${key.toUpperCase()}</label>
                <input type="text" value="${this.state.fields[key]}" 
                       oninput="App.updateField('${key}', this.value)">
            </div>
        `).join('');
    },

    updateField(key, val) {
        this.state.fields[key] = val;
        Storage.save(this.state);
    },

    renderTools() {
        const list = document.getElementById('tool-list');
        list.innerHTML = this.state.tools.map((t, i) => `
            <tr>
                <td>${i+1}</td>
                <td style="font-weight: bold;">${t.nm}</td>
                <td>${t.id}</td>
                <td><input type="text" value="${t.dia}" class="table-input" oninput="App.updateToolDia(${i}, this.value)"></td>
                <td><button onclick="App.deleteTool(${i})">DEL</button></td>
            </tr>
        `).join('');
    },

    updateToolDia(index, val) {
        this.state.tools[index].dia = val;
        Storage.save(this.state);
    },

    deleteTool(index) {
        this.state.tools.splice(index, 1);
        Storage.save(this.state);
        this.renderTools();
    }
};

const UI = {
    openImport() { document.getElementById('modal-import').style.display = 'flex'; },
    closeModal() { document.getElementById('modal-import').style.display = 'none'; },
    processImport() {
        const val = document.getElementById('import-input').value;
        const lines = val.split('\n').filter(l => l.trim());
        lines.forEach(line => {
            App.state.tools.push({ nm: line.trim().toUpperCase(), id: 'T?', dia: '0.0' });
        });
        Storage.save(App.state);
        App.renderTools();
        this.closeModal();
        document.getElementById('import-input').value = '';
    }
};

window.onload = () => App.init();
