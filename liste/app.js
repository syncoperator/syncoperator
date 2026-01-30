"use strict";

/* ========= STORE ========= */
const Store = {
    key: "qs_core_v1",
    state: {
        projects: [],
        current: null
    },

    load() {
        const d = localStorage.getItem(this.key);
        if (d) this.state = JSON.parse(d);
    },

    save() {
        localStorage.setItem(this.key, JSON.stringify(this.state));
    },

    addProject(num, name) {
        this.state.projects.push({ num, name, tools: [] });
        this.save();
    },

    addTool(tool) {
        this.current().tools.push(tool);
        this.save();
    },

    updateTool(oldId, tool) {
        const tools = this.current().tools;
        const i = tools.findIndex(t => t.id === oldId);
        tools[i] = tool;
        this.save();
    },

    removeTool(id) {
        this.current().tools = this.current().tools.filter(t => t.id !== id);
        this.save();
    },

    current() {
        return this.state.projects[this.state.current];
    }
};

/* ========= UI ========= */
const UI = {
    el: id => document.getElementById(id),

    show(id) { this.el(id).classList.add("active"); },
    hide(id) { this.el(id).classList.remove("active"); },

    renderProjects() {
        const l = this.el("project-list");
        l.innerHTML = "";
        Store.state.projects.forEach((p, i) => {
            l.insertAdjacentHTML("beforeend", `
                <div class="card" onclick="Actions.openProject(${i})">
                    <div style="flex:1"><b>${p.num}</b><br><small>${p.name}</small></div>
                </div>
            `);
        });
    },

    renderTools() {
        const l = this.el("tool-list");
        l.innerHTML = "";
        Store.current().tools.forEach(t => {
            l.insertAdjacentHTML("beforeend", `
                <div class="card" data-id="${t.id}">
                    <div class="c-drag">☰</div>
                    <div class="c-id">${t.id}</div>
                    <div class="c-name" onclick="Actions.editTool('${t.id}')">${t.name}</div>
                    <div class="c-dia">${t.dia || "-"}</div>
                    <button class="c-del" onclick="Actions.deleteTool('${t.id}')">✕</button>
                </div>
            `);
        });

        new Sortable(l, {
            handle: ".c-drag",
            animation: 150,
            onEnd() {
                const ids = [...l.children].map(e => e.dataset.id);
                Store.current().tools = ids.map(id =>
                    Store.current().tools.find(t => t.id === id)
                );
                Store.save();
            }
        });
    }
};

/* ========= ACTIONS ========= */
const Actions = {
    openProjectModal() {
        UI.show("modal-project");
    },

    saveProject() {
        const num = UI.el("mp-num").value.trim();
        if (!num) return;
        Store.addProject(num, UI.el("mp-name").value.trim());
        UI.el("mp-num").value = "";
        UI.el("mp-name").value = "";
        UI.hide("modal-project");
        UI.renderProjects();
    },

    openProject(i) {
        Store.state.current = i;
        UI.el("project-num").textContent = Store.current().num;
        UI.el("project-name").textContent = Store.current().name;
        UI.el("v-home").classList.remove("active");
        UI.el("v-project").classList.add("active");
        UI.renderTools();
    },

    goHome() {
        UI.el("v-project").classList.remove("active");
        UI.el("v-home").classList.add("active");
        UI.renderProjects();
    },

    openToolModal() {
        UI.el("mt-old").value = "";
        UI.el("mt-id").value = "";
        UI.el("mt-name").value = "";
        UI.el("mt-dia").value = "";
        UI.show("modal-tool");
    },

    editTool(id) {
        const t = Store.current().tools.find(t => t.id === id);
        UI.el("mt-old").value = t.id;
        UI.el("mt-id").value = t.id;
        UI.el("mt-name").value = t.name;
        UI.el("mt-dia").value = t.dia;
        UI.show("modal-tool");
    },

    saveTool() {
        const tool = {
            id: UI.el("mt-id").value.trim(),
            name: UI.el("mt-name").value.trim(),
            dia: UI.el("mt-dia").value.trim()
        };
        const old = UI.el("mt-old").value;
        old ? Store.updateTool(old, tool) : Store.addTool(tool);
        UI.hide("modal-tool");
        UI.renderTools();
    },

    deleteTool(id) {
        Store.removeTool(id);
        UI.renderTools();
    },

    openImport() {
        UI.show("modal-import");
    },

    importTools() {
        const lines = UI.el("mi-text").value.split("\n");
        let id = null, buf = [];
        lines.forEach(l => {
            l = l.trim();
            if (/^T\d+/i.test(l)) {
                if (id) Store.addTool({ id, name: buf.join(" "), dia: "" });
                id = l.match(/^T\d+/i)[0].toUpperCase();
                buf = [];
            } else if (id) buf.push(l);
        });
        if (id) Store.addTool({ id, name: buf.join(" "), dia: "" });
        UI.el("mi-text").value = "";
        UI.hide("modal-import");
        UI.renderTools();
    },

    exportPDF() {
        const p = Store.current();
        const tpl = UI.el("pdf-template");

        tpl.innerHTML = `
            <h1>${p.num}</h1>
            <p>${p.name}</p>
            <table border="1" width="100%">
                ${p.tools.map(t => `
                    <tr><td>${t.id}</td><td>${t.name}</td><td>${t.dia||"-"}</td></tr>
                `).join("")}
            </table>
        `;

        setTimeout(() => {
            html2pdf().from(tpl).set({ filename: `Setup_${p.num}.pdf` }).save();
        }, 50);
    }
};

/* INIT */
Store.load();
UI.renderProjects();
