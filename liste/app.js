"use strict";

/* ========= STORE ========= */
const Store = {
    key: "core_v2",
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
        this.state.projects.push({
            num,
            name,
            entities: { tools: [] }
        });
        this.save();
    },

    currentProject() {
        return this.state.projects[this.state.current];
    },

    tools() {
        return this.currentProject().entities.tools;
    }
};

/* ========= UI ========= */
const UI = {
    el: id => document.getElementById(id),
    show: id => UI.el(id).classList.add("active"),
    hide: id => UI.el(id).classList.remove("active"),

    renderProjects() {
        const l = UI.el("project-list");
        l.innerHTML = "";
        Store.state.projects.forEach((p, i) => {
            l.insertAdjacentHTML("beforeend", `
                <div class="card" onclick="Actions.openProject(${i})">
                    <b>${p.num}</b><br><small>${p.name}</small>
                </div>
            `);
        });
    },

    renderTools() {
        const l = UI.el("tool-list");
        l.innerHTML = "";
        Store.tools().forEach(t => {
            l.insertAdjacentHTML("beforeend", `
                <div class="card">
                    <div class="c-id">${t.id}</div>
                    <div class="c-name" onclick="Actions.editTool('${t.id}')">${t.name}</div>
                    <div class="c-dia">${t.dia || "-"}</div>
                    <button class="c-del" onclick="Actions.deleteTool('${t.id}')">✕</button>
                </div>
            `);
        });
    }
};

/* ========= PDF MODULE ========= */
const PdfModule = {
    whitePage() {
        return `
            <div style="
                width:794px;
                height:1123px;
                padding:60px;
                font-family:-apple-system,sans-serif;
            ">
            </div>
        `;
    },

    preview() {
        UI.el("pdf-preview").innerHTML = this.whitePage();
        UI.show("modal-pdf");
    },

    export(project) {
        const tpl = UI.el("pdf-template");
        tpl.innerHTML = this.whitePage();

        html2pdf().from(tpl).set({
            filename: `Blank_${project.num}.pdf`,
            html2canvas: { scale: 2 },
            jsPDF: { unit: "mm", format: "a4", orientation: "portrait" }
        }).save();
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
        const p = Store.currentProject();
        UI.el("p-num").textContent = p.num;
        UI.el("p-name").textContent = p.name;
        UI.el("view-home").classList.remove("active");
        UI.el("view-project").classList.add("active");
        UI.renderTools();
    },

    goHome() {
        UI.el("view-project").classList.remove("active");
        UI.el("view-home").classList.add("active");
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
        const t = Store.tools().find(t => t.id === id);
        UI.el("mt-old").value = t.id;
        UI.el("mt-id").value = t.id;
        UI.el("mt-name").value = t.name;
        UI.el("mt-dia").value = t.dia;
        UI.show("modal-tool");
    },

    saveTool() {
        const id = UI.el("mt-id").value.trim();
        if (!id) return;

        const tool = {
            id,
            name: UI.el("mt-name").value.trim(),
            dia: UI.el("mt-dia").value.trim()
        };

        const old = UI.el("mt-old").value;
        const tools = Store.tools();

        if (old) {
            const i = tools.findIndex(t => t.id === old);
            tools[i] = tool;
        } else {
            tools.push(tool);
        }

        Store.save();
        UI.hide("modal-tool");
        UI.renderTools();
    },

    deleteTool(id) {
        const tools = Store.tools();
        const i = tools.findIndex(t => t.id === id);
        tools.splice(i, 1);
        Store.save();
        UI.renderTools();
    },

    openPdfPreview() {
        PdfModule.preview();
    },

    exportPdf() {
        PdfModule.export(Store.currentProject());
    }
};

/* INIT */
Store.load();
UI.renderProjects();
