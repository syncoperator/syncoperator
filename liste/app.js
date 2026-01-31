"use strict";

/* ========= CORE STORE ========= */
const Store = {
    key: "core_v1",
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
        this.state.projects.push({ num, name });
        this.save();
    },

    currentProject() {
        return this.state.projects[this.state.current];
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
    }
};

/* ========= PDF MODULE ========= */
const PdfModule = {
    buildWhitePage(project) {
        return `
            <div style="
                width:794px;
                height:1123px;
                padding:60px;
                font-family:-apple-system,sans-serif;
            ">
                <!-- deliberately empty -->
            </div>
        `;
    },

    openPreview(project) {
        UI.el("pdf-preview").innerHTML = this.buildWhitePage(project);
        UI.show("modal-pdf");
    },

    export(project) {
        const tpl = UI.el("pdf-template");
        tpl.innerHTML = this.buildWhitePage(project);

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
    },

    goHome() {
        UI.el("view-project").classList.remove("active");
        UI.el("view-home").classList.add("active");
        UI.renderProjects();
    },

    openPdfPreview() {
        PdfModule.openPreview(Store.currentProject());
    },

    exportPdf() {
        PdfModule.export(Store.currentProject());
    }
};

/* INIT */
Store.load();
UI.renderProjects();
