const App = {
    state: JSON.parse(localStorage.getItem('sync_v17')) || {
        slots: { "1": [], "2": [] }
    },

    exportPDF(type) {
        const buf = document.getElementById('pdf-buffer');
        const html = (type === 'plan') 
            ? TableRenderer.renderOperationPlan(this.state.slots["1"], this.state.slots["2"])
            : TableRenderer.renderToolList(this.state.slots["1"], this.state.slots["2"]);

        buf.innerHTML = `
            <div class="pdf-page">
                <h1 style="font-size:14px; text-transform:uppercase; border-bottom:2px solid #000; padding-bottom:5px; margin-bottom:15px;">
                    ${type === 'plan' ? 'Operation Plan' : 'Werkzeugliste'}
                </h1>
                ${html}
            </div>`;

        html2pdf().from(buf).set({
            margin: 5,
            filename: `SyncOp_${type}.pdf`,
            jsPDF: { orientation: 'landscape', unit: 'mm', format: 'a4' }
        }).save();
    }
};
