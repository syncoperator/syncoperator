let db = JSON.parse(localStorage.getItem('qs_v20')) || [];
let cur = null;
let sortable = null;

const save = () => localStorage.setItem('qs_v20', JSON.stringify(db));
const showM = id => document.getElementById(id).classList.add('active');
const hideM = id => document.getElementById(id).classList.remove('active');

/* PROJECTS */
function renderP() {
    const l = document.getElementById('list-p');
    l.innerHTML = '';
    db.forEach((p,i)=>{
        l.insertAdjacentHTML('beforeend',`
            <div class="card" onclick="openP(${i})">
                <div style="flex:1"><b>${p.num}</b><br><small>${p.name}</small></div>
                <button class="c-del" onclick="event.stopPropagation();delP(${i})">✕</button>
            </div>
        `);
    });
}

function addP() {
    const num = p_num.value.trim();
    if(!num) return;
    db.push({num,name:p_nam.value,tools:[]});
    save(); renderP(); hideM('m-p');
}

function delP(i){ if(confirm('Löschen?')){ db.splice(i,1); save(); renderP(); } }

function openP(i){
    cur=i;
    v_home.classList.remove('active');
    v_det.classList.add('active');
    d_num.textContent=db[i].num;
    d_nam.textContent=db[i].name;
    renderT();
}

function goHome(){
    v_det.classList.remove('active');
    v_home.classList.add('active');
    renderP();
}

/* TOOLS */
function renderT(){
    const l=list_t;
    l.innerHTML='';
    db[cur].tools.forEach(t=>{
        l.insertAdjacentHTML('beforeend',`
            <div class="card" data-id="${t.id}">
                <div class="c-drag">☰</div>
                <div class="c-id">${t.id}</div>
                <div class="c-name" onclick="editT('${t.id}')">${t.nm}</div>
                <div class="c-diam">${t.dia||'-'}</div>
                <button class="c-del" onclick="delT('${t.id}')">✕</button>
            </div>
        `);
    });

    if(sortable) sortable.destroy();
    sortable=new Sortable(l,{
        handle:'.c-drag',
        animation:150,
        onEnd(){
            const ids=[...l.children].map(c=>c.dataset.id);
            db[cur].tools=ids.map(id=>db[cur].tools.find(t=>t.id===id));
            save();
        }
    });
}

function newTool(){
    t_id_hidden.value='';
    t_id.value='';
    t_nm.value='';
    t_dia.value='';
    showM('m-t');
}

function editT(id){
    const t=db[cur].tools.find(x=>x.id===id);
    t_id_hidden.value=id;
    t_id.value=t.id;
    t_nm.value=t.nm;
    t_dia.value=t.dia;
    showM('m-t');
}

function saveT(){
    const id=t_id.value.trim();
    if(!id) return;
    const t={id,nm:t_nm.value,dia:t_dia.value};
    const old=t_id_hidden.value;
    if(old){
        const i=db[cur].tools.findIndex(x=>x.id===old);
        db[cur].tools[i]=t;
    }else{
        db[cur].tools.push(t);
    }
    save(); renderT(); hideM('m-t');
}

function delT(id){
    db[cur].tools=db[cur].tools.filter(t=>t.id!==id);
    save(); renderT();
}

/* IMPORT */
function doImp(){
    const lines=i_txt.value.split('\n');
    let cid=null, buf=[];
    lines.forEach(l=>{
        l=l.trim(); if(!l) return;
        if(/^T\d+/i.test(l)){
            if(cid) db[cur].tools.push({id:cid,nm:buf.join(' '),dia:''});
            cid=l.match(/^T\d+/i)[0].toUpperCase();
            buf=[];
        }else if(cid){ buf.push(l); }
    });
    if(cid) db[cur].tools.push({id:cid,nm:buf.join(' '),dia:''});
    save(); renderT(); hideM('m-i');
}

/* PDF */
function downloadPDF(){
    const p=db[cur];
    pdf_template.innerHTML=`
        <div class="pdf-border">
            <p class="pdf-proj">${p.name}</p>
            <h1 class="pdf-title">${p.num}</h1>
            <div class="pdf-hr"></div>
            <table class="pdf-table">
                <thead><tr><th>T</th><th>NAME</th><th>Ø</th></tr></thead>
                <tbody>
                    ${p.tools.map(t=>`
                        <tr><td>${t.id}</td><td>${t.nm}</td><td align="right">${t.dia||'-'}</td></tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;
    html2pdf().from(pdf_template).set({filename:`Setup_${p.num}.pdf`}).save();
}

renderP();
