/* ═══════════════════════════════════════════════════════════════════════
   CitiTool · MPF-Generator  (Drop-in Add-on)
   ------------------------------------------------------------------------
   Einbinden mit EINER Zeile, direkt VOR </body> in deiner index.html:

       <script src="mpf-generator.js"></script>

   Das Add-on fügt selbstständig den Button „MPF generieren" in die Toolbar
   ein (neben dem MPF-Import-Button) und öffnet einen Dialog mit Vorschau +
   Download für 1000.MPF (Kanal 1) und 2000.MPF (Kanal 2).

   Es liest die vorhandenen App-Daten:  S.slots, getOp(), slotIsEmpty().
   Kein Eingriff in deinen restlichen Code.

   REGELN (final bestätigt):
   - Status-Nummer = letzte 3 Ziffern des L-Codes (führende Kanal-Ziffer
     entfernt):  L1101→101, L1102→102, L2103→103
   - Register richtet sich nach der SPINDEL des Schritts:
        SP4 → RG704 (Hauptspindel) , SP3 → RG703 (Gegenspindel)
   - Für Schritt i (1-basiert, nur belegte Slots mit Operation):
        IF  <Register der EIGENEN Spindel> == (i==1 ? 1 : Status(code_i))
         <code_i>
         STOPRE
         <Register der NÄCHSTEN Spindel> = (letzter ? 2 : Status(code_{i+1}))
        (beim letzten Schritt = eigenes Register; Übergabe an die nächste
         Spindel, wenn der Folgeschritt auf der anderen Spindel läuft)
   - WAITM(pos,1,2): pos = Slot-Position (1-basiert)
   - GROUP-Argument 3 (spArg): 2 bei SP3, sonst 1
   - L-Code wird aus der Slot-Position abgeleitet: Kanal1→L11NN, Kanal2→L21NN
   - Leere Slots werden übersprungen
   - LF-Zeilenenden; Einsetzung zwischen den Markern
        "; ###_GENERATED_CODE_START_###" und "; ###_GENERATED_CODE_END_###"
═══════════════════════════════════════════════════════════════════════ */
(function () {
  'use strict';

  /* ---------- kleine Helfer (eigene, kollidieren nicht mit der App) ---------- */
  function _q(sel) { return document.querySelector(sel); }
  function _codeState(code) {
    var m = String(code || '').match(/(\d+)/);
    var d = m ? m[1] : '';
    if (!d) return 0;
    return (d.length <= 3) ? (parseInt(d, 10) || 0) : parseInt(d.slice(1), 10);
  }
  function _digits(code) {
    var m = String(code || '').match(/(\d+)/);
    return m ? m[1] : '';
  }
  function _uuid() {
    if (typeof crypto !== 'undefined' && crypto.randomUUID) return crypto.randomUUID();
    var d = Date.now();
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
      var r = (d + Math.random() * 16) % 16 | 0; d = Math.floor(d / 16);
      return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
    });
  }

  /* slotIsEmpty / getOp aus der App nutzen, mit Fallback */
  function _slotEmpty(s) {
    if (typeof slotIsEmpty === 'function') return slotIsEmpty(s);
    return !s || (!s.opId && !s.toolId);
  }
  function _getOp(id) {
    if (typeof getOp === 'function') return getOp(id);
    if (typeof S !== 'undefined' && S.library) return S.library.find(function (o) { return o.id === id; }) || null;
    return null;
  }

  /* ---------- Schritte eines Kanals einsammeln ---------- */
  function collectSteps(kanal) {
    var steps = [];
    var slots = (typeof S !== 'undefined' && S.slots && S.slots[kanal]) ? S.slots[kanal] : [];
    slots.forEach(function (slot, idx) {
      if (_slotEmpty(slot)) return;
      var op = (slot && slot.opId) ? _getOp(slot.opId) : null;
      if (!op) return; // nur Slots mit Operation
      var prefix = (kanal === '2') ? 'L21' : 'L11';
      steps.push({
        pos: idx + 1,
        code: prefix + String(idx + 1).padStart(2, '0'),
        title: (op.title || op.code || '').toUpperCase(),
        spindle: (op.spindle === 'SP3') ? 'SP3' : 'SP4'
      });
    });
    return steps;
  }

  /* ---------- GENERATED-Block bauen (LF) ---------- */
  function buildBlock(kanal) {
    var steps = collectSteps(kanal);
    var uuid = _uuid();
    var out = '';
    steps.forEach(function (step, i) {
      var num = String(i + 1).padStart(3, '0');
      var code = step.code;
      var dg = _digits(code);
      var ifReg = (step.spindle === 'SP3') ? 'RG703' : 'RG704';
      var ifVal = (i === 0) ? 1 : _codeState(code);
      var last = (i === steps.length - 1);
      var assignReg = last ? ifReg : ((steps[i + 1].spindle === 'SP3') ? 'RG703' : 'RG704');
      var assignVal = last ? 2 : _codeState(steps[i + 1].code);
      var spArg = (step.spindle === 'SP3') ? 2 : 1;
      out += 'GROUP_BEGIN(0,"' + num + ': ' + step.title + '",' + spArg + ',0)\n';
      out += ';Id=' + uuid + '\n';
      out += 'WAITM(' + step.pos + ',1,2)\n';
      out += 'STOPRE\n';
      out += 'IF ' + ifReg + '==' + ifVal + '\n';
      out += ' ' + code + '\n';
      out += ' STOPRE\n';
      out += ' ' + assignReg + '=' + assignVal + '\n';
      out += 'ELSE\n';
      out += ' ;DUMMY("' + dg + '")\n';
      out += 'ENDIF\n';
      out += 'GROUP_END(0,' + spArg + ')\n';
    });
    return { text: out, steps: steps };
  }

  /* ---------- Shell (Vorlage) ---------- */
  var MPF_SHELLS = { '1': null, '2': null };

  function splice(shell, block) {
    var A = '; ###_GENERATED_CODE_START_###';
    var B = '; ###_GENERATED_CODE_END_###';
    var si = shell.indexOf(A), ei = shell.indexOf(B);
    if (si === -1 || ei === -1) return shell;
    return shell.slice(0, si + A.length) + '\n' + block + '\n' + shell.slice(ei);
  }

  function defaultShell(kanal) {
    var startN = (kanal === '2') ? 'ZA_START2_PCR2G' : 'ZA_START1_PCR2G';
    var kanalLine = (kanal === '2') ? '; ---------- Kanal: 2 ---------------' : '; ---------- Kanal: 1 ---------------';
    var prologExtra = (kanal === '2') ? '' : ' L1000\n';
    return (
      'EXTERN VERSCHIEBUNG (STRING[20])\n' +
      'EXTERN DUMMY (STRING[25])\n' +
      kanalLine + '\n' +
      'IF ISVAR("DM_GROUP_OFF")\n' +
      ' DM_GROUP_OFF=1\n' +
      'ENDIF\n' +
      'GROUP_BEGIN(0,"INFO",0,0)\n' +
      ';Maschine  : CTX_350_4A\n' +
      ';OPW-Template : ZA_BAR_Z3\n' +
      'GROUP_END(0,0)\n' +
      'GROUP_BEGIN(0,"PROLOG",0,0)\n' +
      prologExtra + ' ;WAITMARKEN\n' +
      'WAITM(10,WAIT_K1,WAIT_K2)\n' +
      ' STOPRE\n' +
      ' ' + startN + '\n' +
      'WAITM(15,WAIT_K1,WAIT_K2)\n' +
      ' STOPRE\n' +
      ' IF RG704<>0 GOTOF LOOPSTART_MAIN\n' +
      ' IF RG700<=0 GOTOF BARCHANGE\n' +
      ';****************************************\n' +
      'LOOPSTART_MAIN:\n' +
      'WAITM(20,WAIT_K1,WAIT_K2)\n' +
      ' $AC_PROG_NET_TIME_TRIGGER=1\n' +
      ' M1 M99\n' +
      ' CHECK_MAG\n' +
      'WAITM(25,WAIT_K1,WAIT_K2)\n' +
      'GROUP_END(0,0)\n' +
      ';****************************************\n' +
      '; ###_GENERATED_CODE_START_###\n' +
      '; ###_GENERATED_CODE_END_###\n' +
      ';****************************************\n' +
      'GROUP_BEGIN(0,"Feierabend",0,0)\n' +
      'SHIFT_END:\n' +
      'WAITM(90,WAIT_K1,WAIT_K2)\n' +
      ' STOPRE\n' +
      ' M30\n' +
      'GROUP_END(0,0)\n' +
      '____\n'
    );
  }

  function generateFile(kanal) {
    var b = buildBlock(kanal);
    var shell = MPF_SHELLS[kanal] || defaultShell(kanal);
    return splice(shell, b.text);
  }

  /* Optional: Shell aus importierter MPF übernehmen.
     Hängt sich (falls vorhanden) an die App-Funktion importMPFFile an,
     ohne sie zu ersetzen. */
  function hookShellCapture() {
    if (typeof window.importMPFFile !== 'function') return;
    if (window.importMPFFile.__mpfHooked) return;
    var orig = window.importMPFFile;
    function wrapped(text, filename, onDone) {
      try {
        var A = '; ###_GENERATED_CODE_START_###';
        var B = '; ###_GENERATED_CODE_END_###';
        if (text && text.indexOf(A) !== -1 && text.indexOf(B) !== -1) {
          var k = (typeof detectKanalFromMPF === 'function')
            ? (detectKanalFromMPF(filename, text) || (typeof S !== 'undefined' ? S.kanal : '1'))
            : '1';
          MPF_SHELLS[k] = text.replace(/\r\n/g, '\n');
        }
      } catch (e) { /* ignore */ }
      return orig.apply(this, arguments);
    }
    wrapped.__mpfHooked = true;
    window.importMPFFile = wrapped;
  }

  /* ---------- Download ---------- */
  function download(filename, content) {
    var blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(a.href);
  }

  /* ---------- Dialog ---------- */
  function openMPFModal() {
    // Nutzt das vorhandene Modal der App, falls verfügbar
    var hasAppModal = (typeof openModal === 'function') && _q('#modal') && _q('#mBody') && _q('#mFoot');
    if (hasAppModal) { openAppModal(); return; }
    openStandaloneModal();
  }

  function buildPane(container, files) {
    var activeK = '1';
    container.innerHTML = '';

    var tabs = document.createElement('div');
    tabs.style.cssText = 'display:flex;border:1px solid #D4D4D4;border-radius:4px;overflow:hidden;margin-bottom:14px;width:max-content;font-family:Inter,system-ui,sans-serif;';
    function mkTab(k, label) {
      var b = document.createElement('button');
      b.type = 'button'; b.textContent = label; b.setAttribute('data-k', k);
      b.style.cssText = 'border:none;border-right:1px solid #E5E5E5;padding:8px 16px;font-size:12.5px;font-weight:600;cursor:pointer;background:#fff;color:#737373;';
      if (k === '2') b.style.borderRight = 'none';
      b.addEventListener('click', function () { activeK = k; render(); });
      return b;
    }
    var t1 = mkTab('1', 'Kanal 1 · 1000.MPF');
    var t2 = mkTab('2', 'Kanal 2 · 2000.MPF');
    tabs.append(t1, t2);

    var info = document.createElement('div');
    info.style.cssText = 'font-family:"JetBrains Mono",monospace;font-size:11px;color:#737373;margin-bottom:10px;line-height:1.5;';

    var pre = document.createElement('pre');
    pre.style.cssText = 'background:#0a0a0a;color:#e5e5e5;padding:14px;border-radius:6px;font-family:"JetBrains Mono",monospace;font-size:10.5px;line-height:1.5;overflow:auto;max-height:48vh;white-space:pre;margin:0;';

    function render() {
      [t1, t2].forEach(function (b) {
        var on = b.getAttribute('data-k') === activeK;
        b.style.background = on ? '#0a0a0a' : '#fff';
        b.style.color = on ? '#fff' : '#737373';
      });
      var steps = buildBlock(activeK).steps;
      var fn = activeK === '1' ? '1000.MPF' : '2000.MPF';
      var note = MPF_SHELLS[activeK] ? 'importierte Vorlage' : 'Standard-Vorlage';
      info.textContent = fn + ' · ' + steps.length + ' Schritt(e) · Register pro Schritt nach Spindel (SP4→RG704, SP3→RG703) · Vorlage: ' + note;
      pre.textContent = files[activeK];
    }
    render();
    container.append(tabs, info, pre);
  }

  function openAppModal() {
    openModal('[MPF] · Generieren', 'NC-Programme erzeugen (1000.MPF / 2000.MPF)');
    var box = _q('#modal .modal-box');
    if (box) box.classList.add('modal-wide');
    var body = _q('#mBody');
    var foot = _q('#mFoot');
    var files = { '1': generateFile('1'), '2': generateFile('2') };
    buildPane(body, files);

    function mkBtn(label, cls, fn) {
      var b = document.createElement('button');
      b.type = 'button'; b.className = 'btn ' + cls; b.textContent = label;
      b.addEventListener('click', fn); return b;
    }
    foot.append(
      mkBtn('Schließen', 'btn-ghost', function () { if (typeof closeModal === 'function') closeModal(); }),
      mkBtn('1000.MPF ⬇', 'btn-ghost', function () { download('1000.MPF', files['1']); }),
      mkBtn('2000.MPF ⬇', 'btn-primary', function () { download('2000.MPF', files['2']); })
    );
  }

  /* Eigenständiges Modal (falls App-Modal fehlt) */
  function openStandaloneModal() {
    var files = { '1': generateFile('1'), '2': generateFile('2') };
    var ov = document.createElement('div');
    ov.style.cssText = 'position:fixed;inset:0;z-index:99999;display:flex;align-items:center;justify-content:center;padding:20px;background:rgba(0,0,0,.45);';
    var box = document.createElement('div');
    box.style.cssText = 'background:#fff;border:1px solid #D4D4D4;border-radius:6px;width:100%;max-width:1100px;max-height:92vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 16px 48px rgba(0,0,0,.18);font-family:Inter,system-ui,sans-serif;';
    var head = document.createElement('div');
    head.style.cssText = 'padding:16px 20px;border-bottom:1px solid #E5E5E5;font-weight:600;font-size:15px;';
    head.textContent = 'MPF generieren — 1000.MPF / 2000.MPF';
    var body = document.createElement('div');
    body.style.cssText = 'padding:18px 20px;overflow:auto;flex:1;';
    var foot = document.createElement('div');
    foot.style.cssText = 'padding:14px 20px;border-top:1px solid #E5E5E5;display:flex;justify-content:flex-end;gap:8px;background:#FAFAFA;';
    function mkBtn(label, primary, fn) {
      var b = document.createElement('button');
      b.type = 'button'; b.textContent = label;
      b.style.cssText = 'font-family:Inter,system-ui,sans-serif;font-size:12px;font-weight:600;padding:8px 14px;border-radius:4px;cursor:pointer;border:1px solid ' + (primary ? '#0a0a0a' : '#D4D4D4') + ';background:' + (primary ? '#0a0a0a' : '#fff') + ';color:' + (primary ? '#fff' : '#0a0a0a') + ';';
      b.addEventListener('click', fn); return b;
    }
    function close() { document.body.removeChild(ov); }
    buildPane(body, files);
    foot.append(
      mkBtn('Schließen', false, close),
      mkBtn('1000.MPF ⬇', false, function () { download('1000.MPF', files['1']); }),
      mkBtn('2000.MPF ⬇', true, function () { download('2000.MPF', files['2']); })
    );
    box.append(head, body, foot);
    ov.append(box);
    ov.addEventListener('click', function (e) { if (e.target === ov) close(); });
    document.body.appendChild(ov);
  }

  /* ---------- Button in die Toolbar einsetzen ---------- */
  function injectButton() {
    if (document.getElementById('btnGenMPF')) return true;
    var importBtn = document.getElementById('btnImportMPF');
    var actions = document.querySelector('.toolbar-actions');
    if (!importBtn && !actions) return false;

    var btn = document.createElement('button');
    btn.id = 'btnGenMPF';
    btn.className = 'btn btn-primary btn-sm';
    btn.title = 'NC-Programme 1000.MPF / 2000.MPF erzeugen';
    btn.innerHTML = '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 2v12"/><path d="M4 2h6l2 2v10"/><path d="M6 6h4M6 9h4M6 11.5h4"/></svg>MPF generieren';
    btn.addEventListener('click', function () {
      hookShellCapture();
      openMPFModal();
    });

    if (importBtn && importBtn.parentNode) {
      importBtn.parentNode.insertBefore(btn, importBtn); // direkt vor dem Import-Button
    } else if (actions) {
      actions.appendChild(btn);
    }
    return true;
  }

  /* ---------- Tastenkürzel (optional): Strg+G ---------- */
  function bindShortcut() {
    document.addEventListener('keydown', function (e) {
      if ((e.ctrlKey || e.metaKey) && (e.key === 'g' || e.key === 'G')) {
        e.preventDefault();
        hookShellCapture();
        openMPFModal();
      }
    });
  }

  /* ---------- Start ---------- */
  function start() {
    hookShellCapture();
    if (!injectButton()) {
      // App-Toolbar evtl. noch nicht da → kurz erneut versuchen
      var tries = 0;
      var iv = setInterval(function () {
        tries++;
        if (injectButton() || tries > 40) clearInterval(iv);
      }, 150);
    }
    bindShortcut();
  }

  // Öffentliche API (falls du es manuell aufrufen willst)
  window.CitiMPF = {
    open: function () { hookShellCapture(); openMPFModal(); },
    generate: generateFile,
    setShell: function (kanal, text) { MPF_SHELLS[kanal] = String(text || '').replace(/\r\n/g, '\n'); },
    _build: buildBlock
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start);
  } else {
    start();
  }
})();

/* ═══════════════════════════════════════════════════════════════════════
   CitiTool · Visual Suite Add-on  (Splash + Premium Logo + MPF/Ablauf Tab)
   ------------------------------------------------------------------------
   Erweitert die App OHNE Eingriff in index.html:
   1) Splash-Intro mit PIN-Gate (Code 3504)
   2) Premium-Logo (Original-Mark mit Verlauf)
   3) Neue Ansicht „MPF / Ablauf" — visuelle Kette + Live-Code, nutzt
      die App-Funktionen S, getOp, L, slotIsEmpty, setView.
   Greift auf window.CitiMPF (oben definiert) für Codegenerierung zu.
═══════════════════════════════════════════════════════════════════════ */
(function () {
  'use strict';

  function ready(fn){ if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',fn); else fn(); }
  function el(tag, attrs, html){ var e=document.createElement(tag); if(attrs)for(var k in attrs)e.setAttribute(k,attrs[k]); if(html!=null)e.innerHTML=html; return e; }

  /* ---------- helpers mirroring app rules ---------- */
  function codeState(c){var m=String(c||'').match(/(\d+)/);var d=m?m[1]:'';if(!d)return 0;return d.length<=3?(parseInt(d,10)||0):parseInt(d.slice(1),10);}
  function slotEmptyG(s){ if(typeof slotIsEmpty==='function')return slotIsEmpty(s); return !s||(!s.opId&&!s.toolId); }
  function getOpG(id){ if(typeof getOp==='function')return getOp(id); return null; }
  function Lcode(k,n){ if(typeof L!=='undefined'&&L[k])return L[k](n); return (k==='2'?'L21':'L11')+String(n).padStart(2,'0'); }

  /* steps for a channel from live S.slots (only slots with an op) */
  function steps(k){
    if(typeof S==='undefined'||!S.slots||!S.slots[k])return [];
    var arr=[];
    S.slots[k].forEach(function(slot,idx){
      if(slotEmptyG(slot))return;
      var op=(slot&&slot.opId)?getOpG(slot.opId):null;
      if(!op)return;
      arr.push({pos:idx+1, code:Lcode(k,idx+1), title:(op.title||op.code||''), spindle:op.spindle==='SP3'?'SP3':'SP4',
                tno:(op.toolNo||''), w1:(slot.w1!=null?slot.w1:''), w2:(slot.w2!=null?slot.w2:1), w3:(slot.w3!=null?slot.w3:2), slotRef:slot, idx:idx});
    });
    return arr.map(function(s,i,a){
      var ifReg=s.spindle==='SP3'?'RG703':'RG704';
      var ifVal=(i===0)?1:codeState(s.code);
      var last=i===a.length-1;
      var nextSp=last?s.spindle:a[i+1].spindle;
      var assignReg=last?ifReg:(nextSp==='SP3'?'RG703':'RG704');
      var assignVal=last?2:codeState(a[i+1].code);
      s.state=codeState(s.code); s.ifReg=ifReg; s.ifVal=ifVal; s.assignReg=assignReg; s.assignVal=assignVal; s.last=last; s.handoff=!last&&s.spindle!==nextSp;
      return s;
    });
  }

  /* ---------- styles ---------- */
  function injectStyles(){
    if(document.getElementById('citiVisualCSS'))return;
    var css = `
    /* premium logo override */
    .brand-mark{background:linear-gradient(135deg,#2c2c2c,#000)!important;overflow:hidden;}
    .brand-mark::before{border-color:#fff!important;border-right-color:transparent!important;}
    /* splash */
    #citiSplash{position:fixed;inset:0;z-index:5000;background:radial-gradient(circle at 50% 38%,#1a1a1a,#0a0a0a 70%);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:26px;transition:opacity .5s,visibility .5s;}
    #citiSplash.hide{opacity:0;visibility:hidden;}
    .cs-logo{width:118px;height:118px;}
    .cs-spin{transform-origin:50% 50%;animation:csSpin 8s linear infinite;}
    @keyframes csSpin{to{transform:rotate(360deg);}}
    .cs-name{font-family:Inter,system-ui,sans-serif;font-size:25px;font-weight:800;letter-spacing:.18em;text-transform:uppercase;color:#fff;}
    .cs-name span{color:#9ca3af;}
    .cs-sub{font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:.3em;text-transform:uppercase;color:#6b7280;margin-top:-16px;}
    .cs-gatelabel{font-family:'JetBrains Mono',monospace;font-size:11px;letter-spacing:.2em;text-transform:uppercase;color:#9ca3af;}
    .cs-pin{display:flex;gap:10px;}
    .cs-dot{width:14px;height:14px;border-radius:50%;border:1.5px solid #4b5563;transition:.15s;}
    .cs-dot.on{background:#fff;border-color:#fff;}
    .cs-keys{display:grid;grid-template-columns:repeat(3,62px);gap:11px;}
    .cs-key{height:62px;border-radius:14px;border:1px solid #2a2a2a;background:#161616;color:#fff;font-family:'JetBrains Mono',monospace;font-size:23px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:.1s;}
    .cs-key:hover{background:#222;} .cs-key:active{transform:scale(.94);background:#2c2c2c;}
    .cs-key.empty{visibility:hidden;} .cs-key.wide{font-size:14px;}
    .cs-gate{display:flex;flex-direction:column;align-items:center;gap:14px;}
    .cs-gate.shake{animation:csShake .4s;}
    @keyframes csShake{0%,100%{transform:translateX(0);}20%{transform:translateX(-12px);}40%{transform:translateX(11px);}60%{transform:translateX(-8px);}80%{transform:translateX(6px);}}
    .cs-hint{font-family:'JetBrains Mono',monospace;font-size:10px;color:#6b7280;height:14px;}
    /* chain view */
    .mpfv-wrap{display:grid;gap:16px;align-items:start;}
    .mpfv-wrap.split{grid-template-columns:1fr 1fr;}
    @media(max-width:900px){.mpfv-wrap.split{grid-template-columns:1fr;}}
    .mpfv-bar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:16px;}
    .mpfv-seg{display:inline-flex;background:var(--bg-2);border:1px solid var(--line);border-radius:var(--r);padding:3px;}
    .mpfv-seg button{border:none;background:transparent;color:var(--ink-4);font-family:var(--sans);font-size:12px;font-weight:600;padding:7px 14px;cursor:pointer;border-radius:3px;display:flex;align-items:center;gap:7px;}
    .mpfv-seg button .mf{font-family:var(--mono);font-size:10px;opacity:.65;}
    .mpfv-seg button.on{background:var(--ink);color:#fff;}
    .mpfv-panel{background:var(--bg);border:1px solid var(--line);border-radius:var(--r-md);overflow:hidden;}
    .mpfv-phead{padding:11px 14px;border-bottom:1px solid var(--line);background:var(--bg-soft);display:flex;align-items:center;gap:10px;}
    .mpfv-ptag{font-family:var(--mono);font-size:9.5px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--ink-4);}
    .mpfv-ptitle{font-size:13px;font-weight:600;}
    .mpfv-legend{display:flex;gap:14px;padding:10px 14px 0;flex-wrap:wrap;}
    .mpfv-ll{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--ink-3);}
    .mpfv-ll .d{width:11px;height:11px;border-radius:3px;}
    .mpfv-chain{padding:12px 14px 16px;display:flex;flex-direction:column;}
    .mpfv-endcap{margin-left:62px;padding:5px 12px;font-family:var(--mono);font-size:10.5px;color:var(--ink-3);}
    .mpfv-step{position:relative;display:flex;}
    .mpfv-rail{flex:0 0 56px;display:flex;flex-direction:column;align-items:center;}
    .mpfv-line{width:2px;flex:1;background:var(--line-2);min-height:8px;}
    .mpfv-node{width:46px;height:46px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-family:var(--mono);font-size:14px;font-weight:700;color:#fff;flex-shrink:0;z-index:1;box-shadow:0 2px 6px rgba(0,0,0,.12);transition:transform .14s;}
    .mpfv-node.sp4{background:#2D6B47;} .mpfv-node.sp3{background:#2D4F8A;}
    .mpfv-step:first-child .mpfv-line.top{background:transparent;}
    .mpfv-step.last .mpfv-line.bot{background:transparent;}
    .mpfv-card{flex:1;margin-left:8px;border:1px solid var(--line);border-radius:var(--r-md);padding:10px 12px;background:var(--bg);transition:.14s;position:relative;animation:mpfvIn .22s ease both;}
    @keyframes mpfvIn{from{opacity:0;transform:translateY(6px);}to{opacity:1;transform:none;}}
    .mpfv-card.sp4{border-left:3px solid #2D6B47;} .mpfv-card.sp3{border-left:3px solid #2D4F8A;}
    .mpfv-card:hover{border-color:var(--line-3);transform:translateX(2px);}
    .mpfv-card.hl{box-shadow:0 0 0 2px var(--ink);border-color:var(--ink);}
    .mpfv-card.hl ~ .mpfv-node,.mpfv-step:hover .mpfv-node{transform:scale(1.08);}
    .mpfv-ctop{display:flex;align-items:center;gap:8px;margin-bottom:7px;}
    .mpfv-handle{cursor:grab;color:var(--ink-5);font-size:13px;width:16px;text-align:center;user-select:none;}
    .mpfv-cstate{font-family:var(--mono);font-size:10px;font-weight:700;color:var(--ink-4);background:var(--bg-2);border:1px solid var(--line-2);border-radius:3px;padding:1px 6px;}
    .mpfv-ctitle{font-weight:600;font-size:13px;flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
    .mpfv-clcode{font-family:var(--mono);font-size:10.5px;font-weight:600;background:var(--bg-2);border:1px solid var(--line-2);padding:1px 6px;border-radius:3px;}
    .mpfv-csp{font-family:var(--mono);font-size:9px;font-weight:700;padding:2px 7px;border-radius:3px;}
    .mpfv-sp4b{background:#EAF6EE;color:#2D6B47;border:1px solid #B5DEC7;} .mpfv-sp3b{background:#E8EFF8;color:#2D4F8A;border:1px solid #B8CEEA;}
    .mpfv-ctool{font-family:var(--mono);font-size:9px;font-weight:600;padding:2px 6px;border-radius:3px;background:var(--bg);border:1px solid var(--line-2);color:var(--ink-3);}
    .mpfv-edit{width:26px;height:24px;border:1px solid var(--line-2);background:var(--bg);border-radius:5px;cursor:pointer;color:var(--ink-4);display:inline-flex;align-items:center;justify-content:center;transition:.14s;}
    .mpfv-edit svg{width:14px;height:14px;} .mpfv-edit:hover{background:var(--ink);color:#fff;border-color:var(--ink);transform:rotate(-8deg);}
    .mpfv-flow{display:flex;align-items:center;gap:7px;font-family:var(--mono);font-size:11px;flex-wrap:wrap;}
    .mpfv-chip{display:inline-flex;align-items:center;gap:4px;padding:3px 8px;border-radius:4px;font-weight:600;white-space:nowrap;}
    .mpfv-if{background:var(--bg-2);border:1px solid var(--line-2);color:var(--ink);}
    .mpfv-set{border:1px solid;} .mpfv-set.r4{background:#EAF6EE;border-color:#B5DEC7;color:#2D6B47;} .mpfv-set.r3{background:#E8EFF8;border-color:#B8CEEA;color:#2D4F8A;}
    .mpfv-r4{color:#2D6B47;} .mpfv-r3{color:#2D4F8A;}
    .mpfv-arr{color:var(--ink-5);font-size:13px;} .mpfv-val{font-weight:700;}
    .mpfv-fertig{color:#92500D;border:1px solid #E8C896;background:#FBF1E0;}
    .mpfv-cfoot{display:flex;align-items:center;gap:8px;margin-top:8px;}
    .mpfv-waitm{font-family:var(--mono);font-size:9.5px;color:var(--ink-5);flex:1;}
    .mpfv-mv{display:flex;gap:3px;}
    .mpfv-mvb{width:24px;height:22px;border:1px solid var(--line-2);background:var(--bg);border-radius:4px;cursor:pointer;color:var(--ink-3);font-size:10px;}
    .mpfv-mvb:hover{background:var(--bg-2);} .mpfv-mvb:disabled{opacity:.25;}
    .mpfv-ho{display:flex;height:20px;}
    .mpfv-ho .rail{flex:0 0 56px;display:flex;justify-content:center;align-items:center;}
    .mpfv-hobadge{font-family:var(--mono);font-size:8.5px;font-weight:700;padding:2px 7px;border-radius:10px;background:#FBF1E0;border:1px solid #E8C896;color:#92500D;white-space:nowrap;}
    .mpfv-add{margin-left:62px;margin-top:4px;border:1px dashed var(--line-2);border-radius:var(--r-md);background:transparent;color:var(--ink-4);font-family:var(--mono);font-size:11px;padding:9px;cursor:pointer;text-transform:uppercase;letter-spacing:.05em;width:calc(100% - 62px);}
    .mpfv-add:hover{border-color:var(--ink);color:var(--ink);}
    .mpfv-empty{padding:28px 14px;text-align:center;color:var(--ink-5);font-family:var(--mono);font-size:11px;}
    .mpfv-codepanel pre{margin:0;background:#0d0d0d;color:#e5e5e5;padding:13px;font-family:var(--mono);font-size:10.5px;line-height:1.55;overflow:auto;white-space:pre;}
    .mpfv-codepanel .ln{display:block;padding:0 4px;border-radius:2px;}
    .mpfv-codepanel .ln.if{color:#9ece6a;} .mpfv-codepanel .ln.set{color:#7aa2f7;} .mpfv-codepanel .ln.code{color:#e0af68;}
    .mpfv-codepanel .ln.grp{color:#bb9af7;} .mpfv-codepanel .ln.cmt{color:#565f89;} .mpfv-codepanel .ln.mark{color:#f7768e;}
    .mpfv-codepanel .ln.hl{background:rgba(255,255,255,.10);}
    `;
    document.head.appendChild(el('style',{id:'citiVisualCSS'},css));
  }

  /* ---------- premium logo SVG into existing brand marks (optional enhance) ---------- */
  function enhanceLogos(){
    // The original uses CSS ::before/::after on .brand-mark; gradient handled by CSS override.
    // Nothing else needed — keeps original DOM intact.
  }

  /* ---------- SPLASH GATE ---------- */
  function buildSplash(){
    if(document.getElementById('citiSplash'))return;
    var sp=el('div',{id:'citiSplash'});
    sp.appendChild(el('div',{class:'cs-logo'},
      '<svg viewBox="0 0 100 100" fill="none">'
      +'<defs><linearGradient id="csM" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#2c2c2c"/><stop offset="1" stop-color="#000"/></linearGradient>'
      +'<linearGradient id="csR" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#fff"/><stop offset="1" stop-color="#cfcfcf"/></linearGradient></defs>'
      +'<rect x="6" y="6" width="88" height="88" rx="22" fill="url(#csM)" stroke="#3a3a3a"/>'
      +'<g class="cs-spin"><circle cx="50" cy="50" r="27" stroke="url(#csR)" stroke-width="6" stroke-linecap="round" stroke-dasharray="127 43" fill="none"/></g>'
      +'<circle cx="50" cy="50" r="6.5" fill="#fff"/></svg>'));
    var tw=el('div',{style:'display:flex;flex-direction:column;align-items:center;gap:6px;'});
    tw.appendChild(el('div',{class:'cs-name'},'Citi<span>Tool</span>'));
    tw.appendChild(el('div',{class:'cs-sub'},'SyncOperator · CNC'));
    sp.appendChild(tw);
    var gate=el('div',{class:'cs-gate',id:'csGate'});
    gate.appendChild(el('div',{class:'cs-gatelabel'},'Zugangscode'));
    var pin=el('div',{class:'cs-pin',id:'csPin'});
    for(var i=0;i<4;i++)pin.appendChild(el('div',{class:'cs-dot'}));
    gate.appendChild(pin);
    var keys=el('div',{class:'cs-keys',id:'csKeys'});
    ['1','2','3','4','5','6','7','8','9','','0','del'].forEach(function(d){
      if(d==='')keys.appendChild(el('div',{class:'cs-key empty'}));
      else if(d==='del')keys.appendChild(el('div',{class:'cs-key wide','data-act':'del'},'⌫'));
      else keys.appendChild(el('div',{class:'cs-key','data-d':d},d));
    });
    gate.appendChild(keys);
    gate.appendChild(el('div',{class:'cs-hint',id:'csHint'}));
    sp.appendChild(gate);
    document.body.appendChild(sp);

    var CODE='3504', buf='';
    function dots(){return sp.querySelectorAll('.cs-dot');}
    function upd(){dots().forEach(function(d,i){d.classList.toggle('on',i<buf.length);});}
    function ok(){ sp.classList.add('hide'); setTimeout(function(){ if(sp.parentNode)sp.parentNode.removeChild(sp); },600); }
    function fail(){ gate.classList.add('shake'); document.getElementById('csHint').textContent='Falscher Code'; setTimeout(function(){gate.classList.remove('shake');buf='';upd();document.getElementById('csHint').textContent='';},500); }
    function push(d){ if(buf.length>=4)return; buf+=d; upd(); if(buf.length===4)setTimeout(function(){ if(buf===CODE)ok(); else fail(); },150); }
    function del(){ buf=buf.slice(0,-1); upd(); }
    keys.addEventListener('click',function(e){var k=e.target.closest('.cs-key');if(!k)return;if(k.getAttribute('data-act')==='del')del();else if(k.getAttribute('data-d')!=null)push(k.getAttribute('data-d'));});
    document.addEventListener('keydown',function(e){ if(!document.getElementById('citiSplash')||sp.classList.contains('hide'))return; if(/^[0-9]$/.test(e.key))push(e.key); else if(e.key==='Backspace')del(); });
  }

  /* ---------- NEW VIEW TAB: MPF / Ablauf ---------- */
  var mpfMode='sync';

  function injectTab(){
    var tabs=document.querySelector('.view-tabs');
    if(!tabs||document.getElementById('mpfvTab'))return false;
    var tab=el('button',{class:'view-tab',id:'mpfvTab','data-view':'mpfv'},
      '<span class="vt-num">04</span><span class="vt-name">MPF / Ablauf</span>');
    tabs.appendChild(tab);
    // pane inside workspace
    var ws=document.querySelector('.workspace');
    if(ws && !document.querySelector('[data-pane="mpfv"]')){
      var pane=el('div',{class:'view-pane','data-pane':'mpfv'});
      pane.innerHTML=
        '<div class="mpfv-bar">'
        +'<div class="mpfv-seg" id="mpfvSeg">'
        +'<button data-m="1">Kanal 1 <span class="mf">1000</span></button>'
        +'<button data-m="2">Kanal 2 <span class="mf">2000</span></button>'
        +'<button data-m="sync" class="on">Synchron <span class="mf">1000+2000</span></button>'
        +'</div><div style="flex:1"></div>'
        +'<button class="btn btn-ghost btn-sm" id="mpfvCodeBtn"><svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 4L2 8l3 4M11 4l3 4-3 4"/></svg>Code anzeigen</button>'
        +'<button class="btn btn-primary btn-sm" id="mpfvDlBtn"><svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M8 2v9M4 7l4 4 4-4M2 14h12"/></svg>MPF generieren</button>'
        +'</div><div class="mpfv-wrap split" id="mpfvCols"></div>';
      // insert as first child of workspace so it sits with other panes
      ws.insertBefore(pane, ws.firstChild);
    }
    // hook tab click into app's setView (extend it)
    tab.addEventListener('click',function(){ showMpfvView(); });
    // also intercept other tabs to hide our pane via app's setView (app handles .active by data-view)
    return true;
  }

  function showMpfvView(){
    // leverage app: set body view + toggle panes/tabs
    document.body.setAttribute('data-view','mpfv');
    document.querySelectorAll('.view-tab').forEach(function(b){b.classList.toggle('active',b.getAttribute('data-view')==='mpfv');});
    document.querySelectorAll('.view-pane').forEach(function(p){p.classList.toggle('active',p.getAttribute('data-pane')==='mpfv');});
    if(typeof S!=='undefined'){ S.view='mpfv'; }
    renderMpfv();
  }

  /* patch app setView so selecting other tabs deactivates our pane and tab properly */
  function patchSetView(){
    if(typeof window.setView!=='function'||window.setView.__mpfvPatched)return;
    var orig=window.setView;
    window.setView=function(v){
      if(v==='mpfv'){ showMpfvView(); return; }
      // ensure our tab/pane removed from active
      var t=document.getElementById('mpfvTab'); if(t)t.classList.remove('active');
      var p=document.querySelector('[data-pane="mpfv"]'); if(p)p.classList.remove('active');
      return orig.apply(this,arguments);
    };
    window.setView.__mpfvPatched=true;
  }

  /* ---------- render chain ---------- */
  function chainHTML(k){
    var st=steps(k);
    if(!st.length)return '<div class="mpfv-empty">Keine belegten Slots in Kanal '+k+'<br>— Operationen im Programmplan anlegen —</div>';
    var h='<div class="mpfv-endcap">▼ Rohteil — '+st[0].ifReg+'==1</div>';
    st.forEach(function(s){
      var spCls=s.spindle==='SP3'?'sp3':'sp4';
      var setCls=s.assignReg==='RG703'?'r3':'r4';
      var ifRegCls=s.ifReg==='RG703'?'mpfv-r3':'mpfv-r4';
      h+='<div class="mpfv-step '+(s.last?'last':'')+'" draggable="true" data-k="'+k+'" data-idx="'+s.idx+'" data-pos="'+s.pos+'">'
        +'<div class="mpfv-rail"><div class="mpfv-line top"></div><div class="mpfv-node '+spCls+'">'+String(s.pos).padStart(2,'0')+'</div><div class="mpfv-line bot"></div></div>'
        +'<div class="mpfv-card '+spCls+'" data-k="'+k+'" data-idx="'+s.idx+'" data-pos="'+s.pos+'">'
          +'<div class="mpfv-ctop">'
            +'<span class="mpfv-handle" title="Ziehen">⠿</span>'
            +'<span class="mpfv-cstate">'+s.state+'</span>'
            +'<span class="mpfv-ctitle">'+escapeH(s.title||s.code)+'</span>'
            +(s.tno?'<span class="mpfv-ctool">'+escapeH(s.tno)+'</span>':'')
            +'<span class="mpfv-clcode">'+s.code+'</span>'
            +'<span class="mpfv-csp '+(s.spindle==='SP3'?'mpfv-sp3b':'mpfv-sp4b')+'">'+s.spindle+'</span>'
            +'<button class="mpfv-edit" data-edit="'+k+':'+s.idx+'" title="Bearbeiten"><svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M11.5 2.5l2 2L6 12l-2.5.5L4 10z"/><path d="M10.5 3.5l2 2"/></svg></button>'
          +'</div>'
          +'<div class="mpfv-flow">'
            +'<span class="mpfv-chip mpfv-if">IF <b class="'+ifRegCls+'">'+s.ifReg+'</b>=='+'<span class="mpfv-val">'+s.ifVal+'</span></span>'
            +'<span class="mpfv-arr">→</span>'
            +'<span class="mpfv-chip mpfv-set '+setCls+'">'+s.assignReg+'=<span class="mpfv-val">'+s.assignVal+'</span></span>'
            +(s.last?'<span class="mpfv-arr">→</span><span class="mpfv-chip mpfv-fertig">Fertigteil</span>':'')
          +'</div>'
          +'<div class="mpfv-cfoot"><span class="mpfv-waitm">WAITM('+(s.w1!==''?s.w1:s.pos)+','+s.w2+','+s.w3+')</span>'
            +'<span class="mpfv-mv"><button class="mpfv-mvb" data-mv="up" data-k="'+k+'" data-idx="'+s.idx+'"'+(s.pos===1?' disabled':'')+'>▲</button><button class="mpfv-mvb" data-mv="down" data-k="'+k+'" data-idx="'+s.idx+'"'+(s.pos===st.length?' disabled':'')+'>▼</button></span>'
          +'</div>'
        +'</div></div>';
      if(s.handoff)h+='<div class="mpfv-ho"><div class="rail"><span class="mpfv-hobadge">⇄ Übergabe → '+s.assignReg+'</span></div><div style="flex:1"></div></div>';
    });
    h+='<button class="mpfv-add" data-add="'+k+'">+ Slot</button>';
    return h;
  }
  function escapeH(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}

  function panelHTML(k){
    var fn=k==='1'?'1000.MPF':'2000.MPF';
    return '<div class="mpfv-panel"><div class="mpfv-phead"><span class="mpfv-ptag">Kanal '+k+'</span><span class="mpfv-ptitle">'+fn+'</span></div>'
      +'<div class="mpfv-legend"><div class="mpfv-ll"><span class="d" style="background:#2D6B47"></span> SP4 → <b style="color:#2D6B47;margin-left:3px">RG704</b></div><div class="mpfv-ll"><span class="d" style="background:#2D4F8A"></span> SP3 → <b style="color:#2D4F8A;margin-left:3px">RG703</b></div></div>'
      +'<div class="mpfv-chain" data-chain="'+k+'">'+chainHTML(k)+'</div></div>';
  }
  function codePanelHTML(k){
    return '<div class="mpfv-panel mpfv-codepanel"><div class="mpfv-phead"><span class="mpfv-ptag">Code</span><span class="mpfv-ptitle">'+(k==='1'?'1000.MPF':'2000.MPF')+'</span></div><pre id="mpfvInlinePre"></pre></div>';
  }

  function renderMpfv(){
    var cols=document.getElementById('mpfvCols');
    if(!cols)return;
    if(mpfMode==='sync'){ cols.className='mpfv-wrap split'; cols.innerHTML=panelHTML('1')+panelHTML('2'); }
    else{ cols.className='mpfv-wrap split'; cols.innerHTML=panelHTML(mpfMode)+codePanelHTML(mpfMode); }
    bindMpfv();
    if(mpfMode!=='sync')paintInline(mpfMode);
    // seg active state
    document.querySelectorAll('#mpfvSeg button').forEach(function(b){b.classList.toggle('on',b.getAttribute('data-m')===mpfMode);});
  }

  var inlineMap={};
  function fileText(k){ return (window.CitiMPF&&window.CitiMPF.generate)?window.CitiMPF.generate(k):''; }
  function blockText(k){
    // build wrapped block from live steps using CitiMPF internal rules via generate() but we want only the marker-wrapped block
    var full=fileText(k);
    return full;
  }
  function paintInto(pre,text){
    inlineMap={};
    pre.innerHTML=''; var cur=0;
    text.split('\n').forEach(function(ln){
      if(ln==='')return;
      var sp=el('span',{class:'ln'});
      var g=ln.match(/GROUP_BEGIN\(0,"(\d+):/); if(g)cur=parseInt(g[1],10);
      var cls='cmt';
      if(/_GENERATED_CODE_(START|END)_/.test(ln)||/^;\*{4,}/.test(ln))cls='mark';
      else if(/GROUP_BEGIN\(0,"\d+:/.test(ln))cls='grp';
      else if(/^IF /.test(ln))cls='if';
      else if(/^\s*RG70[34]=/.test(ln))cls='set';
      else if(/^\s*L\d+\s*$/.test(ln))cls='code';
      else if(/GROUP_(BEGIN|END)/.test(ln))cls='grp';
      sp.classList.add(cls); sp.textContent=ln; sp.setAttribute('data-pos',cur);
      (inlineMap[cur]=inlineMap[cur]||[]).push(sp);
      pre.appendChild(sp);
    });
  }
  function paintInline(k){ var pre=document.getElementById('mpfvInlinePre'); if(pre)paintInto(pre,fileText(k)); }
  function inlineHover(pos){ Object.keys(inlineMap).forEach(function(p){ inlineMap[p].forEach(function(l){ l.classList.toggle('hl', pos!=null && +p===pos); }); }); }

  function bindMpfv(){
    document.querySelectorAll('[data-edit]').forEach(function(elm){
      elm.addEventListener('click',function(e){ e.stopPropagation(); var parts=elm.getAttribute('data-edit').split(':'); openSlotIfAvail(parts[0],+parts[1]); });
    });
    document.querySelectorAll('.mpfv-card').forEach(function(c){
      c.addEventListener('click',function(){ openSlotIfAvail(c.getAttribute('data-k'),+c.getAttribute('data-idx')); });
      c.addEventListener('mouseenter',function(){ var pos=+c.getAttribute('data-pos'); hoverCards(pos); inlineHover(pos); });
      c.addEventListener('mouseleave',function(){ hoverCards(null); inlineHover(null); });
    });
    document.querySelectorAll('[data-add]').forEach(function(b){
      b.addEventListener('click',function(){ var k=b.getAttribute('data-add'); addSlot(k); });
    });
    document.querySelectorAll('.mpfv-mvb').forEach(function(b){
      b.addEventListener('click',function(e){ e.stopPropagation(); moveSlot(b.getAttribute('data-k'),+b.getAttribute('data-idx'),b.getAttribute('data-mv')==='up'?-1:1); });
    });
    bindDnd();
  }
  function hoverCards(pos){ document.querySelectorAll('.mpfv-card').forEach(function(c){ c.classList.toggle('hl', pos!=null && +c.getAttribute('data-pos')===pos); }); }

  function openSlotIfAvail(k,idx){
    if(typeof openSlotEditor==='function'){ openSlotEditor(k,idx); /* re-render after modal closes */ hookModalClose(); }
  }
  function hookModalClose(){
    var modal=document.getElementById('modal'); if(!modal)return;
    var obs=new MutationObserver(function(){ if(!modal.classList.contains('open')){ obs.disconnect(); if(document.body.getAttribute('data-view')==='mpfv')renderMpfv(); } });
    obs.observe(modal,{attributes:true,attributeFilter:['class']});
  }
  function addSlot(k){
    if(typeof S==='undefined')return;
    S.slots[k].push(null);
    if(typeof save==='function')save();
    // open editor for the new slot index
    var idx=S.slots[k].length-1;
    if(typeof openSlotEditor==='function'){ openSlotEditor(k,idx); hookModalClose(); }
    else renderMpfv();
  }
  function moveSlot(k,idx,dir){
    if(typeof S==='undefined')return;
    var arr=S.slots[k]; var j=idx+dir; if(j<0||j>=arr.length)return;
    var t=arr[idx];arr[idx]=arr[j];arr[j]=t;
    if(typeof save==='function')save();
    if(typeof renderSlots==='function')renderSlots();
    renderMpfv();
  }
  var dragSrc=null;
  function bindDnd(){
    document.querySelectorAll('.mpfv-step[draggable="true"]').forEach(function(row){
      row.addEventListener('dragstart',function(e){ dragSrc={k:row.getAttribute('data-k'),idx:+row.getAttribute('data-idx')}; e.dataTransfer.effectAllowed='move'; });
      row.addEventListener('dragover',function(e){ e.preventDefault(); });
      row.addEventListener('drop',function(e){ e.preventDefault(); if(!dragSrc||dragSrc.k!==row.getAttribute('data-k'))return; var k=row.getAttribute('data-k'),ti=+row.getAttribute('data-idx'),fi=dragSrc.idx; if(fi===ti)return; var arr=S.slots[k]; var m=arr.splice(fi,1)[0]; arr.splice(ti,0,m); if(typeof save==='function')save(); if(typeof renderSlots==='function')renderSlots(); renderMpfv(); });
    });
  }

  function bindControls(){
    var seg=document.getElementById('mpfvSeg');
    if(seg&&!seg.__bound){ seg.__bound=true; seg.addEventListener('click',function(e){ var b=e.target.closest('button'); if(!b)return; mpfMode=b.getAttribute('data-m'); renderMpfv(); }); }
    var cb=document.getElementById('mpfvCodeBtn');
    if(cb&&!cb.__bound){ cb.__bound=true; cb.addEventListener('click',function(){ if(window.CitiMPF)window.CitiMPF.open(); }); }
    var db=document.getElementById('mpfvDlBtn');
    if(db&&!db.__bound){ db.__bound=true; db.addEventListener('click',function(){ if(window.CitiMPF)window.CitiMPF.open(); }); }
  }

  /* ---------- start ---------- */
  function start2(){
    injectStyles();
    buildSplash();
    var tries=0;
    var iv=setInterval(function(){
      tries++;
      var done = injectTab();
      patchSetView();
      if(done){ bindControls(); }
      if(done || tries>40) clearInterval(iv);
    },150);
  }

  ready(start2);
})();
