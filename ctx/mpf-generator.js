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
