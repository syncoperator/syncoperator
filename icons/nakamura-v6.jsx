import React,{useState,useRef,useEffect,useCallback}from"react";
const G=`@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@400;500;600;700&display=swap');
*{box-sizing:border-box;margin:0;padding:0;}
body{background:#F5F6FA;-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale;}
:root{
  --bg:#F5F6FA;--p:#FFFFFF;--r:#FAFBFD;--g:#F0F2F7;
  --b1:#E5E8EF;--b2:#D1D5DD;
  --ink:#0A0E1A;--inkS:#4A5168;--inkM:#8B92A8;--inkL:#B8BEC9;
  --blue:#0066FF;--blueD:#0052D4;--blueL:#E8F0FF;--blueG:linear-gradient(135deg,#0066FF,#5B6BFF);
  --green:#00C896;--greenL:#E0FAF1;
  --amber:#FF9500;--amberL:#FFF4E0;
  --red:#FF3B30;--redL:#FFEBEA;
  --violet:#7B5BFF;--violetL:#EFEAFF;
  --rose:#FF2D7B;--roseL:#FFE6F0;
  --cyan:#00C4D4;--cyanL:#E0F9FB;
  --shadow-sm:0 1px 2px rgba(10,14,26,.04),0 1px 1px rgba(10,14,26,.02);
  --shadow:0 4px 12px rgba(10,14,26,.06),0 1px 3px rgba(10,14,26,.04);
  --shadow-lg:0 20px 50px rgba(10,14,26,.12),0 4px 12px rgba(10,14,26,.06);
}
.app{display:flex;flex-direction:column;height:100vh;background:var(--bg);color:var(--ink);overflow:hidden;font-family:'Inter',-apple-system,sans-serif;letter-spacing:-.01em;}
.hdr{flex-shrink:0;background:var(--p);border-bottom:1px solid var(--b1);box-shadow:var(--shadow-sm);}
.hdr1{display:flex;align-items:center;height:58px;}
.logo{display:flex;align-items:center;gap:11px;padding:0 18px;height:100%;border-right:1px solid var(--b1);flex-shrink:0;}
.lm{width:36px;height:36px;border-radius:10px;background:var(--blueG);display:flex;align-items:center;justify-content:center;font-family:'JetBrains Mono',monospace;font-size:15px;font-weight:800;color:#fff;box-shadow:0 4px 12px rgba(0,102,255,.35);}
.lt{font-size:15px;font-weight:700;letter-spacing:-.02em;line-height:1.2;color:var(--ink);}
.ls{font-family:'JetBrains Mono',monospace;font-size:9px;color:var(--inkM);letter-spacing:.5px;margin-top:2px;font-weight:500;}
.spsel{display:flex;height:100%;border-right:1px solid var(--b1);flex-shrink:0;padding:8px;gap:4px;}
.spb{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:0 14px;border:none;cursor:pointer;border-radius:8px;transition:all .2s;gap:2px;background:transparent;}
.spb.on{background:var(--blueL);}
.spid{font-size:13px;font-weight:700;letter-spacing:-.01em;color:var(--inkM);font-family:'JetBrains Mono',monospace;}
.spb.on .spid{color:var(--blue);}
.spsub{font-size:9px;color:var(--inkL);font-weight:500;font-family:'JetBrains Mono',monospace;}
.spb.on .spsub{color:var(--blue);opacity:.7;}
.statrow{display:flex;align-items:center;flex:1;overflow:hidden;padding:0 4px;gap:2px;}
.stat{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:0 14px;height:42px;gap:2px;border-radius:8px;}
.stat:hover{background:var(--r);}
.statv{font-family:'JetBrains Mono',monospace;font-size:15px;font-weight:700;letter-spacing:-.02em;line-height:1;}
.statl{font-size:9px;font-weight:600;letter-spacing:.6px;text-transform:uppercase;color:var(--inkM);font-family:'JetBrains Mono',monospace;}
.bi{width:54px;padding:5px 6px;background:var(--r);border:1px solid var(--b1);border-radius:6px;font-family:'JetBrains Mono',monospace;font-size:13px;font-weight:600;color:var(--ink);outline:none;text-align:center;transition:all .2s;}
.bi:focus{border-color:var(--blue);background:#fff;box-shadow:0 0 0 3px rgba(0,102,255,.1);}
.tabs{display:flex;height:42px;overflow-x:auto;scrollbar-width:none;padding:0 6px;gap:2px;}
.tabs::-webkit-scrollbar{display:none;}
.tab{display:flex;align-items:center;gap:7px;padding:0 14px;border:none;cursor:pointer;font-family:'Inter',sans-serif;font-size:13px;font-weight:600;letter-spacing:-.01em;white-space:nowrap;color:var(--inkM);background:transparent;transition:all .2s;flex-shrink:0;border-radius:8px;margin:4px 0;}
.tab.on{color:var(--blue);background:var(--blueL);}
.tab:not(.on):hover{color:var(--inkS);background:var(--r);}
.tdot{width:6px;height:6px;border-radius:50%;background:currentColor;opacity:.7;}
.main{flex:1;overflow:hidden;display:flex;flex-direction:column;min-height:0;background:var(--bg);}
.simw{flex:1;position:relative;min-height:0;margin:10px;border-radius:14px;overflow:hidden;background:#fff;box-shadow:var(--shadow);}
.hud{position:absolute;top:12px;left:12px;background:rgba(255,255,255,.92);backdrop-filter:blur(12px) saturate(1.5);border:1px solid var(--b1);border-radius:10px;padding:8px 12px;font-family:'JetBrains Mono',monospace;font-size:10px;box-shadow:var(--shadow-sm);}
.hudl{color:var(--inkM);font-size:9px;letter-spacing:.4px;font-weight:600;}
.ctrl{flex-shrink:0;background:var(--p);border-top:1px solid var(--b1);padding:12px;display:flex;flex-direction:column;gap:10px;}
.prow{display:flex;gap:8px;align-items:center;}
.ibtn{width:40px;height:40px;border-radius:10px;border:1px solid var(--b1);background:var(--p);color:var(--inkS);font-size:17px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all .2s;flex-shrink:0;box-shadow:var(--shadow-sm);}
.ibtn:hover{border-color:var(--blue);color:var(--blue);transform:translateY(-1px);box-shadow:var(--shadow);}
.pbtn{flex:1;height:40px;border-radius:10px;border:none;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px;font-family:'Inter',sans-serif;font-size:14px;font-weight:700;letter-spacing:-.01em;transition:all .2s;}
.pbtn.go{background:linear-gradient(135deg,#00C896,#00B383);color:#fff;box-shadow:0 4px 14px rgba(0,200,150,.35);}
.pbtn.go:hover{transform:translateY(-1px);box-shadow:0 6px 20px rgba(0,200,150,.4);}
.pbtn.stop{background:linear-gradient(135deg,#FF3B30,#E03028);color:#fff;box-shadow:0 4px 14px rgba(255,59,48,.35);}
.ctr{padding:0 14px;height:40px;border-radius:10px;background:var(--r);border:1px solid var(--b1);display:flex;align-items:center;font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:600;color:var(--inkS);flex-shrink:0;min-width:78px;justify-content:center;}
.g0b{padding:0 12px;height:40px;border-radius:10px;border:1px solid var(--b1);cursor:pointer;font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:700;flex-shrink:0;transition:all .2s;white-space:nowrap;background:var(--p);}
.g0b.on{background:var(--blueL);border-color:rgba(0,102,255,.2);color:var(--blue);}
.g0b.off{color:var(--inkM);}
.sld{-webkit-appearance:none;width:100%;height:6px;border-radius:3px;background:var(--g);outline:none;cursor:pointer;}
.sld::-webkit-slider-thumb{-webkit-appearance:none;width:18px;height:18px;border-radius:50%;background:var(--blue);cursor:pointer;box-shadow:0 2px 8px rgba(0,102,255,.4);border:2px solid #fff;}
.chips{display:flex;gap:6px;overflow-x:auto;padding-bottom:2px;scrollbar-width:none;}
.chips::-webkit-scrollbar{display:none;}
.chip{display:flex;align-items:center;gap:6px;padding:5px 11px;border-radius:20px;border:1px solid var(--b1);cursor:pointer;flex-shrink:0;transition:all .2s;font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:700;background:var(--p);}
.chip:hover{transform:translateY(-1px);box-shadow:var(--shadow-sm);}
.dot{width:7px;height:7px;border-radius:50%;}
.zbx{position:absolute;bottom:12px;right:12px;display:flex;flex-direction:column;gap:5px;}
.zb{width:32px;height:32px;border-radius:8px;background:rgba(255,255,255,.95);backdrop-filter:blur(8px);border:1px solid var(--b1);color:var(--inkS);font-size:15px;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all .2s;box-shadow:var(--shadow-sm);}
.zb:hover{border-color:var(--blue);color:var(--blue);transform:translateY(-1px);}
.ew{display:flex;height:100%;font-family:'JetBrains Mono',monospace;font-size:13px;line-height:22px;overflow:hidden;border-radius:12px;background:var(--p);box-shadow:var(--shadow-sm);border:1px solid var(--b1);}
.lns{width:48px;flex-shrink:0;background:var(--r);border-right:1px solid var(--b1);overflow-y:hidden;padding-top:12px;user-select:none;}
.ln{height:22px;line-height:22px;padding-right:10px;text-align:right;font-size:11px;color:var(--inkL);font-weight:500;}
.ln.c{color:var(--blue);background:var(--blueL);font-weight:700;}
.ea{flex:1;position:relative;overflow:hidden;background:#fff;}
.eo{position:absolute;inset:0;padding-top:12px;padding-left:14px;overflow:hidden;pointer-events:none;white-space:pre;text-transform:uppercase;}
.el{height:22px;line-height:22px;}
.el.c{background:rgba(0,102,255,.04);}
.eta{position:absolute;inset:0;padding-top:12px;padding-left:14px;font-family:'JetBrains Mono',monospace;font-size:13px;line-height:22px;background:transparent;border:none;outline:none;color:transparent;caret-color:var(--blue);resize:none;white-space:pre;overflow:auto;width:100%;height:100%;text-transform:uppercase;}
.ei{width:68px;flex-shrink:0;background:var(--r);border-left:1px solid var(--b1);display:flex;flex-direction:column;align-items:center;justify-content:flex-end;padding-bottom:16px;gap:3px;font-family:'JetBrains Mono',monospace;color:var(--inkM);font-size:9px;text-align:center;font-weight:600;}
.eiv{font-size:17px;font-weight:700;color:var(--blue);line-height:1;}
.tc{border-radius:14px;border:1px solid var(--b1);overflow:hidden;background:var(--p);box-shadow:var(--shadow-sm);}
.tch{padding:14px 16px;display:flex;align-items:center;gap:12px;border-bottom:1px solid var(--b1);}
.tcb{width:36px;height:36px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0;font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:700;color:#fff;}
.tcbody{padding:14px 16px;display:flex;flex-direction:column;gap:11px;}
.fl{font-family:'Inter',sans-serif;font-size:11px;font-weight:600;letter-spacing:.3px;color:var(--inkM);display:block;margin-bottom:5px;}
.fi{width:100%;padding:9px 12px;border-radius:9px;background:var(--r);border:1px solid var(--b1);font-family:'JetBrains Mono',monospace;font-size:13px;font-weight:600;color:var(--ink);outline:none;box-sizing:border-box;transition:all .2s;}
.fi:focus{border-color:var(--blue);background:#fff;box-shadow:0 0 0 3px rgba(0,102,255,.1);}
.fsel{width:100%;padding:9px 12px;border-radius:9px;background:var(--r);border:1px solid var(--b1);font-family:'Inter',sans-serif;font-size:13px;font-weight:500;color:var(--ink);outline:none;box-sizing:border-box;}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:10px;}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;}
.slbl{font-family:'Inter',sans-serif;font-size:11px;font-weight:700;letter-spacing:.3px;color:var(--inkM);border-bottom:1px solid var(--b1);padding-bottom:6px;margin-bottom:6px;}
.kes{display:flex;gap:6px;overflow-x:auto;padding:10px 12px 6px;align-items:center;scrollbar-width:none;}
.kes::-webkit-scrollbar{display:none;}
.kec{display:flex;align-items:stretch;gap:0;flex-shrink:0;border-radius:11px;transition:all .2s;}
.kecb{padding:9px 13px;cursor:pointer;border-radius:10px 0 0 10px;min-width:158px;transition:all .2s;border:1px solid var(--b1);border-right:none;background:var(--p);}
.kecb.on{background:var(--blueL);border-color:rgba(0,102,255,.3);}
.kecb:hover{background:var(--r);}
.keadd{width:28px;background:var(--greenL);border:1px solid rgba(0,200,150,.25);border-left:none;cursor:pointer;color:var(--green);font-size:16px;font-weight:700;border-radius:0 10px 10px 0;display:flex;align-items:center;justify-content:center;transition:all .2s;}
.keadd:hover{background:var(--green);color:#fff;}
.ker1{display:flex;align-items:center;gap:6px;margin-bottom:3px;}
.ken{width:20px;height:20px;border-radius:5px;display:flex;align-items:center;justify-content:center;flex-shrink:0;font-family:'JetBrains Mono',monospace;font-size:9px;font-weight:700;}
.ket{font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:700;}
.kec2{font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:700;color:var(--ink);text-transform:none;}
.ketags{display:flex;gap:5px;padding-left:26px;}
.ketag{font-family:'JetBrains Mono',monospace;font-size:9px;font-weight:700;padding:2px 6px;border-radius:4px;}
.keset{display:flex;border-top:1px solid var(--b1);overflow-x:auto;background:var(--r);}
.kesc{display:flex;align-items:center;gap:6px;padding:7px 12px;border-right:1px solid var(--b1);flex-shrink:0;}
.kesl{font-size:10px;color:var(--inkM);white-space:nowrap;font-family:'Inter',sans-serif;font-weight:600;}
.kesi{padding:4px 8px;background:var(--p);border:1px solid var(--b1);border-radius:6px;font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:600;color:var(--ink);outline:none;}
.kesi:focus{border-color:var(--blue);}
.ketog{padding:5px 10px;font-size:10px;font-weight:700;border-radius:7px;cursor:pointer;font-family:'JetBrains Mono',monospace;border:1px solid var(--b1);background:var(--p);color:var(--inkM);transition:all .2s;letter-spacing:.3px;}
.ketog.on{background:var(--blueL);border-color:rgba(0,102,255,.3);color:var(--blue);}
.mbg{position:fixed;inset:0;background:rgba(10,14,26,.4);backdrop-filter:blur(20px) saturate(1.8);display:flex;align-items:center;justify-content:center;z-index:200;padding:14px;}
.mbox{background:var(--p);border:1px solid var(--b1);border-radius:20px;width:100%;max-width:480px;max-height:92vh;overflow:auto;box-shadow:var(--shadow-lg);}
.mhdr{display:flex;align-items:center;gap:12px;padding:18px 20px 16px;border-bottom:1px solid var(--b1);position:sticky;top:0;background:var(--p);z-index:2;border-radius:20px 20px 0 0;}
.mbadge{width:42px;height:42px;border-radius:12px;background:var(--blueG);display:flex;align-items:center;justify-content:center;font-family:'JetBrains Mono',monospace;font-size:15px;font-weight:700;color:#fff;box-shadow:0 4px 14px rgba(0,102,255,.35);flex-shrink:0;}
.mtitle{font-family:'Inter',sans-serif;font-size:18px;font-weight:700;color:var(--ink);letter-spacing:-.02em;}
.msub{font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--inkM);margin-top:2px;font-weight:500;}
.mnav{display:flex;gap:5px;margin-left:auto;}
.nb{width:32px;height:32px;border-radius:8px;background:var(--r);border:1px solid var(--b1);color:var(--inkS);cursor:pointer;font-size:15px;display:flex;align-items:center;justify-content:center;transition:all .2s;}
.nb:hover{border-color:var(--blue);color:var(--blue);}
.nb:disabled{opacity:.3;cursor:default;}
.xcb{width:32px;height:32px;border-radius:16px;background:var(--r);border:1px solid var(--b1);color:var(--inkM);cursor:pointer;font-size:17px;display:flex;align-items:center;justify-content:center;transition:all .2s;}
.xcb:hover{background:var(--red);color:#fff;border-color:var(--red);}
.mbody{padding:18px 20px;display:flex;flex-direction:column;gap:14px;}
.mtog{display:flex;gap:6px;}
.mtb{flex:1;padding:11px;border-radius:11px;cursor:pointer;font-family:'Inter',sans-serif;font-size:13px;font-weight:700;border:1.5px solid var(--b1);transition:all .2s;background:var(--p);color:var(--inkM);letter-spacing:-.01em;}
.mtb.g1{background:var(--blueL);border-color:rgba(0,102,255,.3);color:var(--blue);}
.mtb.g0{background:var(--amberL);border-color:rgba(255,149,0,.3);color:var(--amber);}
.mtb.g2,.mtb.g3{background:var(--cyanL);border-color:rgba(0,196,212,.3);color:var(--cyan);}
.sec{border-radius:12px;border:1px solid var(--b1);padding:14px 16px;background:var(--r);}
.sec.v{background:var(--violetL);border-color:rgba(123,91,255,.2);}
.sec.r2{background:var(--roseL);border-color:rgba(255,45,123,.2);}
.stit{font-family:'Inter',sans-serif;font-size:11px;font-weight:700;letter-spacing:.3px;margin-bottom:12px;}
.ag{display:grid;grid-template-columns:1fr 46px 1fr;gap:10px;align-items:end;margin-bottom:10px;}
.db2{width:46px;height:46px;border-radius:11px;cursor:pointer;background:var(--p);border:1.5px solid rgba(123,91,255,.3);color:var(--violet);font-size:18px;display:flex;align-items:center;justify-content:center;margin-top:20px;transition:all .2s;}
.db2:hover{background:var(--violet);color:#fff;}
.pg{display:flex;flex-wrap:wrap;gap:5px;}
.pb{padding:5px 10px;border-radius:7px;cursor:pointer;font-family:'JetBrains Mono',monospace;font-size:10px;font-weight:700;white-space:nowrap;border:1px solid;transition:all .15s;}
.ref{margin-top:9px;padding:9px 12px;border-radius:8px;font-family:'JetBrains Mono',monospace;font-size:10px;line-height:1.8;background:rgba(255,255,255,.6);}
.mfoot{display:flex;gap:9px;padding:0 20px 22px;}
.delbtn{width:48px;height:48px;border-radius:12px;cursor:pointer;background:var(--redL);border:1px solid rgba(255,59,48,.2);color:var(--red);font-size:19px;display:flex;align-items:center;justify-content:center;flex-shrink:0;transition:all .2s;}
.delbtn:hover{background:var(--red);color:#fff;}
.cancbtn{flex:1;height:48px;border-radius:12px;cursor:pointer;background:var(--p);border:1px solid var(--b1);color:var(--inkS);font-family:'Inter',sans-serif;font-size:14px;font-weight:600;transition:all .2s;}
.cancbtn:hover{background:var(--r);}
.savebtn{flex:2;height:48px;border-radius:12px;cursor:pointer;background:var(--blueG);border:none;color:#fff;font-family:'Inter',sans-serif;font-size:14px;font-weight:700;letter-spacing:-.01em;box-shadow:0 4px 14px rgba(0,102,255,.35);transition:all .2s;}
.savebtn:hover{transform:translateY(-1px);box-shadow:0 6px 20px rgba(0,102,255,.45);}
.sbar{flex-shrink:0;height:30px;background:var(--p);border-top:1px solid var(--b1);display:flex;align-items:center;gap:16px;padding:0 14px;font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--inkM);overflow:hidden;font-weight:500;}
.sdot{width:7px;height:7px;border-radius:50%;}
.ita{width:100%;min-height:200px;padding:14px 16px;background:var(--r);border:1px solid var(--b1);border-radius:11px;font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:500;color:var(--ink);outline:none;resize:none;line-height:21px;}
.ita:focus{border-color:var(--blue);background:#fff;}

/* DESKTOP LAYOUT */
@media (min-width:1024px){
  .hdr1{height:62px;}
  .lm{width:42px;height:42px;font-size:17px;}
  .lt{font-size:17px;}
  .ls{font-size:10px;}
  .spid{font-size:15px;}
  .statv{font-size:18px;}
  .statl{font-size:10px;}
  .tab{font-size:14px;padding:0 18px;}
  .tabs{height:48px;}
  
  /* Desktop split-view layouts */
  .dsk-sim{display:grid;grid-template-columns:1fr 380px;gap:12px;height:100%;}
  .dsk-sim-main{display:flex;flex-direction:column;min-height:0;gap:10px;}
  .dsk-sim-side{display:flex;flex-direction:column;gap:12px;overflow-y:auto;padding-right:4px;}
  .dsk-ke{display:grid;grid-template-columns:1fr 480px;gap:12px;height:100%;padding:12px;}
  .dsk-ke-canvas{position:relative;background:var(--p);border-radius:14px;overflow:hidden;box-shadow:var(--shadow);border:1px solid var(--b1);}
  .dsk-ke-side{display:flex;flex-direction:column;gap:10px;min-height:0;}
  .dsk-tools{display:grid;grid-template-columns:280px 1fr;gap:12px;height:100%;padding:12px;}
  .dsk-tools-list{background:var(--p);border-radius:14px;padding:14px;overflow-y:auto;box-shadow:var(--shadow-sm);border:1px solid var(--b1);}
  .dsk-tools-edit{overflow-y:auto;padding-right:4px;}
  .dsk-editor{display:grid;grid-template-columns:1fr 320px;gap:12px;height:100%;padding:12px;}
  
  /* Bigger chips on desktop */
  .kecb{min-width:180px;padding:11px 15px;}
  .kec2{font-size:13px;}
  .ken{width:24px;height:24px;font-size:11px;}
  
  /* Modal wider on desktop */
  .mbox{max-width:560px;}
  
  /* Sidebar panel */
  .side-panel{background:var(--p);border-radius:14px;padding:14px;box-shadow:var(--shadow-sm);border:1px solid var(--b1);}
  .side-title{font-size:12px;font-weight:700;color:var(--inkM);letter-spacing:.5px;text-transform:uppercase;margin-bottom:10px;display:flex;align-items:center;gap:6px;}
  .side-title::before{content:"";width:4px;height:14px;background:var(--blue);border-radius:2px;}
}

/* MOBILE / TABLET — keep current single-view */
@media (max-width:1023px){
  .dsk-sim,.dsk-ke,.dsk-tools,.dsk-editor{display:flex;flex-direction:column;height:100%;}
  .dsk-sim-side,.dsk-ke-side,.dsk-tools-list{display:none;}
  .side-panel{background:transparent;padding:0;box-shadow:none;border:none;}
}
`;
const TC=["#3B82F6","#10B981","#F59E0B","#EF4444","#8B5CF6","#06B6D4","#EC4899","#059669","#F97316","#2563EB","#7C3AED","#0D9488","#D97706","#1E40AF","#065F46"];
const TL={external:"Нар.точение",internal:"Расточная",groove:"Нар.канавка",groove_int:"Вн.канавка",thread:"Резьбовая",drill:"Сверло",mill:"Фреза",brush:"Щётка",stop:"Упор"};
const PRO="%\nO2605(PK231850RO)\nG28U0\nG28W0\nG99M41\nG50S3000\nN1G54M41\nG0G99T0101\n(SCHRUPPSTAHL CNMG 120408MS CA6515 KYOCERA)\nM8\nG96S180M3\nG0Z2\nG0X50\nG1Z.1F.2\nG1X0F.12\nG0X36Z1\nG1,A90\nG1X40.5Z-2,A135\nG1Z-5\nG1,A180\nG1X37.2Z-21.8,A215\nG1Z-23.9\nG1X41.5,R.9\nG1Z-27.9\nG1X45.5,R.9\nG1Z-31.25\nG1X43.15,A210\nG1Z-35.65\nG1X45.3,R.85\nG1Z-51.5F.2\nG1X43.15,A210\nG1Z-56.15\nG1X45.5,R.85\nG1Z-76.6\nG1X49F1\nG0Z2\nG28U0\nM9\nM1\nN2G54M41\nG0G99T0202\n(KOPIERST. AUSSEN FERTIG)\nM8\nG96S180M3\nG0Z2.\nG0X47\nG1Z0F.15\nG1X-0.8\nG0X28Z1\nG1Z0\nG1,A90,R.9F.15\nG1X40Z-2.2,A135,R.9\nG1Z-16\nG1,A180\nG1X36.17Z-21.4,A215F.1\nG1Z-24F.07\nG1X40.94,R.9\nG1Z-28,R.1\nG1X45,R.9F.1\nG1Z-30.7F.15\nG1X43.15,A210\nG1Z-35.75\nG1X45,R.9\nG1Z-51.4F.15\nG1X43.15,A210\nG1Z-56.25\nG1X45,R.9\nG1Z-60\nG1X49F.5\nG0Z2\nG28U0\nM9\nM1\nN3G54M41\nG0G99T0303\n(STECHST. 3MM)\nM8\nG97S700M3\nG0Z2.\nG0X46Z-32\nG1Z-32.25F.3\nG1X45.05F.1\nG1Z-33.25,R.9F.07\nG1X43F.1\nG1Z-35.75F.15\nG1X46Z-35F.5\nG1Z-52F1\nG1Z-52.75F.3\nG1X45F.1\nG1Z-53.75,R.9F.07\nG1X43F.1\nG1Z-56.25F.15\nG1X46Z-55.9F.5\nG28U0\nM9\nM1\nN4G54M41\nG99T0404\n(KOPIERBOHRST. D16)\nM8\nG97S1700M3\nG0Z2.\nG0X20\nG1Z-20.5F1\nG1X24.7F.1\nG1,A180\nG1X26Z-19.4,A155\nG1Z-21.05\nG1,A270\nG1X19.9Z-21.6,A225,R.9\nG1Z-22.8\nG0Z5\nG28U0M9\nM1\nN10G54M41\nG0G99T1010\n(GEWINDEST. M26X1)\nM8\nG97S900M3\nG0Z5\nG0X23.5Z3\nG76P021060Q61R.05\nG76X26.Z-20.5P613Q61F1.\nG0Z5\nG28U0M9\nM1\nM30\n%";
const PRU="%\nO2605(PK231850RU)\nG28U0\nG28B0M97\nG28W0\nG99M441\nG50S3000\nN1G55M441\nG99G0T0101\n(SCHRUPPSTAHL CNMG 120408MS CA6515 KYOCERA G55)\nM8\nG96S160M54\nG0Z-2.\nG0X46.\nG1Z-.1F.2\nG1X10.F.12\nG0X36.\nG1Z0\nG2X37.59Z.264R.9F.06\nG1X39.473Z1.205F.15\nG2X40.Z1.841R.9F.06\nG1Z15.05F.15\nX42.4\nG2X45.Z16.35R1.3F.06\nG1Z18.F.15\nG28U0\nM9\nM1\nN2G55M441\nG99T0202\n(SCHLICHTSTAHL DCMT 11T304GK G55)\nM8\nG96S180M54\nG0Z-2.\nG0X34\nG1Z0F.1\nG1,A90,R.9\nG1X40Z1.8,A45\nG1Z7\nG1,A0\nG1X37.6Z13.3,A330F.08\nG1Z15.1\nG1X44.9,R1\nG1Z16.5\nG1X47F.2\nG0Z-2\nG28U0\nM9\nM1\nN3G55M441\nG0G99T0303\n(GEWINDE M40X1.5 G55)\nM8\nG97S1000M54\nG0Z-5\nG0X41.Z-5\nG76P020060Q92R0.03\nG76X38.16Z14.5P920Q77F1.5\nG0Z-5\nG28U0\nM9\nM1\nN4G55M427\nG0G99T0404\n(BOHRST. D10 HM G55)\nM8\nG96S120M54\nG0Z-2.\nG0X16.5\nG1Z0F.1\nG1,A270,R.7\nG1X13.04Z1.2,A315,R.7\nG1Z1.5\nG1X11.5\nG0Z-5\nG28U0M9\nM1\nN10G54M428\nG99T1010\n(ABSTECHSTAHL 3MM IC328 ISCAR)\nM8\nG97S1000M3\nG0Z-76.5\nG96G99S80M3\nG1X9F.07\nG28U0\nM5\nM1\nM30\n%";

function parse(t){return t.split("\n").map((raw,i)=>{const tr=raw.trim();if(!tr||tr==="%")return{line:i+1,raw,type:"empty",comment:"",params:{}};const comment=(tr.match(/\(([^)]*)\)/)||[])[1]||"";const clean=tr.replace(/\(.*?\)/g,"").replace(/^\//,"").trim();if(!clean)return{line:i+1,raw,type:"comment",comment,params:{}};const params={},parts=clean.split(",");const re=/([A-Z])([-+]?[\d.]+)/g;let m;while((m=re.exec(parts[0]))!==null)params[m[1]]=parseFloat(m[2]);for(let p=1;p<parts.length;p++){const cm=parts[p].trim().match(/^([ARIKar])([-+]?[\d.]*)/i);if(cm)params["_"+cm[1].toUpperCase()]=cm[2]!==""?parseFloat(cm[2]):null;}return{line:i+1,raw,type:"block",comment,params};});}
function getTools(blocks){const seen=new Map();for(let i=0;i<blocks.length;i++){const b=blocks[i];if(b.type!=="block"||b.params.T===undefined)continue;const num=b.params.T>=100?Math.floor(b.params.T/100):Math.floor(b.params.T);const key="T"+String(num).padStart(2,"0");if(seen.has(key))continue;let cmt="";for(let j=i+1;j<Math.min(i+4,blocks.length);j++)if(blocks[j].comment){cmt=blocks[j].comment;break;}seen.set(key,{key,num,comment:cmt});}return[...seen.values()];}
function getPaths(blocks){let x=0,z=0,tool="T01";return blocks.filter(b=>b.type==="block").flatMap(b=>{const p=b.params;if(p.T!==undefined){const n=p.T>=100?Math.floor(p.T/100):Math.floor(p.T);tool="T"+String(n).padStart(2,"0");}const nx=p.X!==undefined?p.X/2:(p.U!==undefined?x+p.U/2:x);const nz=p.Z!==undefined?p.Z:(p.W!==undefined?z+p.W:z);if(nx===x&&nz===z)return[];const isArc=p.G===2||p.G===3;const path={from:{x,z},to:{x:nx,z:nz},rapid:p.G===0,arc:isArc?p.G:0,arcR:p.R||p._R||0,arcI:p.I||p._I||0,arcK:p.K||p._K||0,tool,angle:p._A,radius:p._R,feed:p.F};x=nx;z=nz;return[path];});}
const fN=v=>{if(v==null||v==="")return"";const n=parseFloat(v);return isNaN(n)?"":parseFloat(n.toFixed(4)).toString();};

function SLine({text}){const tr=text.trim();if(!tr||tr==="%")return (<span style={{color:"#B8BEC9"}}>{text||" "}</span>);if(tr[0]==="("||tr[0]==="/")return (<span style={{color:"#9CA3B5",fontStyle:"italic"}}>{text}</span>);const out=[];let i=0;while(i<tr.length){if(tr[i]==="("){const e=tr.indexOf(")",i);out.push(<span key={i} style={{color:"#9CA3B5",fontStyle:"italic"}}>{tr.slice(i,e>=0?e+1:tr.length)}</span>);i=e>=0?e+1:tr.length;}else if(tr[i]===","){let j=i+1;while(j<tr.length&&/[ARar\d.\-+]/.test(tr[j]))j++;out.push(<span key={i} style={{color:"#F59E0B",fontWeight:700}}>{tr.slice(i,j)}</span>);i=j;}else if(/[A-Z]/.test(tr[i])){let j=i+1;while(j<tr.length&&/[\d.\-+]/.test(tr[j]))j++;const l=tr[i],v=tr.slice(i,j);const s=l==="N"?{color:"#8B5CF6",fontWeight:700}:l==="G"?{color:"#3B82F6",fontWeight:700}:l==="M"?{color:"#10B981",fontWeight:700}:l==="T"?{color:"#F59E0B",fontWeight:700}:(l==="F"||l==="S")?{color:"#F43F5E",fontWeight:600}:"XZUWYBCH".includes(l)?{color:"#0A0E1A",fontWeight:600}:{color:"#7A82A0"};out.push(<span key={i} style={s}>{v} </span>);i=j;}else{out.push(<span key={i} style={{color:"#7A82A0"}}>{tr[i]}</span>);i++;}}return (<span>{out}</span>);}

function useCV(draw,deps){const wR=useRef(null),cR=useRef(null),rR=useRef(null);const go=useCallback(()=>{if(rR.current)cancelAnimationFrame(rR.current);rR.current=requestAnimationFrame(()=>{const w=wR.current,c=cR.current;if(!w||!c)return;const W=w.clientWidth,H=w.clientHeight;if(W<4||H<4)return;const dpr=window.devicePixelRatio||1;c.width=W*dpr;c.height=H*dpr;c.style.width=W+"px";c.style.height=H+"px";const ctx=c.getContext("2d");ctx.scale(dpr,dpr);draw(ctx,W,H);});},[...deps]);useEffect(()=>{const w=wR.current;if(!w)return;const obs=new ResizeObserver(go);obs.observe(w);go();return()=>obs.disconnect();},[go]);return{wR,cR,go};}

function CView({paths,vis,cols,showG0}){const draw=useCallback((ctx,W,H)=>{ctx.fillStyle="#FAFBFD";ctx.fillRect(0,0,W,H);const fp=paths.filter(p=>!vis||vis.has(p.tool));const cut=fp.filter(p=>!p.rapid);if(!cut.length){ctx.fillStyle="#B8BEC9";ctx.font="12px Inter";ctx.textAlign="center";ctx.fillText("No contour",W/2,H/2);return;}let maxX=0,minZ=Infinity,maxZ=-Infinity;for(const p of cut){maxX=Math.max(maxX,Math.abs(p.from.x),Math.abs(p.to.x));minZ=Math.min(minZ,p.from.z,p.to.z);maxZ=Math.max(maxZ,p.from.z,p.to.z);}maxX=Math.max(maxX,3);const rZ=Math.max(maxZ-minZ,5);const pad=30,sc=Math.min((W-pad*2)/rZ,(H/2-pad)/maxX)*0.88;const tC=(x,z)=>({cx:pad+(z-minZ)*sc,cy:H/2-x*sc});ctx.strokeStyle="#EEF0F5";ctx.lineWidth=0.5;for(let g=Math.ceil(minZ/10)*10;g<=maxZ;g+=10){const{cx}=tC(0,g);ctx.beginPath();ctx.moveTo(cx,pad/2);ctx.lineTo(cx,H-pad/2);ctx.stroke();ctx.fillStyle="#9CA3B5";ctx.font="8px monospace";ctx.textAlign="center";ctx.fillText(g,cx,H-4);}ctx.strokeStyle="#E5E8EF";ctx.lineWidth=0.8;ctx.setLineDash([5,4]);const cl=tC(0,minZ),cr=tC(0,maxZ);ctx.beginPath();ctx.moveTo(cl.cx,cl.cy);ctx.lineTo(cr.cx,cr.cy);ctx.stroke();ctx.setLineDash([]);if(showG0){fp.filter(p=>p.rapid).forEach(p=>{const a=tC(p.from.x,p.from.z),b=tC(p.to.x,p.to.z);ctx.beginPath();ctx.moveTo(a.cx,a.cy);ctx.lineTo(b.cx,b.cy);ctx.globalAlpha=0.18;ctx.strokeStyle="#8B92A8";ctx.lineWidth=0.7;ctx.setLineDash([3,4]);ctx.stroke();ctx.globalAlpha=1;ctx.setLineDash([]);});}cut.forEach(p=>{const col=cols[p.tool]||"#3B82F6",a=tC(p.from.x,p.from.z),b=tC(p.to.x,p.to.z);ctx.beginPath();ctx.moveTo(a.cx,a.cy);ctx.lineTo(b.cx,b.cy);ctx.strokeStyle=col;ctx.lineWidth=1.8;ctx.stroke();const am=tC(-p.from.x,p.from.z),bm=tC(-p.to.x,p.to.z);ctx.beginPath();ctx.moveTo(am.cx,am.cy);ctx.lineTo(bm.cx,bm.cy);ctx.strokeStyle=col+"44";ctx.lineWidth=1;ctx.stroke();});},[paths,vis,cols,showG0]);const{wR,cR}=useCV(draw,[draw]);return (<div ref={wR} style={{width:"100%",height:"100%",position:"relative"}}><canvas ref={cR} style={{position:"absolute",top:0,left:0}}/></div>);}

function SView({paths,step,dia,cols,showG0}){const wR=useRef(null),cR=useRef(null),rR=useRef(null),vR=useRef({s:1,ox:0,oy:0}),dR=useRef(null);
const draw=useCallback((ctx,W,H)=>{const{s,ox,oy}=vR.current;ctx.fillStyle="#FAFBFD";ctx.fillRect(0,0,W,H);if(!paths.length){ctx.fillStyle="#B8BEC9";ctx.font="13px Inter";ctx.textAlign="center";ctx.fillText("PRESS START",W/2,H/2);return;}let maxX=dia/2+5,minZ=Infinity,maxZ=-Infinity;for(const p of paths){minZ=Math.min(minZ,p.from.z,p.to.z);maxZ=Math.max(maxZ,p.from.z,p.to.z);}minZ-=4;maxZ+=6;const pad=26,base=Math.min((W-pad*2)/(maxZ-minZ),(H/2-pad)/maxX)*0.85,sc=base*s;const tC=(x,z)=>({cx:pad+(z-minZ)*sc+ox,cy:H/2-x*sc+oy});ctx.strokeStyle="#F0F2F7";ctx.lineWidth=0.4;for(let g=Math.ceil(minZ/10)*10;g<=maxZ;g+=10){const{cx}=tC(0,g);ctx.beginPath();ctx.moveTo(cx,0);ctx.lineTo(cx,H);ctx.stroke();ctx.fillStyle="#B8BEC9";ctx.font="8px monospace";ctx.textAlign="center";ctx.fillText(g,cx,H-3);}const bTL=tC(dia/2,minZ+2);const gr=ctx.createLinearGradient(bTL.cx,bTL.cy,bTL.cx,bTL.cy+dia*sc);gr.addColorStop(0,"#EAF0F8");gr.addColorStop(.5,"#F5F8FC");gr.addColorStop(1,"#EAF0F8");ctx.fillStyle=gr;ctx.strokeStyle="#7090C0";ctx.lineWidth=1;ctx.fillRect(bTL.cx,bTL.cy,(maxZ-minZ-4)*sc,dia*sc);ctx.strokeRect(bTL.cx,bTL.cy,(maxZ-minZ-4)*sc,dia*sc);ctx.strokeStyle="#D1D5DD";ctx.lineWidth=0.8;ctx.setLineDash([7,5]);const cl0=tC(0,minZ),cl1=tC(0,maxZ);ctx.beginPath();ctx.moveTo(cl0.cx,cl0.cy);ctx.lineTo(cl1.cx,cl1.cy);ctx.stroke();ctx.setLineDash([]);const limit=step>=0?step:paths.length;let tx=0,tz=0,lc="#3B82F6";for(let i=0;i<Math.min(limit,paths.length);i++){const p=paths[i],a=tC(p.from.x,p.from.z),b=tC(p.to.x,p.to.z),col=cols[p.tool]||"#3B82F6";if(p.rapid){
  if(!showG0){tx=p.to.x;tz=p.to.z;continue;}
  ctx.globalAlpha=0.22;ctx.strokeStyle="#6B7388";ctx.lineWidth=0.8;ctx.setLineDash([4,5]);
  ctx.beginPath();ctx.moveTo(a.cx,a.cy);ctx.lineTo(b.cx,b.cy);ctx.stroke();
  ctx.globalAlpha=1;ctx.setLineDash([]);
} else {
  const lineCol=col+(i<limit-5?"BB":"FF");
  const lw=i===limit-1?2.5:1.5;
  ctx.globalAlpha=1;ctx.strokeStyle=lineCol;ctx.lineWidth=lw;ctx.setLineDash([]);

  const am=tC(-p.from.x,p.from.z),bm=tC(-p.to.x,p.to.z);

  if(p.arc>0&&p.arcR>0){
    // G2/G3 full arc
    const rPx=p.arcR*sc;
    const dx=b.cx-a.cx,dy=b.cy-a.cy,d=Math.hypot(dx,dy);
    if(d>0&&rPx*2>=d){
      const mx2=(a.cx+b.cx)/2,my2=(a.cy+b.cy)/2;
      const h=Math.sqrt(Math.max(0,rPx*rPx-d*d/4));
      const sign=p.arc===2?-1:1;
      const ocx=mx2+(-dy/d)*h*sign, ocy=my2+(dx/d)*h*sign;
      const a1=Math.atan2(a.cy-ocy,a.cx-ocx),a2=Math.atan2(b.cy-ocy,b.cx-ocx);
      ctx.beginPath();ctx.arc(ocx,ocy,rPx,a1,a2,p.arc===3);ctx.stroke();
      // mirror arc
      const dx2=bm.cx-am.cx,dy2=bm.cy-am.cy,d2=Math.hypot(dx2,dy2);
      if(d2>0&&rPx*2>=d2){
        const mx3=(am.cx+bm.cx)/2,my3=(am.cy+bm.cy)/2;
        const h2=Math.sqrt(Math.max(0,rPx*rPx-d2*d2/4));
        const ocx2=mx3+(-dy2/d2)*h2*sign,ocy2=my3+(dx2/d2)*h2*sign;
        const a3=Math.atan2(am.cy-ocy2,am.cx-ocx2),a4=Math.atan2(bm.cy-ocy2,bm.cx-ocx2);
        ctx.globalAlpha=0.2;ctx.lineWidth=0.9;
        ctx.beginPath();ctx.arc(ocx2,ocy2,rPx,a3,a4,p.arc===3);ctx.stroke();
        ctx.globalAlpha=1;ctx.lineWidth=lw;
      }
    } else {
      ctx.beginPath();ctx.moveTo(a.cx,a.cy);ctx.lineTo(b.cx,b.cy);ctx.stroke();
    }
  } else if(p.radius>0&&i<paths.length-1){
    // ,R corner rounding using next segment
    const next=paths[i+1];
    const c=tC(next.to.x,next.to.z);
    const rPx=p.radius*sc;
    const d1=Math.hypot(b.cx-a.cx,b.cy-a.cy);
    const d2=Math.hypot(c.cx-b.cx,c.cy-b.cy);
    const t1=Math.min(rPx/Math.max(d1,1),0.48);
    const t2=Math.min(rPx/Math.max(d2,1),0.48);
    const tx1=b.cx-(b.cx-a.cx)*t1,ty1=b.cy-(b.cy-a.cy)*t1;
    const tx2=b.cx+(c.cx-b.cx)*t2,ty2=b.cy+(c.cy-b.cy)*t2;
    ctx.beginPath();ctx.moveTo(a.cx,a.cy);ctx.lineTo(tx1,ty1);
    ctx.quadraticCurveTo(b.cx,b.cy,tx2,ty2);ctx.stroke();
    // mirror
    const cm=tC(-next.to.x,next.to.z);
    const dm1=Math.hypot(bm.cx-am.cx,bm.cy-am.cy);
    const dm2=Math.hypot(cm.cx-bm.cx,cm.cy-bm.cy);
    const tm1=Math.min(rPx/Math.max(dm1,1),0.48);
    const tm2=Math.min(rPx/Math.max(dm2,1),0.48);
    const tmx1=bm.cx-(bm.cx-am.cx)*tm1,tmy1=bm.cy-(bm.cy-am.cy)*tm1;
    const tmx2=bm.cx+(cm.cx-bm.cx)*tm2,tmy2=bm.cy+(cm.cy-bm.cy)*tm2;
    ctx.globalAlpha=0.2;ctx.lineWidth=0.9;
    ctx.beginPath();ctx.moveTo(am.cx,am.cy);ctx.lineTo(tmx1,tmy1);
    ctx.quadraticCurveTo(bm.cx,bm.cy,tmx2,tmy2);ctx.stroke();
    ctx.globalAlpha=1;ctx.lineWidth=lw;
  } else {
    ctx.beginPath();ctx.moveTo(a.cx,a.cy);ctx.lineTo(b.cx,b.cy);ctx.stroke();
    ctx.globalAlpha=0.2;ctx.lineWidth=0.9;
    ctx.beginPath();ctx.moveTo(am.cx,am.cy);ctx.lineTo(bm.cx,bm.cy);ctx.stroke();
    ctx.globalAlpha=1;ctx.lineWidth=lw;
  }
}
ctx.globalAlpha=1;ctx.setLineDash([]);tx=p.to.x;tz=p.to.z;lc=col;}if(limit>0&&limit<=paths.length){const tp=tC(tx,tz);const gw=ctx.createRadialGradient(tp.cx,tp.cy,0,tp.cx,tp.cy,18);gw.addColorStop(0,lc+"44");gw.addColorStop(1,"transparent");ctx.fillStyle=gw;ctx.beginPath();ctx.arc(tp.cx,tp.cy,18,0,Math.PI*2);ctx.fill();ctx.beginPath();ctx.arc(tp.cx,tp.cy,7,0,Math.PI*2);ctx.fillStyle=lc;ctx.fill();ctx.strokeStyle="#fff";ctx.lineWidth=1.5;ctx.stroke();}if(s!==1){ctx.fillStyle="rgba(8,10,20,.85)";ctx.fillRect(W-50,H-18,46,14);ctx.fillStyle="#6B7388";ctx.font="9px monospace";ctx.textAlign="center";ctx.fillText(s.toFixed(1)+"x",W-27,H-8);}},[paths,step,dia,cols,showG0]);
const go=useCallback(()=>{if(rR.current)cancelAnimationFrame(rR.current);rR.current=requestAnimationFrame(()=>{const w=wR.current,c=cR.current;if(!w||!c)return;const W=w.clientWidth,H=w.clientHeight;if(W<4||H<4)return;const dpr=window.devicePixelRatio||1;c.width=W*dpr;c.height=H*dpr;c.style.width=W+"px";c.style.height=H+"px";const ctx=c.getContext("2d");ctx.scale(dpr,dpr);draw(ctx,W,H);});},[draw]);
useEffect(()=>{const w=wR.current;if(!w)return;const obs=new ResizeObserver(go);obs.observe(w);go();return()=>obs.disconnect();},[go]);
useEffect(()=>{const c=cR.current;if(!c)return;const fn=e=>{e.preventDefault();const r=c.getBoundingClientRect(),mx=e.clientX-r.left,my=e.clientY-r.top,f=e.deltaY<0?1.15:1/1.15,v=vR.current,ns=Math.max(.2,Math.min(25,v.s*f)),ratio=ns/v.s;vR.current={s:ns,ox:mx-ratio*(mx-v.ox),oy:my-ratio*(my-v.oy)};go();};c.addEventListener("wheel",fn,{passive:false});return()=>c.removeEventListener("wheel",fn);},[go]);
useEffect(()=>{const c=cR.current;if(!c)return;let ld=0;const D=t=>Math.hypot(t[0].clientX-t[1].clientX,t[0].clientY-t[1].clientY);const M=(t,r)=>({x:(t[0].clientX+t[1].clientX)/2-r.left,y:(t[0].clientY+t[1].clientY)/2-r.top});const ts=e=>{if(e.touches.length===2){ld=D(e.touches);}else if(e.touches.length===1){const t=e.touches[0];dR.current={x:t.clientX,y:t.clientY,ox:vR.current.ox,oy:vR.current.oy};}};const tm=e=>{e.preventDefault();if(e.touches.length===2){const d=D(e.touches),m=M(e.touches,c.getBoundingClientRect()),f=d/(ld||d),v=vR.current,ns=Math.max(.2,Math.min(25,v.s*f)),r=ns/v.s;vR.current={s:ns,ox:m.x-r*(m.x-v.ox),oy:m.y-r*(m.y-v.oy)};ld=d;go();}else if(e.touches.length===1&&dR.current){const t=e.touches[0];vR.current.ox=dR.current.ox+(t.clientX-dR.current.x);vR.current.oy=dR.current.oy+(t.clientY-dR.current.y);go();}};const te=()=>{dR.current=null;};c.addEventListener("touchstart",ts,{passive:false});c.addEventListener("touchmove",tm,{passive:false});c.addEventListener("touchend",te);return()=>{c.removeEventListener("touchstart",ts);c.removeEventListener("touchmove",tm);c.removeEventListener("touchend",te);};},[go]);
useEffect(()=>{const c=cR.current;if(!c)return;const dn=e=>{dR.current={x:e.clientX,y:e.clientY,ox:vR.current.ox,oy:vR.current.oy};c.style.cursor="grabbing";};const mv=e=>{if(!dR.current)return;vR.current.ox=dR.current.ox+(e.clientX-dR.current.x);vR.current.oy=dR.current.oy+(e.clientY-dR.current.y);go();};const up=()=>{dR.current=null;c.style.cursor="grab";};c.addEventListener("mousedown",dn);window.addEventListener("mousemove",mv);window.addEventListener("mouseup",up);c.style.cursor="grab";return()=>{c.removeEventListener("mousedown",dn);window.removeEventListener("mousemove",mv);window.removeEventListener("mouseup",up);};},[go]);
const zoom=f=>{const w=wR.current;if(!w)return;const cx=w.clientWidth/2,cy=w.clientHeight/2,v=vR.current,ns=Math.max(.2,Math.min(25,v.s*f)),r=ns/v.s;vR.current={s:ns,ox:cx-r*(cx-v.ox),oy:cy-r*(cy-v.oy)};go();};
return(<div ref={wR} style={{width:"100%",height:"100%",position:"relative"}}><canvas ref={cR} style={{position:"absolute",top:0,left:0,cursor:"grab"}}/><div className="zbx">{[{l:"+",f:()=>zoom(1.3)},{l:"-",f:()=>zoom(1/1.3)},{l:"[]",f:()=>{vR.current={s:1,ox:0,oy:0};go();}}].map(({l,f})=>(<button key={l} className="zb" onClick={f}>{l}</button>))}</div></div>);}

function Editor({value,onChange}){const tR=useRef(null),oR=useRef(null),nR=useRef(null);const[cur,setCur]=useState(1);const lines=value.split("\n");const sync=()=>{const t=tR.current;if(t&&oR.current&&nR.current){oR.current.scrollTop=t.scrollTop;nR.current.scrollTop=t.scrollTop;}};const upd=()=>{const t=tR.current;if(t){const ln=(t.value.slice(0,t.selectionStart).match(/\n/g)||[]).length+1;setCur(ln);}};return(<div className="ew"><div ref={nR} className="lns">{lines.map((_,i)=>(<div key={i} className={"ln"+(cur===i+1?" c":"")}>{i+1}</div>))}</div><div className="ea"><div ref={oR} className="eo">{lines.map((l,i)=>(<div key={i} className={"el"+(cur===i+1?" c":"")}><SLine text={l}/></div>))}</div><textarea ref={tR} value={value} className="eta" onChange={e=>onChange(e.target.value)} onScroll={sync} onClick={upd} onKeyUp={upd} onKeyDown={e=>{if(e.key==="Tab"){e.preventDefault();const t=e.target,s=t.selectionStart,en=t.selectionEnd,nv=value.slice(0,s)+"  "+value.slice(en);onChange(nv);requestAnimationFrame(()=>{t.selectionStart=t.selectionEnd=s+2;});}}} spellCheck={false} autoCapitalize="off" autoCorrect="off"/></div><div className="ei"><div>LINE</div><div className="eiv">{cur}</div><div style={{marginTop:5}}>TOT</div><div style={{fontSize:12,fontWeight:600,color:"var(--inkS)"}}>{lines.length}</div></div></div>);}
function chamferToA(deg,dir,sp){const d=Math.abs(parseFloat(deg));if(!d||isNaN(d))return"";if(sp==="RO")return dir==="up"?180-d:180+d;return dir==="up"?d:360-d;}
function buildKG(rows,cfg){if(!rows.length)return"(empty)";const sp2=cfg.sp==="RU"?"M54":"M3",ws=cfg.sp==="RU"?"G55":"G54";const L=[`(KONTUR ${cfg.sp})`,ws,`G0G99${cfg.tool}`,`G96S${cfg.speed}${sp2}`,"M8",`G0Z${fN(rows[0].z)||"0"}`,`G0X${fN(rows[0].x)||"0"}`];for(let i=1;i<rows.length;i++){const r=rows[i],p=rows[i-1];if(r.move==="G0"){L.push(`G0X${fN(r.x)}Z${fN(r.z)}`);continue;}if(r.move==="G2"||r.move==="G3"){let s=r.move;if(r.x!==""&&Math.abs((+r.x||0)-(+p.x||0))>5e-4)s+=`X${fN(r.x)}`;if(r.z!==""&&Math.abs((+r.z||0)-(+p.z||0))>5e-4)s+=`Z${fN(r.z)}`;if(+r.arcR>0)s+=`R${fN(r.arcR)}`;s+=`F${r.feed||cfg.feed}`;L.push(s);continue;}let s="G1";if(r.x!==""&&Math.abs((+r.x||0)-(+p.x||0))>5e-4)s+=`X${fN(r.x)}`;if(r.z!==""&&Math.abs((+r.z||0)-(+p.z||0))>5e-4)s+=`Z${fN(r.z)}`;if(r.angleA!==""&&r.angleA!=null)s+=`,A${r.angleA}`;if(+r.radius>0)s+=`,R${fN(r.radius)}`;s+=`F${r.feed||cfg.feed}`;L.push(s);}L.push(`G0${cfg.sp==="RU"?"W":"Z"}5`,"G28U0","M9","M30");return L.join("\n");}
function parseImp(text){const rows=[];let lx=0,lz=0;for(const raw of text.split("\n")){const line=raw.trim().replace(/\(.*?\)/g,"").toUpperCase();if(!line||line.startsWith("%")||/^[ON]\d/.test(line))continue;const gn=re=>{const m=line.match(re);return m?parseFloat(m[1]):null;};const G=gn(/G(\d+\.?\d*)/),X=gn(/X([-\d.]+)/),Z=gn(/Z([-\d.]+)/),U=gn(/U([-\d.]+)/),W=gn(/W([-\d.]+)/),F=gn(/F([\d.]+)/);const mA=raw.match(/,A([\d.]+)/i),mR=raw.match(/,R([\d.]+)/i);if(X===null&&Z===null&&U===null&&W===null)continue;const xv=X!==null?X:(U!==null?lx*2+U:lx*2),zv=Z!==null?Z:(W!==null?lz+W:lz);rows.push({move:G===0?"G0":"G1",x:fN(xv),z:fN(zv),angleA:mA?parseFloat(mA[1]):"",radius:mR?parseFloat(mR[1]):"",feed:F?fN(F):"",chamferDeg:"",chamferDir:"up"});lx=xv/2;lz=zv;}return rows;}
const mkR=o=>({move:"G1",x:"",z:"",chamferDeg:"",chamferDir:"up",angleA:"",radius:"",feed:"",arcR:"",...o});
const AP={RO:[{l:"45°↗",a:135},{l:"45°↘",a:225},{l:"30°↗",a:150},{l:"30°↘",a:210},{l:"60°↗",a:120},{l:"90°",a:90},{l:"180°",a:180},{l:"270°",a:270}],RU:[{l:"45°↗",a:45},{l:"45°↘",a:315},{l:"30°↗",a:30},{l:"30°↘",a:330},{l:"60°↗",a:60},{l:"90°",a:90},{l:"0°",a:0},{l:"270°",a:270}]};
const EX_RO=[mkR({move:"G0",x:"55",z:"2"}),mkR({x:"55",z:"0",angleA:"90",feed:"0.15"}),mkR({x:"40",z:"-2",chamferDeg:"45",chamferDir:"down",angleA:"135",feed:"0.15"}),mkR({x:"40",z:"-20",angleA:"180",feed:"0.12"}),mkR({x:"45",z:"-22",chamferDeg:"45",chamferDir:"up",angleA:"135",radius:"0.9",feed:"0.12"}),mkR({x:"45",z:"-35",feed:"0.12"}),mkR({x:"50",z:"-36",chamferDeg:"45",chamferDir:"up",angleA:"135",radius:"1",feed:"0.1"}),mkR({x:"50",z:"-55",feed:"0.15"})];
const EX_RU=[mkR({move:"G0",x:"46",z:"-2"}),mkR({x:"36",z:"0",angleA:"90",feed:"0.15"}),mkR({x:"40",z:"1.8",chamferDeg:"45",chamferDir:"up",angleA:"45",feed:"0.15"}),mkR({x:"40",z:"7",angleA:"0",feed:"0.12"}),mkR({x:"37",z:"13",chamferDeg:"30",chamferDir:"down",angleA:"330",feed:"0.08"}),mkR({x:"45",z:"15",radius:"1",feed:"0.1"}),mkR({x:"47",z:"16.5",feed:"0.15"})];

function KECanvas({rows,sp,aIdx,showNums=true,showTags=true}){const draw=useCallback((ctx,W,H)=>{ctx.fillStyle="#FAFBFD";ctx.fillRect(0,0,W,H);const pts=[];let lastX=0,lastZ=0;rows.forEach((r,i)=>{const hasX=r.x!==""&&!isNaN(+r.x);const hasZ=r.z!==""&&!isNaN(+r.z);// Always add point if at least one coord exists OR if it has angleA/radius (modifies previous segment)
const hasData=hasX||hasZ||(r.angleA!==""&&r.angleA!=null)||(+r.radius>0);if(!hasData)return;const xr=hasX?+r.x/2:lastX;const zv=hasZ?+r.z:lastZ;pts.push({...r,xr,zv,oi:i});if(hasX)lastX=xr;if(hasZ)lastZ=zv;});if(!pts.length){ctx.fillStyle="#B8BEC9";ctx.font="12px Inter";ctx.textAlign="center";ctx.fillText("Enter first point...",W/2,H/2);return;}if(pts.length===1){const p=pts[0];const pad=36,sc=8;const cx=W/2,cy=H/2;ctx.beginPath();ctx.arc(cx,cy,10,0,Math.PI*2);ctx.fillStyle="#10B981";ctx.fill();ctx.strokeStyle="#fff";ctx.lineWidth=2;ctx.stroke();ctx.fillStyle="#F59E0B";ctx.font="bold 11px monospace";ctx.textAlign="center";ctx.fillText("Ø"+fN(+p.xr*2)+" Z"+fN(p.zv),cx,cy-16);ctx.fillStyle="#fff";ctx.font="bold 9px monospace";ctx.fillText("1",cx,cy+4);return;}let maxX=0,minZ=Infinity,maxZ=-Infinity;for(const p of pts){maxX=Math.max(maxX,p.xr);minZ=Math.min(minZ,p.zv);maxZ=Math.max(maxZ,p.zv);}maxX=Math.max(maxX,2);const rZ=Math.max(maxZ-minZ,1);const pad=34,sc=Math.min((W-pad*2)/rZ,(H/2-pad)/maxX)*0.86;const tC=(xr,z)=>({cx:pad+(z-minZ)*sc,cy:H/2-xr*sc});const gs=sc>=20?1:sc>=8?2:sc>=3?5:10;ctx.strokeStyle="#F0F2F7";ctx.lineWidth=0.5;for(let g=Math.ceil(minZ/gs)*gs;g<=maxZ;g+=gs){const{cx}=tC(0,g);ctx.beginPath();ctx.moveTo(cx,pad/2);ctx.lineTo(cx,H-pad/2);ctx.stroke();ctx.fillStyle="#B8BEC9";ctx.font="8px monospace";ctx.textAlign="center";ctx.fillText(g,cx,H-4);}ctx.strokeStyle="#EAEDF2";ctx.lineWidth=0.8;ctx.setLineDash([5,4]);const cL=tC(0,minZ),cRv=tC(0,maxZ);ctx.beginPath();ctx.moveTo(cL.cx,cL.cy);ctx.lineTo(cRv.cx,cRv.cy);ctx.stroke();ctx.setLineDash([]);ctx.fillStyle="#9CA3B5";ctx.font="9px monospace";ctx.textAlign="left";ctx.fillText("X",3,H/2-maxX*sc-5);ctx.fillText(sp==="RU"?"Z(+)":"Z(-)",pad+rZ*sc+3,H/2+4);for(let i=1;i<pts.length;i++){const a=pts[i-1],b=pts[i],pa=tC(a.xr,a.zv),pb=tC(b.xr,b.zv),pam=tC(-a.xr,a.zv),pbm=tC(-b.xr,b.zv);const rapid=b.move==="G0",isAct=b.oi===aIdx||a.oi===aIdx;const R=+b.radius;ctx.strokeStyle=isAct?"#F59E0B":rapid?"#D1D5DD":"#3B82F6";ctx.lineWidth=isAct?3:rapid?1:2.2;ctx.setLineDash(rapid?[4,4]:[]);ctx.globalAlpha=rapid?0.4:1;ctx.beginPath();ctx.moveTo(pa.cx,pa.cy);ctx.lineTo(pb.cx,pb.cy);ctx.stroke();ctx.globalAlpha=isAct?0.25:rapid?0.08:0.2;ctx.strokeStyle=isAct?"#F59E0B":"#3B82F6";ctx.lineWidth=1;ctx.setLineDash([]);ctx.beginPath();ctx.moveTo(pam.cx,pam.cy);ctx.lineTo(pbm.cx,pbm.cy);ctx.stroke();ctx.globalAlpha=1;ctx.setLineDash([]);if(!rapid&&sc>1.2){const mx=(pa.cx+pb.cx)/2,my=(pa.cy+pb.cy)/2-10;const tags=[];if(b.angleA!==""&&b.angleA!=null)tags.push([`,A${b.angleA}`,"#8B5CF6"]);if(R>0)tags.push([`,R${b.radius}`,"#F43F5E"]);tags.forEach(([t,c],ti)=>{ctx.fillStyle=c;ctx.font="bold 9px monospace";ctx.textAlign="center";ctx.fillText(t,mx,my-ti*11);});}}for(let i=0;i<pts.length;i++){const p=pts[i],{cx,cy}=tC(p.xr,p.zv),isAct=p.oi===aIdx;if(isAct){const gw=ctx.createRadialGradient(cx,cy,0,cx,cy,16);gw.addColorStop(0,"rgba(245,158,11,.3)");gw.addColorStop(1,"transparent");ctx.fillStyle=gw;ctx.beginPath();ctx.arc(cx,cy,16,0,Math.PI*2);ctx.fill();}ctx.beginPath();ctx.arc(cx,cy,isAct?8:5,0,Math.PI*2);ctx.fillStyle=isAct?"#F59E0B":i===0?"#10B981":i===pts.length-1?"#EF4444":"#3B82F6";ctx.fill();ctx.strokeStyle="#fff";ctx.lineWidth=1.5;ctx.stroke();if(sc>2){ctx.fillStyle=isAct?"#F59E0B":"#6B7388";ctx.font=(isAct?"bold ":"")+"9px monospace";ctx.textAlign="center";ctx.fillText("O"+fN(+p.xr*2)+" Z"+fN(p.zv),cx,cy-(isAct?15:11));}if(showNums){ctx.fillStyle="#fff";ctx.font="bold 8px monospace";ctx.textAlign="center";ctx.fillText(i+1,cx,cy+3);}}},[rows,sp,aIdx]);const{wR,cR}=useCV(draw,[draw]);return (<div ref={wR} style={{width:"100%",height:"100%",position:"relative"}}><canvas ref={cR} style={{position:"absolute",top:0,left:0}}/></div>);}

function PointModal({pt,idx,total,sp,onSave,onDelete,onClose,onPrev,onNext}){
const[v,setV]=useState({...pt});const set=(k,val)=>setV(p=>({...p,[k]:val}));const xRef=useRef(null);
useEffect(()=>{setTimeout(()=>{if(xRef.current){xRef.current.focus();xRef.current.select();}},80);},[]);
const rapid=v.move==="G0";const fields=rapid?["x","z"]:["x","z","chamferDeg","angleA","radius","feed"];const iR=useRef({});
const hk=(e,fk)=>{if(e.key==="Escape"){onClose();return;}if(e.key!=="Enter"&&e.key!=="Tab")return;e.preventDefault();const fi=fields.indexOf(fk);if(fi<fields.length-1){const r=iR.current[fields[fi+1]];if(r){r.focus();r.select();}}else onSave(idx,v);};
const onCh=(k,val)=>{const deg=parseFloat(k==="chamferDeg"?val:v.chamferDeg),dir=k==="chamferDir"?val:v.chamferDir,a=(!isNaN(deg)&&deg>0)?chamferToA(deg,dir,sp):"";setV(p=>({...p,[k]:val,angleA:a}));};
const inp=(fk,opts={})=>(<input ref={el=>{iR.current[fk]=el;if(fk==="x")xRef.current=el;}} type="number" step={opts.step||0.001} value={v[fk]??""} placeholder={opts.ph||"-"} onChange={e=>opts.onChange?opts.onChange(e.target.value):set(fk,e.target.value)} onKeyDown={e=>hk(e,fk)} onFocus={e=>e.target.select()} className="fi" style={{color:opts.color||"var(--ink)",...(opts.style||{})}}/>);
const presets=AP[sp];
return(<div className="mbg" onClick={e=>{if(e.target===e.currentTarget)onClose();}}><div className="mbox">
<div className="mhdr"><div className="mbadge">{idx+1}</div><div style={{flex:1}}><div className="mtitle">PUNKT {idx+1}/{total}</div><div className="msub">{sp} — {rapid?"G0 EILGANG":"G1 VORSCHUB"}{(v.x||v.z)?" — "+(v.x?"Ø"+fN(+v.x):"")+" "+(v.z?"Z"+fN(+v.z):""):""}</div></div><div className="mnav">{[["‹",()=>{onSave(idx,v);onPrev();},idx===0],["›",()=>{onSave(idx,v);onNext();},idx===total-1]].map(([l,fn,dis])=>(<button key={l} className="nb" onClick={fn} disabled={dis}>{l}</button>))}</div><button className="xcb" onClick={onClose}>×</button></div>
<div className="mbody">
<div className="mtog">{[["G1","✂ G1"],["G0","⚡ G0"],["G2","↻ G2"],["G3","↺ G3"]].map(([mv,lb])=>(<button key={mv} onClick={()=>set("move",mv)} className={"mtb "+(v.move===mv?mv.toLowerCase():"")} style={{fontSize:12}}>{lb}</button>))}</div>
<div className="g2"><div><label className="fl">X DURCHMESSER</label>{inp("x",{ph:"Ø mm",style:{fontSize:17,fontWeight:700}})}</div><div><label className="fl" style={{color:sp==="RU"?"var(--amber)":undefined}}>Z {sp==="RU"?"(+GEGEN)":"(-HAUPT)"}</label>{inp("z",{ph:"Z mm",color:sp==="RU"?"var(--amber)":undefined,style:{fontSize:17,fontWeight:700}})}</div></div>
{(v.move==="G2"||v.move==="G3")&&<div style={{background:"rgba(6,182,212,.08)",border:"1px solid rgba(6,182,212,.3)",borderRadius:10,padding:"12px 13px"}}>
<div className="stit" style={{color:"var(--cyan)"}}>↻ {v.move} — RADIUS R</div>
<div style={{display:"grid",gridTemplateColumns:"1fr auto",gap:10,alignItems:"center"}}>{inp("arcR",{ph:"R mm",step:0.1,color:"var(--cyan)",style:{fontSize:16,fontWeight:800}})}<div style={{display:"flex",gap:4,flexWrap:"wrap"}}>{[0.5,1,1.5,2,3,5,10].map(r=>(<button key={r} onClick={()=>set("arcR",r)} className="pb" style={{borderColor:+v.arcR===r?"var(--cyan)":"rgba(6,182,212,.3)",background:+v.arcR===r?"var(--cyan)":"rgba(6,182,212,.08)",color:+v.arcR===r?"#fff":"var(--cyan)"}}>{r}</button>))}</div></div>
<div style={{fontSize:9,color:"var(--inkM)",fontFamily:"Space Mono",marginTop:6,letterSpacing:.5}}>G2=CW (UHRZEIGER) · G3=CCW (GEGENUHRZEIGER)</div>
</div>}
{!rapid&&v.move!=="G2"&&v.move!=="G3"&&<>
<div className="sec v"><div className="stit" style={{color:"var(--violet)"}}>📐 FASE / KONUS — ,A WINKEL</div>
<div className="ag"><div><label className="fl" style={{color:"var(--violet)"}}>ZEICHNUNG °</label>{inp("chamferDeg",{ph:"z.B. 45",step:1,color:"var(--violet)",style:{fontSize:17,fontWeight:800},onChange:v2=>onCh("chamferDeg",v2)})}</div><div><label className="fl" style={{color:"var(--violet)"}}>DIR</label><button className="db2" onClick={()=>onCh("chamferDir",v.chamferDir==="up"?"down":"up")}>{v.chamferDir==="down"?"↘":"↗"}</button></div><div><label className="fl" style={{color:"var(--violet)",fontWeight:800}}>,A RESULT</label>{inp("angleA",{ph:"auto",step:1,color:"var(--violet)",style:{fontSize:17,fontWeight:900}})}</div></div>
<div className="pg">{presets.map(({l,a})=>(<button key={a} onClick={()=>set("angleA",a)} className="pb" style={{borderColor:+v.angleA===a?"var(--violet)":"rgba(139,92,246,.25)",background:+v.angleA===a?"var(--violet)":"rgba(139,92,246,.08)",color:+v.angleA===a?"#fff":"var(--violet)"}}>{l}→,A{a}</button>))}</div>
<div className="ref" style={{background:"rgba(139,92,246,.08)",color:"var(--inkS)"}}>{sp==="RO"?<>45↗=<b style={{color:"var(--violet)"}}>,A135</b> · 45↘=<b style={{color:"var(--violet)"}}>,A225</b> · 30↗=<b style={{color:"var(--violet)"}}>,A150</b> · Abs=<b style={{color:"var(--violet)"}}>,A90</b></>:<>45↗=<b style={{color:"var(--violet)"}}>,A45</b> · 45↘=<b style={{color:"var(--violet)"}}>,A315</b> · 30↗=<b style={{color:"var(--violet)"}}>,A30</b> · Abs=<b style={{color:"var(--violet)"}}>,A90</b></>}</div></div>
<div className="sec r2"><div className="stit" style={{color:"var(--rose)"}}>⌒ VERRUNDUNG — ,R RADIUS</div>
<div style={{display:"grid",gridTemplateColumns:"1fr auto",gap:10,alignItems:"center"}}>{inp("radius",{ph:"0=kein",step:0.1,color:"var(--rose)",style:{fontSize:16,fontWeight:800}})}<div style={{display:"flex",flexWrap:"wrap",gap:4}}>{[0,0.2,0.4,0.5,0.8,1,1.5,2,3].map(r=>(<button key={r} onClick={()=>set("radius",r===0?"":r)} className="pb" style={{borderColor:+v.radius===r?"var(--rose)":"rgba(244,63,94,.25)",background:+v.radius===r?"var(--rose)":"rgba(244,63,94,.08)",color:+v.radius===r?"#fff":"var(--rose)"}}>{r===0?"✕":"R"+r}</button>))}</div></div></div>
<div><label className="fl">VORSCHUB F</label><div style={{display:"grid",gridTemplateColumns:"1fr auto",gap:8,alignItems:"center"}}>{inp("feed",{ph:"std",step:0.01})}<div style={{display:"flex",gap:4,flexWrap:"wrap"}}>{[0.05,0.08,0.1,0.12,0.15,0.2].map(f=>(<button key={f} onClick={()=>set("feed",f)} className="pb" style={{borderColor:+v.feed===f?"var(--blue)":"var(--b1)",background:+v.feed===f?"var(--blueL)":"var(--r)",color:+v.feed===f?"var(--blue)":"var(--inkS)"}}>{f}</button>))}</div></div></div>
</>}</div>
<div className="mfoot">{idx>0&&<button className="delbtn" onClick={()=>onDelete(idx)}>🗑</button>}<button className="cancbtn" onClick={onClose}>ABBRECHEN</button><button className="savebtn" onClick={()=>onSave(idx,v)}>✓ SPEICHERN</button></div>
</div></div>);}

function ImportModal({sp,onImport,onClose}){const[code,setCode]=useState("");const preview=code.trim()?parseImp(code):[];
return(<div className="mbg" onClick={e=>{if(e.target===e.currentTarget)onClose();}}><div className="mbox" style={{maxWidth:520}}>
<div className="mhdr"><div className="mbadge" style={{fontSize:18}}>📥</div><div style={{flex:1}}><div className="mtitle">G-CODE IMPORT</div><div className="msub">,A UND ,R WERDEN ERKANNT</div></div><button className="xcb" onClick={onClose}>×</button></div>
<div style={{padding:"14px 17px",display:"flex",flexDirection:"column",gap:10}}><textarea value={code} onChange={e=>setCode(e.target.value)} className="ita" placeholder={"G1X40.5Z-2,A135F.15\nG1X43Z-5,R.9F.12\n..."}/>{preview.length>0&&<div><div style={{fontSize:10,color:"var(--green)",fontFamily:"Space Mono",fontWeight:700,marginBottom:6}}>{preview.length} PUNKTE ERKANNT</div><div style={{display:"flex",gap:5,flexWrap:"wrap"}}>{preview.slice(0,6).map((r,i)=>(<div key={i} style={{padding:"3px 9px",borderRadius:6,background:"var(--blueL)",border:"1px solid rgba(59,130,246,.3)",fontSize:9,fontFamily:"Space Mono",color:"var(--blue)"}}>{i+1} {r.move} X{r.x} Z{r.z}{r.angleA?" ,A"+r.angleA:""}{r.radius?" ,R"+r.radius:""}</div>))}{preview.length>6&&<div style={{padding:"3px 9px",borderRadius:6,background:"var(--r)",fontSize:9,color:"var(--inkM)"}}>+{preview.length-6}</div>}</div></div>}</div>
<div className="mfoot"><button className="cancbtn" onClick={onClose}>ABBRECHEN</button><button className="savebtn" style={{opacity:preview.length?1:.4,cursor:preview.length?"pointer":"default"}} onClick={()=>{if(preview.length)onImport(preview);}}>{preview.length?preview.length+" PUNKTE IMPORTIEREN":"KEIN CODE"}</button></div>
</div></div>);}

function KEPage({sp}){
const[roR,setRoR]=useState(EX_RO);const[ruR,setRuR]=useState(EX_RU);const[cfg,setCfg]=useState({tool:"T0202",speed:180,feed:"0.15"});const[modal,setModal]=useState(null);const[showImp,setShowImp]=useState(false);const[stab,setStab]=useState("canvas");const[copied,setCopied]=useState(false);const[showNums,setShowNums]=useState(true);const[showTags,setShowTags]=useState(true);const[dragIdx,setDragIdx]=useState(null);const[dragOver,setDragOver]=useState(null);
const rows=sp==="RO"?roR:ruR,setRows=sp==="RO"?setRoR:setRuR;
const save=(idx,v)=>setRows(p=>p.map((r,i)=>i===idx?{...r,...v}:r));
const del=(idx)=>{setRows(p=>p.filter((_,i)=>i!==idx));setModal(null);};
const addAfter=(idx)=>{setRows(p=>{const n=[...p];const prev=p[idx]||{};n.splice(idx+1,0,mkR({x:prev.x||'',z:prev.z||''}));return n;});setTimeout(()=>setModal(idx+1),60);};
const addNew=()=>{let newIdx=0;setRows(p=>{const prev=p[p.length-1]||{};newIdx=p.length;return[...p,mkR({x:prev.x||'',z:prev.z||''})];});setTimeout(()=>setModal(newIdx),60);};
const code=buildKG(rows,{...cfg,sp});
const copy=()=>navigator.clipboard?.writeText(code).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),2000);});
return(<div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",minHeight:0}}>
<div style={{display:"flex",borderBottom:"1px solid var(--b1)",background:"var(--p)",flexShrink:0}}>
{[["canvas","◎ KONTUR"],["code","{ } G-CODE"]].map(([k,l])=>(<button key={k} onClick={()=>setStab(k)} className={"tab"+(stab===k?" on":"")}><span className="tdot" style={{background:stab===k?"var(--blue)":"currentColor"}}/>{l}</button>))}
<button onClick={()=>setShowImp(true)} className="tab" style={{marginLeft:"auto"}}>📥 IMPORT</button>
</div>
{stab==="canvas"&&(<div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",minHeight:0}}>
<div style={{flex:1,position:"relative",minHeight:0}}><div style={{position:"absolute",inset:0}}><KECanvas rows={rows} sp={sp} aIdx={modal} showNums={showNums} showTags={showTags}/></div>
<div className="hud"><div className="hudl">{sp} — Z {sp==="RU"?"+(GEGEN)":"-(HAUPT)"}</div></div></div>
<div style={{background:"var(--p)",borderTop:"1px solid var(--b1)",flexShrink:0}}>
<div className="kes">{rows.map((row,i)=>{const rapid=row.move==="G0",active=modal===i,hasA=row.angleA!==""&&row.angleA!=null,hasR=+row.radius>0;return(<div key={i} className="kec" style={{flexShrink:0}}><div className="kecb" style={active?{background:"var(--blueL)",borderColor:"rgba(0,102,255,.3)"}:{}}><div className="ker1"><div className="ken" style={{background:active?"var(--blue)":rapid?"var(--inkL)":"var(--blueL)",color:active||rapid?"#fff":"var(--blue)"}}>{showNums?(i+1):"·"}</div><button onClick={()=>{if(i>0)setRows(p=>{const n=[...p];const t=n[i-1];n[i-1]=n[i];n[i]=t;return n;});}} disabled={i===0} style={{width:20,height:20,padding:0,border:"none",background:"transparent",cursor:i===0?"default":"pointer",color:i===0?"var(--inkL)":"var(--inkS)",fontSize:14,display:"flex",alignItems:"center",justifyContent:"center",borderRadius:5}}>‹</button><button onClick={()=>{if(i<rows.length-1)setRows(p=>{const n=[...p];const t=n[i+1];n[i+1]=n[i];n[i]=t;return n;});}} disabled={i===rows.length-1} style={{width:20,height:20,padding:0,border:"none",background:"transparent",cursor:i===rows.length-1?"default":"pointer",color:i===rows.length-1?"var(--inkL)":"var(--inkS)",fontSize:14,display:"flex",alignItems:"center",justifyContent:"center",borderRadius:5}}>›</button><span className="ket" onClick={()=>setModal(i)} style={{color:rapid?"var(--inkM)":active?"var(--blue)":"var(--inkS)",cursor:"pointer"}}>{row.move||"G1"}</span><span className="kec2" onClick={()=>setModal(i)} style={{cursor:"pointer"}}>{row.x?"X"+row.x:""}{row.x&&row.z?" ":""}{row.z?"Z"+row.z:""}{!row.x&&!row.z?"—":""}</span></div>{showTags&&<div className="ketags">{hasA&&<span className="ketag" style={{background:"var(--violetL)",color:"var(--violet)"}}>,A{row.angleA}</span>}{hasR&&<span className="ketag" style={{background:"var(--roseL)",color:"var(--rose)"}}>,R{row.radius}</span>}</div>}</div><button className="keadd" onClick={()=>addAfter(i)}>+</button></div>);})}<button onClick={addNew} style={{padding:"8px 13px",borderRadius:10,border:"1.5px dashed rgba(59,130,246,.4)",background:"var(--blueL)",color:"var(--blue)",cursor:"pointer",fontSize:11,fontWeight:700,flexShrink:0,whiteSpace:"nowrap",fontFamily:"Barlow Condensed",letterSpacing:1,textTransform:"uppercase",display:"flex",alignItems:"center",gap:4}}>+ PUNKT</button></div>
<div className="keset">{[["🔧","TOOL",cfg.tool,v=>setCfg(p=>({...p,tool:v})),64],["⚡","RPM",cfg.speed,v=>setCfg(p=>({...p,speed:v})),50],["F","STD",cfg.feed,v=>setCfg(p=>({...p,feed:v})),44]].map(([ic,lb,val,fn,w])=>(<div key={lb} className="kesc"><span>{ic}</span><span className="kesl">{lb}</span><input value={val} onChange={e=>fn(e.target.value)} className="kesi" style={{width:w}}/></div>))}<div style={{display:"flex",alignItems:"center",gap:6,padding:"7px 12px",marginLeft:"auto",flexShrink:0}}><button onClick={()=>setShowNums(v=>!v)} className={"ketog"+(showNums?" on":"")}>№</button><button onClick={()=>setShowTags(v=>!v)} className={"ketog"+(showTags?" on":"")} style={{color:showTags?"var(--violet)":undefined,borderColor:showTags?"rgba(123,91,255,.3)":undefined,background:showTags?"var(--violetL)":undefined}}>,A ,R</button></div></div>
</div></div>)}
{stab==="code"&&(<div style={{flex:1,display:"flex",flexDirection:"column",minHeight:0,padding:10,gap:7}}>
<div style={{display:"flex",gap:8,alignItems:"center",flexShrink:0}}><span style={{fontSize:10,color:"var(--inkM)",fontFamily:"Space Mono",flex:1}}>{sp} — {code.split("\n").length} LINES</span><button onClick={copy} style={{padding:"7px 16px",fontSize:11,fontWeight:800,borderRadius:8,border:"none",cursor:"pointer",background:copied?"var(--green)":"linear-gradient(135deg,var(--blueD),var(--violet))",color:"#fff",fontFamily:"Barlow Condensed",letterSpacing:1.5,textTransform:"uppercase"}}>{copied?"COPIED":"COPY"}</button></div>
<div style={{flex:1,position:"relative",minHeight:0}}><textarea value={code} onChange={e=>{try{const imp=parseImp(e.target.value);if(imp.length>0)setRows(imp);}catch(x){}}} style={{position:"absolute",inset:0,background:"#0D1117",borderRadius:10,padding:"12px 14px",border:"1px solid var(--b1)",fontFamily:"Space Mono,monospace",fontSize:12.5,lineHeight:"22px",color:"#E2E8F0",resize:"none",outline:"none",width:"100%",height:"100%",textTransform:"uppercase",caretColor:"#3B82F6",boxSizing:"border-box"}} spellCheck={false} autoCapitalize="off" autoCorrect="off"/></div>
</div>)}
{modal!==null&&rows[modal]&&<PointModal pt={rows[modal]} idx={modal} total={rows.length} sp={sp} onSave={(i,v)=>save(i,v)} onDelete={del} onClose={()=>setModal(null)} onPrev={()=>setModal(m=>Math.max(0,m-1))} onNext={()=>setModal(m=>Math.min(rows.length-1,m+1))}/>}
{showImp&&<ImportModal sp={sp} onImport={imported=>{setRows(imported);setShowImp(false);setModal(null);}} onClose={()=>setShowImp(false)}/>}
</div>);}

const PT={T01_RO:{name:"Schruppstahl CNMG 120408MS",type:"external",insert:"CNMG 120408MS",material:"CA6515 Kyocera",noseR:0.8,angle:80,overhang:40},T02_RO:{name:"Kopierstahl AUSSEN FERTIG",type:"external",insert:"DCMT 11T304",material:"CA6515 Kyocera",noseR:0.4,angle:55,overhang:38},T03_RO:{name:"Stechstahl 3mm",type:"groove",insert:"3mm Groove",material:"IC328 ISCAR",noseR:0.2,angle:0,overhang:30},T04_RO:{name:"Kopierbohrst. D16",type:"internal",insert:"CCMT 09T304",material:"—",noseR:0.4,angle:80,overhang:50},T10_RO:{name:"Gewindest. M26x1",type:"thread",insert:"60deg Thread",material:"—",noseR:0.1,angle:60,overhang:35},T01_RU:{name:"Schruppstahl CNMG 120408MS",type:"external",insert:"CNMG 120408MS",material:"CA6515 Kyocera",noseR:0.8,angle:80,overhang:40},T02_RU:{name:"Schlichtstahl DCMT 11T304GK",type:"external",insert:"DCMT 11T304GK",material:"CA6515 Kyocera",noseR:0.4,angle:55,overhang:38},T03_RU:{name:"Gewinde M40x1.5",type:"thread",insert:"60deg Thread",material:"—",noseR:0.1,angle:60,overhang:50},T04_RU:{name:"Bohrst. D10 HM",type:"internal",insert:"DCMT 07T204",material:"HM",noseR:0.4,angle:55,overhang:80},T10_RU:{name:"Abstechstahl 3mm IC328",type:"groove",insert:"HGN 3003C IC328",material:"IC328 ISCAR",noseR:0.2,angle:0,overhang:30}};
const getP=(key,sp)=>PT[key+"_"+sp]||PT[key+"_RO"]||null;
const DT={name:"",type:"external",insert:"",material:"",noseR:0.8,angle:80,overhang:40,corrX:0,corrZ:0,radial:0,axial:0};

function ToolCard({tKey,color,data,onChange,preset}){const d={...DT,...data};const set=(k,v)=>onChange({...d,[k]:v});
const Num=({label,k,unit,step=0.001})=>(<div><label className="fl">{label}</label><div style={{display:"flex",gap:4,alignItems:"center"}}><input type="number" step={step} value={d[k]??0} onChange={e=>set(k,parseFloat(e.target.value)||0)} className="fi" style={{flex:1,minWidth:0}} onFocus={e=>{e.target.style.borderColor=color;e.target.style.boxShadow="0 0 0 2px "+color+"22";}} onBlur={e=>{e.target.style.borderColor="var(--b1)";e.target.style.boxShadow="none";}}/>{unit&&<span style={{fontSize:10,color:"var(--inkM)",minWidth:22,fontFamily:"Space Mono"}}>{unit}</span>}</div></div>);
return(<div className="tc"><div className="tch" style={{background:color+"18",borderBottomColor:color+"25"}}><div className="tcb" style={{background:color,boxShadow:"0 2px 8px "+color+"55"}}>{tKey}</div><div style={{flex:1,minWidth:0}}><div style={{fontSize:13,fontWeight:700,color:"var(--ink)",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{d.name||tKey}</div><div style={{fontSize:10,color,fontFamily:"Space Mono",marginTop:2,letterSpacing:.5}}>{TL[d.type]||"—"} / {d.insert||"—"}</div></div>{preset&&<button onClick={()=>onChange({...d,...preset})} style={{padding:"4px 8px",fontSize:10,border:"1px solid "+color+"44",borderRadius:5,background:color+"15",color,cursor:"pointer",fontWeight:700}}>↺</button>}</div><div className="tcbody"><div><label className="fl">NAME</label><input type="text" value={d.name||""} placeholder="Description" onChange={e=>set("name",e.target.value)} className="fi" onFocus={e=>{e.target.style.borderColor=color;}} onBlur={e=>{e.target.style.borderColor="var(--b1)";}}/></div><div className="g2"><div><label className="fl">TYPE</label><select value={d.type||""} onChange={e=>set("type",e.target.value)} className="fsel">{Object.entries(TL).map(([v,l])=>(<option key={v} value={v}>{l}</option>))}</select></div><div><label className="fl">INSERT</label><input type="text" value={d.insert||""} placeholder="CNMG 120408" onChange={e=>set("insert",e.target.value)} className="fi"/></div></div><div className="g3"><Num label="NOSE R" k="noseR" unit="mm" step={0.05}/><Num label="ANGLE" k="angle" unit="°" step={1}/><Num label="OVERHANG" k="overhang" unit="mm" step={0.5}/></div><div><div className="slbl">CORRECTION</div><div className="g2"><Num label="CORR X" k="corrX" unit="mm"/><Num label="CORR Z" k="corrZ" unit="mm"/><Num label="RADIAL" k="radial" unit="mm"/><Num label="AXIAL" k="axial" unit="mm"/></div></div></div></div>);}

export default function App(){
const[sp,setSp]=useState("RO");const[cRO,setCRO]=useState(PRO);const[cRU,setCRU]=useState(PRU);const[tab,setTab]=useState("sim");const[toolsRO,setToolsRO]=useState({});const[toolsRU,setToolsRU]=useState({});const[dia,setDia]=useState(55);const[sel,setSel]=useState("T01");const[step,setStep]=useState(-1);const[run,setRun]=useState(false);const[showG0,setShowG0]=useState(true);const[visRO,setVisRO]=useState(()=>new Set());const[visRU,setVisRU]=useState(()=>new Set());const sRef=useRef(null);
const code=sp==="RO"?cRO:cRU,setCode=sp==="RO"?setCRO:setCRU;
const toolDb=sp==="RO"?toolsRO:toolsRU,setToolDb=sp==="RO"?setToolsRO:setToolsRU;
const vis=sp==="RO"?visRO:visRU,setVis=sp==="RO"?setVisRO:setVisRU;
const blocks=parse(code),toolList=getTools(blocks),allPaths=getPaths(blocks);
const cols={};toolList.forEach((t,i)=>{cols[t.key]=TC[i%TC.length];});
useEffect(()=>{const tl=getTools(parse(sp==="RO"?cRO:cRU));setVis(new Set(tl.map(t=>t.key)));const db=sp==="RO"?{...toolsRO}:{...toolsRU};let ch=false;for(const t of tl){if(!db[t.key]){const p=getP(t.key,sp);db[t.key]={...DT,name:t.comment||t.key,...(p||{})};ch=true;}}if(ch)setToolDb({...db});if(tl.length)setSel(tl[0].key);},[sp,cRO,cRU]);
const simPaths=allPaths.filter(p=>vis.has(p.tool));
const start=()=>{setStep(0);setRun(true);};const stop=()=>{setRun(false);clearInterval(sRef.current);};const reset=()=>{stop();setStep(-1);};
useEffect(()=>{if(run){sRef.current=setInterval(()=>{setStep(s=>{if(s>=simPaths.length){setRun(false);clearInterval(sRef.current);return s;}return s+1;});},50);}return()=>clearInterval(sRef.current);},[run,simPaths.length]);
const cur=step>0&&step<=simPaths.length?simPaths[step-1]:null;
const TABS=[["sim","▶ SIM"],["contour","◎ KONTOUR"],["editor","✏ PROGRAMM"],["tools","🔧 WERKZEUGE"],["ke","📐 KONTUR-ED"]];

return(<><style>{G}</style><div className="app">
<div className="hdr">
<div className="hdr1">
<div className="logo"><div className="lm">N</div><div><div className="lt">Nakamura WT-250 II</div><div className="ls">CNC CONTROL SIMULATOR</div></div></div>
<div className="spsel">{[["RO","-Z MAIN"],["RU","+Z SUB"]].map(([s,h])=>(<button key={s} className={"spb"+(sp===s?" on":"")} onClick={()=>setSp(s)}><span className="spid">{s}</span><span className="spsub">{h}</span></button>))}</div>
<div className="statrow">{[{l:"MOVES",v:allPaths.length,c:"var(--blue)"},{l:"TOOLS",v:toolList.length,c:"var(--amber)"},{l:"BLOCKS",v:blocks.filter(b=>b.type==="block").length,c:"var(--cyan)"}].map(({l,v,c})=>(<div key={l} className="stat"><div className="statv" style={{color:c}}>{v}</div><div className="statl">{l}</div></div>))}<div className="stat"><input type="number" step={0.5} value={dia} onChange={e=>setDia(parseFloat(e.target.value)||0)} className="bi" onFocus={e=>e.target.style.borderColor="var(--blue)"} onBlur={e=>e.target.style.borderColor="var(--b1)"}/>  <div className="statl">Ø MM</div></div></div>
</div>
<div className="tabs">{TABS.map(([k,l])=>(<button key={k} className={"tab"+(tab===k?" on":"")} onClick={()=>setTab(k)}><span className="tdot" style={{background:tab===k?"var(--blue)":"currentColor"}}/>{l}</button>))}</div>
</div>

<div className="main">
{tab==="sim"&&(<div className="dsk-sim" style={{padding:10,minHeight:0}}><div className="dsk-sim-main">
<div className="simw"><SView paths={simPaths} step={step} dia={dia} cols={cols} showG0={showG0}/>
<div className="hud"><div className="hudl">{sp} — {toolList.filter(t=>vis.has(t.key)).length} TOOLS — {simPaths.length} MOVES</div>{cur&&<div style={{display:"flex",gap:6,marginTop:4,flexWrap:"wrap",fontSize:11,fontFamily:"Space Mono",color:"#B0B8D0"}}><span style={{color:cols[cur.tool]||"var(--blue)",fontWeight:700}}>{cur.tool}</span><span>X{(cur.to.x*2).toFixed(2)}</span><span>Z{cur.to.z.toFixed(2)}</span>{cur.angle!=null&&<span style={{color:"var(--violet)"}}>,A{cur.angle}</span>}{cur.radius>0&&<span style={{color:"var(--rose)"}}>,R{cur.radius}</span>}<span style={{color:cur.rapid?"#6B7388":"var(--green)"}}>{cur.rapid?"G0":"G1"}</span></div>}</div></div>
<div className="ctrl">
<div className="prow"><button className="ibtn" onClick={reset}>↺</button><button className={"pbtn "+(run?"stop":"go")} onClick={run?stop:start}><span style={{fontSize:18}}>{run?"⏹":"▶"}</span>{run?"STOP":"START"}</button><div className="ctr">{step>=0?Math.min(step,simPaths.length)+"/"+simPaths.length:"— / —"}</div><button className={"g0b "+(showG0?"on":"off")} onClick={()=>setShowG0(v=>!v)}>G0 {showG0?"ON":"OFF"}</button></div>
<input type="range" className="sld" min={0} max={simPaths.length} value={step>=0?step:0} onChange={e=>{stop();setStep(parseInt(e.target.value));}}/>
<div className="chips">{toolList.map((t,i)=>{const col=cols[t.key]||TC[i%TC.length],isV=vis.has(t.key);return(<div key={t.key} className="chip" onClick={()=>setVis(prev=>{const n=new Set(prev);n.has(t.key)?n.delete(t.key):n.add(t.key);return n;})} style={{background:isV?col+"18":"var(--r)",borderColor:isV?col+"55":"var(--b1)",color:col,opacity:isV?1:.35}}><div className="dot" style={{background:col}}/>{t.key}</div>);})}</div>
</div></div></div>)}

{tab==="contour"&&(<div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",minHeight:0,padding:10,gap:8}}>
<div style={{flex:1,borderRadius:8,overflow:"hidden",border:"1px solid var(--b1)",position:"relative",minHeight:0}}><CView paths={allPaths} vis={vis} cols={cols} showG0={showG0}/><div className="hud"><span style={{color:sp==="RU"?"var(--amber)":"var(--blue)",fontWeight:700,fontSize:10,letterSpacing:1,fontFamily:"Space Mono"}}>{sp} Z{sp==="RU"?"+(GEGEN)":"-(HAUPT)"}</span></div></div>
<div style={{display:"flex",gap:5,flexShrink:0,flexWrap:"wrap",alignItems:"center"}}>{toolList.map((t,i)=>{const col=cols[t.key]||TC[i%TC.length],isV=vis.has(t.key);return(<div key={t.key} className="chip" onClick={()=>setVis(prev=>{const n=new Set(prev);n.has(t.key)?n.delete(t.key):n.add(t.key);return n;})} style={{background:isV?col+"18":"var(--r)",borderColor:isV?col+"55":"var(--b1)",color:col,opacity:isV?1:.35}}><div className="dot" style={{background:col}}/>{t.key} <span style={{fontSize:9,color:"var(--inkM)",maxWidth:80,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{toolDb[t.key]?.insert||t.comment}</span></div>);})} <button className={"g0b "+(showG0?"on":"off")} onClick={()=>setShowG0(v=>!v)}>G0 {showG0?"ON":"OFF"}</button></div>
<div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,flexShrink:0}}>{[{l:"MOVES",v:allPaths.length,c:"var(--blue)"},{l:"G1 CUT",v:allPaths.filter(p=>!p.rapid).length,c:"var(--green)"},{l:"G0 RAPID",v:allPaths.filter(p=>p.rapid).length,c:"var(--inkM)"},{l:",A ANGLES",v:allPaths.filter(p=>p.angle!=null).length,c:"var(--violet)"}].map(({l,v,c})=>(<div key={l} style={{background:"var(--r)",border:"1px solid var(--b1)",borderRadius:8,padding:"9px 11px"}}><div style={{fontSize:22,fontWeight:800,color:c,fontFamily:"Space Mono"}}>{v}</div><div style={{fontSize:9,color:"var(--inkM)",marginTop:2,fontFamily:"Barlow Condensed",letterSpacing:1.5,textTransform:"uppercase"}}>{l}</div></div>))}</div>
</div>)}

{tab==="editor"&&(<div style={{flex:1,padding:10,display:"flex",flexDirection:"column",gap:7,overflow:"hidden",minHeight:0}}>
<div style={{fontSize:10,color:"var(--inkM)",fontFamily:"Space Mono",letterSpacing:.5,flexShrink:0}}>{code.split("\n").length} LINES / {blocks.filter(b=>b.type==="block").length} BLOCKS</div>
<div style={{flex:1,minHeight:0,borderRadius:8,overflow:"hidden"}}><Editor value={code} onChange={setCode}/></div>
</div>)}

{tab==="tools"&&(<div style={{flex:1,display:"flex",overflow:"hidden",minHeight:0}}>
<div style={{flex:1,overflowY:"auto",padding:10,display:"flex",flexDirection:"column",gap:10}}>{toolList.filter(t=>t.key===sel).map((t,i)=>{const col=cols[t.key]||TC[i%TC.length];return(<ToolCard key={t.key} tKey={t.key} color={col} data={toolDb[t.key]||{}} preset={getP(t.key,sp)} onChange={nd=>setToolDb(prev=>({...prev,[t.key]:nd}))}/>);})}{!toolList.find(t=>t.key===sel)&&<div style={{color:"var(--inkM)",fontSize:12,fontFamily:"Space Mono",padding:10}}>SELECT TOOL →</div>}</div>
<div style={{width:185,borderLeft:"1px solid var(--b1)",overflowY:"auto",padding:8,background:"var(--g)"}}><div style={{fontSize:9,color:"var(--inkM)",fontFamily:"Barlow Condensed",letterSpacing:2,marginBottom:8,textTransform:"uppercase",fontWeight:700}}>ALL TOOLS / {sp}</div><div style={{display:"flex",flexDirection:"column",gap:4}}>{toolList.map((t,i)=>{const col=cols[t.key]||TC[i%TC.length],td=toolDb[t.key],active=sel===t.key;return(<div key={t.key} onClick={()=>setSel(t.key)} style={{border:"1px solid "+(active?col:"var(--b1)"),borderRadius:8,padding:"8px 9px",cursor:"pointer",background:active?col+"15":"var(--r)",transition:"all .1s"}}><div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:22,height:22,borderRadius:5,background:col,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><span style={{color:"#fff",fontSize:8,fontWeight:800,fontFamily:"Space Mono"}}>{t.key}</span></div><div style={{flex:1,minWidth:0}}><div style={{fontSize:11,fontWeight:700,color:active?col:"var(--ink)",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{td?.name||t.comment||t.key}</div><div style={{fontSize:9,color:"var(--inkM)",fontFamily:"Space Mono",marginTop:1}}>{TL[td?.type]||"—"}</div></div></div></div>);})}</div></div>
</div>)}

{tab==="ke"&&<KEPage sp={sp}/>}
</div>

<div className="sbar"><div className="sdot" style={{background:sp==="RO"?"var(--blue)":"var(--amber)"}}/><span style={{color:sp==="RO"?"var(--blue)":"var(--amber)",fontWeight:700}}>{sp}</span><span>{allPaths.length} MOVES</span><span>{step>=0?"SIM "+Math.min(step,simPaths.length)+"/"+simPaths.length:"READY"}</span><span style={{marginLeft:"auto",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",opacity:.5}}>{toolList.map(t=>t.key).join(" ")}</span></div>
</div></>);}
