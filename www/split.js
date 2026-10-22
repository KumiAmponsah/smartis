// Run once:  node split.js
const fs = require('fs');

const src = fs.readFileSync('index.html', 'utf8');
if (/href="css\/styles\.css"/.test(src)) { console.log('Already split. Nothing to do.'); process.exit(0); }
fs.copyFileSync('index.html', 'index.original.html');   // backup

const css = [], js = [];
let html = src
  .replace(/<style[^>]*>([\s\S]*?)<\/style>/g, (m, c) => { css.push(c.trim()); return ''; })
  .replace(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g, (m, c) => { js.push(c.trim()); return ''; });

if (!css.length || !js.length) { console.log('Could not find inline CSS/JS.'); process.exit(1); }

/* ───────────── WEB-APP CSS (appended last so it wins) ───────────── */
const WEB_CSS = String.raw`
/* ═══════════ WEB APP EDITION ═══════════ */
html{scroll-behavior:smooth;-webkit-text-size-adjust:100%;}
::selection{background:rgba(52,110,224,.35);}
:focus-visible{outline:2px solid var(--accent);outline-offset:2px;border-radius:6px;}

/* No phone status-bar / nav-button blocks in a browser */
body::before,body::after{display:none !important;}
.content-wrap{padding-top:0 !important;padding-bottom:24px !important;}
.topbar2{top:0 !important;}
.sidebar{padding-top:0 !important;padding-bottom:0 !important;}

/* Menu button works on every screen size (collapses sidebar on desktop) */
.menu-toggle{display:inline-flex !important;align-items:center;justify-content:center;flex-shrink:0;
  width:38px;height:38px;border-radius:10px;transition:all .15s;}
.menu-toggle:hover{border-color:var(--accent);color:var(--accent);}

/* Desktop layout: smooth collapsible sidebar + centred content */
@media(min-width:993px){
  .sidebar{transition:transform .25s ease;}
  .content-wrap{transition:margin-left .25s ease;}
  body.sb-collapsed .sidebar{transform:translateX(-100%);}
  body.sb-collapsed .content-wrap{margin-left:0;}
}
.page-body{width:100%;max-width:1900px;margin:0 auto;}
.page{padding:24px clamp(14px,2.2vw,36px);}

/* Large and ultra-wide screens */
@media(min-width:1500px){
  .sidebar{width:270px;}
  .content-wrap{margin-left:270px;}
  body.sb-collapsed .content-wrap{margin-left:0;}
  .sidebar .tab{font-size:14px !important;padding:12px 20px !important;}
  .card{padding:22px 24px;}
  .dt td{font-size:12px;} .dt th{font-size:10px;}
}
@media(min-width:2000px){.page-body{max-width:2200px;}}

/* Multi-column panels adapt to the space available (fixes cramped tablets) */
[style*="grid-template-columns:1fr 1fr"],
[style*="grid-template-columns:repeat(3,1fr)"]{
  grid-template-columns:repeat(auto-fit,minmax(min(100%,230px),1fr)) !important;
}
.dp > div[style*="grid-template-columns"]{display:grid !important;
  grid-template-columns:repeat(auto-fit,minmax(min(100%,270px),1fr)) !important;}
.g2{grid-template-columns:repeat(auto-fit,minmax(min(100%,360px),1fr)) !important;}
.g3{grid-template-columns:repeat(auto-fit,minmax(min(100%,280px),1fr)) !important;}
.g4{grid-template-columns:repeat(auto-fit,minmax(min(100%,210px),1fr)) !important;}
#rc-cards,#dd-cards,#comp-cards,#lv-cards{grid-template-columns:repeat(auto-fill,minmax(min(100%,230px),1fr)) !important;}
.hs{width:clamp(96px,30%,140px);}

/* Header: keep everything visible on small laptops / tablets */
@media(max-width:1180px){.tb-clock{display:none;}}
@media(max-width:1100px){#theme-label{display:none;} #theme-btn{width:38px;padding:0 !important;}}
.topbar2{flex-wrap:nowrap !important;}
.topbar2 > div:last-child{flex-wrap:nowrap;}
.tb-title{overflow:hidden;text-overflow:ellipsis;}

/* Tables scroll inside their card instead of squashing */
.card > div[style*="overflow-x:auto"]{-webkit-overflow-scrolling:touch;}
@media(max-width:900px){.dt{min-width:620px;}}

/* Modals centred, never taller than the window */
.modal{max-height:92vh !important;}
.modal-overlay{backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px);}
#notif-panel{top:68px !important;right:20px !important;}
@media(max-width:480px){#notif-panel{left:10px;right:10px !important;width:auto !important;top:62px !important;}}

/* Toast messages */
#toast-wrap{position:fixed;right:20px;bottom:20px;z-index:6000;display:flex;flex-direction:column;gap:10px;pointer-events:none;}
.toast{pointer-events:auto;display:flex;align-items:center;gap:10px;min-width:220px;max-width:360px;padding:12px 16px;
  background:var(--surface);color:var(--text);border:1px solid var(--border2);border-left:4px solid var(--accent);
  border-radius:12px;font-size:12px;font-weight:600;box-shadow:0 14px 40px rgba(0,0,0,.4);animation:toastIn .3s ease;}
.toast.warn{border-left-color:var(--warning);} .toast.ok{border-left-color:var(--success);}
.toast.out{opacity:0;transform:translateY(8px);transition:all .3s;}
@keyframes toastIn{from{opacity:0;transform:translateY(12px);}to{opacity:1;transform:none;}}
@media(max-width:560px){#toast-wrap{left:12px;right:12px;bottom:12px;} .toast{max-width:none;}}

@media(prefers-reduced-motion:reduce){*{animation-duration:.01ms !important;transition-duration:.01ms !important;scroll-behavior:auto !important;}}
@media print{#toast-wrap{display:none !important;}}
`;

/* ───────────── WEB-APP JS (appended last) ───────────── */
const WEB_JS = String.raw`
// ══════════ WEB APP EDITION ══════════
function toast(msg, type){
  var w = document.getElementById('toast-wrap');
  if(!w){ w = document.createElement('div'); w.id = 'toast-wrap'; document.body.appendChild(w); }
  var t = document.createElement('div');
  t.className = 'toast ' + (type || '');
  t.innerHTML = '<i class="bi bi-' + (type==='ok'?'check-circle-fill':type==='warn'?'exclamation-triangle-fill':'info-circle-fill') + '"></i><span>' + msg + '</span>';
  w.appendChild(t);
  setTimeout(function(){ t.classList.add('out'); setTimeout(function(){ t.remove(); }, 320); }, 3200);
}
(function(){
  var b = document.body;

  // Menu button: collapses the sidebar on desktop, slides it in on phones
  var mt = document.querySelector('.menu-toggle');
  if(mt){
    mt.removeAttribute('onclick');
    mt.addEventListener('click', function(){
      if(window.innerWidth > 992){
        b.classList.toggle('sb-collapsed');
        try{ localStorage.setItem('smartis-sb', b.classList.contains('sb-collapsed') ? '1' : '0'); }catch(e){}
      } else { b.classList.toggle('sb-open'); }
    });
  }
  try{ if(localStorage.getItem('smartis-sb') === '1') b.classList.add('sb-collapsed'); }catch(e){}

  // Choosing a tab closes the phone menu and returns to the top
  document.querySelectorAll('.sidebar .tab').forEach(function(t){
    t.addEventListener('click', function(){ b.classList.remove('sb-open'); window.scrollTo({top:0}); });
  });

  // Esc closes menus and popups
  document.addEventListener('keydown', function(e){
    if(e.key !== 'Escape') return;
    b.classList.remove('sb-open');
    var n = document.getElementById('notif-panel');
    if(n && n.style.display === 'block') toggleNotifPanel();
    closeLogout(); closeWOModal();
  });

  // Click outside the notification panel closes it
  document.addEventListener('click', function(e){
    var n = document.getElementById('notif-panel');
    if(n && n.style.display === 'block' && !n.contains(e.target) && !e.target.closest('#notif-btn')) toggleNotifPanel();
  });

  // Connection feedback
  window.addEventListener('offline', function(){ toast('You are offline. Showing the last saved data.', 'warn'); });
  window.addEventListener('online',  function(){ toast('Back online.', 'ok'); if(typeof refreshData === 'function') refreshData(); });
})();
`;

/* ───────────── Assemble ───────────── */
fs.mkdirSync('css', { recursive: true });
fs.mkdirSync('js',  { recursive: true });
fs.writeFileSync('css/styles.css', css.join('\n\n') + '\n\n' + WEB_CSS);
fs.writeFileSync('js/app.js',      js.join('\n\n')  + '\n\n' + WEB_JS);

const fav = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Crect width='64' height='64' rx='14' fill='%23346ee0'/%3E%3Ctext x='32' y='44' font-size='36' text-anchor='middle' fill='white'%3E%E2%9A%99%3C/text%3E%3C/svg%3E";
html = html
  .replace('<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover"/>',
           '<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover"/>\n<meta name="description" content="SMARTIS — Fleet maintenance and reliability intelligence for Deeprock Ghana Limited"/>\n<meta name="theme-color" content="#1e1e2e"/>\n<link rel="icon" href="' + fav + '"/>')
  .replace('</head>', '<link rel="stylesheet" href="css/styles.css?v=1"/>\n</head>')
  .replace('</body>', '<script src="js/app.js?v=1"></script>\n</body>')
  .replace('SMARTIS v4 · Single-file build', 'SMARTIS v5 · Web edition')
  .replace(/\n{3,}/g, '\n\n');

fs.writeFileSync('index.html', html);
console.log('Done. Created index.html, css/styles.css, js/app.js (backup: index.original.html)');