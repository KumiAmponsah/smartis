/* ═══════════════════════════════════════════════════════════════════════════
   SMARTIS — Homepage (index.html) behaviour
   • Navbar scroll state, mobile drawer
   • Sign-in entry points that open the animated login overlay
   • Shared particle-network canvas
   • SMARTIS login form logic (any username + password accepted)
   ═══════════════════════════════════════════════════════════════════════════ */
(function () {
  var APP = 'index2.html', KEY = 'smartis-auth';
  var $ = function (i) { return document.getElementById(i); };

  $('yr').textContent = new Date().getFullYear();

  /* ── Navbar background on scroll + mobile menu ── */
  var nav = $('nav');
  function onScroll() { nav.classList.toggle('solid', window.scrollY > 12); }
  onScroll();
  window.addEventListener('scroll', onScroll, { passive: true });

  var burger = $('burger'), drawer = $('drawer');
  burger.addEventListener('click', function () {
    var o = drawer.classList.toggle('open');
    burger.setAttribute('aria-expanded', o);
    burger.innerHTML = '<i class="bi bi-' + (o ? 'x-lg' : 'list') + '"></i>';
  });
  drawer.addEventListener('click', function (e) {
    if (e.target.tagName === 'A') {
      drawer.classList.remove('open');
      burger.innerHTML = '<i class="bi bi-list"></i>';
    }
  });

  /* ── Sign-in entry points ── */
  var authed = false;
  try { authed = sessionStorage.getItem(KEY) === '1' || localStorage.getItem(KEY) === '1'; } catch (e) {}
  document.querySelectorAll('[data-signin]').forEach(function (b) {
    if (authed) {
      var l = b.querySelector('.si-label');
      if (l && !/calculator/i.test(l.textContent)) l.textContent = 'Open dashboard';
    }
    b.addEventListener('click', function () {
      if (authed) { location.href = APP; return; }
      var screen = $('login-screen');
      if (screen) {
        screen.classList.remove('lg-out');
        screen.style.display = '';
        document.body.classList.add('lg-locked');
        setTimeout(function () { var u = $('lg-user'); if (u) u.focus(); }, 120);
      }
    });
  });

  /* ── Shared particle network behind the homepage and sign-in screen ── */
  (function () {
    var cv = $('home-canvas'); if (!cv || !cv.getContext) return;
    var cx = cv.getContext('2d'), W = 0, H = 0, D = 1, pts = [], raf = 0;
    function size() {
      D = Math.min(window.devicePixelRatio || 1, 2);
      W = cv.width = Math.floor(innerWidth * D);
      H = cv.height = Math.floor(innerHeight * D);
      cv.style.width = innerWidth + 'px';
      cv.style.height = innerHeight + 'px';
      var n = Math.min(70, Math.floor(innerWidth * innerHeight / 15000));
      pts = [];
      for (var i = 0; i < n; i++) {
        pts.push({
          x: Math.random() * W, y: Math.random() * H,
          vx: (Math.random() - .5) * .35 * D, vy: (Math.random() - .5) * .35 * D,
          r: (Math.random() * 1.4 + .6) * D
        });
      }
    }
    function draw() {
      cx.clearRect(0, 0, W, H);
      var link = 130 * D;
      for (var i = 0; i < pts.length; i++) {
        var a = pts[i];
        a.x += a.vx; a.y += a.vy;
        if (a.x < 0 || a.x > W) a.vx *= -1;
        if (a.y < 0 || a.y > H) a.vy *= -1;
        cx.beginPath(); cx.arc(a.x, a.y, a.r, 0, 6.283);
        cx.fillStyle = 'rgba(150,180,255,.75)'; cx.fill();
        for (var j = i + 1; j < pts.length; j++) {
          var b = pts[j], dx = a.x - b.x, dy = a.y - b.y, d = Math.sqrt(dx * dx + dy * dy);
          if (d < link) {
            cx.strokeStyle = 'rgba(120,160,255,' + (0.22 * (1 - d / link)) + ')';
            cx.lineWidth = D * .8;
            cx.beginPath(); cx.moveTo(a.x, a.y); cx.lineTo(b.x, b.y); cx.stroke();
          }
        }
      }
      raf = requestAnimationFrame(draw);
    }
    size(); draw();
    window.addEventListener('resize', size);
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
      cancelAnimationFrame(raf); cx.clearRect(0, 0, W, H);
    }
  })();

  /* ── SMARTIS sign-in behaviour and session handling ── */
  (function () {
    var L = $('login-screen'); if (!L) return;
    var g = function (i) { return document.getElementById(i); };

    try {
      if (sessionStorage.getItem(KEY) === '1' || localStorage.getItem(KEY) === '1') {
        L.remove(); return;
      }
    } catch (e) {}
    L.style.display = 'none';

    var user = g('lg-user'), pass = g('lg-pass'), card = g('lg-card'),
        err = g('lg-err'), btn = g('lg-btn'), running = true;

    g('lg-eye').addEventListener('click', function () {
      var show = pass.type === 'password';
      pass.type = show ? 'text' : 'password';
      this.innerHTML = '<i class="bi bi-eye' + (show ? '-slash' : '') + '"></i>';
    });

    function caps(e) { if (e.getModifierState) g('lg-caps').classList.toggle('show', e.getModifierState('CapsLock')); }
    pass.addEventListener('keydown', caps);
    pass.addEventListener('keyup', caps);
    pass.addEventListener('blur', function () { g('lg-caps').classList.remove('show'); });

    function shake(msg, el) {
      err.textContent = msg;
      card.classList.remove('lg-shake');
      void card.offsetWidth;
      card.classList.add('lg-shake');
      if (el) el.focus();
    }
    [user, pass].forEach(function (i) {
      i.addEventListener('input', function () { err.textContent = ''; });
    });
    try {
      var ru = localStorage.getItem('smartis-user');
      if (ru) { user.value = ru; g('lg-keep').checked = true; }
    } catch (e) {}

    g('lg-form').addEventListener('submit', function (e) {
      e.preventDefault();
      var u = user.value.trim(), p = pass.value;
      if (!u) { shake('Please enter your username', user); return; }
      if (!p) { shake('Please enter your password', pass); return; }

      btn.disabled = true;
      g('lg-btn-txt').textContent = 'VERIFYING…';
      g('lg-btn-ico').outerHTML = '<span class="lg-spin" id="lg-btn-ico"></span>';

      setTimeout(function () {
        btn.classList.add('ok');
        g('lg-btn-txt').textContent = 'ACCESS GRANTED';
        g('lg-btn-ico').outerHTML = '<i class="bi bi-check-circle-fill" id="lg-btn-ico"></i>';
        try {
          var keep = g('lg-keep').checked;
          (keep ? localStorage : sessionStorage).setItem(KEY, '1');
          if (keep) localStorage.setItem('smartis-user', u);
          else localStorage.removeItem('smartis-user');
        } catch (e) {}
        setTimeout(function () {
          g('lg-wname').textContent = u.charAt(0).toUpperCase() + u.slice(1);
          card.classList.add('lg-hide');
          g('lg-welcome').classList.add('show');
          setTimeout(function () { location.href = APP; }, 900);
        }, 650);
      }, 800);
    });

    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && L.style.display !== 'none') {
        L.style.display = 'none';
        L.classList.remove('lg-out');
        document.body.classList.remove('lg-locked');
      }
    });
  })();
})();