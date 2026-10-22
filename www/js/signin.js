/* ═══════════════════════════════════════════════════════════════════════
   SMARTIS — Sign-in page logic (standalone)
   Uses the Supabase client initialised in signin.html
   On success, stores auth flag and navigates to index.html
   ═══════════════════════════════════════════════════════════════════════ */

var LG_KEY = 'smartis-auth';
var LG_USER_KEY = 'smartis-user-name';

(function () {
  var L = document.getElementById('login-screen');
  if (!L) return;
  var g = function (id) { return document.getElementById(id); };

  // Already signed in? Skip straight to the dashboard.
  try {
    if (sessionStorage.getItem(LG_KEY) === '1' || localStorage.getItem(LG_KEY) === '1') {
      location.replace('index.html?authed=1');
      return;
    }
  } catch (e) {}

  document.body.classList.add('lg-locked');

  var user = g('lg-user'), pass = g('lg-pass'), card = g('lg-card'),
      err = g('lg-err'), btn = g('lg-btn');

  // Pre-fill username from a previous attempt on this device
  try {
    var ru = localStorage.getItem('smartis-user');
    if (ru) { user.value = ru; g('lg-keep').checked = true; }
  } catch (e) {}

  // Show / hide password
  g('lg-eye').addEventListener('click', function () {
    var show = pass.type === 'password';
    pass.type = show ? 'text' : 'password';
    this.innerHTML = '<i class="bi bi-eye' + (show ? '-slash' : '') + '"></i>';
  });

  // Caps-lock warning
  function caps(e) {
    if (e.getModifierState) g('lg-caps').classList.toggle('show', !!e.getModifierState('CapsLock'));
  }
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

  // ── Submit — verify against Supabase ──
  g('lg-form').addEventListener('submit', async function (e) {
    e.preventDefault();
    var u = user.value.trim(), p = pass.value;
    if (!u) { shake('Please enter your username', user); return; }
    if (!p) { shake('Please enter your password', pass); return; }

    if (!window.sb) {
      shake('Login service unavailable. Please reload the page.', pass);
      return;
    }

    btn.disabled = true;
    g('lg-btn-txt').textContent = 'VERIFYING…';
    g('lg-btn-ico').outerHTML = '<span class="lg-spin" id="lg-btn-ico"></span>';

    try {
      var result = await window.sb.rpc('smartis_login', {
        p_username: u,
        p_password: p
      });

      if (result.error) throw new Error(result.error.message || 'Login failed');
      if (!result.data || result.data.length === 0) {
        throw new Error('Invalid username or password');
      }

      // ✔ Credentials OK
      var account = result.data[0];

      btn.classList.add('ok');
      g('lg-btn-txt').textContent = 'ACCESS GRANTED';
      g('lg-btn-ico').outerHTML = '<i class="bi bi-check-circle-fill" id="lg-btn-ico"></i>';

      try {
        var keep = g('lg-keep').checked;
        (keep ? localStorage : sessionStorage).setItem(LG_KEY, '1');
        if (keep) localStorage.setItem('smartis-user', u);
        else localStorage.removeItem('smartis-user');
        localStorage.setItem(LG_USER_KEY, account.full_name || account.username);
      } catch (e) {}

      setTimeout(function () {
        g('lg-wname').textContent = account.full_name || account.username;
        card.classList.add('lg-hide');
        g('lg-welcome').classList.add('show');

        setTimeout(function () {
          L.classList.add('lg-out');
          document.body.classList.remove('lg-locked');
          running = false;
          setTimeout(function () {
            location.replace('index.html?authed=1');
          }, 500);
        }, 1600);
      }, 650);

    } catch (ex) {
      btn.disabled = false;
      btn.classList.remove('ok');
      g('lg-btn-txt').textContent = 'SIGN IN';
      g('lg-btn-ico').outerHTML = '<i class="bi bi-arrow-right" id="lg-btn-ico"></i>';
      shake(ex.message || 'Login failed. Check your username and password.');
    }
  });

  // ── Animated particle network ──
  var cv = g('lg-canvas'), cx = cv.getContext('2d'),
      W = 0, H = 0, D = 1, pts = [], running = true;

  function size() {
    D = Math.min(window.devicePixelRatio || 1, 2);
    W = cv.width = innerWidth * D;
    H = cv.height = innerHeight * D;
    var n = Math.min(70, Math.floor(innerWidth * innerHeight / 15000));
    pts = [];
    for (var i = 0; i < n; i++) {
      pts.push({
        x: Math.random() * W,
        y: Math.random() * H,
        vx: (Math.random() - .5) * .4 * D,
        vy: (Math.random() - .5) * .4 * D,
        r: (Math.random() * 1.4 + .6) * D
      });
    }
  }

  function draw() {
    if (!running) return;
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
        var b = pts[j];
        var dx = a.x - b.x, dy = a.y - b.y;
        var d = Math.sqrt(dx * dx + dy * dy);
        if (d < link) {
          cx.strokeStyle = 'rgba(120,160,255,' + (0.22 * (1 - d / link)) + ')';
          cx.lineWidth = D * .8;
          cx.beginPath(); cx.moveTo(a.x, a.y); cx.lineTo(b.x, b.y); cx.stroke();
        }
      }
    }
    requestAnimationFrame(draw);
  }

  size();
  window.addEventListener('resize', size);
  draw();
})();