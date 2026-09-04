// =============================================
// COSMIC PIG BBQ - Main JS
// =============================================

// ---- Recipe details: auto-open when hash link is clicked or page loads with hash ----
(function initRecipeToggle() {
  function openRecipeDetails(hash) {
    if (!hash) return;
    var target = document.getElementById(hash.replace('#', ''));
    if (!target) return;
    var details = target.querySelector('.full-recipe-details');
    if (details) {
      details.open = true;
      setTimeout(function() { target.scrollIntoView({ behavior: 'smooth', block: 'start' }); }, 50);
    }
  }

  // On page load with a hash
  if (window.location.hash) openRecipeDetails(window.location.hash);

  // On recipe card "View Recipe" clicks
  document.querySelectorAll('a[href^="#"]').forEach(function(link) {
    link.addEventListener('click', function() {
      openRecipeDetails(this.getAttribute('href'));
    });
  });
})();

// ---- Starfield ----
(function initStars() {
  const canvas = document.getElementById('stars-canvas');
  if (!canvas) return;

  const ctx = canvas.getContext('2d');
  let stars = [];
  let raf;

  function resize() {
    canvas.width  = window.innerWidth;
    canvas.height = window.innerHeight;
  }

  function createStars(count) {
    stars = [];
    for (let i = 0; i < count; i++) {
      stars.push({
        x:       Math.random() * canvas.width,
        y:       Math.random() * canvas.height,
        r:       Math.random() * 1.4 + 0.2,
        alpha:   Math.random(),
        speed:   Math.random() * 0.008 + 0.002,
        dir:     Math.random() > 0.5 ? 1 : -1,
        twinkle: Math.random() * Math.PI * 2,
      });
    }
  }

  function draw() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    stars.forEach(s => {
      s.twinkle += s.speed * s.dir;
      const a = 0.35 + Math.sin(s.twinkle) * 0.35;
      ctx.beginPath();
      ctx.arc(s.x, s.y, s.r, 0, Math.PI * 2);
      ctx.fillStyle = `rgba(255,255,255,${a})`;
      ctx.fill();
    });
    raf = requestAnimationFrame(draw);
  }

  window.addEventListener('resize', () => {
    resize();
    createStars(220);
  });

  resize();
  createStars(220);
  draw();
})();

// ---- Sticky header ----
(function stickyHeader() {
  const header = document.querySelector('.site-header');
  if (!header) return;
  window.addEventListener('scroll', () => {
    header.classList.toggle('scrolled', window.scrollY > 50);
  }, { passive: true });
})();

// ---- Mobile nav ----
(function mobileNav() {
  const btn   = document.querySelector('.hamburger');
  const links = document.querySelector('.nav-links');
  if (!btn || !links) return;

  btn.addEventListener('click', () => {
    const open = links.classList.toggle('open');
    btn.classList.toggle('open', open);
    btn.setAttribute('aria-expanded', open);
  });

  // Close on link click
  links.querySelectorAll('a').forEach(a => {
    a.addEventListener('click', () => {
      links.classList.remove('open');
      btn.classList.remove('open');
      btn.setAttribute('aria-expanded', 'false');
    });
  });
})();

// ---- Active nav link ----
(function activeNav() {
  const page = window.location.pathname.split('/').pop() || 'index.html';
  document.querySelectorAll('.nav-links a').forEach(a => {
    const href = a.getAttribute('href');
    if (href === page || (page === '' && href === 'index.html')) {
      a.classList.add('active');
    }
  });
})();

// ---- Scroll reveal ----
(function scrollReveal() {
  const els = document.querySelectorAll('.reveal');
  if (!els.length) return;

  if (!window.IntersectionObserver) {
    els.forEach(el => el.classList.add('visible'));
    return;
  }

  const io = new IntersectionObserver(entries => {
    entries.forEach(e => {
      if (e.isIntersecting) {
        e.target.classList.add('visible');
        io.unobserve(e.target);
      }
    });
  }, { threshold: 0.12 });

  els.forEach(el => io.observe(el));
})();

// ---- All forms → Google Apps Script → Google Sheet ----
// After deploying scripts/preorder-sheet.gs, paste your Web App URL here:
const FORMS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxTezN1di2XZHe4SDC20xF6vIcVTTATe6C2AzpQ3YEorM5G39sW1ItnYMpLwpgzwkR6/exec';

async function sendToSheet(data) {
  if (FORMS_SCRIPT_URL === 'YOUR_APPS_SCRIPT_URL') return;
  try {
    await fetch(FORMS_SCRIPT_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify(data)
    });
  } catch (_) {
    // no-cors fetch may throw on opaque response; submission still sent
  }
}

try { (function preorderForm() {
  const form = document.getElementById('preorder-form');
  if (!form) return;

  const btn = form.querySelector('button[type="submit"]');

  // Phone auto-format → (801) 358-7820
  const phoneInput = form.querySelector('input[name="phone"]');
  if (phoneInput) {
    phoneInput.addEventListener('input', function() {
      const digits = this.value.replace(/\D/g, '').slice(0, 10);
      let fmt = '';
      if (digits.length > 0) fmt = '(' + digits.slice(0, 3);
      if (digits.length >= 4) fmt += ') ' + digits.slice(3, 6);
      if (digits.length >= 7) fmt += '-' + digits.slice(6);
      this.value = fmt;
    });
  }

  form.addEventListener('submit', async function(e) {
    e.preventDefault();

    const fd = new FormData(form);
    const data = {};
    fd.forEach(function(val, key) {
      if (key === 'rubs') {
        data.rubs = data.rubs ? data.rubs + ', ' + val : val;
      } else {
        data[key] = val;
      }
    });

    btn.textContent = 'Sending…';
    btn.disabled = true;

    data.form_type = 'preorder';
    await sendToSheet(data);

    window.location.href = 'preorder-thanks.html';
  });
})(); } catch (_) {}

// ---- Contact form → Google Sheet ----
try { (function contactForm() {
  const form = document.getElementById('contact-form');
  if (!form) return;
  const btn = form.querySelector('button[type="submit"]');

  form.addEventListener('submit', async function(e) {
    e.preventDefault();
    btn.textContent = 'Sending…';
    btn.disabled = true;

    const fd = new FormData(form);
    const data = { form_type: 'contact' };
    fd.forEach(function(val, key) { data[key] = val; });

    await sendToSheet(data);

    btn.textContent = 'Message Sent! 🚀';
    btn.style.background   = 'rgba(100,200,100,0.2)';
    btn.style.borderColor  = '#6cc26c';
    btn.style.color        = '#6cc26c';
    form.reset();
    setTimeout(function() {
      btn.textContent = 'Send Message 🚀';
      btn.disabled = false;
      btn.style.background  = '';
      btn.style.borderColor = '';
      btn.style.color       = '';
    }, 4000);
  });
})(); } catch (_) {}

// ---- Email signup form → Google Sheet ----
try { (function signupForm() {
  const form = document.getElementById('signup-form');
  if (!form) return;
  const btn = form.querySelector('button[type="submit"]');

  form.addEventListener('submit', async function(e) {
    e.preventDefault();
    btn.textContent = 'Sending…';
    btn.disabled = true;

    const fd = new FormData(form);
    const data = { form_type: 'signup' };
    fd.forEach(function(val, key) { data[key] = val; });

    await sendToSheet(data);

    btn.textContent = 'You\'re in! 🚀';
    form.reset();
    setTimeout(function() {
      btn.textContent = 'Join the Crew 🚀';
      btn.disabled = false;
    }, 4000);
  });
})(); } catch (_) {}
