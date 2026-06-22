// =============================================
// COSMIC PIG BBQ - Main JS
// =============================================

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

// ---- Contact form (Formspree — real action set on form element) ----
// The form uses action="https://formspree.io/f/..." so it submits natively.
// JS handler kept only for UX feedback if JS is available and form has no action.
(function contactForm() {
  const form = document.getElementById('contact-form');
  if (!form) return;
  // Only intercept if no real Formspree ID is set yet
  const action = form.getAttribute('action') || '';
  if (action.includes('YOUR_CONTACT_FORM_ID')) {
    form.addEventListener('submit', e => {
      e.preventDefault();
      const btn = form.querySelector('button[type="submit"]');
      btn.textContent = 'Message Sent! 🚀';
      btn.disabled = true;
      btn.style.background = 'rgba(100,200,100,0.2)';
      btn.style.borderColor = '#6cc26c';
      btn.style.color = '#6cc26c';
      form.reset();
      setTimeout(() => {
        btn.textContent = 'Send Message 🚀';
        btn.disabled = false;
        btn.style = '';
      }, 4000);
    });
  }
})();
