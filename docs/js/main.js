/* ============================================================
   DiegoAutomatiza, ANDAMIO: motor de movimiento
   Un solo reloj (gsap.ticker). Sin requestAnimationFrame propios,
   sin listeners de scroll. Reduced motion: todo quieto y visible.
   ============================================================ */

const html = document.documentElement;
const tieneMotion = html.classList.contains('motion');
const punteroFino = matchMedia('(hover: hover) and (pointer: fine)').matches;

if (!window.gsap) {
  // Sin motor: el contenido ya es visible, solo escondemos el loader.
  const loader = document.getElementById('loader');
  if (loader) loader.hidden = true;
  html.classList.remove('motion');
} else {
  const { gsap } = window;
  const ScrollTrigger = window.ScrollTrigger;
  gsap.registerPlugin(ScrollTrigger);

  /* ---------- Scroll: las anclas navegan con scroll suave por JS; el scroll
     programático del navegador (foco, find-in-page, herramientas de QA) queda
     instantáneo, igual que bajo reduced motion ---------- */
  document.addEventListener('click', (e) => {
    const a = e.target.closest('a[href^="#"]');
    if (!a) return;
    const destino = document.querySelector(a.getAttribute('href'));
    if (!destino) return;
    e.preventDefault();
    destino.scrollIntoView({ behavior: tieneMotion ? 'smooth' : 'auto', block: 'start' });
    if (destino.hasAttribute('tabindex')) destino.focus({ preventScroll: true });
    history.replaceState(null, '', a.getAttribute('href'));
  });

  /* ---------- Reveal: los bloques suben y encajan (opacidad, no visibility,
     para que los links sigan siendo enfocables; el foco los revela al instante) ---------- */
  const reveals = new Map();
  if (tieneMotion) {
    gsap.utils.toArray('[data-reveal]').forEach((el) => {
      // fromTo explícito: el CSS ya pone opacity 0, y gsap.from leería ese 0 como destino.
      const tween = gsap.fromTo(el, { opacity: 0, y: 42 }, {
        opacity: 1,
        y: 0,
        duration: 1.15,
        ease: 'expo.out',
        scrollTrigger: { trigger: el, start: 'top 92%', once: true },
        onComplete: () => gsap.set(el, { clearProps: 'transform,translate,rotate,scale' }),
      });
      reveals.set(el, tween);
    });
    document.addEventListener('focusin', (e) => {
      const host = e.target.closest('[data-reveal]');
      if (!host) return;
      const t = reveals.get(host);
      if (t && parseFloat(getComputedStyle(host).opacity) < 1) {
        t.scrollTrigger && t.scrollTrigger.kill();
        t.progress(1);
      }
      // El foco de teclado debe caer a la vista: centramos el elemento enfocado,
      // no su bloque (el bloque puede ser más alto que la vista).
      const r = e.target.getBoundingClientRect();
      if (r.bottom < 96 || r.top > innerHeight - 64) {
        e.target.scrollIntoView({ block: 'center', behavior: 'instant' });
      }
    });
  }

  /* ---------- Loader: contador de obra que entrega al hero ---------- */
  async function runLoader() {
    const root = document.getElementById('loader');
    if (!root) { return; }
    if (!tieneMotion) {
      root.hidden = true;
      return;
    }
    await Promise.race([document.fonts.ready, new Promise((r) => setTimeout(r, 150))]);
    const count = root.querySelector('[data-loader-count]');
    const bar = root.querySelector('[data-loader-bar]');
    const arranque = buildIntro(); // pausado: se arma mientras la cortina sube
    // En redes lentas la carga ya consumió tiempo: el contador se acorta para no sumar.
    const cargaLenta = performance.now() > 1500;
    const cargaMuyLenta = performance.now() > 2500;
    const o = { v: 0 };
    await gsap.to(o, {
      v: 100,
      duration: cargaMuyLenta ? 0.08 : cargaLenta ? 0.14 : 0.2,
      ease: 'power2.inOut',
      onUpdate: () => {
        if (count) count.textContent = String(Math.round(o.v)).padStart(3, '0');
        if (bar) bar.style.transform = `scaleX(${o.v / 100})`;
      },
    });
    // Salida: la cortina sube como una plataforma y entrega al hero que se arma.
    // Presupuesto del loader: <=1.2 s desde que arranca (decisión de la pasada 3).
    arranque.play();
    await gsap.timeline({ onComplete: limpiarIntro })
      .to(root.querySelectorAll('[data-loader-out]'), { yPercent: -130, duration: cargaMuyLenta ? 0.14 : cargaLenta ? 0.18 : 0.22, ease: 'power2.in', stagger: 0.02 })
      .to(root, { clipPath: 'inset(0% 0% 100% 0%)', duration: cargaMuyLenta ? 0.22 : cargaLenta ? 0.28 : 0.3, ease: 'expo.inOut' }, '-=0.1');
    root.hidden = true;
  }

  // Al terminar la intro se retiran los transform inline: el DOM queda limpio
  // y axe puede verificar el contraste de los titulares sin revisión manual.
  function limpiarIntro() {
    gsap.set('.plataforma-texto, .traba, .rail, .hero-meta, .hero-sub, .hero-acciones, .hero-tags', { clearProps: 'transform,translate,rotate,scale' });
  }

  /* ---------- Intro: el titular se monta pieza a pieza (firma) ---------- */
  function buildIntro() {
    const asienta = ['.hero-meta', '.hero-sub', '.hero-acciones', '.hero-tags'];
    return gsap.timeline({ paused: true, defaults: { ease: 'expo.out' } })
      .set('.plataforma-texto', { visibility: 'visible' })
      .from('.plataforma-texto', { yPercent: 115, duration: 1.1, stagger: 0.14 }, 0)
      .fromTo('.traba', { scaleX: 0 }, { scaleX: 1, duration: 0.55, ease: 'power3.out', stagger: 0.14 }, 0.6)
      .fromTo('.rail', { scaleY: 0 }, { scaleY: 1, duration: 1.5 }, 0)
      .from(asienta.join(','), { y: 26, duration: 0.9, stagger: 0.09 }, 0.35);
  }

  /* ---------- Idle: el andamio se asienta ±2 px (pausado fuera de pantalla) ---------- */
  if (tieneMotion) {
    const plataformas = gsap.utils.toArray('.plataforma');
    if (plataformas.length) {
      let activo = false;
      let enTick = false;
      const tick = (t) => {
        plataformas.forEach((p, i) => {
          const v = Math.sin(t * 0.55 + i * 1.9) * 2;
          p.style.translate = `0 ${v.toFixed(2)}px`;
        });
      };
      ScrollTrigger.create({
        trigger: '.hero',
        start: 'top bottom',
        end: 'bottom top',
        onToggle: (s) => {
          activo = s.isActive;
          if (activo && !enTick) { enTick = true; gsap.ticker.add(tick); }
          else if (!activo && enTick) { enTick = false; gsap.ticker.remove(tick); plataformas.forEach((p) => (p.style.translate = '0 0')); }
        },
      });
    }
  }

  /* ---------- Botones: relleno que inunda desde el lado del puntero ---------- */
  if (tieneMotion && punteroFino) {
    document.querySelectorAll('.btn').forEach((btn) => {
      const fill = btn.querySelector('.btn-relleno');
      if (!fill) return;
      const punto = (e) => {
        const r = btn.getBoundingClientRect();
        return `${e.clientX - r.left}px ${e.clientY - r.top}px`;
      };
      let cerrado = true;
      const abre = (p) => {
        if (cerrado) gsap.set(fill, { clipPath: `circle(0% at ${p})` });
        cerrado = false;
        btn.classList.add('is-lleno');
        gsap.to(fill, { clipPath: `circle(150% at ${p})`, duration: 0.6, ease: 'power3.out', overwrite: true });
      };
      const cierra = (p) => {
        btn.classList.remove('is-lleno');
        gsap.to(fill, { clipPath: `circle(0% at ${p})`, duration: 0.5, ease: 'power3.out', overwrite: true, onComplete: () => { cerrado = true; } });
      };
      btn.addEventListener('pointerenter', (e) => abre(punto(e)));
      btn.addEventListener('pointerleave', (e) => cierra(punto(e)));
      btn.addEventListener('focus', () => { if (btn.matches(':focus-visible')) abre('50% 50%'); });
      btn.addEventListener('blur', () => cierra('50% 50%'));
    });
  }

  runLoader();
}
