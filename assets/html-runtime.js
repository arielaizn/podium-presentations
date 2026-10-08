/* Local, dependency-free Podium presentation controls. */
(() => {
  'use strict';
  const slides = [...document.querySelectorAll('.slide')];
  if (!slides.length) return;
  const prev = document.getElementById('prev-slide');
  const next = document.getElementById('next-slide');
  const status = document.getElementById('slide-status');
  const rtl = document.documentElement.dir === 'rtl';
  let current = 0;
  function fromHash() {
    const n = Number(location.hash.replace(/^#slide-/, ''));
    return Number.isInteger(n) && n > 0 ? Math.min(n - 1, slides.length - 1) : 0;
  }
  function show(n, updateHash = true) {
    current = Math.min(Math.max(n, 0), slides.length - 1);
    slides.forEach((slide, i) => {
      const active = i === current;
      slide.classList.toggle('active', active);
      slide.setAttribute('aria-hidden', String(!active));
      slide.inert = !active;
      if (active) slide.scrollTop = 0;
    });
    if (prev) prev.disabled = current === 0;
    if (next) next.disabled = current === slides.length - 1;
    if (status) status.textContent = `${current + 1} / ${slides.length}`;
    if (updateHash) history.replaceState(null, '', `#slide-${current + 1}`);
    document.dispatchEvent(new CustomEvent('podium:slide', {detail:{index:current,total:slides.length}}));
  }
  prev?.addEventListener('click', () => show(current - 1));
  next?.addEventListener('click', () => show(current + 1));
  document.getElementById('fullscreen')?.addEventListener('click', async () => {
    try { if (!document.fullscreenElement) await document.documentElement.requestFullscreen(); else await document.exitFullscreen(); }
    catch (_) { /* Keep presentation usable if fullscreen is denied. */ }
  });
  document.addEventListener('keydown', event => {
    if (event.altKey || event.ctrlKey || event.metaKey || event.target.closest?.('input,textarea,select,button,a,[contenteditable=true]') || event.target.isContentEditable) return;
    let n;
    if (event.key === (rtl ? 'ArrowLeft' : 'ArrowRight') || event.key === ' ' || event.key === 'PageDown') n = current + 1;
    else if (event.key === (rtl ? 'ArrowRight' : 'ArrowLeft') || event.key === 'PageUp') n = current - 1;
    else if (event.key === 'Home') n = 0;
    else if (event.key === 'End') n = slides.length - 1;
    if (n !== undefined) { event.preventDefault(); show(n); }
  });
  let start = null;
  document.addEventListener('touchstart', event => {
    if (event.touches.length === 1) start = [event.touches[0].clientX,event.touches[0].clientY];
  }, {passive:true});
  document.addEventListener('touchend', event => {
    if (!start || !event.changedTouches.length) return;
    const dx = event.changedTouches[0].clientX-start[0], dy = event.changedTouches[0].clientY-start[1];
    start = null;
    if (Math.abs(dx)>65 && Math.abs(dx)>Math.abs(dy)*1.5) show(current + ((rtl ? dx>0 : dx<0) ? 1 : -1));
  }, {passive:true});
  addEventListener('hashchange', () => show(fromHash(), false));
  show(fromHash(), false);
})();
