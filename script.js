const slides = [...document.querySelectorAll('.slide')];
const currentLabel = document.querySelector('[data-current]');
const totalLabel = document.querySelector('[data-total]');
const progress = document.querySelector('.progress__bar');
let current = 0;

slides.forEach((slide, index) => {
  const marker = slide.querySelector('.slide-number');
  if (marker) marker.textContent = String(index + 1).padStart(2, '0');
});

function clamp(value) {
  return Math.max(0, Math.min(slides.length - 1, value));
}

function show(index, updateHash = true) {
  current = clamp(index);
  slides.forEach((slide, i) => {
    slide.classList.toggle('is-active', i === current);
    slide.setAttribute('aria-hidden', i === current ? 'false' : 'true');
  });
  currentLabel.textContent = String(current + 1).padStart(2, '0');
  totalLabel.textContent = String(slides.length).padStart(2, '0');
  progress.style.width = `${((current + 1) / slides.length) * 100}%`;
  document.title = `${slides[current].dataset.title || '交通安全'}｜${current + 1}/${slides.length}`;
  if (updateHash) history.replaceState(null, '', `#${current + 1}`);
}

function next() { show(current + 1); }
function previous() { show(current - 1); }

document.addEventListener('keydown', (event) => {
  if (['ArrowRight', 'PageDown', ' ', 'Enter'].includes(event.key)) {
    event.preventDefault(); next();
  } else if (['ArrowLeft', 'PageUp', 'Backspace'].includes(event.key)) {
    event.preventDefault(); previous();
  } else if (event.key === 'Home') {
    event.preventDefault(); show(0);
  } else if (event.key === 'End') {
    event.preventDefault(); show(slides.length - 1);
  } else if (event.key.toLowerCase() === 'f') {
    document.fullscreenElement ? document.exitFullscreen() : document.documentElement.requestFullscreen();
  } else if (event.key.toLowerCase() === 'p') {
    window.print();
  }
});

document.querySelector('[data-next]').addEventListener('click', next);
document.querySelector('[data-prev]').addEventListener('click', previous);

let touchStartX = null;
document.addEventListener('touchstart', (event) => { touchStartX = event.changedTouches[0].clientX; }, { passive: true });
document.addEventListener('touchend', (event) => {
  if (touchStartX === null) return;
  const delta = event.changedTouches[0].clientX - touchStartX;
  if (Math.abs(delta) > 60) delta < 0 ? next() : previous();
  touchStartX = null;
}, { passive: true });

const fromHash = Number.parseInt(location.hash.slice(1), 10);
show(Number.isFinite(fromHash) ? fromHash - 1 : 0, false);

