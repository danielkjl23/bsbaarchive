// Shared floating action button (FAB).
// Any page including this script should first define:
//   window.FAB_OPTIONS = [{ id, label, onSelect, onMount }, ...]
// before the <script src="widget.js"> tag. onMount(btn) is optional —
// called once with the created button element, useful for toggling an
// "active/armed" glow state on it from the page's own code.

// Fires `handler` once per tap, on whichever event actually arrives first
// (touchend on touch devices, click on mouse/desktop) — and prevents the
// browser's synthetic click-after-touchend from double-firing it. Also
// preventDefaults touchstart so the tap doesn't scroll/zoom or collapse
// a text selection elsewhere on the page before the tap registers.
function addTapListener(el, handler) {
  let firedByTouch = false;
  el.addEventListener('touchstart', (e) => e.preventDefault(), { passive: false });
  el.addEventListener(
    'touchend',
    (e) => {
      e.preventDefault();
      firedByTouch = true;
      handler(e);
      setTimeout(() => {
        firedByTouch = false;
      }, 500);
    },
    { passive: false }
  );
  el.addEventListener('click', (e) => {
    if (firedByTouch) return;
    handler(e);
  });
}
window.addTapListener = addTapListener;

function initFab() {
  const options = window.FAB_OPTIONS || [];
  if (!options.length) return;

  const wrap = document.createElement('div');
  wrap.className = 'fab-wrap';
  wrap.id = 'fabWrap';

  const optionsList = document.createElement('div');
  optionsList.className = 'fab-options';
  optionsList.id = 'fabOptions';
  optionsList.setAttribute('role', 'menu');

  options.forEach((opt) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'fab-option';
    btn.textContent = opt.label;
    btn.dataset.id = opt.id || '';
    btn.setAttribute('role', 'menuitem');

    // Keep focus (and any active text selection elsewhere on the page)
    // intact when this button is pressed.
    btn.addEventListener('mousedown', (e) => e.preventDefault());

    addTapListener(btn, () => {
      if (typeof opt.onSelect === 'function') {
        opt.onSelect(btn);
      }
    });

    optionsList.appendChild(btn);

    if (typeof opt.onMount === 'function') {
      opt.onMount(btn);
    }
  });

  const icon = document.createElement('button');
  icon.type = 'button';
  icon.className = 'fab-icon';
  icon.id = 'fabToggle';
  icon.setAttribute('aria-label', 'Open chapter options');
  icon.setAttribute('aria-expanded', 'false');
  icon.innerHTML =
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 18 9 12 15 6"></polyline></svg>';

  wrap.appendChild(optionsList);
  wrap.appendChild(icon);
  document.body.appendChild(wrap);
  positionFabToContent(wrap);

  addTapListener(icon, () => {
    const isOpen = wrap.classList.toggle('is-open');
    icon.setAttribute('aria-expanded', String(isOpen));
  });

  document.addEventListener('click', (e) => {
    if (!wrap.contains(e.target)) {
      wrap.classList.remove('is-open');
      icon.setAttribute('aria-expanded', 'false');
    }
  });

  document.addEventListener(
    'touchstart',
    (e) => {
      if (!wrap.contains(e.target)) {
        wrap.classList.remove('is-open');
        icon.setAttribute('aria-expanded', 'false');
      }
    },
    { passive: true }
  );

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      wrap.classList.remove('is-open');
      icon.setAttribute('aria-expanded', 'false');
    }
  });
}

document.addEventListener('DOMContentLoaded', initFab);

// ---- Desktop only: position the FAB so it hugs the content
// container's right edge instead of sitting a fixed distance from
// the browser window's edge (the container's width varies by page
// and viewport, so this is measured rather than hardcoded). Mobile
// is left completely alone — it keeps the existing right: 1.75rem
// from style.css, unchanged. ----
function positionFabToContent(wrap) {
  const desktopQuery = window.matchMedia('(min-width: 601px)');
  const EDGE_GAP = 16; // px, ~1rem gap between the FAB and the content edge
  const CONTAINER_SELECTORS = ['#chapterContentWrap', '.chapter-list', '.cards'];

  function update() {
    if (!desktopQuery.matches) {
      wrap.style.left = '';
      wrap.style.right = '';
      return;
    }

    let containerEl = null;
    for (let i = 0; i < CONTAINER_SELECTORS.length; i++) {
      containerEl = document.querySelector(CONTAINER_SELECTORS[i]);
      if (containerEl) break;
    }

    if (!containerEl) {
      wrap.style.left = '';
      wrap.style.right = '';
      return;
    }

    const rect = containerEl.getBoundingClientRect();
    wrap.style.right = 'auto';
    wrap.style.left = `${Math.round(rect.right + EDGE_GAP)}px`;
  }

  // Same fix as chapter.html's toolbar positioning: throttle the
  // resize-driven recalculation to once per animation frame so a zoom
  // gesture (which fires many 'resize' events back to back) doesn't
  // stack up a forced layout read on every single one of them.
  function rafThrottle(fn) {
    let scheduled = false;
    return (...args) => {
      if (scheduled) return;
      scheduled = true;
      requestAnimationFrame(() => {
        scheduled = false;
        fn(...args);
      });
    };
  }

  update();
  window.addEventListener('resize', rafThrottle(update));
  window.addEventListener('load', update);
}
