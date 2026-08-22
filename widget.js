// Shared floating action button (FAB).
// Any page including this script should first define:
//   window.FAB_OPTIONS = [{ id, label, onSelect, onMount }, ...]
// before the <script src="widget.js"> tag. onMount(btn) is optional —
// called once with the created button element, useful for toggling an
// "active/armed" glow state on it from the page's own code.

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

    // Keep focus (and any active text selection in a contenteditable
    // area elsewhere on the page) intact when this button is pressed.
    btn.addEventListener('mousedown', (e) => e.preventDefault());

    btn.addEventListener('click', () => {
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

  icon.addEventListener('click', () => {
    const isOpen = wrap.classList.toggle('is-open');
    icon.setAttribute('aria-expanded', String(isOpen));
  });

  document.addEventListener('click', (e) => {
    if (!wrap.contains(e.target)) {
      wrap.classList.remove('is-open');
      icon.setAttribute('aria-expanded', 'false');
    }
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      wrap.classList.remove('is-open');
      icon.setAttribute('aria-expanded', 'false');
    }
  });
}

document.addEventListener('DOMContentLoaded', initFab);
