// Wires up chapter cards on module pages (sales.html, entrepreneurial.html,
// product.html) so pressing one navigates to the shared chapter.html
// template, passing along which module/chapter/color/title it should show.

// Shrinks a filled-in chapter title just enough to fit within its
// card (max 2 lines) — since titles aren't known ahead of time, this
// tries the normal size first and only steps down if it overflows.
function fitChapterTitle(el) {
  const maxLines = 2;
  const minScale = 0.72;
  const step = 0.04;
  let scale = 1;
  el.style.setProperty('--title-fit-scale', scale);

  const lineHeight = 1.35;
  const baseSizePx = 0.95 * 16; // matches the 0.95rem base in style.css
  const maxHeight = baseSizePx * lineHeight * maxLines + 2;

  while (scale > minScale && el.scrollHeight > maxHeight) {
    scale = Math.max(minScale, scale - step);
    el.style.setProperty('--title-fit-scale', scale);
  }
}

document.addEventListener('DOMContentLoaded', () => {
  const module = document.body.dataset.module || '';
  const color = document.body.dataset.color || 'blue';
  const moduleTitle = document.body.dataset.moduleTitle || '';

  document.querySelectorAll('.chapter-card').forEach((card) => {
    const num = card.dataset.chapter || '1';
    card.setAttribute('tabindex', '0');
    card.setAttribute('role', 'link');

    const titleSpace = card.querySelector('.chapter-title-space');
    const titleText = (window.CHAPTER_TITLES && window.CHAPTER_TITLES[`${module}-${num}`]) || '';
    if (titleSpace && titleText) {
      titleSpace.textContent = titleText;
      titleSpace.classList.add('has-title');
      fitChapterTitle(titleSpace);
    }

    const go = () => {
      const params = new URLSearchParams({ module, num, color, moduleTitle });
      window.location.href = `chapter.html?${params.toString()}`;
    };

    card.addEventListener('click', go);
    card.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        go();
      }
    });
  });
});
