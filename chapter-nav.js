// Wires up chapter cards on module pages (sales.html, entrepreneurial.html,
// product.html) so pressing one navigates to the shared chapter.html
// template, passing along which module/chapter/color/title it should show.

document.addEventListener('DOMContentLoaded', () => {
  const module = document.body.dataset.module || '';
  const color = document.body.dataset.color || 'blue';
  const moduleTitle = document.body.dataset.moduleTitle || '';

  document.querySelectorAll('.chapter-card').forEach((card) => {
    const num = card.dataset.chapter || '1';
    card.setAttribute('tabindex', '0');
    card.setAttribute('role', 'link');

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
