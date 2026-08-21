const moduleLinks = {
  sales: 'sales.html',
  entrepreneurial: 'entrepreneurial.html',
  product: 'product.html',
};

document.querySelectorAll('.card').forEach((card) => {
  // Make it keyboard-accessible since these are <div>s, not <a> tags
  card.setAttribute('tabindex', '0');
  card.setAttribute('role', 'link');

  const goToModule = () => {
    const module = card.dataset.module;
    const url = moduleLinks[module];
    if (url) {
      window.location.href = url;
    } else {
      console.warn('No page mapped for module:', module);
    }
  };

  card.addEventListener('click', goToModule);

  card.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      goToModule();
    }
  });
});
