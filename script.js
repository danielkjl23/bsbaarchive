document.querySelectorAll('.card').forEach((card) => {
  card.addEventListener('click', () => {
    const module = card.dataset.module;
    console.log('Explore module:', module);
  });
});
