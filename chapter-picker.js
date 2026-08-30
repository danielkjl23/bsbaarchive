// ============================================================
// "Select Chapter Download" popup — shown from the FAB on module
// pages (sales.html, product.html, entrepreneurial.html). Lets the
// user glow-select one or more chapters, then downloads them all as
// a single merged .docx (each chapter gets a "CHAPTER N" heading and
// starts on its own page). Depends on docx-export.js + JSZip being
// loaded first, and on chapter-content.js for window.CHAPTER_CONTENT.
// ============================================================

function chapterPickerGetChapters() {
  return Array.from(document.querySelectorAll('.chapter-card'))
    .map((card) => {
      const num = card.dataset.chapter;
      const titleEl = card.querySelector('.chapter-title-space');
      const title = titleEl ? titleEl.textContent.trim() : '';
      return { num, title };
    })
    .sort((a, b) => Number(a.num) - Number(b.num));
}

// Same "use saved marks if they still match the base text" logic as
// chapter.html's own loadContent(), so a chapter you've highlighted/
// underlined/bolded downloads with those marks intact.
function chapterPickerGetWorkingHtml(module, num) {
  const contentKey = `${module}-${num}`;
  const baseHtml = (window.CHAPTER_CONTENT && window.CHAPTER_CONTENT[contentKey]) || '';
  let workingHtml = baseHtml;
  try {
    const saved = JSON.parse(localStorage.getItem(`bsba-progress:${contentKey}`) || 'null');
    if (saved && saved.baseHtml === baseHtml && typeof saved.workingHtml === 'string') {
      workingHtml = saved.workingHtml;
    }
  } catch (err) {
    // Storage blocked/corrupt — fall back to base content.
  }
  return workingHtml;
}

function chapterPickerIsEmpty(module, num) {
  const baseHtml = (window.CHAPTER_CONTENT && window.CHAPTER_CONTENT[`${module}-${num}`]) || '';
  return baseHtml.trim() === '';
}

window.openChapterDownloadPicker = function (module, moduleTitle) {
  const chapters = chapterPickerGetChapters();
  if (!chapters.length) return;

  const selected = new Set();

  const overlay = document.createElement('div');
  overlay.className = 'chapter-picker-overlay';

  const modal = document.createElement('div');
  modal.className = 'chapter-picker-modal';

  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.className = 'chapter-picker-close';
  closeBtn.setAttribute('aria-label', 'Close');
  closeBtn.innerHTML =
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6 6 18"></path><path d="M6 6l12 12"></path></svg>';

  const heading = document.createElement('div');
  heading.className = 'chapter-picker-heading';
  heading.textContent = 'Select chapters to download';

  const list = document.createElement('div');
  list.className = 'chapter-picker-list';

  chapters.forEach(({ num, title }) => {
    const isEmpty = chapterPickerIsEmpty(module, num);
    const chip = document.createElement('button');
    chip.type = 'button';
    chip.className = 'chapter-picker-chip';
    const label = title ? `Chapter ${num} \u2014 ${title}` : `Chapter ${num}`;
    chip.innerHTML = isEmpty
      ? `<span>${label}</span><span class="chapter-picker-chip-note">In development</span>`
      : `<span>${label}</span>`;

    chip.addEventListener('click', () => {
      if (selected.has(num)) {
        selected.delete(num);
        chip.classList.remove('is-selected');
      } else {
        selected.add(num);
        chip.classList.add('is-selected');
      }
      proceedBtn.disabled = selected.size === 0;
    });

    list.appendChild(chip);
  });

  const proceedBtn = document.createElement('button');
  proceedBtn.type = 'button';
  proceedBtn.className = 'chapter-picker-proceed';
  proceedBtn.disabled = true;
  proceedBtn.innerHTML =
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12"></path><path d="m7 11 5 5 5-5"></path><path d="M5 21h14"></path></svg><span>Download</span>';

  function close() {
    overlay.remove();
    document.removeEventListener('keydown', onKeydown);
  }

  function onKeydown(e) {
    if (e.key === 'Escape') close();
  }

  proceedBtn.addEventListener('click', async () => {
    if (!selected.size || proceedBtn.disabled) return;
    proceedBtn.disabled = true;
    proceedBtn.classList.add('is-loading');

    const nums = Array.from(selected)
      .map(Number)
      .sort((a, b) => a - b);

    try {
      const chaptersData = nums.map((n) => ({
        num: n,
        html: chapterPickerGetWorkingHtml(module, n),
      }));
      const blob = await window.DocxExport.buildMultiChapterBlob(chaptersData);
      const modulePart = moduleTitle.replace(/\s+/g, '-');
      const filename = `${modulePart}-Chapter-${nums.join(',')}.docx`;
      window.DocxExport.triggerDownload(blob, filename);
      close();
    } catch (err) {
      alert('Sorry, something went wrong building the Word file.');
      proceedBtn.disabled = false;
      proceedBtn.classList.remove('is-loading');
    }
  });

  modal.appendChild(closeBtn);
  modal.appendChild(heading);
  modal.appendChild(list);
  modal.appendChild(proceedBtn);
  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  closeBtn.addEventListener('click', close);
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) close();
  });
  document.addEventListener('keydown', onKeydown);
};
