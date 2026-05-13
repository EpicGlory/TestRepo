/* Clean and Green Turf — small interactive scripts */

// Custom multi-select dropdown (used on the "Interested in" form field)
document.querySelectorAll('[data-multi-select]').forEach(el => {
  const trigger = el.querySelector('.multi-select-trigger');
  const panel = el.querySelector('.multi-select-panel');
  const label = el.querySelector('.multi-select-label');
  const checkboxes = el.querySelectorAll('input[type="checkbox"]');
  const defaultLabel = label.textContent;

  function refreshLabel() {
    const checked = Array.from(checkboxes).filter(c => c.checked);
    if (checked.length === 0) {
      label.textContent = defaultLabel;
      label.classList.remove('has-selection');
    } else if (checked.length === 1) {
      label.textContent = checked[0].nextElementSibling.textContent.trim();
      label.classList.add('has-selection');
    } else {
      label.textContent = checked.length + ' services selected';
      label.classList.add('has-selection');
    }
  }

  trigger.addEventListener('click', e => {
    e.preventDefault();
    e.stopPropagation();
    el.classList.toggle('open');
  });

  checkboxes.forEach(cb => {
    cb.addEventListener('change', refreshLabel);
  });

  // Click outside to close
  document.addEventListener('click', e => {
    if (!el.contains(e.target)) el.classList.remove('open');
  });

  // Escape to close
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') el.classList.remove('open');
  });
});
