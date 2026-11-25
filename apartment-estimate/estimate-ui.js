const summaryPrint = document.getElementById('summaryPrint');
const reviewPrint = document.getElementById('print');
const review = document.getElementById('reviewStep');
let preparingPrint = false;

function updatePrintButton() {
  summaryPrint.disabled = preparingPrint || reviewPrint.disabled ||
    typeof result === 'undefined' || !result || choices.size === 0;
}

new MutationObserver(updatePrintButton).observe(reviewPrint, { attributeFilter: ['disabled'] });
updatePrintButton();

summaryPrint.addEventListener('click', async () => {
  updatePrintButton();
  if (summaryPrint.disabled) return;

  preparingPrint = true;
  updatePrintButton();
  try {
    if (review.hidden) {
      // The existing review transition validates inputs and waits for a fresh calculation.
      await go(2);
    } else {
      if (!await calculate()) return;
      renderReview();
    }
    if (!review.hidden && !reviewPrint.disabled && result && choices.size) {
      window.print();
    }
  } finally {
    preparingPrint = false;
    updatePrintButton();
  }
});
