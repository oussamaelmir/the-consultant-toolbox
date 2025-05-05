/* global document, Office */
Office.onReady(async (info) => {
  if (info.host === Office.HostType.PowerPoint) {
    const btn = document.getElementById('getStartedBtn');
    btn?.addEventListener('click', async () => {
      try {
        // This will hide/close the pane just as if the user clicked the X
        await Office.addin.hide();
      } catch (err) {
        console.error('Hide API not available, fallback if needed', err);
      }
    });
  }
});
