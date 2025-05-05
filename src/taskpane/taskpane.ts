/* taskpane.ts /
/ global document, Office */

Office.onReady(async (info) => {
  if (info.host === Office.HostType.PowerPoint) {
  const btn = document.getElementById('getStartedBtn');
  btn?.addEventListener('click', async () => {
  try {
  // Close the task pane as if the user clicked the X
  await Office.addin.hide();
  } catch (error) {
  console.error('Office.addin.hide() failed', error);
  }
  });
  }
  });