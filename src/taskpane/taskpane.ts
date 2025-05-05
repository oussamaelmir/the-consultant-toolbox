/* global document, Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    // Hook up 'Get Started' button
    const btn = document.getElementById('getStartedBtn');
    btn?.addEventListener('click', () => {
      // Navigate to your flags UI
      window.location.href = 'flags.html';
    });

    // Show the rest of your page only after first-run placemat is gone (if desired)
    const app = document.getElementById('app-body');
    const sideload = document.getElementById('sideload-msg');
    if (app && sideload) {
      sideload.style.display = 'none';
      app.style.display = 'flex';
      document.getElementById('run')!.onclick = run;
    }
  }
});

export async function run() {
  const options: Office.SetSelectedDataOptions = { coercionType: Office.CoercionType.Text };
  await Office.context.document.setSelectedDataAsync('Hello World!', options);
}