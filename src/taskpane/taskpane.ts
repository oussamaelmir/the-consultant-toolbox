/* global document, Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    // Show overlay only on first run
    const seen = localStorage.getItem('ctb_firstRun');
    const overlay = document.getElementById('firstRunOverlay')!;
    const btn = document.getElementById('getStartedBtn')!;
    if (!seen) {
      overlay.style.display = 'flex';
    } else {
      overlay.style.display = 'none';
      showAppBody();
    }
    btn.addEventListener('click', () => {
      localStorage.setItem('ctb_firstRun','1');
      overlay.style.display = 'none';
      showAppBody();
      // Navigate to flags picker
      window.location.href = 'flags.html';
    });
  }
});

function showAppBody() {
  document.getElementById('sideload-msg')!.style.display = 'none';
  document.getElementById('app-body')!.style.display = 'flex';
}