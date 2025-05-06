import { createPostItNote } from "../commands/commands"; 
// adjust the path if your commands file lives elsewhere

/* global document, Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    const btn = document.getElementById('getStartedBtn');
    btn?.addEventListener('click', async () => {
      // fake up the Event object with a no-op completed()
      const fakeEvent = {
        completed: () => {
          /* nothing to do */
        }
      } as Office.AddinCommands.Event;

      // call your Post-It creation routine
      try {
        await createPostItNote(fakeEvent);
      } catch (err) {
        console.error('Error in createPostItNote:', err);
      }
    });
  }
});
