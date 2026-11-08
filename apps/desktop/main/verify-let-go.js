import { BrowserWindow } from 'electron';

/**
 * Closes the Recorder and Live Broadcast windows a check has finished with.
 * Each one left open keeps drawing its scenes and running its camera, and a few
 * of them together slow what a later block does until it misses its time limit.
 * The last window destroyed quits the application, so one stays if nothing else is open.
 */
export function letGoOfStudioWindows() {
  const open = BrowserWindow.getAllWindows().filter((w) => !w.isDestroyed());
  const studio = open.filter((w) => /app=(broadcast|recorder)\b/.test(w.webContents.getURL()));
  if (studio.length === open.length) studio.shift();
  for (const w of studio) w.destroy();
}
