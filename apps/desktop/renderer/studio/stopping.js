// Closing the window while a take is still being finished.

/** Waits until the take has finished being written. A window that closes sooner cuts the end off the file. */
export async function whileStopping(phaseOf, { most = 60000, every = 100, wait = (ms) => new Promise((r) => setTimeout(r, ms)) } = {}) {
  for (let waited = 0; phaseOf() === 'stopping' && waited < most; waited += every) await wait(every);
}
