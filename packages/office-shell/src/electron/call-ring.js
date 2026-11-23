// The ring for a call that comes while no window is there to ring.
//
// Talk rings in its own window (renderer/apps/talk/ring.js). With no Talk
// window open there was only the Windows notification, whose one short sound
// is over long before the call stops ringing. So the notifier rings for it
// (notifier.js, a report with `ring`): the same two rising notes, twice, then
// quiet, again and again until the call is answered, turned away or stops
// ringing, or a Talk window opens and rings for itself. Never both.
//
// The main process has no sound of its own, so the notes are played in a
// window that is never shown and is closed when the ring stops
// (notify-electron.js). What that window runs is written here as text, with
// nothing of Electron in it, so a test reads it and runs it against a
// stand-in for the sound device.

/**
 * The ring: its notes as [starts at, lasts, pitch in Hz] in seconds from the
 * start of a phrase, how loud at most (0 to 1), and how often the phrase
 * comes again in milliseconds. The same as Talk's own (SOUNDS.ring), which a
 * test holds it to.
 */
export const RING = { every: 3000, level: 0.22, notes: [[0, 0.18, 784], [0.22, 0.3, 988], [0.7, 0.18, 784], [0.92, 0.3, 988]] };

/** The longest a ring goes on by itself: a call that was never said to be over does not ring for ever. */
export const RING_LIMIT_MS = 90_000;

/** The page the ring is played in: nothing on it. */
export const RING_PAGE = '<!doctype html><html><head><meta charset="utf-8"><title>Ring</title></head><body></body></html>';

/**
 * What the ring's window runs: every phrase of the ring laid out ahead on the
 * sound device's own clock, so nothing depends on a timer in a window nobody
 * sees. Answers { state, phrases, notes }. Closing the window ends the sound.
 */
export function ringScript({ sound = RING, ms = RING_LIMIT_MS } = {}) {
  const phrases = Math.max(1, Math.ceil(ms / sound.every));
  return `(() => {
  const sound = ${JSON.stringify(sound)};
  const ctx = new AudioContext();
  try { const going = ctx.resume(); if (going && going.catch) going.catch(() => {}); } catch (e) { /* it plays when it can */ }
  const from = ctx.currentTime + 0.05;
  let notes = 0;
  for (let k = 0; k < ${phrases}; k += 1) {
    const start = from + (k * sound.every) / 1000;
    for (const [at, lasts, hz] of sound.notes) {
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = 'sine';
      osc.frequency.value = hz;
      // Up and down again over a moment, so a note neither clicks in nor out.
      gain.gain.setValueAtTime(0, start + at);
      gain.gain.linearRampToValueAtTime(sound.level, start + at + 0.02);
      gain.gain.setValueAtTime(sound.level, start + at + Math.max(0.02, lasts - 0.05));
      gain.gain.linearRampToValueAtTime(0, start + at + lasts);
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.start(start + at);
      osc.stop(start + at + lasts + 0.02);
      notes += 1;
    }
  }
  window.__ring = ctx;
  return { state: ctx.state, phrases: ${phrases}, notes };
})()`;
}
