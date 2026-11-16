// One spoken line's turn in the studio (spoken-run.js): its voice played into
// the output, or, when it has none, the line held up for its reading time.
//
// A line has no voice when it was made for Show only (it is timed from its
// words, not spoken). If the choice changes to one that reads aloud while
// that line is being shown, it is still shown for its time, silently: an empty
// sound played would end at once and count as a voice in the output.

/** Waits `ms`, or until `alive()` says the reading was stopped. */
export const wait = (ms, alive) => new Promise((resolve) => {
  const started = performance.now();
  const tick = () => {
    if (!alive() || performance.now() - started >= ms) { resolve(); return; }
    setTimeout(tick, Math.min(100, Math.max(10, ms - (performance.now() - started))));
  };
  tick();
});

/**
 * A line's turn: played when it is to be read aloud and has a voice, else
 * waited out for its seconds. Answers whether its voice was played.
 * @param {object} p
 * @param {object} p.media   StudioMedia (speak)
 * @param {{ samples: Float32Array, rate: number, seconds: number }} p.sound
 * @param {boolean} p.play   whether the choice reads it aloud
 * @param {() => boolean} p.alive
 */
export async function sayLine({ media, sound, play, alive }) {
  if (play && sound?.samples?.length) {
    await media.speak(sound.samples, sound.rate);
    return true;
  }
  await wait((sound?.seconds || 0) * 1000, alive);
  return false;
}
