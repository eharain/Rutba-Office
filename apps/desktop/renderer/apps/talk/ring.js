// The sounds Talk makes (Office 2.27.0, the owner's yes to Talk rings of
// 22 Nov 2026): a ring while a call comes in, a soft tone while one's own
// call rings at the other side, and one short tone for being asked into a
// meeting. Made here from plain notes, so Office carries no sound file.
//
// Talk shows no Windows notification and flashes nothing itself; the
// notifier in the main process does that, from what main/talk-notify.js
// says. A sound is the Talk window's own, and the person can turn it off
// (Settings, "Play a sound for calls").

/**
 * Each sound: its notes as [starts at, lasts, pitch in Hz] in seconds from
 * the start of a phrase, how loud at most (0 to 1), and how often the phrase
 * comes again in milliseconds (0: once).
 */
export const SOUNDS = {
  // Two rising notes, twice, then quiet.
  ring: { every: 3000, level: 0.22, notes: [[0, 0.18, 784], [0.22, 0.3, 988], [0.7, 0.18, 784], [0.92, 0.3, 988]] },
  // One low note, as a telephone says "it is ringing there".
  calling: { every: 4000, level: 0.07, notes: [[0, 1, 440]] },
  // Three notes going up, once.
  chime: { every: 0, level: 0.16, notes: [[0, 0.16, 659], [0.18, 0.16, 784], [0.36, 0.32, 988]] },
};

/**
 * Which sound a Talk window should be making now, or null.
 *
 * @param {object} o
 * @param {boolean} o.on whether the person wants sounds (Talk's state `ring`)
 * @param {boolean} [o.silent] a build that never makes a sound (the window checks)
 * @param {boolean} o.incoming a call is coming in and its card is showing
 * @param {string|null} [o.call] the state of the call this window is in ('ringing' while it rings at the other side)
 */
export function soundFor({ on, silent = false, incoming, call = null }) {
  if (!on || silent) return null;
  if (incoming) return 'ring';
  if (call === 'ringing') return 'calling';
  return null;
}

/**
 * @param {object} [o]
 * @param {() => AudioContext} [o.makeContext] made at the first sound, never before
 * @param {Function} [o.setTimer] setInterval, or a test's own
 * @param {Function} [o.clearTimer]
 */
export function createRinger({ makeContext = () => new (globalThis.AudioContext || globalThis.webkitAudioContext)(), setTimer = (fn, ms) => setInterval(fn, ms), clearTimer = (id) => clearInterval(id) } = {}) {
  let ctx = null;
  let timer = null;
  let playing = null;
  const sounding = new Set();

  const context = () => {
    if (!ctx) ctx = makeContext();
    // A window that has not been clicked in yet may hold its sound back until asked.
    try { const going = ctx.resume?.(); going?.catch?.(() => {}); } catch { /* it plays when it can */ }
    return ctx;
  };

  /** One phrase of a sound, from now. */
  const phrase = (sound) => {
    const c = context();
    const from = c.currentTime;
    for (const [at, lasts, hz] of sound.notes) {
      const osc = c.createOscillator();
      const gain = c.createGain();
      osc.type = 'sine';
      osc.frequency.value = hz;
      // Up and down again over a moment, so a note neither clicks in nor out.
      gain.gain.setValueAtTime(0, from + at);
      gain.gain.linearRampToValueAtTime(sound.level, from + at + 0.02);
      gain.gain.setValueAtTime(sound.level, from + at + Math.max(0.02, lasts - 0.05));
      gain.gain.linearRampToValueAtTime(0, from + at + lasts);
      osc.connect(gain);
      gain.connect(c.destination);
      osc.onended = () => sounding.delete(osc);
      sounding.add(osc);
      osc.start(from + at);
      osc.stop(from + at + lasts + 0.02);
    }
  };

  const hush = () => {
    for (const osc of [...sounding]) { try { osc.stop(); } catch { /* it had not started, or has ended */ } }
    sounding.clear();
  };

  const stop = () => {
    if (timer != null) { clearTimer(timer); timer = null; }
    playing = null;
    hush();
  };

  return {
    /**
     * A sound that comes again until stopped ('ring', 'calling'). Asked for
     * again while it plays, it goes on as it was. False when it could not be
     * made: a computer with no sound, which is no fault to show anyone.
     */
    play(name) {
      const sound = SOUNDS[name];
      if (!sound || !sound.every) return false;
      if (playing === name) return true;
      stop();
      try { phrase(sound); } catch { return false; }
      playing = name;
      timer = setTimer(() => { try { phrase(sound); } catch { stop(); } }, sound.every);
      return true;
    },

    /** A sound made once ('chime'), over whatever else is playing. */
    once(name) {
      const sound = SOUNDS[name];
      if (!sound) return false;
      try { phrase(sound); } catch { return false; }
      return true;
    },

    stop,
    playing: () => playing,

    /** The window is closing: nothing sounds on, and the sound device is let go. */
    close() {
      stop();
      try { const going = ctx?.close?.(); going?.catch?.(() => {}); } catch { /* gone already */ }
      ctx = null;
    },
  };
}
