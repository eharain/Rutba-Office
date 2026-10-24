// Where the show is up to.
//
// A presenter view is two windows: the audience sees the slide, the speaker
// sees the notes, the next slide and the clock. They have to agree about which
// slide is showing, and two renderers cannot tell each other anything — the
// only thing they share is this process.
//
// So the position lives here, and every change is broadcast. It is deliberately
// the smallest possible thing: one integer, one boolean and a timestamp. Making
// it any cleverer would mean two copies of the deck's state, which is exactly
// the bug this exists to avoid.
//
// The presenter's clock is its own: `pause` holds it and lets it go on, and
// `resetClock` puts it back to nought; neither touches the show or its slide.
// `running` is the show itself, which ends when it is false.

export function createPresentService({ broadcast, now = Date.now }) {
  /** @type {{ id: string|null, index: number, running: boolean, startedAt: number|null, pausedAt: number|null, blank: boolean }} */
  let state = { id: null, index: 0, running: false, startedAt: null, pausedAt: null, blank: false };

  return {
    state: () => state,

    /**
     * Move the show. Anything not named is left alone, so a presenter window
     * can advance the slide without knowing whether the clock is running.
     */
    set: (patch = {}) => {
      const { pause, resetClock, ...rest } = patch;
      const next = { ...state, ...rest };
      const at = now();

      // Starting the show starts the clock, once. Restarting it from the
      // beginning is a new show and a new clock; advancing is not.
      if (rest.running === true && !state.running) { next.startedAt = at; next.pausedAt = null; }
      if (rest.running === false) { next.startedAt = null; next.pausedAt = null; }
      if (rest.restart) {
        next.startedAt = at;
        next.pausedAt = null;
        next.index = 0;
        delete next.restart;
      }
      // Pause holds the clock where it is, and the show goes on; going on
      // again takes up the time from there.
      if (pause === true && next.startedAt != null && next.pausedAt == null) next.pausedAt = at;
      if (pause === false && next.pausedAt != null) {
        next.startedAt += at - next.pausedAt;
        next.pausedAt = null;
      }
      // Reset: the clock back to nought, held or going as it was. The slide stays.
      if (resetClock && next.startedAt != null) {
        next.startedAt = at;
        if (next.pausedAt != null) next.pausedAt = at;
      }

      state = next;
      broadcast?.('present:state', state);
      return state;
    },
  };
}
