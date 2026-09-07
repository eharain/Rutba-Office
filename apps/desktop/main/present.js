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

export function createPresentService({ broadcast }) {
  /** @type {{ id: string|null, index: number, running: boolean, startedAt: number|null, blank: boolean }} */
  let state = { id: null, index: 0, running: false, startedAt: null, blank: false };

  return {
    state: () => state,

    /**
     * Move the show. Anything not named is left alone, so a presenter window
     * can advance the slide without knowing whether the clock is running.
     */
    set: (patch = {}) => {
      const next = { ...state, ...patch };

      // Starting the show starts the clock, once. Restarting it from the
      // beginning is a new show and a new clock; advancing is not.
      if (patch.running === true && !state.running) next.startedAt = Date.now();
      if (patch.running === false) next.startedAt = null;
      if (patch.restart) {
        next.startedAt = Date.now();
        next.index = 0;
        delete next.restart;
      }

      state = next;
      broadcast?.('present:state', state);
      return state;
    },
  };
}
