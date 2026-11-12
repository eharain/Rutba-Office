// When the studio's window holds a close for its own question (Stop
// recording?, End broadcast?) rather than letting it go.
//
// A close the window asked for itself, once the question is answered or the
// takes were handed over, must not be held: the page's phase is read from
// the last render, and at a hand-over that is still "stopping", so the
// window's own close could be cancelled by its own guard, leaving the
// Recorder open over a recording Video already had (2.25.0 gate: the takes
// landed, the window did not close).
//
// Kept apart from the window so it can be tested without one.

/** Whether closing now is held to ask first: while live or recording, unless this window is closing itself. */
export const closeHeld = ({ phase, live = false, closing = false } = {}) => !closing && (Boolean(live) || phase === 'recording' || phase === 'stopping');

/**
 * Whether a "took" that came after the Recorder gave up waiting may still
 * close it: only while nothing runs. By then the page has rendered, so its
 * phase is current, and a Recorder recording again (or counting down to it,
 * or stopping) is not closed under the person.
 */
export const lateCloseAllowed = (phase) => phase === 'setup';
