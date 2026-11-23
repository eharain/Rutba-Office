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
// A new session is written nowhere (D-2026-11-22-11), so one the person has
// arranged is asked about before it is left: at the window's close, and
// before a saved setup or a setup file takes its place.
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

/**
 * Whether leaving the session now would lose what was arranged in it: a new
 * session the person has changed. A saved setup is written as it changes, the
 * tools' own is theirs, and a new session nobody touched holds nothing to keep.
 */
export const sessionUnkept = ({ setup = null, changed = false } = {}) => Boolean(setup && setup.clean && !setup.tools && changed);

/**
 * Whether a change to the stack is the person arranging: one Undo would take
 * back (`record`), that left the stack different. Picking a scene, a sound
 * turned down and a source found again are not.
 */
export const arranged = ({ record = true, before, next } = {}) => Boolean(record) && JSON.stringify(before) !== JSON.stringify(next);

/**
 * What the Setups menu offers for making a setup: in a new session, saving it
 * under a name ('keep'); with a saved setup open, a new one beside it ('new').
 * Never both: in a new session the two read alike and do opposite things,
 * one keeping what is arranged and the other replacing it.
 */
export const setupChoices = (saved) => (saved ? ['new'] : ['keep']);

/**
 * Whether a start says where the saved setups are: a new session nobody has
 * changed, with setups saved, while nothing runs and the window is the
 * person's own. `shown` is the window's own switch: on at its start, off once
 * a setup was opened, the session reset, or the line closed.
 */
export const offersSaved = ({ shown = false, setup = null, changed = false, count = 0, busy = false } = {}) => Boolean(shown) && Boolean(setup && setup.clean && !setup.tools) && !changed && count > 0 && !busy;
