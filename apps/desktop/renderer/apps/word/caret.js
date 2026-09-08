// The caret between keystrokes.
//
// Every edit is cancelled in the browser and sent to the engine with the caret
// read from the DOM. Between the sending and the engine's answer the DOM caret
// is stale: the insertion it would have made never happened, so it still sits
// where the last keystroke began. A second keystroke sent with that caret
// lands *before* the first — "Typed." arrives as "ye.dpT" at a fast pace, on
// a slow round trip, or in a long document that takes a while to paint. The
// owner saw it as "the editor does not respond properly".
//
// The rule: when the caret is exactly where this editor last left it — sent
// with an edit whose answer has not been painted yet, or placed from the
// engine's answer — the reader has not moved it, and the engine's own caret,
// already past whatever is in flight, is the right one; no position is sent.
// A caret anywhere else was moved by a click or a key, and is sent.

export function samePoint(a, b) {
  return Boolean(a && b) && a.block === b.block && a.offset === b.offset;
}

export function samePosition(a, b) {
  if (!a?.focus || !b?.focus) return false;
  return samePoint(a.anchor || a.focus, b.anchor || b.focus) && samePoint(a.focus, b.focus);
}

/**
 * The selection operation to send ahead of an edit, or null when the engine's
 * caret already says it.
 *
 * `sent` is the position sent with the last edit, cleared once that edit's
 * answer has been painted; `placed` is the position the page last put the
 * caret at, from the engine.
 */
export function selectionToSend(pos, { sent = null, placed = null } = {}) {
  if (!pos?.focus) return null;
  if (samePosition(pos, sent) || samePosition(pos, placed)) return null;
  return { op: 'setSelection', anchor: pos.anchor || pos.focus, focus: pos.focus };
}
