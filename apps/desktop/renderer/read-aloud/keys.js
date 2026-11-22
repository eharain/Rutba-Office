// The keyboard while Documents reads aloud. The sentence being read is the
// selection on the page, so a key that types would put itself in the
// sentence's place, and one press of Space, meant as a pause, would take a
// sentence out of the document. While the selection is the reader's own the
// words stay as they are: Space pauses and reads on, and anything else that
// would change them is told to wait. Once the person has put the caret
// somewhere themselves (a click, an arrow key), the next edit is theirs: the
// reading stops and the edit goes where they put it.

/** The keys that move the caret: after one, the selection is the person's, not the reader's. */
const MOVES = new Set(['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End', 'PageUp', 'PageDown']);

/** Whether a key moves the caret. */
export function movesCaret(key) {
  return MOVES.has(key);
}

/**
 * What an edit asked of the page does while its words are read aloud.
 *
 * @param {object} o
 * @param {boolean} o.reading    Read Aloud is on (reading or paused)
 * @param {boolean} o.own        the selection is the one the reader put there
 * @param {string} [o.inputType] the `beforeinput` kind
 * @param {string|null} [o.data] what it would type
 * @returns {'edit'|'pause'|'keep'|'stop'} `edit`: as ever. `pause`: pause, or read on. `keep`: the words
 *   stay as they are. `stop`: the reading stops, and the edit is made.
 */
export function inputWhileReading({ reading = false, own = false, inputType = '', data = null } = {}) {
  if (!reading) return 'edit';
  if (!own) return 'stop';
  if (inputType === 'insertText' && data === ' ') return 'pause';
  return 'keep';
}
