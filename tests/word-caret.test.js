// Typing faster than the engine answers must not reorder the letters.
//
// The document editor cancels every keystroke in the browser and sends it to
// the engine with the caret read from the DOM. Until the engine's answer is
// painted, that caret is stale — the browser never made the insertion — so a
// keystroke sent with it lands before the previous one. "Typed." typed at
// sixty milliseconds a letter on a loaded machine came back "ye.dpT".
//
// The rule in word/caret.js: a caret exactly where the editor last sent or
// placed it is not resent; the engine's own caret is ahead of it and right.

import test from 'node:test';
import assert from 'node:assert/strict';
import { selectionToSend, samePosition } from '../apps/desktop/renderer/apps/word/caret.js';

/** The smallest engine: one paragraph, one caret, the two operations typing uses. */
function engine(text, caret) {
  const state = { text, caret };
  return {
    state,
    apply(ops) {
      for (const op of ops) {
        if (op.op === 'setSelection') state.caret = op.focus.offset;
        else if (op.op === 'insertText') {
          state.text = state.text.slice(0, state.caret) + op.text + state.text.slice(state.caret);
          state.caret += op.text.length;
        }
      }
      return { block: 0, offset: state.caret };
    },
  };
}

const at = (offset) => ({ focus: { block: 0, offset } });

/** Types `letters` with the DOM caret frozen at `start` — nothing painted until the end. */
function typeWithoutPaint(letters, start, { withRule }) {
  const eng = engine('Quarterly Review', start);
  let sent = null;
  for (const ch of letters) {
    const pos = at(start);
    const selection = withRule ? selectionToSend(pos, { sent, placed: null }) : { op: 'setSelection', anchor: pos.focus, focus: pos.focus };
    eng.apply([...(selection ? [selection] : []), { op: 'insertText', text: ch }]);
    sent = pos;
  }
  return eng.state.text;
}

test('letters typed before the first answer is painted stay in order', () => {
  assert.equal(typeWithoutPaint(' Typed.', 16, { withRule: true }), 'Quarterly Review Typed.');
});

test('without the rule the same keystrokes come back reversed — the defect this pins', () => {
  assert.equal(typeWithoutPaint(' Typed.', 16, { withRule: false }), 'Quarterly Review.depyT ');
});

test('a paint in the middle hands the caret back to the engine without a jump', () => {
  const eng = engine('ab', 2);
  let sent = null;
  let placed = null;
  const key = (pos, ch) => {
    const selection = selectionToSend(pos, { sent, placed });
    const answer = eng.apply([...(selection ? [selection] : []), { op: 'insertText', text: ch }]);
    sent = pos;
    return answer;
  };
  const first = key(at(2), 'c');
  key(at(2), 'd'); // still unpainted: the DOM caret is where 'c' began
  // The first answer is painted: the caret goes where the engine said after 'c'.
  placed = { focus: first };
  sent = null;
  key(at(first.offset), 'e'); // the DOM caret is exactly where it was placed
  assert.equal(eng.state.text, 'abcde');
});

test('a caret the reader moved is always sent', () => {
  const placed = { focus: { block: 0, offset: 5 } };
  assert.deepEqual(selectionToSend(at(2), { sent: null, placed }), { op: 'setSelection', anchor: { block: 0, offset: 2 }, focus: { block: 0, offset: 2 } });
  // Clicking back onto the very spot the last edit was sent from, after it painted, is a move too.
  assert.ok(selectionToSend(at(5), { sent: at(5), placed: { focus: { block: 0, offset: 6 } } }) === null, 'unpainted: the engine is ahead');
  assert.ok(selectionToSend(at(5), { sent: null, placed: { focus: { block: 0, offset: 6 } } }) !== null, 'painted: the click is real');
});

test('a range is compared by both ends', () => {
  const range = { anchor: { block: 0, offset: 1 }, focus: { block: 0, offset: 4 } };
  assert.ok(samePosition(range, { anchor: { block: 0, offset: 1 }, focus: { block: 0, offset: 4 } }));
  assert.ok(!samePosition(range, { focus: { block: 0, offset: 4 } }));
  // Replacing a selection: the first keystroke sends the range, the next — the browser left the range standing — does not.
  assert.equal(selectionToSend(range, { sent: range }), null);
});
