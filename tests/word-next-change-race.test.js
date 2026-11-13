// Review → Next pressed again before the change it selected is painted.
//
// Next sends the page's caret ahead of the engine's nextChange, so a caret
// the reader moved is where the search starts. Between the engine's answer
// and the page painting it, the page still shows the caret from before, and
// a second Next sent with that caret starts the search from there again: it
// finds the first change a second time and Accept This Change then takes the
// wrong one (07.0250 and 07.0251, 2.26.0 record 1, under load only).
//
// The fix is the rule typing already follows (word/caret.js): a caret exactly
// where this editor last sent or placed it is not resent. Next, Tab and the
// equation paste now go through it (caretAhead in word.js).

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { selectionToSend } from '../apps/desktop/renderer/apps/word/caret.js';

const here = path.dirname(fileURLToPath(import.meta.url));

/** The review check's document: "back " and "shop " taken out, " sharp" typed. */
function reviewed() {
  const view = openDocx(buildDocx({ paragraphs: ['Count the till before you lock the door.', 'Open the shop at nine', 'Lock the back door'] }));
  view.setTrackChanges(true, 'Kim');
  view.setSelection({ block: 2, offset: 9 }, { block: 2, offset: 14 });
  view.deleteSelection();
  view.collapseTo({ block: 1, offset: 21 });
  view.insertText(' sharp');
  view.setSelection({ block: 1, offset: 9 }, { block: 1, offset: 14 });
  view.deleteSelection();
  view.collapseTo({ block: 0, offset: 0 });
  return view;
}

/** The ops a page sends, applied as the main process does. */
function apply(view, ops) {
  for (const op of ops) {
    if (op.op === 'setSelection') view.setSelection(op.anchor, op.focus);
    else if (op.op === 'nextChange') view.nextChange(op.direction);
  }
  const { from, to } = view.selection;
  return [from.block, from.offset, to.offset];
}

const start = { anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } };

/** Next pressed `times` times with nothing painted between: the page caret stays at the start. */
function pressNextUnpainted(times, { withRule }) {
  const view = reviewed();
  let sent = null;
  const placed = null;
  const seen = [];
  for (let i = 0; i < times; i++) {
    const pos = start;
    const ahead = withRule ? selectionToSend(pos, { sent, placed }) : { op: 'setSelection', anchor: pos.anchor, focus: pos.focus };
    if (withRule) sent = pos;
    seen.push(apply(view, [...(ahead ? [ahead] : []), { op: 'nextChange', direction: 1 }]));
  }
  return { seen, view };
}

test('the changes in the check document, in reading order', () => {
  const list = reviewed().trackedChangeList();
  assert.deepEqual(list.map((c) => [c.kind, c.block, c.from, c.to]), [
    ['delete', 1, 9, 9],
    ['insert', 1, 16, 22],
    ['delete', 2, 9, 9],
  ]);
});

test('Next pressed twice before the first answer is painted selects each change in turn', () => {
  const { seen } = pressNextUnpainted(2, { withRule: true });
  assert.deepEqual(seen, [[1, 9, 9], [1, 16, 22]]);
});

test('without the rule the second Next finds the first change again: the record\'s failure', () => {
  const { seen } = pressNextUnpainted(2, { withRule: false });
  assert.deepEqual(seen, [[1, 9, 9], [1, 9, 9]]);
});

test('Accept This Change after the two presses keeps the insertion and leaves both deletions', () => {
  const { view } = pressNextUnpainted(2, { withRule: true });
  view.acceptChanges();
  const list = view.trackedChangeList();
  assert.deepEqual(list.map((c) => [c.kind, c.block]), [['delete', 1], ['delete', 2]]);
  assert.equal(view.blocks[1].text, 'Open the at nine sharp');
});

test('a Next after the paint, with the caret where the page put it, goes on to the next change', () => {
  const view = reviewed();
  const first = apply(view, [selectionToSend(start, { sent: null, placed: null }), { op: 'nextChange', direction: 1 }]);
  // Painted: the page caret is where the engine said, and nothing is sent.
  const placed = { anchor: { block: 1, offset: 9 }, focus: { block: 1, offset: 9 } };
  assert.equal(selectionToSend(placed, { sent: null, placed }), null);
  const second = apply(view, [{ op: 'nextChange', direction: 1 }]);
  assert.deepEqual([first, second], [[1, 9, 9], [1, 16, 22]]);
});

test('a caret the reader moved before Next is sent, and the search starts there', () => {
  const view = reviewed();
  apply(view, [{ op: 'nextChange', direction: 1 }]);
  const placed = { anchor: { block: 1, offset: 9 }, focus: { block: 1, offset: 9 } };
  const clicked = { anchor: { block: 2, offset: 0 }, focus: { block: 2, offset: 0 } };
  const ahead = selectionToSend(clicked, { sent: null, placed });
  assert.ok(ahead, 'a click elsewhere is sent');
  assert.deepEqual(apply(view, [ahead, { op: 'nextChange', direction: 1 }]), [2, 9, 9]);
});

test('Next, Tab and the equation paste send the caret by the rule, not straight from the page', () => {
  const src = fs.readFileSync(path.join(here, '../apps/desktop/renderer/apps/word.js'), 'utf8');
  assert.match(src, /apply\(\.\.\.caretAhead\(\), \{ op: 'nextChange'/);
  assert.match(src, /apply\(\.\.\.caretAhead\(\), \{ op: 'tabCell'/);
  assert.match(src, /apply\(\.\.\.caretAhead\(\), \{ op: 'pasteRuns'/);
  // The rule's bookkeeping: what it sends is remembered until the answer is painted.
  assert.match(src, /const caretAhead = useCallback\(\(\) => \{[\s\S]{0,400}selectionToSend\(pos, \{ sent: sentCaret\.current, placed: placedCaret\.current \}\)[\s\S]{0,120}sentCaret\.current = pos;/);
});
