/**
 * Tracked changes one at a time (2.18.0, A3.6).
 *
 * In 2.17.1 Next and Previous went to a paragraph and put the caret at its
 * start, and Accept or Reject This Change took every change in that
 * paragraph. Now Next and Previous select each change in turn, This Change
 * acts on the selected change only, and Accept All and Reject All work with
 * the caret anywhere.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';

/** Three tracked changes in the first line and one in the third, by Kim. */
function reviewed() {
  const view = openDocx(buildDocx({ paragraphs: ['Open the shop at nine', 'Count the float', 'Lock the back door'] }));
  view.setTrackChanges(true, 'Kim');
  view.collapseTo({ block: 0, offset: 21 });
  view.insertText(' sharp');
  view.collapseTo({ block: 0, offset: 0 });
  view.insertText('Please ');
  view.setSelection({ block: 0, offset: 16 }, { block: 0, offset: 21 });
  view.deleteSelection();
  view.setSelection({ block: 2, offset: 9 }, { block: 2, offset: 14 });
  view.deleteSelection();
  return view;
}

test('the changes are listed in reading order, each with its words and author', () => {
  const view = reviewed();
  const list = view.trackedChangeList();
  assert.deepEqual(list.map((c) => [c.kind, c.block, c.from, c.to]), [
    ['insert', 0, 0, 7],
    ['delete', 0, 16, 16],
    ['insert', 0, 23, 29],
    ['delete', 2, 9, 9],
  ]);
  assert.deepEqual(list.map((c) => c.text ?? null), [null, 'shop ', null, 'back ']);
  assert.ok(list.every((c) => c.author === 'Kim'));
  assert.equal(new Set(list.map((c) => c.id)).size, 4, 'each its own id');
});

test('Next selects each change in turn, round to the first; Previous goes back', () => {
  const view = reviewed();
  view.collapseTo({ block: 1, offset: 0 });
  const seen = [];
  for (let i = 0; i < 5; i++) {
    const c = view.nextChange(1);
    seen.push([c.block, c.from, c.to]);
    assert.deepEqual([view.selection.from, view.selection.to], [{ block: c.block, offset: c.from }, { block: c.block, offset: c.to }], 'the change is selected');
  }
  assert.deepEqual(seen, [[2, 9, 9], [0, 0, 7], [0, 16, 16], [0, 23, 29], [2, 9, 9]]);
  const back = view.nextChange(-1);
  assert.deepEqual([back.block, back.from, back.to], [0, 23, 29]);
  view.nextChange(-1);
  const again = view.nextChange(-1);
  assert.deepEqual([again.block, again.from], [0, 0]);
});

test('Accept This Change keeps the selected change only', () => {
  const view = reviewed();
  view.collapseTo({ block: 0, offset: 0 });
  view.nextChange(1);
  view.nextChange(1);
  view.nextChange(1);
  assert.equal(view.copyText(), ' sharp');
  view.acceptChanges();
  assert.deepEqual(view.trackedChangeList().map((c) => c.kind + '@' + c.block), ['insert@0', 'delete@0', 'delete@2'], 'the others still waiting');
  assert.equal(view.block(0).text, 'Please Open the at nine sharp');
  assert.ok(!view.block(0).runs.find((r) => r.text.includes('sharp')).ins, 'its words now plain');
});

test('Reject This Change puts back only the selected deletion', () => {
  const view = reviewed();
  view.collapseTo({ block: 0, offset: 0 });
  view.nextChange(1);
  view.nextChange(1);
  view.rejectChanges();
  assert.equal(view.block(0).text, 'Please Open the shop at nine sharp');
  assert.equal(view.trackedChangeList().length, 3);
  assert.ok(view.trackedChangeList().some((c) => c.kind === 'delete' && c.block === 2), 'the deletion in the third line still there');
  view.undo();
  assert.equal(view.trackedChangeList().length, 4, 'one undo puts it back');
});

test('Accept All and Reject All work with the caret in a line with no change', () => {
  const accepted = reviewed();
  accepted.collapseTo({ block: 1, offset: 3 });
  accepted.acceptChanges({ all: true });
  assert.equal(accepted.trackedChangeList().length, 0);
  assert.deepEqual(accepted.blocks.map((b) => b.text), ['Please Open the at nine sharp', 'Count the float', 'Lock the door']);
  const rejected = reviewed();
  rejected.collapseTo({ block: 1, offset: 3 });
  rejected.rejectChanges({ all: true });
  assert.deepEqual(rejected.blocks.map((b) => b.text), ['Open the shop at nine', 'Count the float', 'Lock the back door']);
});

test('a change of formatting is a change of its own', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Count the float'] }));
  view.setTrackChanges(true, 'Kim');
  view.setSelection({ block: 0, offset: 10 }, { block: 0, offset: 15 });
  view.toggleFormat('b');
  view.collapseTo({ block: 0, offset: 0 });
  view.insertText('Now ');
  assert.deepEqual(view.trackedChangeList().map((c) => [c.kind, c.from, c.to]), [['insert', 0, 4], ['format', 14, 19]]);
  view.collapseTo({ block: 0, offset: 0 });
  view.nextChange(1);
  view.nextChange(1);
  view.rejectChanges();
  assert.deepEqual(view.trackedChangeList().map((c) => c.kind), ['insert'], 'the bold taken back, the insertion kept waiting');
  assert.ok(!view.block(0).runs.some((r) => r.bold));
});
