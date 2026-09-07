/**
 * Merged cells: creating and removing a region, not just reading one.
 *
 * The engine already READS and PAINTS a `<mergeCells>` block from a customer's
 * file — a merge is drawn once from its top-left cell, spanning the range. What
 * was missing was the design tool: a user could not make a merge or take one
 * apart. These tests pin the Excel semantics that make it feel right —
 *
 *   - only the top-left value survives a merge; the rest are cleared;
 *   - a single cell is not a region, so merging one does nothing;
 *   - a selection that would cut an existing merge in half is refused;
 *   - the whole thing is ONE undo step that puts back the merge map AND the
 *     cell values it cleared;
 *
 * — and, above all, that a merge survives the round trip through a real .xlsx.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const book = (rows = [['A', 'B', 'C'], ['x', 'y', 'z']], merges) =>
  buildXlsx({ sheets: [{ name: 'S', rows, ...(merges ? { merges } : {}) }] });

const open = (bytes) => new SheetView(bytes);

/** Select the rectangle from (r1,c1) to (r2,c2). */
const selectRange = (view, r1, c1, r2, c2) => {
  view.select(r1, c1);
  view.select(r2, c2, { extend: true });
};

// ── merging ──────────────────────────────────────────────────────────────────

test('merging a range makes one region and clears the non-anchor cells', () => {
  const view = open(book());
  selectRange(view, 0, 0, 0, 2); // A1:C1
  view.mergeSelection();

  const m = view.mergeAt(0, 0);
  assert.ok(m, 'a merge now exists at the anchor');
  assert.equal(m.ref, 'A1:C1');
  // Every covered cell resolves to the same region, not just the anchor.
  assert.equal(view.mergeAt(0, 2)?.ref, 'A1:C1');

  assert.equal(view.editValue(0, 0), 'A', 'the anchor value is kept');
  assert.equal(view.editValue(0, 1), '', 'the second cell was cleared');
  assert.equal(view.editValue(0, 2), '', 'the third cell was cleared');
});

test('the anchor value is preserved even when the anchor was empty of style', () => {
  const view = open(book([['keep', 'drop', 'drop']]));
  selectRange(view, 0, 0, 0, 2);
  view.mergeSelection();
  assert.equal(view.editValue(0, 0), 'keep');
});

test('a merge survives save and reopens as a real region', () => {
  const view = open(book());
  selectRange(view, 0, 0, 1, 1); // A1:B2
  view.mergeSelection();

  const reopened = open(view.save());
  const m = reopened.mergeAt(0, 0);
  assert.ok(m, 'the reopened file finds the merge');
  assert.equal(m.ref, 'A1:B2');
  assert.equal(reopened.editValue(0, 0), 'A', 'the anchor value round-tripped');
  assert.equal(reopened.editValue(0, 1), '', 'the cleared cell stayed cleared');
  assert.equal(reopened.editValue(1, 0), '', 'x was cleared as a non-anchor cell');
});

// ── unmerging ────────────────────────────────────────────────────────────────

test('unmerge removes a region the selection touches', () => {
  const view = open(book());
  selectRange(view, 0, 0, 0, 2);
  view.mergeSelection();
  assert.ok(view.mergeAt(0, 0));

  view.select(0, 0);
  view.unmergeSelection();
  assert.equal(view.mergeAt(0, 0), null, 'the region is gone');
});

test('an unmerge survives save', () => {
  const view = open(book([['a', 'b', 'c']], ['A1:C1']));
  assert.ok(view.mergeAt(0, 0), 'the built-in merge is read');
  view.select(0, 1); // any cell the region covers
  view.unmergeSelection();

  const reopened = open(view.save());
  assert.equal(reopened.mergeAt(0, 0), null, 'no merge after save+reopen');
});

// ── the refusals ─────────────────────────────────────────────────────────────

test('merging a single cell does nothing', () => {
  const view = open(book());
  view.select(1, 1);
  view.mergeSelection();
  assert.equal(view.mergeAt(1, 1), null, 'no region was created');
  assert.equal(view.canUndo, false, 'and nothing was recorded to undo');
});

test('a merge that would split an existing region is refused', () => {
  // B1:C1 is merged; A1:B1 overlaps B1 but not C1, so it would cut the region.
  const view = open(book([['a', 'b', 'c'], ['d', 'e', 'f']], ['B1:C1']));
  selectRange(view, 0, 0, 0, 1); // A1:B1
  assert.throws(() => view.mergeSelection(), /existing merged region/);
  // The refusal changed nothing.
  assert.equal(view.mergeAt(0, 1)?.ref, 'B1:C1', 'the original merge is intact');
  assert.equal(view.mergeAt(0, 0), null, 'and no new one was made');
});

test('a bigger merge absorbs a region wholly inside it', () => {
  const view = open(book([['a', 'b', 'c'], ['d', 'e', 'f']], ['B1:C1']));
  selectRange(view, 0, 0, 1, 2); // A1:C2 fully contains B1:C1
  view.mergeSelection();
  assert.equal(view.mergeAt(0, 0)?.ref, 'A1:C2');
  // The absorbed inner merge is gone, not left as a second region.
  const regions = view.merges.get('S');
  assert.equal(regions.length, 1, 'exactly one region remains');
});

// ── undo ─────────────────────────────────────────────────────────────────────

test('undo restores both the merge map and the cleared values', () => {
  const view = open(book());
  selectRange(view, 0, 0, 0, 2);
  view.mergeSelection();
  assert.ok(view.mergeAt(0, 0));
  assert.equal(view.editValue(0, 1), '');

  view.undo();
  assert.equal(view.mergeAt(0, 0), null, 'the merge is undone');
  assert.equal(view.editValue(0, 0), 'A', 'the anchor is still itself');
  assert.equal(view.editValue(0, 1), 'B', 'the cleared value came back');
  assert.equal(view.editValue(0, 2), 'C', 'and the third one too');
});

test('redo puts the merge and the clearing back', () => {
  const view = open(book());
  selectRange(view, 0, 0, 0, 2);
  view.mergeSelection();
  view.undo();
  view.redo();
  assert.equal(view.mergeAt(0, 0)?.ref, 'A1:C1', 'the merge is back');
  assert.equal(view.editValue(0, 1), '', 'and the clearing with it');
});

test('undo of an unmerge brings the region back', () => {
  const view = open(book([['a', 'b', 'c']], ['A1:C1']));
  view.select(0, 0);
  view.unmergeSelection();
  assert.equal(view.mergeAt(0, 0), null);
  view.undo();
  assert.equal(view.mergeAt(0, 0)?.ref, 'A1:C1', 'the region is restored');
});

// ── what the toolbar reads ───────────────────────────────────────────────────

test('render() reports whether the selection is merged', () => {
  const view = open(book());
  selectRange(view, 0, 0, 0, 2);
  view.mergeSelection();

  view.select(0, 0); // click the merged cell
  assert.equal(view.render().merged, true, 'a merged cell reads as merged');

  view.unmergeSelection();
  assert.equal(view.render().merged, false, 'and not once it is split');
});

test('a merge marks the document dirty even with nothing to clear', () => {
  const view = open(book([['only', '', '']]));
  selectRange(view, 0, 0, 0, 2); // nothing to clear past the anchor
  view.mergeSelection();
  assert.equal(view.isDirty, true, 'the mergeCells change alone is unsaved work');
  view.undo();
  assert.equal(view.isDirty, false, 'and undoing it clears the marker');
});

// ── fidelity: a file with no merge is not touched ────────────────────────────

test('a workbook that is never merged round-trips byte-identical', () => {
  const bytes = book();
  const view = open(bytes);
  const out = view.save();
  assert.deepEqual(Buffer.from(out), Buffer.from(bytes), 'no merge, no rewrite');
});
