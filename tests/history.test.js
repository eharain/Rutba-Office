/**
 * Undo — the granularity, not the mechanism.
 *
 * Storing and restoring a snapshot is the easy half. What makes undo usable is
 * where the steps FALL: one per word, not one per character; one per paste,
 * however many paragraphs it spans; a caret move ending the run. Those rules are
 * what these tests pin down, because they are the ones a refactor would quietly
 * break without any test failing.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { History, DEFAULT_COALESCE_MS } from '@rutba/editing';
import { SheetView } from '@rutba/sheet-view';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { comparePackages } from '@rutba/ooxml';
import { buildComplexWorkbook } from './fixtures/complex-workbook.js';
import { buildComplexDocument } from './fixtures/complex-document.js';

// Unprotected: these tests are about UNDO, and thawing the sheet first would
// put a history entry and a dirty flag between every test and its subject.
const WORKBOOK = buildComplexWorkbook({ protect: false });
const LETTER = buildComplexDocument();

/** A clock we drive by hand, so coalescing is tested rather than timed. */
const clock = () => {
  let t = 1_000_000;
  return { now: () => t, advance: (ms) => { t += ms; } };
};

// ------------------------------------------------------------- the history --

test('same-group edits inside the window are one step', () => {
  const c = clock();
  const h = new History({ now: c.now });

  assert.equal(h.record({ state: 'a', group: 'type:0' }), true, 'first is a new step');
  c.advance(50);
  assert.equal(h.record({ state: 'b', group: 'type:0' }), false, 'merged');
  c.advance(50);
  assert.equal(h.record({ state: 'c', group: 'type:0' }), false, 'merged');
  assert.equal(h.depth, 1);

  // and the step still holds the state from BEFORE the whole run
  assert.equal(h.undo('current').state, 'a');
});

test('a pause ends the run even when nothing else changed', () => {
  const c = clock();
  const h = new History({ now: c.now });
  h.record({ state: 'a', group: 'type:0' });
  c.advance(DEFAULT_COALESCE_MS + 1);
  h.record({ state: 'b', group: 'type:0' });
  assert.equal(h.depth, 2, 'walking away mid-sentence starts a new step');
});

test('a different group, or none at all, always starts a new step', () => {
  const c = clock();
  const h = new History({ now: c.now });
  h.record({ state: 'a', group: 'type:0' });
  h.record({ state: 'b', group: 'type:1' });
  h.record({ state: 'c', group: null });
  h.record({ state: 'd', group: null });
  assert.equal(h.depth, 4);
});

test('break() ends the run, which is what a caret move does', () => {
  const c = clock();
  const h = new History({ now: c.now });
  h.record({ state: 'a', group: 'type:0' });
  h.break();
  h.record({ state: 'b', group: 'type:0' });
  assert.equal(h.depth, 2);
});

test('a new edit discards the redo branch', () => {
  const h = new History();
  h.record({ state: 'a' });
  h.record({ state: 'b' });
  h.undo('c');
  assert.equal(h.canRedo, true);
  h.record({ state: 'd' });
  assert.equal(h.canRedo, false, 'the branch you did not take is gone');
});

test('the stack is bounded, oldest first', () => {
  const h = new History({ limit: 3 });
  for (const state of ['a', 'b', 'c', 'd', 'e']) h.record({ state });
  assert.equal(h.depth, 3);
  assert.equal(h.undo('now').state, 'e');
  assert.equal(h.undo('now').state, 'd');
  assert.equal(h.undo('now').state, 'c');
  assert.equal(h.undo('now'), null, 'a and b fell off the bottom');
});

test('undo and redo round-trip through the same states', () => {
  const h = new History();
  h.record({ state: 's0', label: 'typing' });
  h.record({ state: 's1', label: 'paste' });

  assert.equal(h.undoLabel, 'paste');
  assert.equal(h.undo('s2').state, 's1');
  assert.equal(h.undo('s1').state, 's0');
  assert.equal(h.canUndo, false);

  assert.equal(h.redo('s0').state, 's1');
  assert.equal(h.redo('s1').state, 's2');
  assert.equal(h.canRedo, false);
});

// --------------------------------------------------------------- documents --

test('a run of keystrokes is one undo, and it puts the caret back', () => {
  const view = openDocx(LETTER);
  const target = view.render().blocks.find((b) => !b.structural && b.text.length > 5).index;
  const before = view.render().blocks[target].text;

  view.setSelection({ block: target, offset: 3 });
  for (const ch of 'Hello') view.insertText(ch);
  assert.equal(view.history.depth, 1, 'five keystrokes, one undo');
  assert.notEqual(view.render().blocks[target].text, before);

  assert.equal(view.undo(), true);
  assert.equal(view.render().blocks[target].text, before);
  assert.deepEqual(view.render().selection.focus, { block: target, offset: 3 },
    'undo that leaves the caret elsewhere makes the user hunt for the change');

  assert.equal(view.redo(), true);
  assert.ok(view.render().blocks[target].text.includes('Hello'));
});

test('moving the caret between runs makes two undos', () => {
  const view = openDocx(LETTER);
  const target = view.render().blocks.find((b) => !b.structural && b.text.length > 5).index;

  view.setSelection({ block: target, offset: 0 });
  view.insertText('A');
  view.setSelection({ block: target, offset: 4 });
  view.insertText('B');
  assert.equal(view.history.depth, 2);
});

test('a multi-paragraph paste is a single undo', () => {
  const view = openDocx(LETTER);
  const target = view.render().blocks.find((b) => !b.structural && b.text.length > 5).index;
  const before = view.render().blocks.length;

  view.setSelection({ block: target, offset: 0 });
  view.pasteText('one\ntwo\nthree');
  assert.equal(view.history.depth, 1);
  assert.equal(view.render().blocks.length, before + 2);

  view.undo();
  assert.equal(view.render().blocks.length, before, 'all three paragraphs came back out');
});

test('typing over a selection is one undo, not two', () => {
  // insertText delegates to deleteSelection; without the depth counter the
  // first undo would restore a state the user never saw.
  const view = openDocx(LETTER);
  const target = view.render().blocks.find((b) => !b.structural && b.text.length > 8).index;
  const before = view.render().blocks[target].text;

  view.setSelection({ block: target, offset: 0 }, { block: target, offset: 5 });
  view.insertText('X');
  assert.equal(view.history.depth, 1);
  view.undo();
  assert.equal(view.render().blocks[target].text, before);
});

test('typing then undoing leaves the file byte-identical', () => {
  // The whole fidelity claim, applied to undo: an edit that was taken back must
  // not leave a rewritten part behind.
  const view = openDocx(LETTER);
  const target = view.render().blocks.find((b) => !b.structural && b.text.length > 5).index;
  view.setSelection({ block: target, offset: 1 });
  view.insertText('scratch');
  view.undo();

  const diff = comparePackages(LETTER, view.save());
  assert.deepEqual(diff.changed, []);
});

// ------------------------------------------------------------------ sheets --

test('a cell edit undoes, and takes its dependent formulas with it', () => {
  const view = SheetView.open(WORKBOOK, { viewportWidth: 900, viewportHeight: 500 });
  const before = view.displayValue(1, 1).text;
  const dependent = view.displayValue(1, 2).text;   // =B2*1.17
  assert.ok(view.editValue(1, 2).startsWith('='), 'the fixture has a dependent formula');

  view.select(1, 1);
  view.setCell(1, 1, '1');
  assert.notEqual(view.displayValue(1, 2).text, dependent, 'the formula reacted');

  view.undo();
  assert.equal(view.displayValue(1, 1).text, before);
  assert.equal(view.displayValue(1, 2).text, dependent, 'and reacted back');
});

test('undoing the only edit makes the document clean again', () => {
  const view = SheetView.open(WORKBOOK, { viewportWidth: 900, viewportHeight: 500 });
  assert.equal(view.isDirty, false);
  view.select(1, 1);
  view.setCell(1, 1, '5');
  assert.equal(view.isDirty, true);
  view.undo();
  assert.equal(view.isDirty, false, 'the Save button must stop offering to save nothing');
  view.redo();
  assert.equal(view.isDirty, true);
});

test('retyping one cell is one undo; moving on starts another', () => {
  const view = SheetView.open(WORKBOOK, { viewportWidth: 900, viewportHeight: 500 });
  view.select(20, 1);
  view.setCell(20, 1, '1');
  view.setCell(20, 1, '12');
  view.setCell(20, 1, '123');
  assert.equal(view.history.depth, 1);

  view.select(21, 1);
  view.setCell(21, 1, 'x');
  assert.equal(view.history.depth, 2);
});

test('a pasted block is one undo across every cell it filled', () => {
  const view = SheetView.open(WORKBOOK, { viewportWidth: 900, viewportHeight: 500 });
  view.select(30, 0);
  view.pasteText('a\tb\tc\nd\te\tf');
  assert.equal(view.history.depth, 1);
  assert.equal(view.displayValue(30, 0).text, 'a');
  assert.equal(view.displayValue(31, 2).text, 'f');

  view.undo();
  assert.equal(view.displayValue(30, 0).text, '');
  assert.equal(view.displayValue(31, 2).text, '');
});

test('clearing a range is one undo and every value comes back', () => {
  const view = SheetView.open(WORKBOOK, { viewportWidth: 900, viewportHeight: 500 });
  view.select(0, 0);
  view.select(2, 2, { extend: true });
  const before = [];
  for (let r = 0; r <= 2; r++) for (let c = 0; c <= 2; c++) before.push(view.editValue(r, c));

  view.clearSelection();
  assert.equal(view.history.depth, 1);
  view.undo();

  const after = [];
  for (let r = 0; r <= 2; r++) for (let c = 0; c <= 2; c++) after.push(view.editValue(r, c));
  assert.deepEqual(after, before);
});

test('editing then undoing back to the start leaves the workbook untouched', () => {
  const view = SheetView.open(WORKBOOK, { viewportWidth: 900, viewportHeight: 500 });
  view.select(1, 1);
  view.setCell(1, 1, '424242');
  view.select(2, 2);
  view.setCell(2, 2, 'scratch');
  view.undo();
  view.undo();

  const diff = comparePackages(WORKBOOK, view.save());
  assert.deepEqual(diff.changed, []);
});

test('the frame tells the toolbar what the buttons should say', () => {
  const view = SheetView.open(WORKBOOK, { viewportWidth: 900, viewportHeight: 500 });
  assert.deepEqual(view.render().history, {
    canUndo: false, canRedo: false, undoLabel: null, redoLabel: null, depth: 0, redoDepth: 0,
  });

  view.select(1, 1);
  view.setCell(1, 1, '7');
  const h = view.render().history;
  assert.equal(h.canUndo, true);
  assert.equal(h.canRedo, false);
  assert.match(h.undoLabel, /^edit B2$/);
});
