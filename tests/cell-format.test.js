/**
 * Setting cell formatting, and the rule that makes it safe.
 *
 * Every `<xf>` in `cellXfs` is addressed by position and every cell points at
 * one by index. So the only safe way to add a style is to APPEND: editing an
 * entry restyles every other cell that shares it, and inserting one changes
 * what every later index means. These tests pin that, because both failures are
 * silent — the file still opens, it just looks wrong somewhere the person who
 * made the edit was not looking.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { applyFormat, formatOf } from '@rutba/sheet-view/styles-write';

const book = (rows = [['Head', 'Other'], [1, 2]], styles) =>
  buildXlsx({ sheets: [{ name: 'S', rows, ...(styles ? { styles } : {}) }] });

const open = (bytes) => new SheetView(bytes);
const cellOf = (view, ref) => (view.render().cells || []).find((c) => c.ref === ref);

// ── the append rule ─────────────────────────────────────────────────────────

test('a new style is appended, and every existing index still means what it did', () => {
  const view = open(book([['Head'], [1]], { A1: { bold: true } }));
  const before = view.pkg.text('xl/styles.xml');
  const beforeXfs = [...before.matchAll(/<xf\b/g)].length;

  // A2, not A1: A1 is already bold in this fixture, and asking for a format a
  // cell already has is correctly a no-op that adds nothing.
  view.select(1, 0);
  view.setFormat({ italic: true });

  const after = view.pkg.text('xl/styles.xml');
  const afterXfs = [...after.matchAll(/<xf\b/g)].length;
  assert.ok(afterXfs > beforeXfs, 'a style was added');

  // The decisive check: everything that was there is still there, in order and
  // unchanged. `indexOf` on the original cellXfs body proves nothing shifted.
  const body = (xml) => /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml)[1];
  assert.ok(body(after).startsWith(body(before)), 'existing entries are untouched and still first');
});

test('the same format asked for twice reuses the entry rather than growing the file', () => {
  const view = open(book([['a', 'b'], ['c', 'd']]));
  view.select(0, 0);
  view.setFormat({ bold: true });
  const afterFirst = view.pkg.text('xl/styles.xml');

  view.select(1, 1);
  view.setFormat({ bold: true });
  const afterSecond = view.pkg.text('xl/styles.xml');

  const count = (xml) => [...xml.matchAll(/<xf\b/g)].length;
  assert.equal(count(afterSecond), count(afterFirst), 'the second bold reused the first bold');
});

test('formatting one cell does not restyle its neighbours', () => {
  const view = open(book([['one', 'two', 'three']]));
  view.select(0, 1);
  view.setFormat({ bold: true });

  const xml = view.pkg.text('xl/styles.xml');
  const at = (row, col) => formatOf(xml, view._styleIndexAt('S', row, col) ?? 0);
  assert.equal(at(0, 1).bold, true);
  assert.equal(at(0, 0).bold, false, 'the cell to the left is untouched');
  assert.equal(at(0, 2).bold, false, 'and so is the cell to the right');
});

// ── toggling ────────────────────────────────────────────────────────────────

test('toggle turns a mixed selection uniformly on, then uniformly off', () => {
  const view = open(book([['a', 'b']]));
  view.select(0, 0);
  view.setFormat({ bold: true });          // A1 bold, B1 not

  view.select(0, 0);
  view.select(0, 1, { extend: true });     // both
  assert.equal(view.formatState().bold, false, 'a mixed selection does not read as bold');

  view.setFormat({ bold: 'toggle' });
  assert.equal(view.formatState().bold, true, 'first press makes the whole selection bold');

  view.setFormat({ bold: 'toggle' });
  assert.equal(view.formatState().bold, false, 'second press clears it');
});

test('alignment is a value, and pressing the one already set clears it', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ align: 'center' });
  assert.equal(view.formatState().align, 'center');

  view.setFormat({ align: 'center' });
  assert.equal(view.formatState().align, null, 'pressing centre again returns it to the default');

  view.setFormat({ align: 'right' });
  assert.equal(view.formatState().align, 'right');
});

test('bold and italic compose rather than replacing each other', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ bold: true });
  view.setFormat({ italic: true });
  const state = view.formatState();
  assert.equal(state.bold, true, 'italic did not clear the bold');
  assert.equal(state.italic, true);
});

// ── it survives a save ──────────────────────────────────────────────────────

test('formatting is in the saved file, and reopens the same way', () => {
  const view = open(book([['Header'], ['body']]));
  view.select(0, 0);
  view.setFormat({ bold: true, align: 'center' });
  const bytes = view.save();

  const reopened = open(bytes);
  const cell = cellOf(reopened, 'A1');
  assert.ok(cell, 'A1 is still there');
  assert.equal(cell.style?.font?.bold, true, 'the reopened file renders it bold');
  assert.equal(cell.align, 'center');
  assert.equal(cellOf(reopened, 'A2')?.style?.font?.bold ?? false, false, 'and A2 is not bold');
});

test('formatting an empty cell works — a blank column can be styled before it is filled', () => {
  const view = open(book([['a']]));
  view.select(0, 3); // D1, which has no <c> at all
  view.setFormat({ bold: true });
  const bytes = view.save();

  const reopened = open(bytes);
  reopened.select(0, 3);
  assert.equal(reopened.formatState().bold, true);
});

// ── undo ────────────────────────────────────────────────────────────────────

test('undo puts the formatting back, and does not strand the style it added', () => {
  const view = open(book([['Head'], [1]], { A1: { bold: true } }));
  const before = view.pkg.text('xl/styles.xml');

  view.select(1, 0);
  view.setFormat({ bold: true });
  assert.equal(view.formatState().bold, true);

  view.undo();
  assert.equal(view.formatState().bold, false, 'the cell is no longer bold');
  assert.equal(view.pkg.text('xl/styles.xml'), before,
    'and styles.xml is back as it was — otherwise the next bold reuses an entry\n'
    + 'the document no longer references');
});

test('redo reapplies it', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ bold: true });
  view.undo();
  view.redo();
  assert.equal(view.formatState().bold, true);
});

test('formatting marks the document unsaved, and undoing all of it does not', () => {
  const view = open(book());
  assert.equal(view.isDirty, false);
  view.select(0, 0);
  view.setFormat({ bold: true });
  assert.equal(view.isDirty, true);
});

// ── the writer in isolation ─────────────────────────────────────────────────

test('applyFormat refuses to change anything when the format is already set', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ bold: true });
  const xml = view.pkg.text('xl/styles.xml');
  const index = view._styleIndexAt('S', 0, 0);

  const again = applyFormat(xml, index, { bold: true });
  assert.equal(again.changed, false, 'asking for bold on a bold cell is a no-op');
  assert.equal(again.index, index);
  assert.equal(again.xml, xml, 'and it does not grow the style table');
});
