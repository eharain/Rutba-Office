/**
 * Cell appearance: font family and size, colours, fills and borders.
 *
 * The append-never-rewrite rule from `cell-format.test.js` applies to `<fonts>`,
 * `<fills>` and `<borders>` exactly as it does to `<cellXfs>` — all four are
 * addressed by index, so editing an entry restyles every cell that shares it.
 *
 * Two traps here are specific to appearance and both produce a file that opens
 * and looks wrong rather than one that fails:
 *
 *   - a solid fill shows its FOREGROUND colour, so setting `bgColor` gives a
 *     cell that is still white and a bug report saying "the colour did nothing";
 *   - a colour must be eight hex digits with alpha first, and a malformed one
 *     is a file Excel declines to open.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { argb, formatOf } from '@rutba/sheet-view/styles-write';

const book = (rows = [['Head', 'Other'], [1, 2]]) =>
  buildXlsx({ sheets: [{ name: 'S', rows }] });

const open = (bytes) => new SheetView(bytes);
const stateAt = (view, row, col) => {
  view.select(row, col);
  return view.formatState();
};

// ── colours ─────────────────────────────────────────────────────────────────

test('a colour is normalised to eight hex digits, alpha first', () => {
  assert.equal(argb('#1A2B3C'), 'FF1A2B3C');
  assert.equal(argb('1a2b3c'), 'FF1A2B3C');
  assert.equal(argb('FF1A2B3C'), 'FF1A2B3C');
});

test('a malformed colour is refused rather than guessed', () => {
  // Excel declines to open a file with a bad colour in styles.xml, which is a
  // far worse outcome than a click that does nothing.
  assert.throws(() => argb('red'), /not a colour/);
  assert.throws(() => argb('#12345'), /not a colour/);
  assert.throws(() => argb(''), /not a colour/);
});

// ── font family, size and colour ────────────────────────────────────────────

test('font family and size are set, and survive a save', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ fontName: 'Georgia', fontSize: 18 });

  const reopened = open(view.save());
  const at = stateAt(reopened, 0, 0);
  assert.equal(at.fontName, 'Georgia');
  assert.equal(at.fontSize, 18);
});

test('font colour is set without disturbing the weight', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ bold: true });
  view.setFormat({ fontColour: '#C00000' });

  const at = view.formatState();
  assert.equal(at.fontColour, 'FFC00000');
  assert.equal(at.bold, true, 'setting a colour did not clear the bold');
});

test('changing only the size keeps the family', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ fontName: 'Georgia', fontSize: 11 });
  view.setFormat({ fontSize: 20 });

  const at = view.formatState();
  assert.equal(at.fontName, 'Georgia', 'the family was carried across');
  assert.equal(at.fontSize, 20);
});

// ── fills ───────────────────────────────────────────────────────────────────

test('a fill paints the foreground colour, which is the one that shows', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ fill: '#FFE699' });

  const xml = view.pkg.text('xl/styles.xml');
  assert.match(xml, /<fgColor rgb="FFFFE699"\/>/, 'a solid fill shows fgColor');
  assert.equal(view.formatState().fill, 'FFFFE699');
});

test('a fill survives a save and reopens as itself', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ fill: '#FFE699' });
  const reopened = open(view.save());
  assert.equal(stateAt(reopened, 0, 0).fill, 'FFFFE699');
  assert.equal(stateAt(reopened, 1, 0).fill, null, 'and the cell below is unfilled');
});

test('the same fill twice reuses the entry', () => {
  const view = open(book([['a', 'b']]));
  view.select(0, 0);
  view.setFormat({ fill: '#DDEEFF' });
  const after1 = view.pkg.text('xl/styles.xml');
  view.select(0, 1);
  view.setFormat({ fill: '#DDEEFF' });
  const after2 = view.pkg.text('xl/styles.xml');

  const fills = (x) => [...x.matchAll(/<fill>/g)].length;
  assert.equal(fills(after2), fills(after1));
});

// ── borders ─────────────────────────────────────────────────────────────────

test('a border edge is set, and the others are left alone', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ border: { bottom: { style: 'thin', colour: '#000000' } } });

  const at = view.formatState();
  assert.equal(at.border.bottom?.style, 'thin');
  assert.equal(at.border.top, undefined, 'only the edge asked for was drawn');
});

test('adding a second edge keeps the first', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ border: { left: { style: 'thin', colour: '#000000' } } });
  view.setFormat({ border: { bottom: { style: 'medium', colour: '#C00000' } } });

  const at = view.formatState();
  assert.equal(at.border.left?.style, 'thin', 'the left rule was not thrown away');
  assert.equal(at.border.bottom?.style, 'medium');
  assert.equal(at.border.bottom?.colour, 'FFC00000');
});

test('an edge is cleared by naming it null', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ border: { bottom: { style: 'thin', colour: '#000000' } } });
  view.setFormat({ border: { bottom: null } });
  assert.equal(view.formatState().border.bottom, undefined);
});

test('a border survives a save', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ border: { top: { style: 'double', colour: '#333333' } } });
  const reopened = open(view.save());
  assert.equal(stateAt(reopened, 0, 0).border.top?.style, 'double');
});

// ── the whole lot at once, which is what a header row really is ─────────────

test('a header row takes weight, colour, fill and a rule in one pass', () => {
  const view = open(book([['Region', 'Revenue'], ['North', 120]]));
  view.select(0, 0);
  view.select(0, 1, { extend: true });
  view.setFormat({
    bold: true,
    fontColour: '#FFFFFF',
    fill: '#305496',
    align: 'center',
    border: { bottom: { style: 'medium', colour: '#1F3864' } },
  });

  const reopened = open(view.save());
  for (const col of [0, 1]) {
    const at = stateAt(reopened, 0, col);
    assert.equal(at.bold, true, 'bold at column ' + col);
    assert.equal(at.fontColour, 'FFFFFFFF');
    assert.equal(at.fill, 'FF305496');
    assert.equal(at.align, 'center');
    assert.equal(at.border.bottom?.style, 'medium');
  }
  // Both header cells are identical, so they must share ONE style entry.
  const a = reopened._styleIndexAt('S', 0, 0);
  const b = reopened._styleIndexAt('S', 0, 1);
  assert.equal(a, b, 'two cells with the same appearance share one style');

  assert.equal(stateAt(reopened, 1, 0).fill, null, 'the body row is untouched');
});

test('the whole style table stays append-only through all of it', () => {
  const view = open(book([['a', 'b'], ['c', 'd']]));
  const bodyOf = (xml, tag) => new RegExp('<' + tag + '\\b[^>]*>([\\s\\S]*?)</' + tag + '>').exec(xml)[1];
  view.select(0, 0);
  view.setFormat({ bold: true });
  const before = view.pkg.text('xl/styles.xml');

  view.select(1, 1);
  view.setFormat({ fill: '#FFE699', fontColour: '#C00000', border: { top: { style: 'thin' } } });
  const after = view.pkg.text('xl/styles.xml');

  for (const tag of ['fonts', 'fills', 'borders', 'cellXfs']) {
    assert.ok(bodyOf(after, tag).startsWith(bodyOf(before, tag)),
      tag + ' grew at the end; nothing already in it moved or changed');
  }
});

test('undo puts appearance back, table and all', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ bold: true });          // so styles.xml exists and is non-trivial
  const before = view.pkg.text('xl/styles.xml');

  view.select(1, 0);
  view.setFormat({ fill: '#FFE699', border: { bottom: { style: 'thin' } } });
  assert.equal(view.formatState().fill, 'FFFFE699');

  view.undo();
  assert.equal(view.formatState().fill, null);
  assert.equal(view.pkg.text('xl/styles.xml'), before,
    'the fill and border entries it appended went with it');
});
