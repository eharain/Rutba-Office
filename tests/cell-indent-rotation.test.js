/**
 * Indent and text orientation: the two cell-alignment properties the Home
 * tab greyed out. Both live in the `<alignment>` element beside horizontal
 * and vertical alignment and wrap, under the same append-never-rewrite rule
 * as `cell-appearance.test.js`: setting one appends an `<xf>`, and nothing
 * already in `cellXfs` changes meaning.
 *
 * The vocabulary is Excel's own. `indent` counts in Excel's units and rides
 * with a left, right or distributed alignment (a cell without one is given
 * left, as Excel writes when the arrow is pressed). `textRotation` is 1..90
 * anticlockwise, 91..180 clockwise by the value less 90, and 255 for letters
 * stacked upright; zero is horizontal and carries no attr.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const book = (rows = [['Head', 'Other'], [1, 2]]) =>
  buildXlsx({ sheets: [{ name: 'S', rows }] });

const open = (bytes) => new SheetView(bytes);
const stateAt = (view, row, col) => {
  view.select(row, col);
  return view.formatState();
};

/** The `<xf>` a cell currently points at, as raw XML. */
const xfOf = (view, sheet, row, col) => {
  const xml = view.pkg.text('xl/styles.xml');
  const idx = view._styleIndexAt(sheet, row, col) ?? 0;
  const body = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml)[1];
  const xfs = [...body.matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map((m) => m[0]);
  return xfs[idx];
};

// ── indent ───────────────────────────────────────────────────────────────────

test('an indent is written beside a left alignment, and survives a save', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ indent: 2 });

  assert.equal(view.formatState().indent, 2, 'set in-session');
  assert.equal(view.formatState().align, 'left', 'an indent rides with a left alignment, as Excel writes one');
  assert.match(xfOf(view, 'S', 0, 0), /indent="2"/);

  const reopened = open(view.save());
  assert.equal(stateAt(reopened, 0, 0).indent, 2, 'and after reopen');
  assert.equal(stateAt(reopened, 1, 0).indent, 0, 'the cell below is untouched');
});

test('the arrows step each cell from its own indent, never below zero', () => {
  const view = open(book([['a', 'x'], ['b', 'y']]));
  view.select(0, 0);
  view.setFormat({ indent: 3 });          // A1 at 3, A2 at 0

  view.select(0, 0);
  view.select(1, 0, { extend: true });
  assert.equal(view.formatState().indent, null, 'two indents have none to show');
  view.setFormat({ indentBy: 1 });        // A1 → 4, A2 → 1
  assert.equal(stateAt(view, 0, 0).indent, 4);
  assert.equal(stateAt(view, 1, 0).indent, 1);

  view.select(0, 0);
  view.select(1, 0, { extend: true });
  view.setFormat({ indentBy: -1 });
  view.setFormat({ indentBy: -1 });       // A1 → 2, A2 → 0, not −1
  assert.equal(stateAt(view, 0, 0).indent, 2);
  assert.equal(stateAt(view, 1, 0).indent, 0);
  assert.doesNotMatch(xfOf(view, 'S', 1, 0), /indent=/, 'zero carries no attr');
});

test('a right alignment keeps its indent; centring drops it, as Excel does', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ align: 'right' });
  view.setFormat({ indent: 1 });
  assert.equal(view.formatState().align, 'right', 'right is an alignment an indent rides with');
  assert.match(xfOf(view, 'S', 0, 0), /horizontal="right"/);
  assert.match(xfOf(view, 'S', 0, 0), /indent="1"/);

  view.setFormat({ align: 'center' });
  assert.equal(view.formatState().indent, 0, 'a centred cell has no indent');
  assert.doesNotMatch(xfOf(view, 'S', 0, 0), /indent=/);
});

// ── text orientation ─────────────────────────────────────────────────────────

for (const [label, rotation] of [['anticlockwise', 45], ['clockwise', 135], ['up', 90], ['down', 180], ['stacked', 255]]) {
  test('text rotation ' + label + ' writes textRotation and survives a save', () => {
    const view = open(book());
    view.select(0, 0);
    view.setFormat({ rotation });

    assert.equal(view.formatState().rotation, rotation, 'set in-session');
    assert.match(xfOf(view, 'S', 0, 0), new RegExp('textRotation="' + rotation + '"'));

    const reopened = open(view.save());
    assert.equal(stateAt(reopened, 0, 0).rotation, rotation, 'and after reopen');
    assert.equal(stateAt(reopened, 1, 0).rotation, 0, 'the cell below is untouched');
  });
}

test('horizontal text takes the rotation off, and the style keeps what is left', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ rotation: 90, indent: 1 });
  assert.equal(view.styleFor(0, 0).align.rotation, 90, 'the painter reads the rotation');
  assert.equal(view.styleFor(0, 0).align.indent, 1, 'and the indent');

  view.setFormat({ rotation: 0 });
  assert.equal(view.formatState().rotation, 0);
  assert.doesNotMatch(xfOf(view, 'S', 0, 0), /textRotation/, 'horizontal carries no attr');
  assert.equal(view.formatState().indent, 1, 'the indent survived');
});

// ── undo ─────────────────────────────────────────────────────────────────────

test('undo puts the indent and rotation back, style table and all', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ bold: true });             // so styles.xml exists and is non-trivial
  const before = view.pkg.text('xl/styles.xml');

  view.select(1, 0);
  view.setFormat({ indent: 2, rotation: 45 });
  assert.equal(view.formatState().indent, 2);
  assert.equal(view.formatState().rotation, 45);

  view.undo();
  assert.equal(view.formatState().indent, 0);
  assert.equal(view.formatState().rotation, 0);
  assert.equal(view.pkg.text('xl/styles.xml'), before,
    'the alignment xf it appended went with it');
});
