/**
 * Vertical alignment and wrap text: two cell-alignment properties that were
 * already READ and PAINTED but could not be SET.
 *
 * They live in the same `<alignment>` element as horizontal alignment, and the
 * same append-never-rewrite rule from `cell-appearance.test.js` applies: setting
 * one appends a new `<xf>`, and every entry already in `cellXfs` keeps its
 * meaning so no unrelated cell is restyled.
 *
 * The OOXML vocabulary is used verbatim — top/center/bottom for vertical, and
 * wrapText="1" for wrap — so nothing is translated on the way in or out.
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

// ── vertical alignment ───────────────────────────────────────────────────────

for (const valign of ['top', 'center', 'bottom']) {
  test('vertical ' + valign + ' is set and survives a save', () => {
    const view = open(book());
    view.select(0, 0);
    view.setFormat({ valign });

    assert.equal(view.formatState().valign, valign, 'set in-session');
    assert.match(xfOf(view, 'S', 0, 0), new RegExp('vertical="' + valign + '"'));

    const reopened = open(view.save());
    assert.equal(stateAt(reopened, 0, 0).valign, valign, 'and after reopen');
    assert.equal(stateAt(reopened, 1, 0).valign, null, 'the cell below is untouched');
  });
}

test('clearing valign removes the vertical attr', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ valign: 'center' });
  assert.match(xfOf(view, 'S', 0, 0), /vertical="center"/);

  view.setFormat({ valign: null });
  assert.equal(view.formatState().valign, null);
  assert.doesNotMatch(xfOf(view, 'S', 0, 0), /vertical=/, 'the attr is gone from the cell');
});

// ── wrap text ────────────────────────────────────────────────────────────────

test('wrap text is turned on, writes wrapText="1", and survives a save', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ wrap: true });

  assert.equal(view.formatState().wrap, true, 'on in-session');
  assert.match(xfOf(view, 'S', 0, 0), /wrapText="1"/);

  const reopened = open(view.save());
  assert.equal(stateAt(reopened, 0, 0).wrap, true, 'and after reopen');
});

test('wrap text is turned off and the attr is removed', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ wrap: true });
  view.setFormat({ wrap: false });

  assert.equal(view.formatState().wrap, false);
  assert.doesNotMatch(xfOf(view, 'S', 0, 0), /wrapText/, 'off carries no attr, not wrapText="0"');

  const reopened = open(view.save());
  assert.equal(stateAt(reopened, 0, 0).wrap, false, 'off survives a save');
});

// ── the append rule ──────────────────────────────────────────────────────────

test('setting valign and wrap keeps the style table append-only', () => {
  const view = open(book([['a', 'b'], ['c', 'd']]));
  const bodyOf = (xml, tag) => new RegExp('<' + tag + '\\b[^>]*>([\\s\\S]*?)</' + tag + '>').exec(xml)[1];
  view.select(0, 0);
  view.setFormat({ bold: true });
  const before = view.pkg.text('xl/styles.xml');

  view.select(1, 1);
  view.setFormat({ valign: 'top', wrap: true });
  const after = view.pkg.text('xl/styles.xml');

  assert.ok(bodyOf(after, 'cellXfs').startsWith(bodyOf(before, 'cellXfs')),
    'cellXfs grew at the end; nothing already in it moved or changed');
});

// ── reconciliation across a selection ────────────────────────────────────────

test('a mixed selection reports valign null', () => {
  const view = open(book([['top', 'x'], ['plain', 'y']]));
  view.select(0, 0);
  view.setFormat({ valign: 'top' });          // A1 is top-aligned, A2 is not

  view.select(0, 0);
  view.select(1, 0, { extend: true });
  assert.equal(view.formatState().valign, null, 'two vertical alignments have none to show');
});

test('a selection where only some cells wrap reports wrap false', () => {
  const view = open(book([['a', 'x'], ['b', 'y']]));
  view.select(0, 0);
  view.setFormat({ wrap: true });             // A1 wraps, A2 does not

  view.select(0, 0);
  view.select(1, 0, { extend: true });
  assert.equal(view.formatState().wrap, false, 'a partly-wrapping selection reads as off');
});

// ── independence from the rest of the alignment element ──────────────────────

test('setting valign leaves horizontal align, font and fill alone', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ align: 'center', bold: true, fill: '#FFE699' });
  view.setFormat({ valign: 'bottom' });

  const at = view.formatState();
  assert.equal(at.valign, 'bottom', 'the vertical alignment was set');
  assert.equal(at.align, 'center', 'the horizontal alignment survived');
  assert.equal(at.bold, true, 'the weight survived');
  assert.equal(at.fill, 'FFFFE699', 'the fill survived');

  // And the reverse: horizontal alignment does not disturb an existing vertical.
  view.setFormat({ align: 'right' });
  const after = view.formatState();
  assert.equal(after.valign, 'bottom', 'the vertical alignment was not thrown away');
  assert.equal(after.align, 'right');
});

// ── undo ─────────────────────────────────────────────────────────────────────

test('undo puts valign and wrap back, style table and all', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ bold: true });             // so styles.xml exists and is non-trivial
  const before = view.pkg.text('xl/styles.xml');

  view.select(1, 0);
  view.setFormat({ valign: 'center', wrap: true });
  assert.equal(view.formatState().valign, 'center');
  assert.equal(view.formatState().wrap, true);

  view.undo();
  assert.equal(view.formatState().valign, null);
  assert.equal(view.formatState().wrap, false);
  assert.equal(view.pkg.text('xl/styles.xml'), before,
    'the alignment xf it appended went with it');
});
