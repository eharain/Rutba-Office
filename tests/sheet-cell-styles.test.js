// Home → Cell Styles: theme colours and named styles.
//
// A style used to be a handful of fixed colours painted on the cells: nothing
// in the file said "this is Heading 1", and a heading stayed navy whatever
// theme the workbook wore. Now the heading styles take their colours from the
// theme, as Excel's do, and a style is kept in the file under Excel's own
// name, with each cell that wears it pointing at it, so the cells can be
// restyled together when the style changes.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { CELL_STYLES } from '../packages/sheet-view/src/cell-styles.js';
import { applyFormat, namedStyles, formatOf } from '../packages/sheet-view/src/styles-write.js';

const BOOK = () => buildXlsx({ sheets: [{ name: 'Report', rows: [['Sales', 'Costs', 'Notes'], [10, 4, 'ok'], [20, 9, 'late'], ['Total', 13, '']] }] });
const styles = (view) => view.pkg.text('xl/styles.xml');
const shown = (view, row, col) => view.styles.byStyleIndex[view._shownStyleIndex(view.activeSheet, row, col) ?? 0];
const wear = (view, row, col, name, toCol = col) => {
  view.select(row, col);
  if (toCol !== col) view.selection.extendTo(row, toCol);
  view.applyCellStyle({ name });
};

test('a heading style is written in theme colours, and the file names it as Excel does', () => {
  const view = new SheetView(BOOK());
  wear(view, 0, 0, 'Heading 1', 1);
  const xml = styles(view);
  assert.match(xml, /<font><b\/>.*?<color theme="3"\/>.*?<\/font>/, 'the dark text colour of the theme, not a fixed navy');
  assert.match(xml, /<bottom style="medium"><color theme="4"\/><\/bottom>/, 'the border is the first accent');
  assert.deepEqual(namedStyles(xml), [{ name: 'Normal', xfId: 0, builtinId: 0 }, { name: 'Heading 1', xfId: 1, builtinId: 16 }]);
  assert.match(xml, /<cellStyleXfs count="2">.*<xf numFmtId="0" fontId="\d+" fillId="0" borderId="\d+" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"\/><\/cellStyleXfs>/, 'the style speaks for font and border only');
  assert.match(xml, /<cellXfs count="\d+">.*<xf [^>]*xfId="1"[^>]*\/><\/cellXfs>/, 'the cell points at its style');
  assert.equal(view.formatState().style, 'Heading 1');
  assert.equal(shown(view, 0, 0).font.bold, true);

  // The same style on more cells makes nothing new.
  wear(view, 0, 2, 'heading 1');
  assert.equal(namedStyles(styles(view)).length, 2);
  assert.equal(view._shownStyleIndex('Report', 0, 2), view._shownStyleIndex('Report', 0, 0));

  const again = SheetView.open(view.save());
  again.select(0, 1);
  assert.equal(again.formatState().style, 'Heading 1', 'the name is in the saved file');
  again.select(1, 1);
  assert.equal(again.formatState().style, 'Normal');
});

test('cells in a heading style change colour with the theme', () => {
  const view = new SheetView(BOOK());
  wear(view, 0, 0, 'Heading 1');
  wear(view, 3, 0, 'Total', 1);
  const before = { font: shown(view, 0, 0).font.colour, edge: shown(view, 0, 0).border.bottom.colour, total: shown(view, 3, 0).border.top.colour };
  view.setWorkbookTheme({ theme: 'harbour' });
  const after = { font: shown(view, 0, 0).font.colour, edge: shown(view, 0, 0).border.bottom.colour, total: shown(view, 3, 0).border.top.colour };
  assert.deepEqual(after, { font: '#1f4e79', edge: '#1f6fb2', total: '#1f6fb2' });
  assert.notDeepEqual(before, after);
  view.undo();
  assert.equal(shown(view, 0, 0).font.colour, before.font, 'undo brings the old theme and the old colours back');
});

test('a style leaves alone what it does not speak for, and the fixed-colour styles stay fixed', () => {
  const view = new SheetView(BOOK());
  view.select(1, 0);
  view.setFormat({ numberFormat: '0.00', fill: '#FFF2CC', align: 'center' });
  wear(view, 1, 0, 'Heading 2');
  const at = formatOf(styles(view), view._shownStyleIndex('Report', 1, 0));
  assert.deepEqual([at.numberFormat, at.fill, at.align, at.bold, at.fontSize], ['0.00', 'FFFFF2CC', 'center', true, 13]);

  wear(view, 2, 2, 'Bad');
  assert.match(styles(view), /<fgColor rgb="FFFFC7CE"\/>/);
  view.setWorkbookTheme({ theme: 'harbour' });
  assert.equal(shown(view, 2, 2).fill.colour.toLowerCase(), '#ffc7ce');

  wear(view, 1, 1, 'Currency');
  assert.equal(view.displayValue(1, 1).text, '£4.00');
  assert.equal(view.formatState().style, 'Currency');
  assert.throws(() => view.applyCellStyle({ name: 'Sparkly' }), /no cell style called Sparkly/);
});

test('Normal takes a style off again', () => {
  const view = new SheetView(BOOK());
  wear(view, 0, 0, 'Heading 1');
  wear(view, 0, 0, 'Normal');
  const at = formatOf(styles(view), view._shownStyleIndex('Report', 0, 0));
  assert.deepEqual([at.bold, at.styleXf, at.border], [false, 0, {}]);
  assert.equal(view.formatState().style, 'Normal');
});

test('a style made over from one cell restyles every cell that wears it, in one undo step', () => {
  const view = new SheetView(BOOK());
  wear(view, 0, 0, 'Heading 1', 2);
  wear(view, 3, 0, 'Total');
  // One heading cell is given a fill and a size by hand, as a person tries a look.
  view.select(0, 1);
  view.setFormat({ fill: '#DDEBF7', fontSize: 20 });
  assert.equal(view.formatState().style, 'Heading 1', 'a cell changed by hand still wears its style');
  assert.equal(shown(view, 0, 0).fill ?? null, null);

  assert.equal(view.updateCellStyle(), 'Heading 1');
  for (const col of [0, 1, 2]) {
    assert.equal(shown(view, 0, col).fill.colour.toLowerCase(), '#ddebf7', 'column ' + col + ' has the fill');
    assert.equal(shown(view, 0, col).font.sizePt, 20);
  }
  assert.equal(shown(view, 3, 0).fill ?? null, null, 'a cell in another style is as it was');
  assert.equal(shown(view, 1, 0).fill ?? null, null);
  assert.match(styles(view), /<cellStyleXfs count="3"><xf [^>]*\/><xf numFmtId="0" fontId="\d+" fillId="[1-9]\d*" borderId="\d+" applyNumberFormat="0" applyFill="1" /, 'the style now speaks for the fill too');

  // A cell given the style afterwards wears the new look.
  wear(view, 2, 0, 'Heading 1');
  assert.equal(shown(view, 2, 0).fill.colour.toLowerCase(), '#ddebf7');
  assert.equal(shown(view, 2, 0).font.sizePt, 20);

  const reopened = SheetView.open(view.save());
  assert.equal(shown(reopened, 0, 2).font.sizePt, 20, 'saved, and read back the same');

  view.undo(); // the style on A3
  view.undo(); // the style made over
  assert.equal(shown(view, 0, 0).fill ?? null, null, 'undo takes the new look off every cell');
  assert.equal(shown(view, 0, 1).fill.colour.toLowerCase(), '#ddebf7', 'and leaves the one changed by hand');

  view.select(1, 1);
  assert.throws(() => view.updateCellStyle(), /wears no named style/);
});

test('an edge in a theme colour keeps it when another edge of the cell changes', () => {
  const view = new SheetView(BOOK());
  wear(view, 0, 0, 'Heading 1');
  view.select(0, 0);
  view.setFormat({ border: { top: { style: 'thin', colour: '#000000' } } });
  assert.match(styles(view), /<top style="thin"><color rgb="FF000000"\/><\/top><bottom style="medium"><color theme="4"\/><\/bottom>/);
});

test('every built-in style can be given and saved, and a colour that is not one is refused', () => {
  const view = new SheetView(BOOK());
  CELL_STYLES.forEach((style, i) => {
    view.select(10 + i, 0);
    view.applyCellStyle({ name: style.name });
    assert.equal(view.formatState().style, style.name);
  });
  const again = SheetView.open(view.save());
  assert.deepEqual(namedStyles(styles(again)).map((s) => s.name), CELL_STYLES.map((s) => s.name));
  assert.throws(() => applyFormat(styles(view), 0, { fontColour: { theme: 40 } }), /not a theme colour/);
});
