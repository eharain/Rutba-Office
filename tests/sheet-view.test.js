/**
 * The rendering layer.
 *
 * The parts users feel every second — how a number reads, where the cursor goes,
 * what the formula bar shows — tested as pure logic, so the browser shell can
 * stay thin enough to verify by reading it.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  formatValue, readNumberFormats, BUILTIN_FORMATS, splitSections,
  readStyles, readMergedCells, readTheme, applyTint, readColourElement, THEME_SLOTS,
  SheetGeometry, charWidthToPixels, pointsToPixels,
  Selection, ref, colName,
  SheetView, coerceInput,
} from '@rutba/sheet-view';
import { OoxmlPackage, Workbook, comparePackages } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';
import { ERR } from '@rutba/formula';
import { buildComplexWorkbook, FRAGILE_PARTS, FRAGILE_SHEET_ELEMENTS } from './fixtures/complex-workbook.js';

const FIXTURE = buildComplexWorkbook();

// -------------------------------------------------------- number formats --

test('the built-in format table matches the spec', () => {
  assert.equal(BUILTIN_FORMATS[0], 'General');
  assert.equal(BUILTIN_FORMATS[4], '#,##0.00');
  assert.equal(BUILTIN_FORMATS[9], '0%');
  assert.equal(BUILTIN_FORMATS[14], 'mm-dd-yy');
  assert.equal(BUILTIN_FORMATS[49], '@');
});

test('General renders a number the way a spreadsheet does', () => {
  assert.equal(formatValue(1661.4, 'General').text, '1661.4');
  assert.equal(formatValue(0.30000000000000004, 'General').text, '0.3', 'float noise is hidden');
  assert.equal(formatValue(1000000, 'General').text, '1000000');
});

test('thousands, decimals and padding', () => {
  assert.equal(formatValue(1234.5, '#,##0.00').text, '1,234.50');
  assert.equal(formatValue(1234.5, '#,##0').text, '1,235', 'rounds to the pattern');
  assert.equal(formatValue(0.5, '0.00').text, '0.50');
  assert.equal(formatValue(5, '000').text, '005');
  assert.equal(formatValue(1234567, '#,##0,').text, '1,235', 'a trailing comma scales by thousands');
});

test('percentages multiply by 100', () => {
  assert.equal(formatValue(0.155, '0%').text, '16%');
  assert.equal(formatValue(0.155, '0.00%').text, '15.50%');
});

test('a literal in the format is printed, and currency reads correctly', () => {
  assert.equal(formatValue(1850, '#,##0.00" PKR"').text, '1,850.00 PKR');
  assert.equal(formatValue(1850, '"$"#,##0.00').text, '$1,850.00');
});

test('the negative section formats the ABSOLUTE value', () => {
  // the sign comes from the format; otherwise this renders "(-500)"
  assert.equal(formatValue(-500, '#,##0;(#,##0)').text, '(500)');
  assert.equal(formatValue(500, '#,##0;(#,##0)').text, '500');
  assert.equal(formatValue(-500, '#,##0').text, '-500', 'one section keeps the minus sign');
});

test('sections select on positive, negative, zero and text', () => {
  const code = '#,##0.00;[Red](#,##0.00);"-";@';
  assert.equal(formatValue(1234.5, code).text, '1,234.50');
  const neg = formatValue(-1234.5, code);
  assert.equal(neg.text, '(1,234.50)');
  assert.equal(neg.colour, 'red', 'colour is a hint, not printed');
  assert.equal(formatValue(0, code).text, '-');
  assert.equal(formatValue('n/a', code).text, 'n/a');
  assert.deepEqual(splitSections(code).length, 4);
});

test('a semicolon inside quotes does not split a section', () => {
  assert.equal(splitSections('"a;b"#,##0').length, 1);
  assert.equal(formatValue(5, '"a;b"0').text, 'a;b5');
});

test('dates render, and m means month or minute by context', () => {
  const serial = 46254; // 2026-08-20
  assert.equal(formatValue(serial, 'yyyy-mm-dd').text, '2026-08-20');
  assert.equal(formatValue(serial, 'd-mmm-yy').text, '20-Aug-26');
  assert.equal(formatValue(serial, 'mmmm').text, 'August');
  assert.equal(formatValue(serial, 'dddd').text, 'Thursday');

  // the classic ambiguity: mm after h is minutes, mm before dd is months
  assert.equal(formatValue(serial + 0.5, 'h:mm').text, '12:00');
  assert.equal(formatValue(serial, 'mm-dd').text, '08-20');
  assert.equal(formatValue(serial + 0.5, 'h:mm AM/PM').text, '12:00 PM');
  assert.equal(formatValue(serial + 0.25, 'h:mm AM/PM').text, '6:00 AM');
});

test('errors and booleans ignore the format', () => {
  assert.equal(formatValue(ERR.DIV0(), '#,##0.00').text, '#DIV/0!');
  assert.equal(formatValue(true, '#,##0.00').text, 'TRUE');
  assert.equal(formatValue('', '#,##0.00').text, '');
});

test('cell formats are read from the workbook style table', () => {
  const pkg = OoxmlPackage.read(FIXTURE);
  const { byStyleIndex } = readNumberFormats(pkg);
  // the fixture's cellXfs: index 4 is numFmtId 3 (#,##0), index 5 is custom 164
  assert.equal(byStyleIndex[4], '#,##0');
  assert.equal(byStyleIndex[5], '#,##0.00" PKR"');
  assert.equal(byStyleIndex[0], 'General');
});

// ------------------------------------------------------------- geometry --

test('column widths convert from character units by the format\'s own formula', () => {
  // 9.140625 is the value a file stores for a default column — it already
  // includes the 5px padding that the UI's "8.43 characters" does not.
  assert.equal(charWidthToPixels(9.140625), 64, 'the default column is 64px');
  assert.equal(charWidthToPixels(8.43), 59, 'the UI number is not the stored one');
  assert.equal(charWidthToPixels(18.5), 129);
  assert.equal(pointsToPixels(15), 20, 'the default row is 20px');
});

test('geometry reads widths and heights out of the sheet', () => {
  const xml = OoxmlPackage.read(FIXTURE).text('xl/worksheets/sheet1.xml');
  const geo = SheetGeometry.fromSheetXml(xml);
  assert.equal(geo.colWidth(0), charWidthToPixels(18.5), 'column A is customised');
  assert.equal(geo.colWidth(1), charWidthToPixels(12.75));
  assert.equal(geo.colWidth(9), geo.defaultColWidth, 'uncustomised columns use the default');
  assert.equal(geo.rowHeight(0), pointsToPixels(22.5), 'the header row is taller');
  assert.equal(geo.rowHeight(1), geo.defaultRowHeight);
});

test('offsets and hit-testing agree with each other', () => {
  const geo = new SheetGeometry();
  geo.colWidths.set(0, 200);
  assert.equal(geo.colOffset(0), 0);
  assert.equal(geo.colOffset(1), 200);
  assert.equal(geo.colOffset(2), 200 + geo.defaultColWidth);
  assert.equal(geo.colAt(0), 0);
  assert.equal(geo.colAt(199), 0);
  assert.equal(geo.colAt(200), 1);
  assert.equal(geo.rowAt(geo.defaultRowHeight * 5), 5);
});

test('a hidden column takes no space', () => {
  const geo = new SheetGeometry();
  geo.hiddenCols.add(1);
  assert.equal(geo.colWidth(1), 0);
  assert.equal(geo.colOffset(2), geo.defaultColWidth, 'the hidden column is skipped');
});

test('virtualisation renders a window, not a million rows', () => {
  const geo = new SheetGeometry();
  const vp = geo.viewport({ scrollX: 0, scrollY: 0, width: 900, height: 500 });
  assert.ok(vp.rows < 60, 'a 500px viewport is tens of rows, not thousands: ' + vp.rows);
  assert.ok(vp.cols < 30);

  // scrolled far down, it still returns a small window at the right place
  const deep = geo.viewport({ scrollX: 0, scrollY: 20000, width: 900, height: 500 });
  assert.ok(deep.rows < 60);
  assert.equal(deep.firstRow, geo.rowAt(20000) - 3, 'overscan of 3');
  assert.ok(deep.firstRow > 900);
});

test('scrollToShow brings a cell into view from either direction', () => {
  const geo = new SheetGeometry();
  const below = geo.scrollToShow({ row: 100, col: 0, scrollX: 0, scrollY: 0, width: 900, height: 500 });
  assert.ok(below.scrollY > 0);
  const above = geo.scrollToShow({ row: 0, col: 0, scrollX: 0, scrollY: 5000, width: 900, height: 500 });
  assert.equal(above.scrollY, 0);
  const noMove = geo.scrollToShow({ row: 2, col: 2, scrollX: 0, scrollY: 0, width: 900, height: 500 });
  assert.deepEqual(noMove, { scrollX: 0, scrollY: 0 });
});

// ------------------------------------------------------------ selection --

test('references convert both ways', () => {
  assert.equal(colName(0), 'A');
  assert.equal(colName(25), 'Z');
  assert.equal(colName(26), 'AA');
  assert.equal(ref(0, 0), 'A1');
  assert.equal(ref(6, 1), 'B7');
});

test('arrows move, shift extends, and the anchor stays put', () => {
  const s = Selection.at(5, 5);
  s.move('right');
  assert.deepEqual(s.active, { row: 5, col: 6 });
  assert.ok(s.isSingle);

  s.move('down', { extend: true });
  s.move('right', { extend: true });
  assert.deepEqual(s.range, { top: 5, bottom: 6, left: 6, right: 7 });
  assert.equal(s.toString(), 'G6:H7');

  s.move('left');
  assert.ok(s.isSingle, 'a plain arrow collapses the selection');
});

test('selection cannot move above or left of the sheet', () => {
  const s = Selection.at(0, 0);
  s.move('up');
  s.move('left');
  assert.deepEqual(s.active, { row: 0, col: 0 });
});

test('ctrl+arrow runs to the edge of a data block', () => {
  // A1:A5 filled, A6:A9 blank, A10 filled
  const filled = new Set(['0,0', '1,0', '2,0', '3,0', '4,0', '9,0']);
  const isFilled = (r, c) => filled.has(r + ',' + c);
  const bounds = { maxRow: 100, maxCol: 10 };

  const s = Selection.at(0, 0);
  s.move('down', { jump: true, isFilled, bounds });
  assert.equal(s.active.row, 4, 'runs to the end of the filled block');

  s.move('down', { jump: true, isFilled, bounds });
  assert.equal(s.active.row, 9, 'then skips the gap to the next filled cell');

  s.move('down', { jump: true, isFilled, bounds });
  assert.equal(s.active.row, 100, 'then to the sheet edge');
});

test('tab and enter wrap inside a multi-cell selection', () => {
  const s = Selection.at(0, 0);
  s.extendTo(1, 1); // A1:B2 — the extent is B2 but the ACTIVE cell stays at A1
  assert.deepEqual(s.active, { row: 0, col: 0 }, 'dragging does not move where you type');
  assert.deepEqual(s.extent, { row: 1, col: 1 });

  s.tab();
  assert.deepEqual(s.active, { row: 0, col: 1 });
  s.tab();
  assert.deepEqual(s.active, { row: 1, col: 0 }, 'wraps to the next row of the block');
  s.tab();
  assert.deepEqual(s.active, { row: 1, col: 1 });
  s.tab();
  assert.deepEqual(s.active, { row: 0, col: 0 }, 'wraps around the block');
  assert.deepEqual(s.range, { top: 0, bottom: 1, left: 0, right: 1 }, 'the block itself never moves');
});

test('enter returns to the column the tab run started from', () => {
  const s = Selection.at(3, 1); // B4
  s.tab();
  s.tab();
  assert.deepEqual(s.active, { row: 3, col: 3 }, 'tabbed across to D4');
  s.enter();
  assert.deepEqual(s.active, { row: 4, col: 1 }, 'enter drops to B5, not D5');
});

// ----------------------------------------------------------- sheet view --

test('a real workbook opens into a renderable view', () => {
  const view = SheetView.open(FIXTURE);
  assert.deepEqual(view.sheetNames(), ['Stock', 'Summary']);
  assert.equal(view.activeSheet, 'Stock');

  const frame = view.render();
  assert.ok(frame.cells.length > 0);
  assert.ok(frame.columns.some((c) => c.name === 'A'));
  assert.ok(frame.rows.some((r) => r.label === '1'));
  assert.equal(frame.selection.active.ref, 'A1');
});

test('the grid shows formatted values, the formula bar shows the formula', () => {
  const view = SheetView.open(FIXTURE);
  // C2 is =B2*1.17 with the custom "PKR" format on style 5
  assert.equal(view.displayValue(1, 2).text, '1,661.40 PKR');
  assert.equal(view.editValue(1, 2), '=B2*1.17');

  // B2 is 1420 with #,##0
  assert.equal(view.displayValue(1, 1).text, '1,420');
  assert.equal(view.editValue(1, 1), '1420');

  view.select(1, 2);
  const frame = view.render();
  assert.equal(frame.formulaBar, '=B2*1.17', 'never the calculated value');
  const cell = frame.cells.find((c) => c.ref === 'C2');
  assert.equal(cell.text, '1,661.40 PKR', 'never the formula');
  assert.equal(cell.isFormula, true);
});

test('editing recalculates dependents immediately', () => {
  const view = SheetView.open(FIXTURE);
  view.unprotect(); // the fixture ships PROTECTED, as a customer file can
  view.select(1, 1); // B2
  view.beginEdit({ replace: true, initial: '2000' });
  view.commitEdit();

  assert.equal(view.displayValue(1, 1).text, '2,000');
  assert.equal(view.displayValue(1, 2).text, '2,340.00 PKR', 'C2 followed B2');
  assert.deepEqual(view.selection.active, { row: 2, col: 1 }, 'commit moved down');
});

test('beginEdit replaces on typing but keeps content for amendment', () => {
  const view = SheetView.open(FIXTURE);
  view.select(1, 2);
  view.beginEdit({ replace: false });
  assert.equal(view.editing.draft, '=B2*1.17', 'F2 keeps the formula to amend');

  view.cancelEdit();
  view.beginEdit({ replace: true, initial: '9' });
  assert.equal(view.editing.draft, '9', 'typing replaces');
  view.cancelEdit();
  assert.equal(view.editValue(1, 2), '=B2*1.17', 'cancel leaves the cell alone');
});

test('typed input is coerced the way a spreadsheet coerces it', () => {
  assert.equal(coerceInput('42'), 42);
  assert.equal(coerceInput('  42  '), 42);
  assert.equal(coerceInput('=A1+1'), '=A1+1');
  assert.equal(coerceInput("'0042"), '0042', 'a leading apostrophe keeps it text');
  assert.equal(coerceInput('TRUE'), true);
  assert.equal(coerceInput('SKU-1001'), 'SKU-1001');
  assert.equal(coerceInput(''), '');
});

test('the status line summarises a multi-cell selection', () => {
  const view = SheetView.open(FIXTURE);
  view.select(1, 1);
  assert.equal(view.render().status, null, 'a single cell has no summary');

  view.select(2, 1, { extend: true }); // B2:B3 = 1420, 860
  const status = view.render().status;
  assert.equal(status.count, 2);
  assert.equal(status.numeric, 2);
  assert.equal(status.sum, 2280);
  assert.equal(status.average, 1140);
  assert.equal(status.min, 860);
  assert.equal(status.max, 1420);
});

test('copy renders displayed text, paste writes typed values', () => {
  const view = SheetView.open(FIXTURE);
  view.select(1, 0);
  view.select(2, 1, { extend: true });
  assert.equal(view.copyText(), 'Steel bracket 40mm\t1,420\nSteel bracket 60mm\t860');

  view.selectSheet('Summary');
  view.select(4, 0);
  view.pasteText('SKU-9\t99\nSKU-8\t88');
  assert.equal(view.displayValue(4, 0).text, 'SKU-9');
  assert.equal(view.displayValue(5, 1).text, '88');
  assert.deepEqual(view.selection.range, { top: 4, bottom: 5, left: 0, right: 1 });
});

test('saving goes through the preserving path and caches formula results', () => {
  const view = SheetView.open(FIXTURE);
  view.unprotect(); // editing a protected sheet now refuses, as Excel does
  view.select(1, 1);
  view.beginEdit({ replace: true, initial: '2000' });
  view.commitEdit();
  view.select(0, 7);
  view.beginEdit({ replace: true, initial: '=SUM(B2:B3)' });
  view.commitEdit();

  assert.ok(view.isDirty);
  const out = view.save();
  assert.ok(!view.isDirty);

  const diff = comparePackages(FIXTURE, out);
  assert.deepEqual(diff.changed.map((c) => c.name), ['xl/worksheets/sheet1.xml']);
  assert.equal(diff.removed.length, 0);
  for (const part of FRAGILE_PARTS) assert.ok(diff.identical.includes(part), part + ' disturbed by a save');

  const sheetXml = OoxmlPackage.read(out).text('xl/worksheets/sheet1.xml');
  for (const el of FRAGILE_SHEET_ELEMENTS) assert.ok(sheetXml.includes('<' + el), el + ' lost on save');
  // The unprotect above was a DELIBERATE removal, not a preservation failure.
  assert.ok(!sheetXml.includes('<sheetProtection'), 'unprotect removed exactly the element it names');
  assert.match(sheetXml, /<c r="B2" s="4"><v>2000<\/v><\/c>/, 'the edited cell kept its style');
  assert.match(sheetXml, /<c r="H1"[^>]*><f>SUM\(B2:B3\)<\/f><v>2860<\/v><\/c>/, 'the new formula was cached');
});

test('a saved view reopens with the same values', () => {
  const view = SheetView.open(FIXTURE);
  view.unprotect();
  view.select(1, 1);
  view.beginEdit({ replace: true, initial: '2000' });
  view.commitEdit();
  const reopened = SheetView.open(view.save());
  assert.equal(reopened.displayValue(1, 1).text, '2,000');
  assert.equal(reopened.displayValue(1, 2).text, '2,340.00 PKR');
  assert.equal(reopened.editValue(1, 2), '=B2*1.17');
});

test('the cursor can move past the end of the data', () => {
  // Found by driving the real shell: navigation had been bounded by the USED
  // range, so a sheet with data in A:F refused to let the cursor reach G.
  const view = SheetView.open(FIXTURE);
  view.unprotect();
  const used = view.bounds;
  assert.ok(used.maxCol < 10, 'the fixture only uses a few columns');

  view.select(0, 0);
  for (let i = 0; i < 12; i++) view.moveSelection('right');
  assert.deepEqual(view.selection.active, { row: 0, col: 12 }, 'stopped at the edge of the data');

  for (let i = 0; i < 30; i++) view.moveSelection('down');
  assert.equal(view.selection.active.row, 30);

  // and typing out there works
  view.setCell(30, 12, 'new territory');
  assert.equal(view.displayValue(30, 12).text, 'new territory');
});

test('switching sheets resets the view to that sheet', () => {
  const view = SheetView.open(FIXTURE);
  view.select(5, 5);
  view.selectSheet('Summary');
  assert.equal(view.activeSheet, 'Summary');
  assert.deepEqual(view.selection.active, { row: 0, col: 0 });
  assert.equal(view.displayValue(0, 0).text, 'Summary');
  assert.throws(() => view.selectSheet('Nope'), /no such sheet/);
});

// ------------------------------------------------------------- drawings --

test('a sheet renders its chart, its shape and its picture', () => {
  const view = SheetView.open(FIXTURE, { viewportWidth: 1600, viewportHeight: 1400 });
  const drawings = view.render().drawings;
  assert.deepEqual(drawings.map((d) => d.kind), ['chart', 'shape', 'image']);

  // all three produce SVG — none falls back to the "preserved but not drawn" box
  assert.ok(drawings.every((d) => d.svg && !d.unsupported), JSON.stringify(drawings.map((d) => d.unsupported)));
  assert.match(drawings[0].svg, /Stock on hand/);
  assert.match(drawings[1].svg, /Confirm with supplier/);
  assert.match(drawings[2].svg, /<image[^>]*href="data:image\/png;base64,/);

  // and each is positioned and sized from its anchor, in grid pixels
  for (const d of drawings) {
    assert.ok(Number.isFinite(d.x) && Number.isFinite(d.y));
    assert.ok(d.width > 0 && d.height > 0);
  }
});

test('a drawing outside the viewport is not rendered', () => {
  const small = SheetView.open(FIXTURE, { viewportWidth: 200, viewportHeight: 120 });
  const wide = SheetView.open(FIXTURE, { viewportWidth: 1600, viewportHeight: 1400 });
  assert.ok(small.render().drawings.length < wide.render().drawings.length);
});

test('drawing parts survive a save untouched', () => {
  // We re-draw a chart in our own palette but never rewrite its part, so a file
  // that round-trips through us keeps the customer's own styling.
  const view = SheetView.open(FIXTURE, { viewportWidth: 1600, viewportHeight: 1400 });
  view.unprotect();
  view.setCell(1, 0, '999');
  const saved = view.save();
  const diff = comparePackages(FIXTURE, saved);
  for (const name of ['xl/drawings/drawing1.xml', 'xl/charts/chart1.xml', 'xl/media/image1.png']) {
    assert.ok(!diff.changed.includes(name), name + ' was rewritten');
  }
});


// ---------------------------------------------------------------------------
// Cell appearance — why a real workbook stopped rendering as plain text.
// ---------------------------------------------------------------------------

test('theme colour INDICES are not the theme document order', () => {
  // The single most common way themed colours come out wrong: <a:clrScheme>
  // lists dk1 first, but index 0 is lt1. Get it backwards and every themed
  // header is black on white or white on white.
  const pkg = OoxmlPackage.read(FIXTURE);
  const theme = readTheme(pkg);
  assert.equal(THEME_SLOTS[0], 'lt1');
  assert.equal(THEME_SLOTS[1], 'dk1');
  assert.equal(theme[0], '#ffffff', 'index 0 is the light background');
  assert.equal(theme[1], '#000000', 'index 1 is the dark text');
  assert.equal(theme[4], '#4472c4', 'index 4 is accent1');
});

test('tint moves luminance, it does not mix with white', () => {
  // These are the values Excel itself produces for "Blue, Accent 1, Darker 50%"
  // and "Lighter 80%". A linear blend toward white gives neither.
  assert.equal(applyTint('#4472c4', -0.5), '#203864');
  assert.equal(applyTint('#4472c4', 0.8), '#dae3f3');
  assert.equal(applyTint('#4472c4', 0), '#4472c4');
  assert.equal(applyTint(null, 0.5), null);
});

test('the system colours resolve to nothing rather than to a guess', () => {
  // indexed 64/65 are "whatever the reader's ink is". Guessing black would put
  // black text on a black fill in a dark theme.
  assert.equal(readColourElement('<color indexed="64"/>'), null);
  assert.equal(readColourElement('<color indexed="65"/>'), null);
  assert.equal(readColourElement('<color indexed="10"/>'), '#ff0000');
  assert.equal(readColourElement('<color rgb="FF1F5F8B"/>'), '#1f5f8b');
  assert.equal(readColourElement(null), null);
});

test('a style index carries the font, fill, border and alignment together', () => {
  const pkg = OoxmlPackage.read(FIXTURE);
  const { byStyleIndex } = readStyles(pkg, { BUILTIN_FORMATS });

  const header = byStyleIndex[2];
  assert.equal(header.font.bold, true);
  assert.equal(header.fill.colour, '#d9e1f2');
  assert.equal(header.border.top.widthPx, 1);
  assert.equal(header.border.bottom.colour, '#7f7f7f');

  const banner = byStyleIndex[6];
  assert.equal(banner.font.sizePt, 14);
  assert.equal(banner.font.colour, '#ffffff', 'theme 0 is the light background');
  assert.equal(banner.fill.colour, '#4472c4', 'theme 4 is accent1');
  assert.equal(banner.align.horizontal, 'center');
  assert.equal(banner.align.vertical, 'center');

  const note = byStyleIndex[7];
  assert.equal(note.fill.colour, '#dae3f3', 'accent1 at tint 0.8');
  assert.equal(note.align.wrap, true);
  assert.equal(note.align.indent, 1);
  assert.equal(note.border.bottom.style, 'double');
});

test('applyFont="0" means inherit, which is not the same as font zero', () => {
  const pkg = OoxmlPackage.read(FIXTURE);
  const { byStyleIndex } = readStyles(pkg, { BUILTIN_FORMATS });
  assert.equal(byStyleIndex[8].font, null, 'the fontId is present but not applied');
});

test('patternType none is no fill, and gray125 is not a solid', () => {
  const pkg = OoxmlPackage.read(FIXTURE);
  const { byStyleIndex } = readStyles(pkg, { BUILTIN_FORMATS });
  assert.equal(byStyleIndex[0].fill, null);
  assert.equal(byStyleIndex[4].fill, null, 'a bordered cell with no fill has none');
});

test('a merged range is drawn once, by its top-left cell', () => {
  const view = SheetView.open(FIXTURE, { viewportWidth: 1400, viewportHeight: 900 });
  const cells = view.render().cells;

  const banner = cells.find((c) => c.ref === 'A20');
  assert.ok(banner, 'the top-left cell of the merge is there');
  assert.equal(banner.merged.ref, 'A20:D20');
  assert.equal(banner.merged.cols, 4);
  assert.equal(banner.text, 'Warehouse stock — WH1');

  // the covered cells are not empty — they are not present at all, or the grid
  // lines would run through the middle of the title
  for (const ref of ['B20', 'C20', 'D20']) {
    assert.equal(cells.find((c) => c.ref === ref), undefined, ref + ' should not be drawn');
  }

  // and it is as wide as the columns it spans
  const geo = view.geo;
  const expected = geo.colOffset(3) + geo.colWidth(3) - geo.colOffset(0);
  assert.equal(banner.width, expected);
});

test('a merge spanning rows is as tall as the rows it covers', () => {
  const view = SheetView.open(FIXTURE, { viewportWidth: 1400, viewportHeight: 900 });
  const note = view.render().cells.find((c) => c.ref === 'A22');
  assert.equal(note.merged.rows, 2);
  assert.equal(note.height, view.geo.rowHeight(21) + view.geo.rowHeight(22));
  assert.equal(view.render().cells.find((c) => c.ref === 'A23'), undefined);
});

test('a styled but empty cell is still drawn', () => {
  // A shaded header with nothing typed in it is a real thing; skipping it leaves
  // a hole in the band.
  const view = SheetView.open(FIXTURE, { viewportWidth: 1400, viewportHeight: 900 });
  const drawn = view.render().cells.filter((c) => c.text === '' && c.style);
  assert.ok(drawn.every((c) => c.style.fill || c.style.border),
    'only decorated empty cells are drawn, not every blank on the sheet');
});

test('the file\'s own alignment beats the number format\'s default', () => {
  const view = SheetView.open(FIXTURE, { viewportWidth: 1400, viewportHeight: 900 });
  const cells = view.render().cells;
  // a number would right-align by default; the banner says centre
  assert.equal(cells.find((c) => c.ref === 'A20').align, 'center');
  // and a number with nothing said about it still right-aligns
  assert.equal(cells.find((c) => c.ref === 'B2').align, 'right');
});

test('reading styles rewrites nothing', () => {
  const view = SheetView.open(FIXTURE, { viewportWidth: 1400, viewportHeight: 900 });
  view.render();
  assert.deepEqual(comparePackages(FIXTURE, view.save()).changed, []);
});

test('an edited cell keeps its style index', () => {
  // The reason editing a formatted workbook is safe: we never write styles.xml,
  // and the cell keeps pointing at the same entry.
  const view = SheetView.open(FIXTURE, { viewportWidth: 1400, viewportHeight: 900 });
  view.unprotect();
  const before = view.styleFor(0, 0);
  view.select(0, 0);
  view.setCell(0, 0, 'Renamed');
  assert.deepEqual(view.styleFor(0, 0), before);

  const saved = view.save();
  const reopened = SheetView.open(saved, { viewportWidth: 1400, viewportHeight: 900 });
  assert.deepEqual(reopened.styleFor(0, 0), before);
  assert.equal(reopened.displayValue(0, 0).text, 'Renamed');
  assert.ok(!comparePackages(FIXTURE, saved).changed.includes('xl/styles.xml'));
});

test('merged ranges are read as ranges, not resolved into a grid', () => {
  const merges = readMergedCells('<worksheet><mergeCells count="2">'
    + '<mergeCell ref="A1:D1"/><mergeCell ref="B3:B6"/></mergeCells></worksheet>');
  assert.deepEqual(merges, [
    { top: 0, left: 0, bottom: 0, right: 3, ref: 'A1:D1' },
    { top: 2, left: 1, bottom: 5, right: 1, ref: 'B3:B6' },
  ]);
  assert.deepEqual(readMergedCells('<worksheet/>'), []);
  assert.deepEqual(readMergedCells(null), []);
});

// ------------------------------------------- autosum, find and the spread --

const simpleBook = (rows) => buildXlsx({ sheets: [{ name: 'Data', rows }] });

test('AutoSum infers the run above, then the run to the left', () => {
  const view = SheetView.open(simpleBook([[1], [2], [3]]));
  view.select(3, 0);
  view.autoSum('SUM');
  assert.equal(view.editValue(3, 0), '=SUM(A1:A3)');
  assert.equal(view.calc.getValue('Data', 3, 0), 6);

  const wide = SheetView.open(simpleBook([[10, 20, 30]]));
  wide.select(0, 3);
  wide.autoSum('AVERAGE');
  assert.equal(wide.editValue(0, 3), '=AVERAGE(A1:C1)');
  assert.equal(wide.calc.getValue('Data', 0, 3), 20);
});

test('AutoSum with nothing adjacent to sum is a no-op, not =SUM()', () => {
  const view = SheetView.open(simpleBook([[1]]));
  view.select(5, 5);
  view.autoSum('SUM');
  assert.equal(view.editValue(5, 5), '');
  assert.equal(view.isDirty, false);
});

test('AutoSum over a selection totals each column below it, as one undo step', () => {
  const view = SheetView.open(simpleBook([[1, 2], [3, 4]]));
  view.select(0, 0);
  view.select(1, 1, { extend: true });
  view.autoSum('SUM');
  assert.equal(view.editValue(2, 0), '=SUM(A1:A2)');
  assert.equal(view.editValue(2, 1), '=SUM(B1:B2)');
  assert.equal(view.calc.getValue('Data', 2, 0), 4);
  assert.equal(view.calc.getValue('Data', 2, 1), 6);

  view.undo();
  assert.equal(view.editValue(2, 0), '', 'both formulas leave in one undo');
  assert.equal(view.editValue(2, 1), '');
});

test('find walks forward from the cursor and wraps; a miss reports itself', () => {
  const view = SheetView.open(simpleBook([['alpha', 'beta'], ['gamma', 'alpha']]));
  assert.equal(view.findNext('alpha'), true);
  assert.equal(view.selection.active.ref ?? ref(view.selection.active.row, view.selection.active.col), 'B2');
  assert.equal(view.findNext('alpha'), true);
  assert.equal(ref(view.selection.active.row, view.selection.active.col), 'A1', 'wraps past the end');
  assert.equal(view.findNext('nothing-here'), false);
  assert.equal(view.findNext(''), false, 'an empty needle matches nothing');
});

test('replace edits the input — a formula stays a formula — and replace-all is one undo', () => {
  const view = SheetView.open(simpleBook([['cost', 10], ['cost plus', '=B1*2']]));
  const replaced = view.replaceAll('cost', 'price');
  assert.equal(replaced, 2);
  assert.equal(view.displayValue(0, 0).text, 'price');
  assert.equal(view.displayValue(1, 0).text, 'price plus');
  view.undo();
  assert.equal(view.displayValue(0, 0).text, 'cost', 'both cells return in one step');

  // Replacing inside a formula's text re-parses it rather than freezing its value.
  view.select(0, 0);
  const b2 = SheetView.open(simpleBook([[5, '=A1*2']]));
  b2.replaceAll('A1', 'A1+1');
  assert.equal(b2.editValue(0, 1), '=A1+1*2');
  assert.equal(b2.calc.getValue('Data', 0, 1), 7);
});

test('replaceNext replaces the current match and moves to the next', () => {
  const view = SheetView.open(simpleBook([['aa', 'bb'], ['aa', 'cc']]));
  view.select(0, 0);
  assert.equal(view.replaceNext('aa', 'zz'), true);
  assert.equal(view.displayValue(0, 0).text, 'zz');
  assert.equal(ref(view.selection.active.row, view.selection.active.col), 'A2', 'cursor moved to the next match');
});

test('fill continues a numeric series and copies a lone value', () => {
  const view = SheetView.open(simpleBook([[1], [2]]));
  view.select(0, 0);
  view.select(1, 0, { extend: true });
  view.fill({ top: 0, left: 0, bottom: 4, right: 0 });
  assert.equal(view.calc.getValue('Data', 2, 0), 3);
  assert.equal(view.calc.getValue('Data', 3, 0), 4);
  assert.equal(view.calc.getValue('Data', 4, 0), 5);
  assert.deepEqual(view.selection.range, { top: 0, bottom: 4, left: 0, right: 0 }, 'the selection becomes the target');

  const lone = SheetView.open(simpleBook([[7]]));
  lone.select(0, 0);
  lone.fill({ top: 0, left: 0, bottom: 2, right: 0 });
  assert.equal(lone.calc.getValue('Data', 1, 0), 7, 'a single number copies');
  assert.equal(lone.calc.getValue('Data', 2, 0), 7);
});

test('fill shifts a formula to where it lands, in every direction', () => {
  const view = SheetView.open(simpleBook([[10, 20, '=A1*2']]));
  view.select(0, 2);
  view.fill({ top: 0, left: 2, bottom: 2, right: 2 });
  assert.equal(view.editValue(1, 2), '=A2*2');
  assert.equal(view.editValue(2, 2), '=A3*2');

  const right = SheetView.open(simpleBook([[10], [20], ['=A1+A2']]));
  right.select(2, 0);
  right.fill({ top: 2, left: 0, bottom: 2, right: 2 });
  assert.equal(right.editValue(2, 1), '=B1+B2');
  assert.equal(right.editValue(2, 2), '=C1+C2');
});

test('fill increments text ending in a number, cycles mixed sources, and undoes as one step', () => {
  const view = SheetView.open(simpleBook([['Item 1']]));
  view.select(0, 0);
  view.fill({ top: 0, left: 0, bottom: 2, right: 0 });
  assert.equal(view.displayValue(1, 0).text, 'Item 2');
  assert.equal(view.displayValue(2, 0).text, 'Item 3');

  view.undo();
  assert.equal(view.displayValue(1, 0).text, '', 'the whole fill leaves in one undo');
  assert.equal(view.displayValue(2, 0).text, '');

  const mixed = SheetView.open(simpleBook([['red'], ['blue']]));
  mixed.select(0, 0);
  mixed.select(1, 0, { extend: true });
  mixed.fill({ top: 0, left: 0, bottom: 3, right: 0 });
  assert.equal(mixed.displayValue(2, 0).text, 'red', 'text cycles');
  assert.equal(mixed.displayValue(3, 0).text, 'blue');
});

test('fill refuses a target that is not a one-axis extension', () => {
  const view = SheetView.open(simpleBook([[1]]));
  view.select(0, 0);
  assert.throws(() => view.fill({ top: 0, left: 0, bottom: 2, right: 2 }), /one axis/);
  assert.throws(() => view.fill({ top: -1, left: 0, bottom: 0, right: 0 }), /bad fill target/);
});

test('fill upward continues the series backwards', () => {
  const view = SheetView.open(simpleBook([[], [], [10], [20]]));
  view.select(2, 0);
  view.select(3, 0, { extend: true });
  view.fill({ top: 0, left: 0, bottom: 3, right: 0 });
  assert.equal(view.calc.getValue('Data', 1, 0), 0, 'one step back from 10');
  assert.equal(view.calc.getValue('Data', 0, 0), -10);
});

test('a paste of our own copy is rich: formulas shift and formatting travels', () => {
  const view = SheetView.open(simpleBook([[1, '=A1*10'], [2, '=A2*10']]));
  view.select(0, 0);
  view.select(1, 1, { extend: true });
  view.setFormat({ bold: true });
  view.markClipboard();

  view.select(4, 2); // paste anchor C5
  view.pasteText(view.clipboard.text);
  assert.equal(view.calc.getValue('Data', 4, 2), 1, 'the value came');
  assert.equal(view.editValue(4, 3), '=C5*10', 'the formula shifted to where it landed');
  assert.equal(view.editValue(5, 3), '=C6*10');
  assert.equal(
    view._styleIndexAt('Data', 4, 2),
    view._styleIndexAt('Data', 0, 0),
    'the style index travelled',
  );
  assert.deepEqual(view.selection.range, { top: 4, left: 2, bottom: 5, right: 3 }, 'the paste is selected');

  view.undo();
  assert.equal(view.editValue(4, 2), '', 'one undo step');
  assert.equal(view.editValue(5, 3), '');
});

test('a paste whose text is not our copy stays plain', () => {
  const view = SheetView.open(simpleBook([[1, 2]]));
  view.select(0, 0);
  view.select(0, 1, { extend: true });
  view.markClipboard();

  view.select(3, 0);
  view.pasteText('something\telse');
  assert.equal(view.displayValue(3, 0).text, 'something');
  assert.equal(view.displayValue(3, 1).text, 'else', 'tab-separated text pastes as cells');
});

test('header selection covers the used range, and shift-extends', () => {
  const view = SheetView.open(simpleBook([[1, 2, 3], [4, 5, 6]]));
  view.selectColumn(1);
  assert.deepEqual(view.selection.range, { top: 0, bottom: 1, left: 1, right: 1 });
  view.selectColumn(2, { extend: true });
  assert.deepEqual(view.selection.range, { top: 0, bottom: 1, left: 1, right: 2 });
  view.selectRow(0);
  assert.deepEqual(view.selection.range, { top: 0, bottom: 0, left: 0, right: 2 });
});

test('sort orders rows by the active column, blanks sinking in both directions', () => {
  const view = SheetView.open(simpleBook([['b', 20], ['', 99], ['a', 10], ['c', 30]]));
  view.select(0, 0);
  view.select(3, 1, { extend: true });
  view.sortSelection({ ascending: true });
  assert.equal(view.displayValue(0, 0).text, 'a');
  assert.equal(view.displayValue(1, 0).text, 'b');
  assert.equal(view.displayValue(2, 0).text, 'c');
  assert.equal(view.displayValue(3, 1).text, '99', 'the blank-keyed row sank');
  assert.equal(view.displayValue(0, 1).text, '10', 'the row moved together');

  view.select(0, 0);
  view.select(3, 1, { extend: true });
  view.sortSelection({ ascending: false });
  assert.equal(view.displayValue(0, 0).text, 'c');
  assert.equal(view.displayValue(2, 0).text, 'a');
  assert.equal(view.displayValue(3, 1).text, '99', 'blanks sink descending too');

  view.undo();
  assert.equal(view.displayValue(0, 0).text, 'a', 'one undo step per sort');
});

test('a single cell expands to its data block and skips a detected header', () => {
  const view = SheetView.open(simpleBook([['name', 'qty'], ['pear', 5], ['apple', 2]]));
  view.select(1, 0);
  view.sortSelection({ ascending: true });
  assert.equal(view.displayValue(0, 0).text, 'name', 'the header stayed put');
  assert.equal(view.displayValue(1, 0).text, 'apple');
  assert.equal(view.displayValue(2, 0).text, 'pear');
  assert.equal(view.displayValue(1, 1).text, '2', 'rows moved whole');
});

test('sort shifts formulas with their rows', () => {
  const view = SheetView.open(simpleBook([['b', '=A1&"!"'], ['a', '=A2&"!"']]));
  view.select(0, 0);
  view.select(1, 1, { extend: true });
  view.sortSelection({ ascending: true });
  assert.equal(view.displayValue(0, 0).text, 'a');
  assert.equal(view.editValue(0, 1), '=A1&"!"', 'the moved formula still points at its own row');
  assert.equal(view.displayValue(0, 1).text, 'a!');
  assert.equal(view.displayValue(1, 1).text, 'b!');
});

test('Goal Seek finds the input that reaches the target, as one undoable edit', () => {
  // B1 = A1 * 3 + 10; what must A1 be for B1 to reach 25?
  const view = SheetView.open(simpleBook([[2, '=A1*3+10']]));
  const { value } = view.goalSeek({ set: 'B1', to: 25, by: 'A1' });
  assert.equal(Math.abs(value - 5) < 1e-6, true, 'found 5');
  assert.equal(view.calc.getValue('Data', 0, 0), value);
  assert.equal(Math.abs(view.calc.getValue('Data', 0, 1) - 25) < 1e-6, true);

  view.undo();
  assert.equal(view.calc.getValue('Data', 0, 0), 2, 'one undo restores the original input');
  assert.equal(view.calc.getValue('Data', 0, 1), 16);
});

test('Goal Seek refuses what it cannot honestly do, changing nothing', () => {
  const view = SheetView.open(simpleBook([[2, '=A1*3']]));
  assert.throws(() => view.goalSeek({ set: 'A1', to: 5, by: 'A1' }), /must hold a formula/);
  assert.throws(() => view.goalSeek({ set: 'B1', to: 5, by: 'B1' }), /plain value/);
  assert.throws(() => view.goalSeek({ set: 'nope', to: 5, by: 'A1' }), /needs a cell/);
  // A target the formula never reaches: B2 ignores A1 entirely.
  const flat = SheetView.open(simpleBook([[2, '=5+0']]));
  assert.throws(() => flat.goalSeek({ set: 'B1', to: 99, by: 'A1' }), /could not find/);
  assert.equal(flat.calc.getValue('Data', 0, 0), 2, 'the changing cell was restored');
  assert.equal(flat.isDirty, false, 'a failed seek leaves no unsaved mark');
});

test('the format painter tiles the picked-up look across the target', () => {
  const view = SheetView.open(simpleBook([[1, 2], [3, 4], [5, 6], [7, 8]]));
  view.select(0, 0);
  view.setFormat({ bold: true });
  view.select(0, 1);
  view.setFormat({ italic: true });

  // Pick up the two-cell pattern of row 1, paint it over rows 3-4 (2×2).
  view.select(0, 0);
  view.select(0, 1, { extend: true });
  view.markFormatBrush();
  view.select(2, 0);
  view.select(3, 1, { extend: true });
  view.paintFormat();

  const boldIdx = view._styleIndexAt('Data', 0, 0);
  const italicIdx = view._styleIndexAt('Data', 0, 1);
  assert.equal(view._styleIndexAt('Data', 2, 0), boldIdx, 'pattern column 1');
  assert.equal(view._styleIndexAt('Data', 2, 1), italicIdx, 'pattern column 2');
  assert.equal(view._styleIndexAt('Data', 3, 0), boldIdx, 'tiled down');
  assert.equal(view._styleIndexAt('Data', 3, 1), italicIdx);

  view.undo();
  assert.equal(view._styleIndexAt('Data', 2, 0), null, 'one undo step clears the paint');
  assert.equal(view._styleIndexAt('Data', 3, 1), null);

  // Painting with no brush, or one whose sheet is gone, is a quiet no-op.
  const bare = SheetView.open(simpleBook([[1]]));
  bare.paintFormat();
  assert.equal(bare.isDirty, false);
});

test('the canvas total spreads with the selection and the scroll position', () => {
  const view = SheetView.open(simpleBook([[1, 2], [3, 4]]));
  const before = view.render().total;

  view.select(120, 20);
  const withSelection = view.render().total;
  assert.ok(withSelection.height > before.height, 'selecting far below grows the sheet');
  assert.ok(withSelection.width > before.width, 'selecting far right widens it');

  view.select(0, 0);
  view.scrollTo(0, withSelection.height + 500);
  const scrolled = view.render().total;
  assert.ok(scrolled.height > withSelection.height, 'scrolling toward the edge keeps extending it');
});

// -------------------------------------------------------------- defined names --

test('define name: the selection gets a name, formulas find it, undo removes it', () => {
  const view = SheetView.open(simpleBook([[10], [20], [30]]));
  view.select(0, 0);
  view.select(2, 0, { extend: true });
  view.defineName('Sales');

  const names = view.render().names;
  assert.equal(names.length, 1);
  assert.equal(names[0].name, 'Sales');
  assert.equal(names[0].ref, 'Data!$A$1:$A$3', 'the selection travels as a sheet-qualified absolute ref');

  view.select(0, 2);
  view.setCell(0, 2, '=SUM(Sales)');
  assert.equal(view.calc.getValue('Data', 0, 2), 60);

  assert.ok(view.isDirty, 'a defined name is a real change to save');
  view.undo(); // the formula
  view.undo(); // the name
  assert.equal(view.render().names.length, 0, 'undo restores workbook.xml without the name');
  assert.equal(view.calc.names.size, 0, 'the calc model let go of it too');

  view.redo();
  assert.equal(view.render().names.length, 1, 'redo brings it back');
  view.setCell(0, 3, '=SUM(Sales)');
  assert.equal(view.calc.getValue('Data', 0, 3), 60, 'and a fresh formula finds it again');
});

test('define name refuses what Excel refuses, without recording an undo step', () => {
  const view = SheetView.open(simpleBook([[1]]));
  assert.throws(() => view.defineName('my name'), /not a valid name/);
  assert.throws(() => view.defineName('A1'), /cell reference/);
  assert.throws(() => view.defineName(''), /not a valid name/);
  assert.throws(() => view.defineName('Good', 'not a ref'), /is not a range/);
  assert.equal(view.canUndo, false, 'a refused edit leaves no history');
  assert.equal(view.isDirty, false);
});

test('delete name: the name goes, dependents show #NAME?, and it survives a save round trip', () => {
  const view = SheetView.open(simpleBook([[5], [6]]));
  view.select(0, 0);
  view.select(1, 0, { extend: true });
  view.defineName('Pair');
  view.setCell(0, 2, '=SUM(Pair)');
  assert.equal(view.calc.getValue('Data', 0, 2), 11);

  // The saved file carries the name — a reopened view still resolves it.
  const reopened = SheetView.open(view.save());
  assert.equal(reopened.render().names[0].name, 'Pair');
  assert.equal(reopened.calc.getValue('Data', 0, 2), 11);

  reopened.deleteName('Pair');
  assert.equal(reopened.render().names.length, 0);
  reopened.setCell(1, 2, '=SUM(Pair)');
  assert.equal(reopened.calc.getValue('Data', 1, 2).type, '#NAME?');
  assert.throws(() => reopened.deleteName('Pair'), /no defined name/);
});

test('goto name selects the target range, crossing sheets when it must', () => {
  const buf = buildXlsx({
    sheets: [
      { name: 'Front', rows: [[1]] },
      { name: 'Back', rows: [[1, 2], [3, 4]] },
    ],
    definedNames: [{ name: 'There', ref: 'Back!$A$1:$B$2' }],
  });
  const view = SheetView.open(buf);
  assert.equal(view.activeSheet, 'Front');
  view.gotoName('there'); // case-insensitive, as Excel is
  assert.equal(view.activeSheet, 'Back');
  assert.equal(view.selection.toString(), 'A1:B2');
  assert.throws(() => view.gotoName('Nowhere'), /no defined name/);
});

// ------------------------------------------------------------ data validation --

const ruledBook = () => buildXlsx({
  sheets: [{
    name: 'Data',
    rows: [['Yes', 1], [null, 2]],
    validations: [
      { type: 'list', sqref: 'A1:A5', formula1: '"Yes,No,Maybe"', promptTitle: 'Answer', prompt: 'Yes, No or Maybe' },
      { type: 'whole', operator: 'between', sqref: 'B1:B5', formula1: '1', formula2: '10',
        errorTitle: 'Score', error: 'a whole number from 1 to 10' },
    ],
  }],
});

test('data validation judges direct entry the way its author wrote it', () => {
  const view = SheetView.open(ruledBook());

  // The list rule, membership case-insensitive as Excel compares text.
  assert.equal(view.checkValidation(0, 0, 'Maybe'), null);
  assert.equal(view.checkValidation(0, 0, 'no'), null);
  assert.match(view.checkValidation(0, 0, 'perhaps'), /Yes, No, Maybe/);

  // The whole-number rule, refusing in the author's own words.
  assert.equal(view.checkValidation(0, 1, '5'), null);
  assert.equal(view.checkValidation(0, 1, '11'), 'Score: a whole number from 1 to 10');
  assert.match(view.checkValidation(0, 1, '2.5'), /whole number/);
  assert.match(view.checkValidation(0, 1, 'abc'), /Score/);

  // A blank is allowed (the rule said allowBlank), a formula is judged by its
  // RESULT at entry — later recalculation is not re-checked, as in Excel.
  assert.equal(view.checkValidation(0, 1, ''), null);
  assert.equal(view.checkValidation(0, 1, '=2+3'), null);
  assert.match(view.checkValidation(0, 1, '=20*2'), /Score/);

  // Outside every ruled range, anything goes.
  assert.equal(view.checkValidation(0, 3, 'whatever'), null);
});

test('the frame carries the active cell rule: the list, and the author prompt', () => {
  const view = SheetView.open(ruledBook());
  view.select(0, 0);
  let v = view.render().validation;
  assert.deepEqual(v.list, ['Yes', 'No', 'Maybe']);
  assert.equal(v.promptTitle, 'Answer');

  view.select(0, 1);
  v = view.render().validation;
  assert.equal(v.type, 'whole');
  assert.equal(v.list, null, 'a numeric rule offers no dropdown');

  view.select(9, 9);
  assert.equal(view.render().validation, null);
});

test('a list fed by a RANGE reads live cell values, and rules survive a save', () => {
  const buf = buildXlsx({
    sheets: [{
      name: 'Data',
      rows: [['x', 'Red'], ['y', 'Green'], ['z', 'Blue']],
      validations: [{ type: 'list', sqref: 'A1:A3', formula1: '$B$1:$B$3' }],
    }],
  });
  const view = SheetView.open(buf);
  view.select(0, 0);
  assert.deepEqual(view.render().validation.list, ['Red', 'Green', 'Blue']);
  assert.equal(view.checkValidation(0, 0, 'Green'), null);
  assert.match(view.checkValidation(0, 0, 'Purple'), /Red, Green, Blue/);

  // Edit the source list and the rule follows it — it reads the live model.
  view.setCell(1, 1, 'Emerald');
  assert.equal(view.checkValidation(0, 0, 'Emerald'), null);
  assert.match(view.checkValidation(0, 0, 'Green'), /Red, Emerald, Blue/);

  // The block rides the round trip untouched; a reopened view still enforces.
  const reopened = SheetView.open(view.save());
  assert.match(reopened.checkValidation(0, 0, 'Purple') ?? '', /Red/);
});

// ---------------------------------------------------- conditional formatting --

test('readConditionalFormatting reads the rule vocabulary, sorted by priority', async () => {
  const { readConditionalFormatting } = await import('@rutba/sheet-view');
  const xml = '<worksheet><sheetData/>'
    + '<conditionalFormatting sqref="B2:B9 D4">'
    + '<cfRule type="cellIs" dxfId="1" priority="2" operator="between"><formula>5</formula><formula>$E$1</formula></cfRule>'
    + '<cfRule type="colorScale" priority="1"><colorScale>'
    + '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>'
    + '<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/>'
    + '</colorScale></cfRule>'
    + '</conditionalFormatting>'
    + '<conditionalFormatting sqref="A1:A3">'
    + '<cfRule type="containsText" dxfId="0" priority="3" text="due" stopIfTrue="1"/>'
    + '</conditionalFormatting>'
    + '</worksheet>';
  const rules = readConditionalFormatting(xml);
  assert.equal(rules.length, 3);
  assert.equal(rules[0].type, 'colorScale', 'priority 1 sorts first');
  assert.deepEqual(rules[0].colours, ['#f8696b', '#ffeb84', '#63be7b']);
  assert.equal(rules[0].cfvos[1].type, 'percentile');
  assert.equal(rules[1].type, 'cellIs');
  assert.deepEqual(rules[1].formulas, ['5', '$E$1']);
  assert.equal(rules[1].ranges.length, 2, 'sqref can name several ranges');
  assert.deepEqual(rules[1].ranges[1], { top: 3, left: 3, bottom: 3, right: 3 });
  assert.equal(rules[2].text, 'due');
  assert.equal(rules[2].stopIfTrue, true);
});

test('a colour scale and a data bar paint from the file, straight to the frame', () => {
  const wb = Workbook.open(simpleBook([[10, 10], [50, 50], [90, 90]]));
  const spliced = wb.pkg.text('xl/worksheets/sheet1.xml').replace('</sheetData>', '</sheetData>'
    + '<conditionalFormatting sqref="A1:A3"><cfRule type="colorScale" priority="1"><colorScale>'
    + '<cfvo type="min"/><cfvo type="max"/>'
    + '<color rgb="FF000000"/><color rgb="FF64C864"/>'
    + '</colorScale></cfRule></conditionalFormatting>'
    + '<conditionalFormatting sqref="B1:B3"><cfRule type="dataBar" priority="2"><dataBar>'
    + '<cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/>'
    + '</dataBar></cfRule></conditionalFormatting>');
  wb.pkg.write_('xl/worksheets/sheet1.xml', spliced);
  const view = SheetView.open(wb.save());

  const cells = new Map(view.render().cells.map((c) => [c.ref, c]));
  assert.equal(cells.get('A1').style.fill.colour, '#000000', 'the minimum wears the first stop');
  assert.equal(cells.get('A3').style.fill.colour, '#64c864', 'the maximum wears the last');
  assert.equal(cells.get('A2').style.fill.colour, '#326432', 'the middle interpolates');

  assert.equal(cells.get('B1').bar.fraction, 0);
  assert.equal(cells.get('B3').bar.fraction, 1);
  assert.equal(cells.get('B2').bar.fraction, 0.5);
  assert.equal(cells.get('B2').bar.colour, '#638ec6');

  // Painting is a reading: the parts are untouched, so fidelity holds.
  assert.equal(view.isDirty, false);
});

test('cellIs, expression, top10 and duplicates fire through dxf looks', () => {
  const view = SheetView.open(simpleBook([[1, 'x'], [8, 'y'], [12, 'x'], [20, 'z']]));
  view.styles.dxfs = [
    { font: { bold: false, italic: false, strike: false, colour: '#9c0006' }, fill: { pattern: 'solid', colour: '#ffc7ce' } },
    { font: null, fill: { pattern: 'solid', colour: '#c6efce' } },
  ];
  const A = { top: 0, left: 0, bottom: 3, right: 0 };
  const B = { top: 0, left: 1, bottom: 3, right: 1 };
  view.conditionals.set('Data', [
    { type: 'cellIs', priority: 1, dxfId: 0, operator: 'greaterThan', formulas: ['10'], cfvos: [], colours: [], ranges: [A], stopIfTrue: false },
    { type: 'expression', priority: 2, dxfId: 1, formulas: ['MOD(A1,2)=0'], cfvos: [], colours: [], ranges: [A], stopIfTrue: false },
    { type: 'duplicateValues', priority: 3, dxfId: 1, formulas: [], cfvos: [], colours: [], ranges: [B], stopIfTrue: false },
  ]);

  const cache = new Map();
  assert.equal(view._conditionalStyle(0, 0, cache), null, '1 fails both rules');
  assert.equal(view._conditionalStyle(1, 0, cache).fill.colour, '#c6efce', '8 is even');
  const both = view._conditionalStyle(2, 0, cache);
  assert.equal(both.fill.colour, '#ffc7ce', '12 matches BOTH; the higher priority keeps the fill');
  assert.equal(both.font.colour, '#9c0006');
  assert.equal(view._conditionalStyle(0, 1, cache).fill.colour, '#c6efce', 'x appears twice');
  assert.equal(view._conditionalStyle(1, 1, cache), null, 'y is unique');

  // The frame wears the merged look, and an expression shifts like a copy:
  // MOD(A1,2) judged at row 1 reads A2.
  const cells = new Map(view.render().cells.map((c) => [c.ref, c]));
  assert.equal(cells.get('A2').style.fill.colour, '#c6efce', '8 is even at its own row');
  assert.equal(cells.get('A4').style.fill.colour, '#ffc7ce', '20 matches both; priority keeps the fill');
  assert.equal(cells.get('A3').style.font.colour, '#9c0006');
});

// ---------------------------------------------------------------- comments --

test('cell comments show as notes: ref, author, text — threaded shadows unnamed', () => {
  const wb = Workbook.open(simpleBook([['total', 100]]));
  wb.pkg.addPart('xl/worksheets/_rels/sheet1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>'
    + '</Relationships>',
    'application/vnd.openxmlformats-package.relationships+xml');
  wb.pkg.addPart('xl/comments1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + '<authors><author>Ejaz</author><author>tc={GUID}</author></authors>'
    + '<commentList>'
    + '<comment ref="B1" authorId="0"><text><r><t>Check &amp; confirm</t></r><r><t> this figure</t></r></text></comment>'
    + '<comment ref="C3" authorId="1"><text><t>a threaded shadow</t></text></comment>'
    + '</commentList></comments>',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml');
  const view = SheetView.open(wb.save());

  const cells = new Map(view.render().cells.map((c) => [c.ref, c]));
  assert.deepEqual(cells.get('B1').note, { author: 'Ejaz', text: 'Check & confirm this figure' },
    'runs concatenate and entities unescape');
  assert.ok(cells.get('C3'), 'an EMPTY commented cell still renders, or the mark has nowhere to sit');
  assert.equal(cells.get('C3').note.author, null, 'a tc= shadow author is a marker, not a name');
  assert.equal(cells.get('A1').note ?? null, null);
});

// ------------------------------------------------------- fill: dates and names --

test('a single DATE drags to the next days — the format is what makes it a date', async () => {
  const { dateToSerial } = await import('@rutba/formula');
  const d = (y, m, day) => dateToSerial(new Date(Date.UTC(y, m - 1, day)));
  const view = SheetView.open(simpleBook([[1]]));
  view.select(0, 0);
  view.setCell(0, 0, d(2026, 8, 20));
  view.setFormat({ numberFormat: 'yyyy-mm-dd' });
  view.select(0, 0);
  view.fill({ top: 0, left: 0, bottom: 3, right: 0 });
  assert.equal(view.calc.getValue('Data', 1, 0), d(2026, 8, 21));
  assert.equal(view.calc.getValue('Data', 3, 0), d(2026, 8, 23));
  assert.equal(view.displayValue(3, 0).text, '2026-08-23', 'the format travelled with the fill');

  // The same single value WITHOUT a date format copies, as it always did.
  const plain = SheetView.open(simpleBook([[5]]));
  plain.select(0, 0);
  plain.fill({ top: 0, left: 0, bottom: 2, right: 0 });
  assert.equal(plain.calc.getValue('Data', 2, 0), 5);
});

test('same-day-of-month dates step by months, clamped as Excel clamps', async () => {
  const { dateToSerial } = await import('@rutba/formula');
  const d = (y, m, day) => dateToSerial(new Date(Date.UTC(y, m - 1, day)));

  const view = SheetView.open(simpleBook([[1], [1]]));
  view.select(0, 0);
  view.select(1, 0, { extend: true });
  view.setCell(0, 0, d(2026, 1, 15));
  view.setCell(1, 0, d(2026, 2, 15));
  view.setFormat({ numberFormat: 'yyyy-mm-dd' });
  view.fill({ top: 0, left: 0, bottom: 4, right: 0 });
  assert.equal(view.displayValue(2, 0).text, '2026-03-15');
  assert.equal(view.displayValue(4, 0).text, '2026-05-15', 'the 15th of every month, not 15th + 31 days');

  // Jan 31, Feb 28 continues as month ends where the month is short.
  const ends = SheetView.open(simpleBook([[1], [1]]));
  ends.select(0, 0);
  ends.select(1, 0, { extend: true });
  ends.setCell(0, 0, d(2026, 1, 31));
  ends.setCell(1, 0, d(2026, 2, 28));
  ends.setFormat({ numberFormat: 'yyyy-mm-dd' });
  ends.fill({ top: 0, left: 0, bottom: 3, right: 0 });
  assert.equal(ends.displayValue(2, 0).text, '2026-03-31');
  assert.equal(ends.displayValue(3, 0).text, '2026-04-30', 'April has no 31st');
});

test('month and weekday names continue their list, wrapping, keeping case', () => {
  const view = SheetView.open(simpleBook([['Jan', 'MON', 'nov'], ['', 'WED', '']]));
  view.select(0, 0);
  view.fill({ top: 0, left: 0, bottom: 2, right: 0 });
  assert.equal(view.calc.getValue('Data', 1, 0), 'Feb');
  assert.equal(view.calc.getValue('Data', 2, 0), 'Mar');

  // Two sources set the stride: MON, WED -> FRI, SUN, TUE — wrapping the week.
  view.select(0, 1);
  view.select(1, 1, { extend: true });
  view.fill({ top: 0, left: 1, bottom: 4, right: 1 });
  assert.equal(view.calc.getValue('Data', 2, 1), 'FRI', 'the stride is two, the case is the source\u0027s');
  assert.equal(view.calc.getValue('Data', 4, 1), 'TUE', 'the week wraps');

  // Lower case wraps the year and stays lower.
  view.select(0, 2);
  view.fill({ top: 0, left: 2, bottom: 3, right: 2 });
  assert.equal(view.calc.getValue('Data', 1, 2), 'dec');
  assert.equal(view.calc.getValue('Data', 2, 2), 'jan');
});

// ------------------------------------------------------------------- tables --

const tableBook = () => {
  const wb = Workbook.open(simpleBook([['Item', 'Qty'], ['Ink', 4], ['Paper', 6]]));
  wb.pkg.addPart('xl/worksheets/_rels/sheet1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>'
    + '</Relationships>',
    'application/vnd.openxmlformats-package.relationships+xml');
  wb.pkg.addPart('xl/tables/table1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Stock" displayName="Stock" ref="A1:B3">'
    + '<autoFilter ref="A1:B3"/>'
    + '<tableColumns count="2"><tableColumn id="1" name="Item"/><tableColumn id="2" name="Qty"/></tableColumns>'
    + '<tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/>'
    + '</table>',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml');
  return wb.save();
};

test('a workbook table resolves structured references and rides the frame', () => {
  const view = SheetView.open(tableBook());

  view.setCell(0, 3, '=SUM(Stock[Qty])'); // D1
  assert.equal(view.calc.getValue('Data', 0, 3), 10);
  view.setCell(2, 1, 16); // Paper: 6 -> 16
  assert.equal(view.calc.getValue('Data', 0, 3), 20, 'the sum tracks edits through the table');

  view.setCell(1, 3, '=Stock[@Qty]*3'); // D2 — the named form reaches @ from beside the table
  assert.equal(view.calc.getValue('Data', 1, 2, undefined) ?? '', '', 'nothing was written to C2');
  assert.equal(view.calc.getValue('Data', 1, 3), 12);

  const frame = view.render();
  assert.equal(frame.tables.length, 1);
  assert.equal(frame.tables[0].name, 'Stock');
  assert.deepEqual(frame.tables[0].columns, ['Item', 'Qty']);
  assert.equal(frame.tables[0].bottom, 2);

  // The table part is preserved untouched — reading it is not a write.
  const reopened = SheetView.open(view.save());
  assert.equal(reopened.calc.getValue('Data', 0, 3), 20, 'the reference survives the round trip');
});

test('a table filter hides rows in the FILE, and clearing brings them back', () => {
  const view = SheetView.open(tableBook());
  view.applyFilter('Stock', 'Qty', ['4']);
  assert.equal(view.geo.rowHeight(2), 0, 'the excluded Paper row collapsed');
  assert.ok(view.geo.rowHeight(1) > 0, 'the Ink row still shows');

  // The state is exactly what Excel stores: filterColumn in the table part,
  // hidden on the sheet row — so the file opens identically filtered anywhere.
  const saved = Workbook.open(view.save());
  assert.match(saved.pkg.text('xl/tables/table1.xml'),
    /<filterColumn colId="1"><filters><filter val="4"\/><\/filters><\/filterColumn>/);
  assert.match(saved.pkg.text('xl/worksheets/sheet1.xml'), /<row[^>]*r="3"[^>]*hidden="1"/);
  const reopened = SheetView.open(view.save());
  assert.equal(reopened.geo.rowHeight(2), 0, 'a reopened workbook is still filtered');

  view.undo();
  assert.ok(view.geo.rowHeight(2) > 0, 'undo unhid the row');
  assert.equal(view.workbook.tableFilters('xl/tables/table1.xml').size, 0, 'and the filter state went with it');
  view.redo();
  assert.equal(view.geo.rowHeight(2), 0, 'redo re-filtered');

  view.applyFilter('Stock', 'Qty', null);
  assert.ok(view.geo.rowHeight(2) > 0, 'clearing is applying null');

  // The dropdown list is per-person cursor state, distinct values sorted.
  view.openFilterPanel('Stock', 'Qty');
  assert.deepEqual(view.filterPanel.values, ['4', '6']);
  assert.equal(view.filterPanel.selected, null, 'no filter means everything passes');
  view.openFilterPanel(null);
  assert.equal(view.filterPanel, null);

  assert.throws(() => view.applyFilter('Ghost', 'Qty', ['4']), /no table/);
  assert.throws(() => view.applyFilter('Stock', 'Ghost', ['4']), /no column/);

  // The frame says which columns wear a funnel.
  view.applyFilter('Stock', 'Item', ['Ink']);
  const table = view.render().tables[0];
  assert.deepEqual(table.filtered, [0]);
});

test('a spilled formula shows its ghosts in the grid, and survives a save', () => {
  const view = SheetView.open(simpleBook([[1]]));
  view.select(0, 1);
  view.setCell(0, 1, '=SEQUENCE(3)');
  const cells = new Map(view.render().cells.map((c) => [c.ref, c]));
  assert.equal(cells.get('B1').text, '1');
  assert.equal(cells.get('B2').text, '2', 'the ghost paints');
  assert.equal(cells.get('B3').text, '3');
  assert.equal(view.editValue(1, 1), '', 'a ghost has no input of its own');

  const reopened = SheetView.open(view.save());
  assert.equal(reopened.calc.getValue('Data', 2, 1), 3, 'the formula spills again on open');
});

// ----------------------------------------------------------------- scenarios --

test('a scenario stores the selection, applies as one undo, and rides the file', () => {
  const view = SheetView.open(simpleBook([[100, 0.2], [null, '=A1*(1+B1)']]));
  assert.equal(view.calc.getValue('Data', 1, 1), 120);

  // Capture the current inputs of A1:B1 as the base case.
  view.select(0, 0);
  view.select(0, 1, { extend: true });
  view.defineScenario('Base', 'as sold today');

  // Change the inputs and capture again as the optimistic case.
  view.setCell(0, 0, 150);
  view.setCell(0, 1, 0.5);
  view.defineScenario('Best');
  assert.equal(view.calc.getValue('Data', 1, 1), 225);

  // Showing a scenario lands its inputs as ONE ordinary edit.
  view.showScenario('Base');
  assert.equal(view.calc.getValue('Data', 1, 1), 120, 'the base inputs came back');
  view.undo();
  assert.equal(view.calc.getValue('Data', 1, 1), 225, 'one undo step restores the case before');

  // The scenarios live in the sheet part and survive a round trip.
  const reopened = SheetView.open(view.save());
  assert.deepEqual(reopened.scenarios().map((s) => s.name).sort(), ['Base', 'Best']);
  assert.equal(reopened.scenarios().find((s) => s.name === 'Base').comment, 'as sold today');
  reopened.showScenario('Best');
  assert.equal(reopened.calc.getValue('Data', 1, 1), 225);

  // Deleting forgets it; re-defining a name replaces it.
  reopened.deleteScenario('Best');
  assert.deepEqual(reopened.scenarios().map((s) => s.name), ['Base']);
  assert.throws(() => reopened.showScenario('Best'), /no scenario/);
  assert.throws(() => reopened.defineScenario(''), /needs a name/);

  // A formula cell cannot be a changing cell — that is what Goal Seek is for.
  reopened.select(1, 1);
  assert.throws(() => reopened.defineScenario('Broken'), /plain values/);
});

// -------------------------------------------------------------- frozen panes --

test('freeze panes writes the pane element, rides every frame, and thaws', () => {
  const rows = Array.from({ length: 60 }, (_, i) => ['r' + (i + 1), i + 1]);
  rows[0] = ['Header', 'Qty'];
  const view = SheetView.open(simpleBook(rows));

  view.freezePanes(1, 1);
  let frame = view.render();
  assert.equal(frame.frozen.rows, 1);
  assert.equal(frame.frozen.cols, 1);
  assert.ok(frame.frozen.height > 0, 'the band size travels in pixels');

  // Scrolled far down, the frozen row is STILL in the frame for the client
  // to pin — and so is the frozen column when scrolled right.
  view.scrollTo(0, 5000);
  frame = view.render();
  assert.ok(frame.viewport.firstRow > 1, 'the viewport left row 1 behind');
  assert.ok(frame.cells.some((c) => c.ref === 'A1'), 'the frozen corner cell rides along');
  assert.ok(frame.cells.some((c) => c.ref === 'B1'), 'the frozen row rides along');
  assert.ok(frame.rows.some((r) => r.index === 0), 'so does its header band entry');

  // What Excel stores is what we store: a pane element in the sheet view.
  const saved = Workbook.open(view.save());
  assert.match(saved.pkg.text('xl/worksheets/sheet1.xml'),
    /<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"\/>/);
  const reopened = SheetView.open(view.save());
  assert.deepEqual(reopened.frozenPane(), { rows: 1, cols: 1 });

  // Undo restores the part without the pane; thawing writes zeros.
  view.undo();
  assert.deepEqual(view.frozenPane(), { rows: 0, cols: 0 });
  view.redo();
  assert.deepEqual(view.frozenPane(), { rows: 1, cols: 1 });
  view.freezePanes(0, 0);
  assert.deepEqual(view.frozenPane(), { rows: 0, cols: 0 });
  assert.equal(view.render().frozen.height, 0);

  assert.throws(() => view.freezePanes(-1, 0), /whole counts/);
});

// ------------------------------------------------------------- pivot tables --

/**
 * A workbook with a REAL pivot: a source sheet, a report sheet, the table
 * definition, and the cache pair — assembled by hand because the creator
 * cannot build one and a pivot is exactly the shape a customer file has.
 */
const pivotBook = ({ compact = true, colFields = true } = {}) => {
  const wb = Workbook.open(buildXlsx({
    sheets: [
      {
        name: 'Data',
        rows: [
          ['Region', 'Product', 'Qty'],
          ['North', 'Widget', 5],
          ['North', 'Gadget', 3],
          ['South', 'Widget', 2],
          ['South', 'Gadget', 7],
        ],
      },
      // The materialised cells a pivot leaves behind, as Excel wrote them.
      { name: 'Report', rows: [['Row Labels', 'Gadget', 'Widget', 'Grand Total']] },
    ],
  }));
  wb.pkg.addPart('xl/worksheets/_rels/sheet2.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>'
    + '</Relationships>',
    'application/vnd.openxmlformats-package.relationships+xml');
  wb.pkg.addPart('xl/pivotTables/_rels/pivotTable1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/>'
    + '</Relationships>',
    'application/vnd.openxmlformats-package.relationships+xml');
  wb.pkg.addPart('xl/pivotCache/pivotCacheDefinition1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="4">'
    + '<cacheSource type="worksheet"><worksheetSource ref="A1:C5" sheet="Data"/></cacheSource>'
    + '<cacheFields count="3">'
    + '<cacheField name="Region" numFmtId="0"><sharedItems count="2"><s v="North"/><s v="South"/></sharedItems></cacheField>'
    + '<cacheField name="Product" numFmtId="0"><sharedItems count="2"><s v="Gadget"/><s v="Widget"/></sharedItems></cacheField>'
    + '<cacheField name="Qty" numFmtId="0"><sharedItems containsString="0" containsNumber="1" minValue="2" maxValue="7"/></cacheField>'
    + '</cacheFields></pivotCacheDefinition>',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml');
  wb.pkg.addPart('xl/pivotTables/pivotTable1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    + ' name="SalesPivot" cacheId="1" compact="' + (compact ? 1 : 0) + '" outline="' + (compact ? 1 : 0) + '">'
    + '<location ref="A1:D4" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>'
    + '<pivotFields count="3">'
    + '<pivotField axis="axisRow" showAll="0"/>'
    + (colFields ? '<pivotField axis="axisCol" showAll="0"/>' : '<pivotField showAll="0"/>')
    + '<pivotField dataField="1" showAll="0"/>'
    + '</pivotFields>'
    + '<rowFields count="1"><field x="0"/></rowFields>'
    + (colFields ? '<colFields count="1"><field x="1"/></colFields>' : '')
    + '<dataFields count="1"><dataField name="Sum of Qty" fld="2" baseField="0" baseItem="0"/></dataFields>'
    + '</pivotTableDefinition>',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml');
  return wb.save();
};

test('a pivot is READ: its source, its axes, and what it summarises', () => {
  const view = SheetView.open(pivotBook());
  const p = view.render().pivots;
  assert.equal(p.length, 1);
  assert.equal(p[0].name, 'SalesPivot');
  assert.equal(p[0].sheet, 'Report');
  assert.equal(p[0].source, 'Data!A1:C5');
  assert.deepEqual(p[0].fields, ['Region', 'Product', 'Qty']);
  assert.deepEqual(p[0].rows, ['Region']);
  assert.deepEqual(p[0].cols, ['Product']);
  assert.deepEqual(p[0].values, ['Sum of Qty']);
  assert.equal(p[0].unsupported, null);
});

test('refreshing a pivot recomputes it from the LIVE source', () => {
  const view = SheetView.open(pivotBook());
  view.selectSheet('Report');
  const shape = view.refreshPivot('SalesPivot');
  assert.deepEqual(shape, { rows: 4, cols: 4 });

  const at = (ref) => view.displayValue(...parseRefPairLocal(ref)).text;
  assert.equal(at('A1'), 'Row Labels');
  assert.equal(at('B1'), 'Gadget');
  assert.equal(at('C1'), 'Widget');
  assert.equal(at('D1'), 'Grand Total');
  assert.equal(at('A2'), 'North');
  assert.equal(at('B2'), '3', 'North/Gadget');
  assert.equal(at('C2'), '5', 'North/Widget');
  assert.equal(at('D2'), '8', 'the row total');
  assert.equal(at('A3'), 'South');
  assert.equal(at('B3'), '7');
  assert.equal(at('D3'), '9');
  assert.equal(at('A4'), 'Grand Total');
  assert.equal(at('B4'), '10', 'every Gadget');
  assert.equal(at('D4'), '17', 'the grand total');

  // THE POINT: edit the source, refresh, and the pivot moves. Before this
  // existed the materialised cells kept last week's totals, silently.
  view.selectSheet('Data');
  view.setCell(1, 2, 50); // North/Widget 5 -> 50
  view.selectSheet('Report');
  view.refreshPivot('SalesPivot');
  assert.equal(at('C2'), '50');
  assert.equal(at('D2'), '53');
  assert.equal(at('D4'), '62');

  // One undo step covers the whole rectangle.
  view.undo();
  assert.equal(at('C2'), '5', 'the previous refresh came back whole');
  assert.equal(at('D4'), '17');
});

test('a pivot that grows takes its stored location with it, and shrinks back', () => {
  const view = SheetView.open(pivotBook());

  // Data BELOW the declared source range is not in the pivot — the source is
  // A1:C5 and that is the whole contract. Excel behaves the same way, and a
  // reader who assumes otherwise is the person this test exists for.
  view.selectSheet('Data');
  view.setCell(5, 0, 'East');
  view.setCell(5, 1, 'Widget');
  view.setCell(5, 2, 4);
  view.selectSheet('Report');
  assert.equal(view.refreshPivot('SalesPivot').rows, 4, 'the out-of-range row was not counted');

  // A third region INSIDE the range does grow it.
  view.selectSheet('Data');
  view.setCell(4, 0, 'East'); // the South/Gadget row becomes East/Gadget
  view.selectSheet('Report');
  const grown = view.refreshPivot('SalesPivot');
  assert.equal(grown.rows, 5, 'three regions plus header plus grand total');
  assert.equal(view.displayValue(1, 0).text, 'East', 'sorted into place');

  const saved = Workbook.open(view.save());
  assert.match(saved.pkg.text('xl/pivotTables/pivotTable1.xml'), /ref="A1:D5"/,
    'the stored location followed the shape');
  assert.match(saved.pkg.text('xl/pivotCache/pivotCacheDefinition1.xml'), /refreshOnLoad="1"/,
    'and Excel is told to rebuild its own cache from the same source');

  // Put the region back: the pivot shrinks and clears up after itself.
  view.selectSheet('Data');
  view.setCell(4, 0, 'South');
  view.selectSheet('Report');
  view.refreshPivot('SalesPivot');
  assert.equal(view.displayValue(4, 0).text, '', 'the row it no longer covers was cleared');
  assert.equal(view.displayValue(3, 0).text, 'Grand Total');
});

test('a pivot this editor cannot recompute says why, and refuses rather than guessing', () => {
  const wb = Workbook.open(pivotBook());
  // Group the Region field, as Excel does for dates and number bands.
  wb.pkg.write_('xl/pivotCache/pivotCacheDefinition1.xml',
    wb.pkg.text('xl/pivotCache/pivotCacheDefinition1.xml')
      .replace('<sharedItems count="2"><s v="North"/><s v="South"/></sharedItems>',
        '<fieldGroup par="3"><rangePr groupBy="months"/></fieldGroup>'));
  const view = SheetView.open(wb.save());
  const p = view.render().pivots[0];
  assert.match(p.unsupported, /grouped/);
  assert.equal(p.ref, null, 'an unrefreshable pivot offers no rectangle to act on');
  view.selectSheet('Report');
  assert.throws(() => view.refreshPivot('SalesPivot'), /grouped/);
  assert.throws(() => view.refreshPivot('Ghost'), /no pivot table/);
  assert.equal(view.isDirty, false, 'a refused refresh changed nothing');
});

/** A1 -> [row, col], local to these tests. */
function parseRefPairLocal(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return [Number(m[2]) - 1, col - 1];
}

test('a pivot can be CREATED over a range, and undo takes its parts away', () => {
  const view = SheetView.open(buildXlsx({
    sheets: [{
      name: 'Data',
      rows: [
        ['Region', 'Product', 'Qty'],
        ['North', 'Widget', 5],
        ['North', 'Gadget', 3],
        ['South', 'Widget', 2],
        ['South', 'Gadget', 7],
      ],
    }],
  }));

  const shape = view.createPivot({
    name: 'ByRegion',
    source: { top: 0, left: 0, bottom: 4, right: 2 },
    target: { sheet: 'Data', row: 6, col: 0 },
    rowFields: ['Region'],
    colFields: ['Product'],
    dataFields: [{ field: 'Qty' }],
  });
  assert.equal(shape.sheet, 'Data');
  assert.deepEqual([shape.rows, shape.cols], [4, 4]);

  // It rendered where it was asked to, computed from the source.
  assert.equal(view.displayValue(6, 1).text, 'Gadget');
  assert.equal(view.displayValue(7, 0).text, 'North');
  assert.equal(view.displayValue(7, 3).text, '8');
  assert.equal(view.displayValue(9, 3).text, '17', 'the grand total');

  // It is a REAL pivot: the parts exist, wired both ways, and it reads back.
  const p = view.pivots();
  assert.equal(p.length, 1);
  assert.equal(p[0].name, 'ByRegion');
  assert.equal(p[0].unsupported, null);
  assert.deepEqual(p[0].rowFields.map((f) => p[0].cacheFields[f].name), ['Region']);

  const saved = Workbook.open(view.save());
  assert.ok(saved.pkg.has('xl/pivotTables/pivotTable1.xml'));
  assert.ok(saved.pkg.has('xl/pivotCache/pivotCacheDefinition1.xml'));
  assert.ok(saved.pkg.has('xl/pivotCache/pivotCacheRecords1.xml'));
  assert.match(saved.pkg.text('xl/workbook.xml'), /<pivotCaches><pivotCache cacheId="1"/,
    'the workbook lists its cache');
  assert.match(saved.pkg.text('[Content_Types].xml'), /pivotTable\+xml/);
  // The reopened file still knows what it is, and refreshes.
  const reopened = SheetView.open(view.save());
  assert.equal(reopened.pivots().length, 1);
  reopened.refreshPivot('ByRegion');
  assert.equal(reopened.displayValue(9, 3).text, '17');

  // Undo removes the parts as well as the cells — no ghost pivot left.
  view.undo();
  assert.equal(view.pivots().length, 0, 'the definition went with it');
  assert.equal(view.pkg.has('xl/pivotTables/pivotTable1.xml'), false, 'and so did its part');
  assert.equal(view.displayValue(7, 0).text, '', 'and the cells it painted');
  view.redo();
  assert.equal(view.pivots().length, 1, 'redo puts the whole thing back');
  assert.equal(view.displayValue(9, 3).text, '17');
  assert.ok(Workbook.open(view.save()).pkg.has('xl/pivotCache/pivotCacheRecords1.xml'));
});

test('creating a pivot refuses what it cannot honestly build', () => {
  const view = SheetView.open(simpleBook([['Region', 'Qty'], ['North', 5]]));
  const source = { top: 0, left: 0, bottom: 1, right: 1 };
  const base = { source, target: { sheet: 'Data', row: 4, col: 0 }, rowFields: ['Region'] };
  assert.throws(() => view.createPivot({ ...base, name: '', dataFields: [{ field: 'Qty' }] }), /needs a name/);
  assert.throws(() => view.createPivot({ ...base, name: 'P', dataFields: [] }), /at least one value/);
  assert.throws(() => view.createPivot({ ...base, name: 'P', dataFields: [{ field: 'Ghost' }] }), /no source column/);
  assert.throws(() => view.createPivot({
    ...base, name: 'P', colFields: ['Region'], dataFields: [{ field: 'Qty' }],
  }), /both axes/);
  assert.throws(() => view.createPivot({
    name: 'P', source: { top: 0, left: 0, bottom: 0, right: 1 },
    target: { sheet: 'Data', row: 4, col: 0 }, rowFields: ['Region'], dataFields: [{ field: 'Qty' }],
  }), /header row and at least one row/);
  assert.equal(view.isDirty, false, 'a refused creation changed nothing');
  assert.equal(view.canUndo, false, 'and left no history');
});

// ---------------------------------------------------------------- data tables --

test('a one-way data table runs its formula across substituted inputs', () => {
  // A loan: B1 rate, B2 principal, B3 the interest. The table tries rates.
  const view = SheetView.open(simpleBook([
    [null, 0.1],
    [null, 1000],
    [null, '=B2*B1'],
  ]));
  assert.equal(view.calc.getValue('Data', 2, 1), 100);

  // Candidates down D2:D4, the formula referenced at E1.
  view.setCell(0, 4, '=B3');
  view.setCell(1, 3, 0.05);
  view.setCell(2, 3, 0.2);
  view.setCell(3, 3, 0.5);
  const done = view.dataTable({
    range: { top: 0, left: 3, bottom: 3, right: 4 },
    colInput: 'B1',
  });
  assert.equal(done.cells, 3);
  assert.equal(view.calc.getValue('Data', 1, 4), 50);
  assert.equal(view.calc.getValue('Data', 2, 4), 200);
  assert.equal(view.calc.getValue('Data', 3, 4), 500);

  // The model is exactly as it was: a half-substituted sheet would be a
  // wrong sheet that looks authoritative.
  assert.equal(view.calc.getValue('Data', 0, 1), 0.1);
  assert.equal(view.calc.getValue('Data', 2, 1), 100);

  // One undo step for the whole table.
  view.undo();
  assert.equal(view.calc.getValue('Data', 1, 4), '');
});

test('a two-way data table substitutes on both edges', () => {
  // B1 rate, B2 years, B3 = rate * years * 100.
  const view = SheetView.open(simpleBook([
    [null, 1],
    [null, 1],
    [null, '=B1*B2*100'],
  ]));
  // Corner formula at D1, years across E1:F1, rates down D2:D3.
  view.setCell(0, 3, '=B3');
  view.setCell(0, 4, 2);
  view.setCell(0, 5, 3);
  view.setCell(1, 3, 5);
  view.setCell(2, 3, 10);
  view.dataTable({
    range: { top: 0, left: 3, bottom: 2, right: 5 },
    rowInput: 'B2',
    colInput: 'B1',
  });
  assert.equal(view.calc.getValue('Data', 1, 4), 1000, '5 x 2 x 100');
  assert.equal(view.calc.getValue('Data', 1, 5), 1500, '5 x 3 x 100');
  assert.equal(view.calc.getValue('Data', 2, 4), 2000, '10 x 2 x 100');
  assert.equal(view.calc.getValue('Data', 2, 5), 3000);
  assert.equal(view.calc.getValue('Data', 0, 1), 1, 'the inputs went back');
});

test('a data table refuses what it cannot honestly run', () => {
  const view = SheetView.open(simpleBook([[null, 5], [null, '=B1*2']]));
  const range = { top: 0, left: 3, bottom: 2, right: 4 };
  assert.throws(() => view.dataTable({ range, colInput: '' }), /row input cell, a column input cell, or both/);
  assert.throws(() => view.dataTable({ range, colInput: 'nonsense' }), /must be a cell/);
  assert.throws(() => view.dataTable({ range, colInput: 'B2' }), /plain value, not a formula/);
  assert.throws(() => view.dataTable({
    range: { top: 0, left: 3, bottom: 0, right: 4 }, colInput: 'B1',
  }), /two rows and two columns/);
  // No formula in the top row where a column-input table needs one.
  assert.throws(() => view.dataTable({ range, colInput: 'B1' }), /formula in the top row/);
  assert.equal(view.canUndo, false, 'a refused table left no history');
  assert.equal(view.calc.getValue('Data', 0, 1), 5, 'and did not disturb the model');
});

test('a data table candidate that breaks the formula shows WHICH error, not a blank', () => {
  const view = SheetView.open(simpleBook([[null, 4], [null, '=100/B1']]));
  view.setCell(0, 4, '=B2');   // E1 references the formula
  view.setCell(1, 3, 2);       // D2: fine
  view.setCell(2, 3, 0);       // D3: divides by zero
  view.dataTable({ range: { top: 0, left: 3, bottom: 2, right: 4 }, colInput: 'B1' });
  assert.equal(view.calc.getValue('Data', 1, 4), 50);
  assert.equal(view.calc.getValue('Data', 2, 4), '#DIV/0!',
    'the failing input is visible in the table, as text');
  assert.equal(view.calc.getValue('Data', 0, 1), 4, 'and the model went back');
});

// ------------------------------------------------------------ sheet protection --

test('protection guards locked cells — and every gesture passes one gate', () => {
  const view = SheetView.open(simpleBook([[1, 2], [3, 4]]));

  // Unlock B1 the way a person does: format it, THEN protect the sheet.
  view.select(0, 1);
  view.setFormat({ locked: false });
  view.protect();
  assert.equal(view.render().protection.sheet, true);

  // A locked cell refuses, with the cell named; the unlocked one edits.
  assert.throws(() => view.setCell(0, 0, 99), /protected and A1 is locked/);
  view.setCell(0, 1, 99);
  assert.equal(view.calc.getValue('Data', 0, 1), 99);

  // The single gate covers gestures that never call setCell directly.
  view.select(1, 0); // A2, locked
  assert.throws(() => view.clearSelection(), /locked/);
  view.select(0, 0);
  view.select(1, 1, { extend: true });
  // Descending, so the rows genuinely reorder — a sort that changes nothing
  // writes nothing, and a gate nothing reaches proves nothing.
  assert.throws(() => view.sortSelection({ ascending: false }), /locked/);
  assert.throws(() => view.fill({ top: 0, left: 0, bottom: 3, right: 1 }), /locked/);

  // Structural edits refuse outright on a protected sheet.
  assert.throws(() => view.insertRows(1, 1), /unprotect it before changing its structure/);

  // A refused gesture recorded nothing: undo history holds only real edits.
  view.undo();
  assert.equal(view.calc.getValue('Data', 0, 1), 2, 'the one real edit undid');
  assert.equal(view.canUndo, true, 'protect itself is undoable');

  // Unprotect thaws everything, and the round trip carries the state.
  view.unprotect();
  view.setCell(0, 0, 7);
  assert.equal(view.calc.getValue('Data', 0, 0), 7);
});

test('protection survives the file: written as Excel writes it, read back enforced', () => {
  const view = SheetView.open(simpleBook([[1]]));
  view.protect();
  const saved = Workbook.open(view.save());
  assert.match(saved.pkg.text('xl/worksheets/sheet1.xml'),
    /<sheetProtection sheet="1" objects="1" scenarios="1"\/>/);

  const reopened = SheetView.open(view.save());
  assert.equal(reopened.protection().sheet, true);
  assert.throws(() => reopened.setCell(0, 0, 2), /locked/);
  reopened.unprotect();
  reopened.setCell(0, 0, 2);
  assert.equal(reopened.calc.getValue('Data', 0, 0), 2);
});

test('a password seals protection against everyone but Excel', () => {
  const wb = Workbook.open(simpleBook([[1]]));
  wb.pkg.write_('xl/worksheets/sheet1.xml',
    wb.pkg.text('xl/worksheets/sheet1.xml').replace('</sheetData>',
      '</sheetData><sheetProtection algorithmName="SHA-512" hashValue="abc=" saltValue="s=" spinCount="100000" sheet="1"/>'));
  const view = SheetView.open(wb.save());
  assert.equal(view.render().protection.hasPassword, true);
  assert.throws(() => view.setCell(0, 0, 2), /locked/);
  assert.throws(() => view.unprotect(), /password.*Excel/i);
});

test('formatting an unlocked cell no longer drops its protection child', () => {
  const view = SheetView.open(simpleBook([[5]]));
  view.select(0, 0);
  view.setFormat({ locked: false });
  // The regression this pins: the xf rebuild used to emit only the alignment
  // child, so ANY later format edit silently re-locked the cell.
  view.setFormat({ bold: true });
  assert.equal(view.formatState().locked, false, 'bold did not re-lock the cell');
  view.protect();
  view.setCell(0, 0, 8);
  assert.equal(view.calc.getValue('Data', 0, 0), 8, 'still editable under protection');

  const reopened = SheetView.open(view.save());
  assert.equal(reopened.isCellLocked(0, 0), false, 'the unlock survived the round trip');
});

// ------------------------------------------------------------- table styling --

test('a table paints its banding — in our palette, never over a louder voice', () => {
  const view = SheetView.open(tableBook()); // Stock, A1:B3, TableStyleMedium2 + row stripes
  const cells = new Map(view.render().cells.map((c) => [c.ref, c]));

  // The header row wears the accent with bold white text.
  assert.ok(cells.get('A1').style.fill, 'the header row is filled');
  assert.equal(cells.get('A1').style.font.bold, true);
  assert.equal(cells.get('A1').style.font.colour, '#ffffff');

  // Row stripes: the first body row is plain, the second banded.
  assert.equal(cells.get('A2').style?.fill ?? null, null, 'body row one is plain');
  assert.ok(cells.get('A3').style.fill, 'body row two wears the stripe');
  assert.notEqual(cells.get('A3').style.fill.colour, cells.get('A1').style.fill.colour,
    'the stripe is a tint, not the header accent');

  // A cell OUTSIDE the table is untouched.
  view.setCell(0, 3, 'x');
  const again = new Map(view.render().cells.map((c) => [c.ref, c]));
  assert.equal(again.get('D1').style ?? null, null);

  // The cell's own fill beats the banding: nothing the file said is hidden.
  view.select(2, 0);
  view.setFormat({ fill: '#ff0000' });
  const styled = new Map(view.render().cells.map((c) => [c.ref, c]));
  assert.equal(styled.get('A3').style.fill.colour, '#ff0000');
});

// -------------------------------------------------------- clipboard interop --

test('copy renders an HTML table other applications understand', () => {
  const view = SheetView.open(simpleBook([['Item', 12.5], ['Total', 25]]));
  view.select(0, 0);
  view.setFormat({ bold: true, fill: '#ffe0e0' });
  view.select(0, 0);
  view.select(1, 1, { extend: true });
  const html = view.copyHtml();
  assert.match(html, /<table><tr><td style="font-weight:bold;background:#ffe0e0">Item<\/td>/);
  assert.match(html, /text-align:right">12.5<\/td>/, 'numbers carry their alignment');
  assert.match(html, /<td[^>]*>25<\/td>/);

  // Markup in a cell is text, never live HTML on someone else s clipboard.
  view.setCell(1, 0, '<b>&raw');
  const again = view.copyHtml();
  assert.match(again, /&lt;b&gt;&amp;raw/);
});

test('an HTML table pastes by CELL — values with tabs and newlines stay whole', () => {
  const view = SheetView.open(simpleBook([[1]]));
  view.select(1, 0);
  view.pasteText('Item\tQty\nInk\t4', // the text flavour, deliberately misleading
    '<table><tbody>'
    + '<tr><td>Item</td><td>Qty</td></tr>'
    + '<tr><td><b>Ink</b>, black<br>refill</td><td style="text-align:right">4</td></tr>'
    + '</tbody></table>');
  assert.equal(view.calc.getValue('Data', 1, 0), 'Item');
  assert.equal(view.calc.getValue('Data', 2, 0), 'Ink, black\nrefill', 'tags stripped, the break kept, ONE cell');
  assert.equal(view.calc.getValue('Data', 2, 1), 4, 'the number came through as a number');
  assert.equal(view.selection.toString(), 'A2:B3', 'the paste selected what it filled');

  // One undo covers the grid.
  view.undo();
  assert.equal(view.calc.getValue('Data', 1, 0), '');

  // Our OWN copy still takes the rich road — the mark outranks the HTML.
  view.setCell(0, 1, '=A1*2');
  view.select(0, 1);
  view.markClipboard();
  const mark = view.clipboard.text;
  view.select(2, 1);
  view.pasteText(mark, '<table><tr><td>should not be used</td></tr></table>');
  assert.equal(view.editValue(2, 1), '=A3*2', 'the formula shifted — that is the rich path');

  // HTML with no table falls back to the text flavour.
  view.select(4, 0);
  view.pasteText('plain', '<div>no table here</div>');
  assert.equal(view.calc.getValue('Data', 4, 0), 'plain');
});

// ------------------------------------------- the sheet-level autofilter --

test('Data - Filter toggles a sheet autofilter over the data region, and off unhides', () => {
  const view = SheetView.open(simpleBook([['Item', 'Qty'], ['Ink', 4], ['Paper', 6]]));
  view.select(1, 0);
  view.toggleAutoFilter();

  // The file carries exactly what Excel writes: <autoFilter ref> in the
  // sheet part, in schema position.
  const saved = Workbook.open(view.save());
  assert.match(saved.pkg.text('xl/worksheets/sheet1.xml'), /<autoFilter ref="A1:B3"\/>/);

  // The frame carries it as the #sheet pseudo-table, so the grid's funnels
  // and panel serve it with the code the tables already use.
  const entry = view.render().tables.find((t) => t.name === '#sheet');
  assert.ok(entry, 'the pseudo-entry rides the frame');
  assert.deepEqual(entry.columns, ['Item', 'Qty']);
  assert.deepEqual([entry.top, entry.left, entry.bottom, entry.right], [0, 0, 2, 1]);

  view.toggleAutoFilter();
  assert.doesNotMatch(Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml'), /<autoFilter/);
  assert.equal(view.render().tables.find((t) => t.name === '#sheet'), undefined);
});

test('a sheet filter hides rows in the FILE and clears like a table filter', () => {
  const view = SheetView.open(simpleBook([['Item', 'Qty'], ['Ink', 4], ['Paper', 6]]));
  view.select(0, 0);
  view.toggleAutoFilter();
  view.applyFilter('#sheet', 'Qty', ['4']);
  assert.equal(view.geo.rowHeight(2), 0, 'the excluded Paper row collapsed');
  assert.ok(view.geo.rowHeight(1) > 0, 'the Ink row still shows');

  const sheetXml = Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');
  assert.match(sheetXml, /<autoFilter ref="A1:B3"><filterColumn colId="1"><filters><filter val="4"\/><\/filters><\/filterColumn><\/autoFilter>/);
  assert.match(sheetXml, /<row[^>]*r="3"[^>]*hidden="1"/);

  const reopened = SheetView.open(view.save());
  assert.equal(reopened.geo.rowHeight(2), 0, 'a reopened workbook is still filtered');
  assert.ok(reopened.render().tables.find((t) => t.name === '#sheet').filtered.includes(1));

  view.applyFilter('#sheet', 'Qty', null);
  assert.ok(view.geo.rowHeight(2) > 0, 'clearing brings the row back');
  assert.match(Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml'),
    /<autoFilter ref="A1:B3"\/>/, 'an emptied filter collapses to the bare element');
});

test('turning the filter OFF unhides what it had hidden, and undo restores both parts', () => {
  const view = SheetView.open(simpleBook([['Item', 'Qty'], ['Ink', 4], ['Paper', 6]]));
  view.select(0, 0);
  view.toggleAutoFilter();
  view.applyFilter('#sheet', 'Qty', ['4']);
  assert.equal(view.geo.rowHeight(2), 0);

  view.toggleAutoFilter(); // off - funnels leave, rows come back
  assert.ok(view.geo.rowHeight(2) > 0, 'off unhides');
  assert.doesNotMatch(Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml'), /<autoFilter/);

  view.undo(); // the toggle-off
  assert.equal(view.geo.rowHeight(2), 0, 'undo brings the filter and its hidden row back');
  assert.match(Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml'), /<filterColumn colId="1"/);
});

test('the filter refuses an empty cell with nothing around it', () => {
  const view = SheetView.open(simpleBook([['Item', 'Qty'], ['Ink', 4]]));
  view.select(9, 9);
  assert.throws(() => view.toggleAutoFilter(), /Select a cell inside your data/);
});

test('the fixture sheet autofilter reads through the same digest', () => {
  const view = SheetView.open(buildComplexWorkbook());
  const entry = view.render().tables.find((t) => t.name === '#sheet');
  assert.ok(entry, 'the complex fixture ships <autoFilter ref="A1:D3"/> and it shows');
  assert.deepEqual([entry.top, entry.bottom], [0, 2]);
});

// ------------------------------------------------- calculated columns --

test('a formula typed into a table data column fills the column, shifted per row', () => {
  const view = SheetView.open(tableBook());
  view.setCell(1, 2, '=B2*2'); // C2, beside the Stock table's Qty column? C is OUTSIDE A1:B3
  assert.equal(view.editValue(2, 2), '', 'outside the table nothing propagates');

  view.setCell(1, 1, '=A2&"!"'); // B2 IS in the table's data band (rows 2-3)
  assert.equal(view.editValue(2, 1), '6', 'a column holding values keeps them - no overwrite');

  view.undo();
});

test('an empty table column takes the formula everywhere as one undo step', () => {
  const wb = Workbook.open(simpleBook([['Item', 'Qty', 'Total'], ['Ink', 4, null], ['Paper', 6, null]]));
  wb.pkg.addPart('xl/worksheets/_rels/sheet1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>'
    + '</Relationships>',
    'application/vnd.openxmlformats-package.relationships+xml');
  wb.pkg.addPart('xl/tables/table1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Stock" displayName="Stock" ref="A1:C3">'
    + '<tableColumns count="3"><tableColumn id="1" name="Item"/><tableColumn id="2" name="Qty"/><tableColumn id="3" name="Total"/></tableColumns>'
    + '</table>',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml');
  const view = SheetView.open(wb.save());

  view.setCell(1, 2, '=B2*10'); // C2 - the Total column is empty below
  assert.equal(view.editValue(2, 2), '=B3*10', 'the formula filled down, shifted');
  assert.equal(view.calc.getValue('Data', 2, 2), 60);

  view.undo();
  assert.equal(view.editValue(1, 2), '', 'one undo removes the whole fill');
  assert.equal(view.editValue(2, 2), '');

  view.redo();
  assert.equal(view.calc.getValue('Data', 1, 2), 40, 'redo brings the column back');
  assert.equal(view.calc.getValue('Data', 2, 2), 60);
});

// ------------------------------------------------------------ icon sets --

test('an icon set paints from the file: which icon, which cell, straight to the frame', () => {
  const wb = Workbook.open(simpleBook([[10], [50], [90]]));
  const spliced = wb.pkg.text('xl/worksheets/sheet1.xml').replace('</sheetData>', '</sheetData>'
    + '<conditionalFormatting sqref="A1:A3"><cfRule type="iconSet" priority="1">'
    + '<iconSet iconSet="3Arrows">'
    + '<cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/>'
    + '</iconSet></cfRule></conditionalFormatting>');
  wb.pkg.write_('xl/worksheets/sheet1.xml', spliced);
  const view = SheetView.open(wb.save());

  const cells = new Map(view.render().cells.map((c) => [c.ref, c]));
  assert.deepEqual(cells.get('A1').icon, { set: '3Arrows', index: 0, count: 3 }, 'the low value earns the down arrow');
  assert.deepEqual(cells.get('A2').icon, { set: '3Arrows', index: 1, count: 3 });
  assert.deepEqual(cells.get('A3').icon, { set: '3Arrows', index: 2, count: 3 }, 'the high value earns the up arrow');
  assert.equal(cells.get('A1').text, '10', 'the number still shows beside the icon');

  // Painting is a reading: the parts are untouched, so fidelity holds.
  assert.equal(view.isDirty, false);
});

test('showValue="0" hides the number, reverse flips the order, the value itself is untouched', () => {
  const wb = Workbook.open(simpleBook([[10], [90]]));
  const spliced = wb.pkg.text('xl/worksheets/sheet1.xml').replace('</sheetData>', '</sheetData>'
    + '<conditionalFormatting sqref="A1:A2"><cfRule type="iconSet" priority="1">'
    + '<iconSet iconSet="3TrafficLights1" showValue="0" reverse="1">'
    + '<cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/>'
    + '</iconSet></cfRule></conditionalFormatting>');
  wb.pkg.write_('xl/worksheets/sheet1.xml', spliced);
  const view = SheetView.open(wb.save());

  const cells = new Map(view.render().cells.map((c) => [c.ref, c]));
  assert.equal(cells.get('A1').text, '', 'showValue 0 paints the icon alone');
  assert.equal(cells.get('A1').icon.index, 2, 'reverse: the LOW value wears the last icon');
  assert.equal(cells.get('A2').icon.index, 0);
  assert.equal(view.displayValue(0, 0).text, '10', 'only the paint hides it - the value stands');
  assert.equal(view.editValue(0, 0), '10');
});

test('a pivot with TWO row fields nests them in order - the path the panel now offers', () => {
  const view = SheetView.open(buildXlsx({
    sheets: [{
      name: 'Data',
      rows: [
        ['Region', 'Product', 'Rep', 'Qty'],
        ['North', 'Widget', 'Amy', 5],
        ['North', 'Widget', 'Bob', 3],
        ['North', 'Gadget', 'Amy', 2],
        ['South', 'Widget', 'Bob', 7],
      ],
    }],
  }));
  view.createPivot({
    name: 'Nested',
    source: { top: 0, left: 0, bottom: 4, right: 3 },
    target: { sheet: 'Data', row: 6, col: 0 },
    rowFields: ['Region', 'Product'],
    colFields: ['Rep'],
    dataFields: [{ field: 'Qty' }],
  });

  const p = view.pivots();
  assert.equal(p.length, 1);
  assert.deepEqual(p[0].rowFields.map((f) => p[0].cacheFields[f].name), ['Region', 'Product'],
    'both row fields, in the order the picker gave them');

  // The grand total is the whole Qty column, wherever the layout put it.
  const texts = view.render().cells.filter((c) => c.row >= 6).map((c) => c.text);
  assert.ok(texts.includes('17'), 'the grand total appears in the computed grid');
  assert.ok(texts.some((t) => t.includes('Widget') || t.includes('Gadget')),
    'the second row field labels its groups');

  // And the whole apparatus reopens from the saved file.
  const reopened = SheetView.open(view.save());
  assert.equal(reopened.pivots()[0].unsupported, null);
  reopened.refreshPivot('Nested');
});

// ------------------------------------------ the t="array" spill marker --

test('a saved spill carries the array marker, and it follows the spill as it grows', () => {
  const view = SheetView.open(simpleBook([[1, 2, 3]]));
  view.setCell(1, 0, '=TRANSPOSE(A1:C1)'); // A2 spills A2:A4

  let xml = Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');
  assert.match(xml, /<f t="array" ref="A2:A4">TRANSPOSE\(A1:C1\)<\/f>/,
    'the anchor writes the ARRAY form - one meaning in every Excel');
  assert.match(Workbook.open(view.save()).pkg.text('xl/workbook.xml'), /fullCalcOnLoad="1"/,
    'ghosts cache no values, so the file asks the next consumer to compute');

  view.setCell(1, 0, '=TRANSPOSE(A1:B1)'); // narrower source - shorter spill
  xml = Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');
  assert.match(xml, /<f t="array" ref="A2:A3">/, 'the ref follows the spill');

  view.setCell(1, 0, '=A1*2'); // no longer an array at all
  xml = Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /t="array"/, 'a scalar formula sheds the marker');
});

test('a saved spill REOPENS as a spill - the ghosts do not block what made them', () => {
  const view = SheetView.open(simpleBook([[5, 6, 7]]));
  view.setCell(1, 0, '=TRANSPOSE(A1:C1)');
  assert.equal(view.calc.getValue('Data', 3, 0), 7, 'spilled before the round trip');

  const reopened = SheetView.open(view.save());
  assert.equal(reopened.calc.getValue('Data', 1, 0), 5);
  assert.equal(reopened.calc.getValue('Data', 3, 0), 7, 'and after it');
  assert.equal(reopened.editValue(2, 0), '', 'a ghost is still a ghost - no input of its own');
});

test('an Excel-written array file loads without its cached ghost values blocking the array', () => {
  // Excel writes the anchor with t="array" AND caches every covered cell's
  // value as a plain <c><v>. Loading those as constants would turn the file
  // into #SPILL! - they are results, not inputs.
  const wb = Workbook.open(simpleBook([[10, 20]]));
  const xml = wb.pkg.text('xl/worksheets/sheet1.xml')
    .replace(/<row r="1"[^>]*>/, '$&')
    .replace('</sheetData>',
      '<row r="2"><c r="A2"><f t="array" ref="A2:B2">TRANSPOSE(TRANSPOSE(A1:B1))</f><v>10</v></c>'
      + '<c r="B2"><v>20</v></c></row></sheetData>');
  wb.pkg.write_('xl/worksheets/sheet1.xml', xml);

  const view = SheetView.open(wb.save());
  assert.equal(view.calc.getValue('Data', 1, 0), 10, 'the anchor computes');
  assert.equal(view.calc.getValue('Data', 1, 1), 20, 'the covered cell is the SPILL, not a constant');
  assert.equal(view.editValue(1, 1), '', 'so it has no input of its own');
  view.setCell(0, 1, 99);
  assert.equal(view.calc.getValue('Data', 1, 1), 99, 'and it tracks the source like the living formula it is');
});

// -------------------------------------------- the {=TABLE()} marker --

test('a data table writes the dataTable marker, so Excel gets a LIVE table', () => {
  const view = SheetView.open(simpleBook([[10], [], ['', '=A1*2'], [5], [7]]));
  const out = view.dataTable({ range: { top: 2, left: 0, bottom: 4, right: 1 }, colInput: 'A1' });
  assert.equal(out.live, true);

  assert.equal(view.displayValue(3, 1).text, '10', 'candidate 5 through =A1*2');
  assert.equal(view.displayValue(4, 1).text, '14');

  const xml = Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');
  assert.match(xml, /<c r="B4"[^>]*><f t="dataTable" ref="B4:B5" dt2D="0" dtr="0" r1="A1"\/><v>10<\/v><\/c>/,
    'the master carries the marker AND its cached value');
  assert.match(Workbook.open(view.save()).pkg.text('xl/workbook.xml'), /fullCalcOnLoad="1"/);

  // The marked file REOPENS as values - the marker is Excel's to compute,
  // and a bodiless formula must never come back as a garbage '=' input.
  const reopened = SheetView.open(view.save());
  assert.equal(reopened.displayValue(3, 1).text, '10');
  assert.equal(reopened.editValue(3, 1), '10', 'the master reads as its value, not as a formula');

  view.undo();
  assert.equal(view.displayValue(3, 1).text, '', 'undo takes the whole table away');
  assert.doesNotMatch(Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml'), /dataTable/,
    'marker included');
});

test('a two-way data table writes dt2D with both input cells', () => {
  const view = SheetView.open(simpleBook([
    [2, 3],            // A1 = row input, B1 = col input
    ['=A1*B1', 10, 20], // corner formula, candidates across
    [100], [200],       // candidates down
  ]));
  view.dataTable({ range: { top: 1, left: 0, bottom: 3, right: 2 }, rowInput: 'A1', colInput: 'B1' });
  assert.equal(view.displayValue(2, 1).text, '1000', '10 x 100');
  assert.equal(view.displayValue(3, 2).text, '4000', '20 x 200');

  const xml = Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');
  assert.match(xml, /<f t="dataTable" ref="B3:C4" dt2D="1" dtr="0" r1="A1" r2="B1"\/>/);
});

test('a sparse table - a non-formula column in the top row - stays an honest snapshot', () => {
  const view = SheetView.open(simpleBook([[10], [], ['', '=A1*2', 'note'], [5]]));
  const out = view.dataTable({ range: { top: 2, left: 0, bottom: 3, right: 2 }, colInput: 'A1' });
  assert.equal(out.live, false, 'the computed cells do not fill the interior, so no marker');
  assert.doesNotMatch(Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml'), /dataTable/);
  assert.equal(view.displayValue(3, 1).text, '10', 'the values still land');
});

// -------------------------------------- inbound clipboard formatting --

test('a foreign HTML paste brings its formatting: bold, colours, fills, alignment', () => {
  const view = SheetView.open(simpleBook([['x']]));
  view.select(2, 0);
  view.pasteText('Total\t120\nnote\t7',
    '<table><tr>'
    + '<td style="font-weight:bold;color:#C00000;text-align:right">Total</td>'
    + '<td style="background-color:#FFF2CC">120</td></tr>'
    + '<tr><td><i>note</i></td><td>7</td></tr></table>');

  assert.equal(view.calc.getValue('Data', 2, 0), 'Total');
  assert.equal(view.calc.getValue('Data', 2, 1), 120, 'numbers coerce as they always did');

  const a3 = view.styleFor(2, 0);
  assert.equal(a3.font.bold, true, 'font-weight:bold landed');
  assert.equal(a3.font.colour.toLowerCase(), '#c00000', 'the colour landed');
  assert.equal(a3.align.horizontal, 'right', 'the alignment landed');
  assert.equal(view.styleFor(2, 1).fill.colour.toLowerCase(), '#fff2cc', 'the fill landed');
  assert.equal(view.styleFor(3, 0).font.italic, true, 'the <i> tag form works too');
  assert.equal(view.styleFor(3, 1)?.font?.bold ?? false, false, 'a plain cell stays plain');

  // One undo takes the paste away - values and looks together.
  view.undo();
  assert.equal(view.calc.getValue('Data', 2, 0) ?? '', '', 'the value went');
  assert.equal(view.styleFor(2, 0)?.font?.bold ?? false, false, 'and the look went with it');
});

test('a white background and a default colour are noise, not choices - not written', () => {
  const view = SheetView.open(simpleBook([['x']]));
  view.select(1, 0);
  assert.equal(view.pkg.has('xl/styles.xml'), false, 'this workbook ships without a styles part');
  view.pasteText('a\tb',
    '<table><tr><td style="background-color:#ffffff;color:#000000">a</td><td>b</td></tr></table>');
  assert.equal(view.pkg.has('xl/styles.xml'), false,
    'nothing format-shaped was said, so no styles part was conjured to say it');
});

// -------------------------------------------------- timePeriod rules --

test('timePeriod rules fire on the clock the view was given - the last CF kind paints', async () => {
  const { dateToSerial } = await import('@rutba/formula');
  // A fixed Wednesday: 2026-09-02. Weeks for these rules run Sunday-Saturday,
  // so this week began Sunday 2026-08-30.
  const now = () => new Date(Date.UTC(2026, 8, 2, 10, 30));
  const today = Math.floor(dateToSerial(new Date(Date.UTC(2026, 8, 2))));
  const view = new SheetView(simpleBook([
    [today], [today - 1], [today + 1],   // A1 today, A2 yesterday, A3 tomorrow
    [today - 6], [today - 8],            // A4 in the last 7 days, A5 out of them
    [today - 2], [today - 4],            // A6 Monday of THIS week, A7 Saturday of LAST week
    [today - 25],                        // A8 2026-08-08 - last month
  ]), { now });
  view.styles.dxfs = [{ font: null, fill: { pattern: 'solid', colour: '#ffeb9c' } }];
  const rule = (timePeriod) => ({
    type: 'timePeriod', timePeriod, priority: 1, dxfId: 0,
    formulas: [], cfvos: [], colours: [], ranges: [{ top: 0, left: 0, bottom: 7, right: 0 }],
    stopIfTrue: false,
  });
  const firesAt = (timePeriod, row) => {
    view.conditionals.set('Data', [rule(timePeriod)]);
    const cell = view.render().cells.find((c) => c.row === row && c.col === 0);
    return Boolean(cell?.style?.fill);
  };

  assert.equal(firesAt('today', 0), true);
  assert.equal(firesAt('today', 1), false);
  assert.equal(firesAt('yesterday', 1), true);
  assert.equal(firesAt('tomorrow', 2), true);
  assert.equal(firesAt('last7Days', 3), true, 'six days back is inside the window');
  assert.equal(firesAt('last7Days', 4), false, 'eight days back is not');
  assert.equal(firesAt('thisWeek', 5), true, 'Monday belongs to the Sunday-bounded week');
  assert.equal(firesAt('thisWeek', 6), false, 'Saturday before it does not');
  assert.equal(firesAt('lastWeek', 6), true);
  assert.equal(firesAt('lastMonth', 7), true);
  assert.equal(firesAt('thisMonth', 0), true);
  assert.equal(firesAt('nextMonth', 0), false);
});

// ------------------------------------------ inserting shapes and charts --

test('a shape inserts over the selection, paints, round-trips and undoes clean', () => {
  const view = SheetView.open(simpleBook([['x']]));
  view.select(1, 1);
  view.select(4, 4, { extend: true });
  view.insertShape({ geometry: 'roundRect', text: 'Approved' });

  // The whole apparatus exists: part, content type, rels entry, sheet ref.
  const saved = Workbook.open(view.save());
  assert.ok(saved.pkg.has('xl/drawings/drawing1.xml'));
  const drawing = saved.pkg.text('xl/drawings/drawing1.xml');
  assert.match(drawing, /<a:prstGeom prst="roundRect">/);
  assert.match(drawing, /<a:t>Approved<\/a:t>/);
  assert.match(saved.pkg.text('xl/worksheets/sheet1.xml'), /<drawing r:id="rId1"/);
  assert.match(saved.pkg.text('[Content_Types].xml'), /drawing\+xml/);

  // It paints - the same read path that has always drawn customer shapes.
  const painted = view.render().drawings.find((d) => d.svg && /Shape/.test(d.name ?? ''));
  assert.ok(painted, 'the inserted shape reaches the frame with an svg');

  // The reopened file still carries it, and a second shape shares the part.
  const reopened = SheetView.open(view.save());
  reopened.insertShape({ geometry: 'star5' });
  const two = Workbook.open(reopened.save()).pkg.text('xl/drawings/drawing1.xml');
  assert.match(two, /roundRect/);
  assert.match(two, /star5/);

  // Undo removes the parts it created - no ghost drawing behind a live rel.
  view.undo();
  assert.equal(view.pkg.has('xl/drawings/drawing1.xml'), false, 'the created part went');
  assert.doesNotMatch(Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml'), /<drawing /,
    'and the sheet no longer references it');
  assert.equal(view.render().drawings.length, 0, 'nothing paints');
  view.redo();
  assert.ok(view.pkg.has('xl/drawings/drawing1.xml'), 'redo brings the whole apparatus back');
  assert.equal(view.render().drawings.length, 1);
});

test('a chart inserts from the data region: real references, matching caches, painted', () => {
  const view = SheetView.open(simpleBook([
    ['Month', 'Sales', 'Cost'],
    ['Jan', 100, 60],
    ['Feb', 120, 80],
  ]));
  view.select(1, 1); // anywhere in the block
  view.insertChart({ title: 'Performance' });

  const saved = Workbook.open(view.save());
  const chart = saved.pkg.text('xl/charts/chart1.xml');
  assert.match(chart, /<c:f>Data!\$A\$2:\$A\$3<\/c:f>/, 'categories reference the first column');
  assert.match(chart, /<c:f>Data!\$B\$2:\$B\$3<\/c:f>/, 'the first series references its column');
  assert.match(chart, /<c:v>120<\/c:v>/, 'the cache matches the sheet');
  assert.match(chart, /<c:axId /, 'the axes Excel requires are there');
  assert.match(saved.pkg.text('xl/drawings/_rels/drawing1.xml.rels'), /chart1\.xml/);

  const painted = view.render().drawings.find((d) => d.svg);
  assert.ok(painted, 'the chart paints from its own part');

  view.undo();
  assert.equal(view.pkg.has('xl/charts/chart1.xml'), false);
  assert.equal(view.pkg.has('xl/drawings/drawing1.xml'), false);
});

test('inserting refuses what it cannot honestly do', () => {
  const view = SheetView.open(simpleBook([['lonely']]));
  assert.throws(() => view.insertShape({ geometry: 'cloudCallout' }), /not a shape this editor draws/);
  view.select(5, 5);
  assert.throws(() => view.insertChart({}), /block of data/);
  const protectedView = SheetView.open(buildComplexWorkbook());
  assert.throws(() => protectedView.insertShape({ geometry: 'rect' }), /protected/);
  assert.equal(protectedView.history.past.length, 0, 'a refusal leaves no undo step');
});

test('every chart kind inserts through the one writer, paints, and reopens', () => {
  const KINDS = [
    ['bar', /<c:barChart><c:barDir val="bar"\/>/],
    ['line', /<c:lineChart>/],
    ['area', /<c:areaChart>/],
    ['pie', /<c:pieChart>/],
    ['doughnut', /<c:doughnutChart>[\s\S]*<c:holeSize/],
  ];
  for (const [kind, pattern] of KINDS) {
    const view = SheetView.open(simpleBook([
      ['Month', 'Sales'], ['Jan', 100], ['Feb', 120],
    ]));
    view.select(0, 0);
    view.insertChart({ kind, title: kind + ' test' });
    const xml = Workbook.open(view.save()).pkg.text('xl/charts/chart1.xml');
    assert.match(xml, pattern, kind + ' writes its own plot element');
    const painted = SheetView.open(view.save()).render().drawings.find((d) => d.svg);
    assert.ok(painted, kind + ' paints after a round trip');
  }
  const view = SheetView.open(simpleBook([['a', 1], ['b', 2]]));
  view.select(0, 0);
  assert.throws(() => view.insertChart({ kind: 'radar' }), /not a chart kind/);
});

// ------------------------------------- authoring conditional formatting --

test('a greater-than rule authored over the selection paints, saves, and undoes clean', () => {
  const view = SheetView.open(simpleBook([[50], [150], [99]]));
  view.select(0, 0);
  view.select(2, 0, { extend: true });
  // The styles part is ENSURED outside the edit (an empty style table is
  // harmless furniture, not undoable work), so the byte comparison starts
  // after it exists.
  view._ensureStylesPart();
  const before = view.save();
  view.addConditionalRule({ kind: 'cellIs', operator: 'greaterThan', values: ['100'], look: 'red' });

  // It PAINTS immediately - the same evaluator that has always painted rules.
  const cells = new Map(view.render().cells.map((c) => [c.ref, c]));
  assert.equal(cells.get('A2').style.fill.colour.toLowerCase(), '#ffc7ce', '150 fires');
  assert.equal(cells.get('A1')?.style?.fill ?? null, null, '50 does not');

  // The file carries exactly what Excel's dialog writes.
  const saved = Workbook.open(view.save());
  assert.match(saved.pkg.text('xl/worksheets/sheet1.xml'),
    /<conditionalFormatting sqref="A1:A3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100<\/formula><\/cfRule><\/conditionalFormatting>/);
  assert.match(saved.pkg.text('xl/styles.xml'),
    /<dxfs count="1"><dxf><font><color rgb="FF9C0006"\/><\/font><fill><patternFill><bgColor rgb="FFFFC7CE"\/><\/patternFill><\/fill><\/dxf><\/dxfs>/);

  // Undo restores both parts byte-identically.
  view.undo();
  assert.deepEqual([...view.save()], [...before], 'the rule and its dxf both left');
});

test('text-contains, duplicates, colour scale and data bar all author their Excel shapes', () => {
  const view = SheetView.open(simpleBook([['alpha', 1], ['beta', 5], ['alpha', 9]]));
  view.select(0, 0);
  view.select(2, 0, { extend: true });
  view.addConditionalRule({ kind: 'containsText', text: 'alp', look: 'yellow' });
  view.addConditionalRule({ kind: 'duplicates', look: 'green' });
  view.select(0, 1);
  view.select(2, 1, { extend: true });
  view.addConditionalRule({ kind: 'colorScale' });
  view.addConditionalRule({ kind: 'dataBar' });

  const xml = Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');
  assert.match(xml, /type="containsText"[^>]*operator="containsText" text="alp"/);
  assert.match(xml, /<formula>NOT\(ISERROR\(SEARCH\("alp",A1\)\)\)<\/formula>/);
  assert.match(xml, /<cfRule type="duplicateValues" dxfId="1" priority="2"\/>/);
  assert.match(xml, /<colorScale><cfvo type="min"\/><cfvo type="percentile" val="50"\/><cfvo type="max"\/>/);
  assert.match(xml, /<dataBar><cfvo type="min"\/><cfvo type="max"\/><color rgb="FF638EC6"\/><\/dataBar>/);

  const cells = new Map(view.render().cells.map((c) => [c.ref, c]));
  assert.ok(cells.get('A1').style.fill, 'contains-text fires on alpha');
  assert.ok(cells.get('B2').bar, 'the data bar rides the frame');

  // Distinct looks earned distinct dxf entries; a repeat would reuse.
  assert.match(Workbook.open(view.save()).pkg.text('xl/styles.xml'), /<dxfs count="2">/);
});

test('clear rules: inside the selection, or the whole sheet - and refusals are honest', () => {
  const view = SheetView.open(simpleBook([[1], [2], [3], [4]]));
  view.select(0, 0);
  view.select(1, 0, { extend: true });
  view.addConditionalRule({ kind: 'cellIs', operator: 'greaterThan', values: [1], look: 'red' });
  view.select(2, 0);
  view.select(3, 0, { extend: true });
  view.addConditionalRule({ kind: 'cellIs', operator: 'lessThan', values: [4], look: 'green' });

  view.select(0, 0);
  view.select(1, 0, { extend: true });
  assert.equal(view.clearConditionalRules(), 1, 'only the block inside the selection went');
  assert.match(Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml'), /lessThan/);

  view.clearConditionalRules({ all: true });
  assert.doesNotMatch(Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml'), /<conditionalFormatting/);

  assert.throws(() => view.addConditionalRule({ kind: 'cellIs', operator: 'greaterThan', values: [], look: 'red' }),
    /needs a value/);
  assert.throws(() => view.addConditionalRule({ kind: 'nope' }), /not a rule/);
  assert.equal(view.history.past.length, 4, 'refusals recorded nothing new');
});

test('authoring onto the protected fixture refuses before anything records', () => {
  const view = SheetView.open(buildComplexWorkbook());
  assert.throws(() => view.addConditionalRule({ kind: 'dataBar' }), /protected/);
  assert.throws(() => view.clearConditionalRules({ all: true }), /protected/);
  assert.equal(view.history.past.length, 0);
});

// ------------------------------------------ authoring data validation --

test('an authored list rule enforces immediately, offers its dropdown, and undoes clean', () => {
  const view = SheetView.open(simpleBook([['pick']]));
  view.select(1, 0);
  const before = view.save();
  view.addValidationRule({ kind: 'list', items: ['Red', 'Amber', 'Green'], error: 'Pick a colour from the list' });

  const xml = Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');
  assert.match(xml, /<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" error="Pick a colour from the list" sqref="A2"><formula1>"Red,Amber,Green"<\/formula1><\/dataValidation><\/dataValidations>/);

  // Enforcement is live - the same gate that has always guarded direct entry.
  assert.match(view.checkValidation(1, 0, 'Blue'), /Pick a colour/);
  assert.equal(view.checkValidation(1, 0, 'Amber'), null);
  view.undo(); // the rule itself
  assert.deepEqual([...view.save()], [...before], 'the rule left byte-identically');
});

test('whole-number rules write their operator forms, and a second rule joins the one block', () => {
  const view = SheetView.open(simpleBook([[1, 2]]));
  view.select(1, 0);
  view.addValidationRule({ kind: 'whole', operator: 'between', values: [1, 10] });
  view.select(1, 1);
  view.addValidationRule({ kind: 'whole', operator: 'greaterThanOrEqual', values: [0], error: 'No negatives' });

  const xml = Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');
  assert.match(xml, /<dataValidations count="2">/, 'one block, two entries');
  assert.match(xml, /type="whole" allowBlank[^>]*sqref="A2"><formula1>1<\/formula1><formula2>10<\/formula2>/);
  assert.match(xml, /operator="greaterThanOrEqual"[^>]*sqref="B2"><formula1>0<\/formula1>/);

  assert.notEqual(view.checkValidation(1, 0, '11'), null, '11 is outside 1..10');
  assert.match(view.checkValidation(1, 1, '-1'), /No negatives/);
  assert.equal(view.checkValidation(1, 1, '5'), null);
});

test('clear validation: in the selection, on the sheet, and the honest refusals', () => {
  const view = SheetView.open(simpleBook([[1, 2]]));
  view.select(1, 0);
  view.addValidationRule({ kind: 'whole', operator: 'between', values: [1, 10] });
  view.select(1, 1);
  view.addValidationRule({ kind: 'list', items: ['a'] });

  view.select(1, 0);
  assert.equal(view.clearValidationRules(), 1);
  const xml = Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');
  assert.match(xml, /<dataValidations count="1">/, 'the block recounts');
  assert.doesNotMatch(xml, /type="whole"/);

  view.clearValidationRules({ all: true });
  assert.doesNotMatch(Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml'),
    /<dataValidations/, 'an emptied block leaves entirely');

  assert.throws(() => view.addValidationRule({ kind: 'list', items: [] }), /needs its choices/);
  assert.throws(() => view.addValidationRule({ kind: 'whole', operator: 'between', values: [1] }), /two numbers/);
  assert.throws(() => view.addValidationRule({ kind: 'time' }), /not a rule/);

  const sealed = SheetView.open(buildComplexWorkbook());
  assert.throws(() => sealed.addValidationRule({ kind: 'list', items: ['x'] }), /protected/);
  assert.equal(sealed.history.past.length, 0);
});
