/**
 * Page Layout → Sheet Right-to-Left: a sheet that reads from the right, as
 * an Arabic or Hebrew workbook's does — read from `<sheetView rightToLeft>`
 * and written there, the frame saying so for the window to draw the grid
 * mirrored, and a cell aligned left or right by its own format kept on that
 * side of the mirrored cell, general alignment mirroring with the sheet.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';

const open = (bytes) => new SheetView(bytes ?? buildXlsx({ sheets: [{ name: 'بيانات', rows: [['الاسم', 'المبلغ'], ['أحمد', 120], ['سارة', 75]] }, { name: 'Other', rows: [['x']] }] }), { viewportWidth: 900, viewportHeight: 600 });
const sheetXml = (view, n = 1) => OoxmlPackage.read(view.save()).text(`xl/worksheets/sheet${n}.xml`);

test('a sheet turned right to left writes rightToLeft on its sheetView, and only that sheet', () => {
  const view = open();
  assert.equal(view.rightToLeft(), false);
  assert.equal(view.render().rtl, false);
  view.setRightToLeft(true);
  assert.equal(view.rightToLeft(), true);
  assert.equal(view.render().rtl, true, 'the frame says so, for the window to mirror the grid');
  assert.match(sheetXml(view, 1), /<sheetView\b[^>]*\brightToLeft="1"/);
  assert.doesNotMatch(sheetXml(view, 2), /rightToLeft/, 'the other sheet still reads from the left');
  view.setRightToLeft(false);
  assert.doesNotMatch(sheetXml(view, 1), /rightToLeft/, 'the attribute goes, as Excel leaves it off');
});

test('a workbook saved right to left opens right to left, the sheet\'s own setting followed sheet by sheet', () => {
  const first = open();
  first.setRightToLeft(true);
  const again = open(first.save());
  assert.equal(again.rightToLeft(), true);
  again.selectSheet('Other');
  assert.equal(again.rightToLeft(), false);
  assert.equal(again.render().rtl, false);
});

test('right to left, a cell aligned left by its format stays on the left of the mirrored cell; general text and numbers mirror', () => {
  const view = open();
  view.select(1, 0);
  view.setFormat({ align: 'left' });
  view.select(2, 0);
  view.setFormat({ align: 'right' });
  view.setRightToLeft(true);
  const cell = (row, col) => view.render().cells.find((c) => c.row === row && c.col === col);
  // The grid is drawn mirrored: the engine's left is the screen's right.
  assert.equal(cell(1, 0).align, 'right', 'aligned left: drawn at the mirrored right, the screen\'s left');
  assert.equal(cell(2, 0).align, 'left', 'aligned right: the other way');
  assert.equal(cell(0, 0).align, 'left', 'general text: left in the engine, the right of the screen as Excel puts it');
  assert.equal(cell(1, 1).align, 'right', 'general numbers: right in the engine, the left of the screen');
  view.setRightToLeft(false);
  assert.equal(cell(1, 0).align, 'left');
  assert.equal(cell(2, 0).align, 'right');
});
