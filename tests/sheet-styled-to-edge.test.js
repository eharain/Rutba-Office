/**
 * A sheet styled to its edge, written in one pass.
 *
 * An .ods from a statistics office shades and borders its tables across the
 * sheet's whole width and sizes a thousand columns. Opening one took from
 * twenty seconds to two minutes: each column width re-read every column
 * before it, each cell's number format was looked for among every other
 * cell's, and each cell's look re-parsed its row. The widths and heights are
 * now written with the sheet, a cell's own style is looked up, and a row's
 * looks are written in one pass. What the workbook holds is the same.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';

const sheetXml = (bytes) => OoxmlPackage.read(Buffer.from(bytes)).text('xl/worksheets/sheet1.xml');
const styleOf = (xml, ref) => (new RegExp(`<c r="${ref}"[^>]*?\\bs="(\\d+)"`).exec(xml) || [, '0'])[1];

test('row heights are written with the sheet, a tall empty row past the last row\'s words too', () => {
  const xml = sheetXml(buildXlsx({ sheets: [{ name: 'S', rows: [['a'], [], ['c']], heights: { 0: 30, 1: 18.75, 5: 42 } }] }));
  assert.match(xml, /<row r="1" ht="30" customHeight="1">/);
  assert.match(xml, /<row r="2" ht="18.75" customHeight="1"><\/row>/, 'an empty row keeps its height');
  assert.match(xml, /<row r="3"><c r="A3"/, 'a row of no stated height has none');
  assert.match(xml, /<row r="6" ht="42" customHeight="1"><\/row><\/sheetData>/);
});

test('a width in 256ths of a character is written as it is; another is rounded as before', () => {
  const xml = sheetXml(buildXlsx({ sheets: [{ name: 'S', rows: [['a']], widths: { A: 9.140625, B: 12.3456 } }] }));
  assert.match(xml, /<col min="1" max="1" width="9.140625" customWidth="1"\/>/);
  assert.match(xml, /<col min="2" max="2" width="12.35" customWidth="1"\/>/);
});

test('a cell\'s own style and a range over it: the one declared last wins, either way round', () => {
  const pct = { numFmt: '0%' };
  const money = { numFmt: '#,##0.00' };
  const rows = [[1, 2], [3, 4]];
  const cellLast = sheetXml(buildXlsx({ sheets: [{ name: 'S', rows, styles: { 'A1:B2': money, A1: pct } }] }));
  const rangeLast = sheetXml(buildXlsx({ sheets: [{ name: 'S', rows, styles: { A1: pct, 'A1:B2': money } }] }));
  assert.notEqual(styleOf(cellLast, 'A1'), styleOf(cellLast, 'B1'), 'the cell declared after the range keeps its own');
  assert.equal(styleOf(cellLast, 'B2'), styleOf(cellLast, 'B1'));
  assert.equal(styleOf(rangeLast, 'A1'), styleOf(rangeLast, 'B1'), 'the range declared after the cell covers it');
});

test('a row\'s cells restyled in one pass: existing cells keep their contents, missing ones are made in column order', () => {
  const wb = Workbook.open(buildXlsx({ sheets: [{ name: 'S', rows: [[null, 'b', null, '=B1', null, 'f']] }] }));
  const { part } = wb._sheetPart('S');
  part.setCellStyles(0, new Map([[5, 3], [0, 2], [3, 4], [2, 1], [7, 6]]));
  const inner = part._rowAt(0).inner;
  assert.deepEqual([...inner.matchAll(/<c r="([A-Z]+)1"/g)].map((m) => m[1]), ['A', 'B', 'C', 'D', 'F', 'H'], 'in column order');
  assert.match(inner, /<c r="A1" s="2"\/>/);
  assert.match(inner, /<c r="C1" s="1"\/>/);
  assert.match(inner, /<c r="D1" s="4"><f>B1<\/f><\/c>/, 'the formula kept');
  assert.match(inner, /<c r="H1" s="6"\/>/);
  assert.equal(wb.getCell('S', 'B1'), 'b', 'an untouched cell untouched');
  assert.equal(wb.getCell('S', 'F1'), 'f');
  // One cell, as before: the same pass.
  part.setCellStyle(0, 1, 9);
  assert.match(part._rowAt(0).inner, /<c r="B1"[^>]*\bs="9"/);
  part.setCellStyle(0, 1, null);
  assert.doesNotMatch(part._rowAt(0).inner, /<c r="B1"[^>]*\bs=/, 'null takes the style away');
});
