// Words with looks of their own in a cell: read from a rich string, written
// back as one, kept when the words change, given the cell's Bold, and left
// as they are by Save.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook, readRuns } from '../packages/ooxml/src/workbook.js';
import { buildXlsx } from '../packages/ooxml/src/build.js';
import { SheetView } from '../packages/sheet-view/src/index.js';

const RICH = '<r><t>Plain </t></r><r><rPr><i/><color rgb="FFFF0000"/><sz val="12"/></rPr><t>red</t></r><r><rPr><b/><color theme="1"/></rPr><t xml:space="preserve"> end</t></r>';

test('a rich string read as runs: italic, colour, size, bold, a theme colour kept as written', () => {
  assert.deepEqual(readRuns(RICH), [
    { text: 'Plain ' },
    { text: 'red', italic: true, size: 12, color: '#FF0000' },
    { text: ' end', bold: true, colorXml: '<color theme="1"/>' },
  ]);
  assert.equal(readRuns('<t>just words</t>'), null);
});

/** A workbook whose A1 holds RICH as an inline string. */
function richBook() {
  const wb = Workbook.open(buildXlsx({ sheets: [{ name: 'S', rows: [['x', 'y']] }] }));
  wb.setCellRuns('S', 0, 0, readRuns(RICH));
  return wb.save();
}

test('runs written into a cell read back the same, the theme colour as it was', () => {
  const wb = Workbook.open(richBook());
  assert.deepEqual(wb.cellRuns('S', 0, 0), readRuns(RICH));
  assert.equal(wb.getCell('S', 'A1'), 'Plain red end');
  assert.equal(wb.cellRuns('S', 0, 1), null);
});

test('words retyped keep each run round the change; Bold on the cell reaches every run; Save leaves them; Undo puts them back', () => {
  const v = SheetView.open(Buffer.from(richBook()));
  v.setCell(0, 0, 'Plain red ending');
  const typed = v._cellRuns('S', 0, 0);
  assert.deepEqual(typed.map((r) => [r.text, Boolean(r.italic), Boolean(r.bold)]), [['Plain ', false, false], ['red', true, false], [' ending', false, true]]);
  v.select(0, 0);
  v.setFormat({ bold: true });
  assert.ok(v._cellRuns('S', 0, 0).every((r) => r.bold));
  const back = SheetView.open(Buffer.from(v.save()));
  assert.deepEqual(back._cellRuns('S', 0, 0).map((r) => [r.text, Boolean(r.italic), Boolean(r.bold)]), [['Plain ', false, true], ['red', true, true], [' ending', false, true]]);
  v.undo();
  v.undo();
  assert.equal(v._cellRuns('S', 0, 0).map((r) => r.text).join(''), 'Plain red end');
});

test('a number or a formula typed over rich words is no longer words with runs', () => {
  const v = SheetView.open(Buffer.from(richBook()));
  v.setCell(0, 0, '42');
  const back = SheetView.open(Buffer.from(v.save()));
  assert.equal(back._cellRuns('S', 0, 0), null);
  assert.equal(back.displayValue(0, 0).text, '42');
});
