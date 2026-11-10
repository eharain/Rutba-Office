// Rows and columns inserted and deleted here against what Excel did to the
// same workbook (tools/excel-structure.mjs wrote down Excel's results).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook, recalculateWorkbook } from '@rutba/ooxml';
import { EDITS, applyEdit, describe } from '../tools/excel-structure.mjs';

const bytes = fs.readFileSync(new URL('./fixtures/excel-structure.xlsx', import.meta.url));
const answers = JSON.parse(fs.readFileSync(new URL('./fixtures/excel-structure.json', import.meta.url), 'utf8'));

test('the workbook Excel wrote reads the same here', () => {
  assert.deepEqual(JSON.parse(JSON.stringify(describe(bytes))), answers.base);
});

EDITS.forEach((edit, i) => {
  test(`as Excel: ${edit.name}`, () => {
    const ours = JSON.parse(JSON.stringify(describe(applyEdit(bytes, edit))));
    assert.deepEqual(ours, answers.edits[i].excel);
  });
});

// What the comparison above cannot reach: files Excel would write but this
// fixture does not hold, read back by our own reader.

/** The fixture with its second sheet's cells replaced by the given XML. */
function withCells(sheetData) {
  const wb = Workbook.open(bytes);
  const part = wb.partNameFor('Other');
  wb.pkg.write_(part, wb.pkg.text(part).replace(/<sheetData>[\s\S]*<\/sheetData>/, () => '<sheetData>' + sheetData + '</sheetData>'));
  return wb.save();
}
const SHARED = '<row r="1"><c r="A1"><v>1</v></c><c r="C1"><f t="shared" ref="C1:C3" si="9">A1*10</f><v>10</v></c></row>'
  + '<row r="2"><c r="A2"><v>2</v></c><c r="C2"><f t="shared" si="9"/><v>20</v></c><c r="D2"><f>A2+1</f><v>3</v></c></row>'
  + '<row r="3"><c r="A3"><v>3</v></c><c r="C3"><f t="shared" si="9"/><v>30</v></c></row>';
const cellsOf = (wb, name) => Object.fromEntries([...wb.cells(name)].map((c) => [c.ref, c.value]));

test('a shared child is a formula and follows the cell it reads', () => {
  const wb = Workbook.open(withCells(SHARED));
  assert.equal(cellsOf(wb, 'Other').C3, '=A3*10');
  wb.setCell('Other', 'A3', 100);
  const { output } = recalculateWorkbook(wb.save());
  const part = Workbook.open(output);
  const xml = part.pkg.text(part.partNameFor('Other'));
  assert.match(/<c r="C3"[^>]*>[\s\S]*?<\/c>/.exec(xml)[0], /<v>1000<\/v>/);
});

test('inserting above a shared child leaves the next formula its own cell', () => {
  const wb = Workbook.open(withCells(SHARED));
  wb.insertRows('Other', 0, 1);
  const again = Workbook.open(wb.save());
  const xml = again.pkg.text(again.partNameFor('Other'));
  assert.equal(xml.includes('&lt;v&gt;'), false);
  const cells = cellsOf(again, 'Other');
  assert.equal(cells.C3, '=A3*10');
  assert.equal(cells.D3, '=A3+1');
  assert.equal(cells.C4, '=A4*10');
});

test('inserting a column keeps identical styled blanks in their own places', () => {
  const wb = Workbook.open(withCells('<row r="1"><c r="A1" s="1"/><c r="B1" s="1"/><c r="C1" s="1"/></row>'));
  wb.insertCols('Other', 0, 1);
  const again = Workbook.open(wb.save());
  const xml = again.pkg.text(again.partNameFor('Other'));
  assert.deepEqual([...xml.matchAll(/<c r="([A-Z]+1)"/g)].map((m) => m[1]), ['B1', 'C1', 'D1']);
});

test('a table deleted whole leaves no link, part or content type, and undo brings them back', () => {
  const wb = Workbook.open(bytes);
  const [t] = wb.tables();
  const parts = Object.fromEntries(wb.structuralParts().map((n) => [n, wb.pkg.text(n)]));
  wb._removeTablePart(t);
  assert.equal(wb.pkg.has(t.part), false);
  assert.equal(wb.tables().length, 0);
  assert.equal(wb.pkg.text('[Content_Types].xml').includes(t.part), false);
  assert.equal(wb.pkg.rels(wb.partNameFor(t.sheet)).some((r) => String(r.Type).endsWith('/table')), false);
  wb.restoreParts(parts);
  assert.deepEqual(wb.tables().map((x) => x.ref), [t.ref]);
  assert.ok(wb.pkg.text('[Content_Types].xml').includes(t.part));
});

test('an edit drops the calculation list that names cells which moved, link and content type too', () => {
  assert.ok(Workbook.open(bytes).pkg.has('xl/calcChain.xml'), 'Excel wrote one into the fixture');
  const untouched = Workbook.open(Workbook.open(bytes).save());
  assert.ok(untouched.pkg.has('xl/calcChain.xml'), 'a workbook nobody edited keeps it');
  const after = Workbook.open(applyEdit(bytes, EDITS[0]));
  assert.equal(after.pkg.has('xl/calcChain.xml'), false);
  assert.equal(after.pkg.text('[Content_Types].xml').includes('calcChain'), false);
  assert.equal(after.pkg.rels('xl/workbook.xml').some((r) => /calcChain/.test(r.Target)), false);
});
