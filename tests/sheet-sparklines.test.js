// Insert → Sparklines: a line or column spark, written as Excel writes one —
// an `x14:sparklineGroups` extension at the end of the worksheet — merged
// into whatever `extLst` the sheet already carries, after whatever else is
// in it. One group holds one sparkline per row of the target cells (or per
// column, when the target is itself a row); each sparkline's own data range
// is what the values a cell shows come from.

import test from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildXlsx } from '@rutba/ooxml/build';
import { Workbook } from '@rutba/ooxml/workbook';
import { SheetView } from '@rutba/sheet-view';

const RICH = join(dirname(fileURLToPath(import.meta.url)), 'fixtures', 'rich');

const book = () => buildXlsx({
  sheets: [{
    name: 'Sales',
    rows: [
      ['Item', 'Q1', 'Q2', 'Q3', '', ''],
      ['Pens', 5, 8, 3, '', ''],
      ['Ink', 2, 6, 9, '', ''],
      ['Paper', 7, 1, 4, '', ''],
    ],
  }],
});

// `pkg.text` reads the stored bytes; a loaded part may hold un-flushed
// edits, so this reads it back the way `save()` would (see
// `Workbook.snapshotParts`).
const sheetXml = (wb, name = 'Sales') => {
  const part = wb.partNameFor(name);
  return wb.snapshotParts([part])[part];
};

test('a line group over three rows is written exactly as Excel writes one, at the end of the sheet, and reads back', () => {
  const wb = Workbook.open(book());
  wb.addSparklines('Sales', { type: 'line', data: 'B2:D4', at: 'E2:E4' });

  const xml = sheetXml(wb);
  assert.match(xml, /<\/sheetData><extLst><ext uri="\{05C60535-1F16-4fd2-B633-F4F36F0B64E0\}" xmlns:x14="http:\/\/schemas\.microsoft\.com\/office\/spreadsheetml\/2009\/9\/main"><x14:sparklineGroups xmlns:xm="http:\/\/schemas\.microsoft\.com\/office\/excel\/2006\/main">/,
    'the extension sits right after sheetData, with both namespaces Excel declares');
  assert.ok(xml.endsWith('</x14:sparklineGroups></ext></extLst></worksheet>'), 'extLst is the very last child');
  assert.match(xml, /<x14:sparklineGroup displayEmptyCellsAs="gap">/, 'a line sparkline carries no type= — line is the default');
  assert.match(xml, /<x14:colorSeries rgb="FF376092"\/><x14:colorNegative rgb="FFD00000"\/><x14:colorAxis rgb="FF000000"\/><x14:colorMarkers rgb="FFD00000"\/><x14:colorFirst rgb="FFD00000"\/><x14:colorLast rgb="FFD00000"\/><x14:colorHigh rgb="FFD00000"\/><x14:colorLow rgb="FFD00000"\/>/,
    "Excel's own defaults, nobody having chosen otherwise");
  assert.match(xml, /<x14:sparkline><xm:f>Sales!B2:D2<\/xm:f><xm:sqref>E2<\/xm:sqref><\/x14:sparkline>/);
  assert.match(xml, /<x14:sparkline><xm:f>Sales!B3:D3<\/xm:f><xm:sqref>E3<\/xm:sqref><\/x14:sparkline>/);
  assert.match(xml, /<x14:sparkline><xm:f>Sales!B4:D4<\/xm:f><xm:sqref>E4<\/xm:sqref><\/x14:sparkline>/);

  const groups = wb.sparklineGroups('Sales');
  assert.deepEqual(groups, [{
    type: 'line',
    colour: '376092',
    sparklines: [
      { data: 'Sales!B2:D2', at: 'E2' },
      { data: 'Sales!B3:D3', at: 'E3' },
      { data: 'Sales!B4:D4', at: 'E4' },
    ],
  }]);
});

test('values for a sparkline resolve through the view, cell by cell', () => {
  const view = new SheetView(book());
  view.select(1, 4); // E2
  view.addSparklines({ type: 'line', data: 'B2:D4', at: 'E2:E4' });

  const [group] = view.sparklineGroups('Sales');
  assert.equal(group.sparklines.length, 3);
  const at2 = group.sparklines.find((s) => s.at === 'E2');
  const [sheetName, rangeText] = at2.data.split('!');
  const [a, b] = rangeText.split(':');
  const values = [];
  const fromCol = a.charCodeAt(0) - 65;
  const toCol = b.charCodeAt(0) - 65;
  const row = Number(a.slice(1)) - 1;
  for (let c = fromCol; c <= toCol; c++) values.push(view.calc.getValue(sheetName, row, c));
  assert.deepEqual(values, [5, 8, 3], "E2's sparkline reads Pens's row");
});

test('a column group added beside it merges into the same extLst, as a second x14:sparklineGroup', () => {
  const wb = Workbook.open(book());
  wb.addSparklines('Sales', { type: 'line', data: 'B2:D4', at: 'E2:E4' });
  wb.addSparklines('Sales', { type: 'column', data: 'C2:D4', at: 'F2:F4' });

  const xml = sheetXml(wb);
  assert.equal((xml.match(/<extLst>/g) || []).length, 1, 'one extLst');
  assert.equal((xml.match(/<ext uri="\{05C60535-1F16-4fd2-B633-F4F36F0B64E0\}"/g) || []).length, 1, 'one sparkline ext');
  assert.equal((xml.match(/<x14:sparklineGroup\b/g) || []).length, 2, 'two groups in the one collection');
  assert.match(xml, /<x14:sparklineGroup type="column" displayEmptyCellsAs="gap">/, 'the column group says so');

  const groups = wb.sparklineGroups('Sales');
  assert.deepEqual(groups.map((g) => g.type), ['line', 'column']);
  assert.deepEqual(groups[1].sparklines, [
    { data: 'Sales!C2:D2', at: 'F2' },
    { data: 'Sales!C3:D3', at: 'F3' },
    { data: 'Sales!C4:D4', at: 'F4' },
  ]);
});

test('removal: one cell out of a group, then the emptied group, then the emptied ext', () => {
  const wb = Workbook.open(book());
  wb.addSparklines('Sales', { type: 'line', data: 'B2:D4', at: 'E2:E4' });
  wb.addSparklines('Sales', { type: 'column', data: 'B2:D4', at: 'F2:F4' });

  // One cell.
  let removed = wb.removeSparklines('Sales', 'E3');
  assert.equal(removed, 1);
  let groups = wb.sparklineGroups('Sales');
  assert.deepEqual(groups[0].sparklines.map((s) => s.at), ['E2', 'E4'], 'the other two of the line group stay');
  assert.equal(groups.length, 2, 'both groups still stand');

  // The rest of that group, by range — the group itself goes, the other rides through.
  removed = wb.removeSparklines('Sales', 'E1:E10');
  assert.equal(removed, 2);
  groups = wb.sparklineGroups('Sales');
  assert.deepEqual(groups.map((g) => g.type), ['column'], 'the emptied line group is gone; the column group is untouched');
  assert.match(sheetXml(wb), /<extLst>/, 'the extLst itself stays — the column group still needs it');

  // The last group, emptied — the whole extension goes with it.
  removed = wb.removeSparklines('Sales', 'F2:F4');
  assert.equal(removed, 3);
  assert.deepEqual(wb.sparklineGroups('Sales'), []);
  assert.doesNotMatch(sheetXml(wb), /<extLst>/, 'nothing left to carry an extLst for');

  // Removing where there is nothing reports nothing, and touches nothing.
  assert.equal(wb.removeSparklines('Sales', 'A1'), 0);
});

test('a sheet that already carries an extLst with another extension keeps it, before ours', () => {
  const wb = Workbook.open(book());
  const part = wb.partNameFor('Sales');
  const other = '<extLst><ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><x14ac:mark/></ext></extLst>';
  wb.pkg.write_(part, wb.pkg.text(part).replace('</worksheet>', other + '</worksheet>'));

  wb.addSparklines('Sales', { type: 'line', data: 'B2:D2', at: 'E2' });

  const xml = sheetXml(wb);
  assert.equal((xml.match(/<extLst>/g) || []).length, 1, 'merged into the ONE extLst, not a second one');
  const otherAt = xml.indexOf('{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}');
  const sparkAt = xml.indexOf('{05C60535-1F16-4fd2-B633-F4F36F0B64E0}');
  assert.ok(otherAt >= 0 && sparkAt > otherAt, 'the sparkline ext lands AFTER the one already there');
  assert.match(xml, /<x14ac:mark\/>/, 'the other extension rides through untouched');
});

test('a sparkline survives save() and reopening the file', () => {
  const view = new SheetView(book());
  view.addSparklines({ type: 'line', data: 'B2:D4', at: 'E2:E4' });
  view.addSparklines({ type: 'column', data: 'B2:D2', at: 'F2' });

  const reopened = SheetView.open(view.save());
  const groups = reopened.sparklineGroups('Sales');
  assert.equal(groups.length, 2);
  assert.deepEqual(groups[0].sparklines.map((s) => s.at), ['E2', 'E3', 'E4']);
  assert.deepEqual(groups[1].sparklines, [{ data: 'Sales!B2:D2', at: 'F2' }]);
});

test('an undo through the view puts a sparkline back, and a protected sheet refuses one', () => {
  const view = new SheetView(book());
  view.addSparklines({ type: 'line', data: 'B2:D4', at: 'E2:E4' });
  assert.equal(view.sparklineGroups('Sales').length, 1);
  view.undo();
  assert.equal(view.sparklineGroups('Sales').length, 0, 'undo takes the whole edit back');

  view.protect();
  assert.throws(() => view.addSparklines({ type: 'line', data: 'B2:D4', at: 'E2:E4' }), /protected/);
});

test('mismatched shapes throw, and nothing is written', () => {
  const wb = Workbook.open(book());
  assert.throws(() => wb.addSparklines('Sales', { type: 'line', data: 'B2:D4', at: 'E2:E3' }), /same number of rows/);
  assert.throws(() => wb.addSparklines('Sales', { type: 'line', data: 'B2:C2', at: 'E2:G2' }), /same number of columns/);
  assert.throws(() => wb.addSparklines('Sales', { type: 'line', data: 'B2:D4', at: 'E2:F3' }), /single row or a single column/);
  assert.throws(() => wb.addSparklines('Sales', { type: 'bar', data: 'B2:D2', at: 'E2' }), /"line" or "column"/);
  assert.doesNotMatch(sheetXml(wb), /sparklineGroups/, 'a refused call writes nothing');
});

test('the showcase fixture: if Excel wrote sparklines into it, they read back as it wrote them', { skip: !existsSync(join(RICH, 'showcase.xlsx')) && 'fixture not generated' }, () => {
  const wb = Workbook.open(readFileSync(join(RICH, 'showcase.xlsx')));
  const sheetName = wb.sheetNames().find((n) => wb.sparklineGroups(n).length);
  if (!sheetName) return; // this build of the fixture carries none — nothing to assert
  const [group] = wb.sparklineGroups(sheetName);
  assert.equal(group.type, 'line');
  assert.equal(group.colour, '376092');
  assert.ok(group.sparklines.length >= 10, 'one sparkline per data row');
  assert.ok(group.sparklines.every((s) => /^[A-Z]+\d+$/.test(s.at)), 'each sparkline sits in one cell');
  assert.ok(group.sparklines.every((s) => /!/.test(s.data)), "each sparkline's data is sheet-qualified, as Excel wrote it");
});
