/**
 * Calculation over real workbooks.
 *
 * This is the gap FORMAT-FIDELITY.md listed as the top functional item: until
 * now formulas were written and preserved but never evaluated, so a bound sheet
 * was only correct once someone opened it in Excel.
 *
 * The bar these tests hold: recalculation must be as preserving as every other
 * edit. It may write cached values into formula cells and nothing else.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  OoxmlPackage, Workbook, recalculateWorkbook, inspectCalculation,
  toSpreadsheet, parseDefinedNameRange, comparePackages,
} from '@rutba/ooxml';
import { isError } from '@rutba/formula';
import { buildComplexWorkbook, FRAGILE_PARTS, FRAGILE_SHEET_ELEMENTS } from './fixtures/complex-workbook.js';

const FIXTURE = buildComplexWorkbook();

test('a real workbook loads into the calculation model', () => {
  const wb = Workbook.open(FIXTURE);
  const { sheet, formulas } = toSpreadsheet(wb);

  assert.deepEqual(sheet.sheetNames(), ['Stock', 'Summary']);
  assert.deepEqual(formulas.map((f) => f.ref).sort(), ['C2', 'C3']);
  // values arrived, including ones behind the shared string table
  assert.equal(sheet.getByRef('Stock', 'A2'), 'Steel bracket 40mm');
  assert.equal(sheet.getByRef('Stock', 'B2'), 1420);
});

test('defined names come across, including the range form', () => {
  assert.deepEqual(parseDefinedNameRange('Stock!$A$1:$C$3'), {
    sheet: 'Stock', start: { row: 0, col: 0 }, end: { row: 2, col: 2 },
  });
  assert.deepEqual(parseDefinedNameRange("'My Sheet'!$B$2"), {
    sheet: 'My Sheet', start: { row: 1, col: 1 }, end: { row: 1, col: 1 },
  });
  assert.equal(parseDefinedNameRange('not a range'), null);

  // EXISTING_RANGE is Stock!$A$1:$C$3 — the whole block, so the sum covers
  // both the "On hand" and the calculated "With tax" columns.
  const { sheet } = toSpreadsheet(Workbook.open(FIXTURE));
  sheet.setByRef('Stock', 'E1', '=SUM(EXISTING_RANGE)');
  sheet.recalculate();
  assert.equal(
    Math.round(sheet.getByRef('Stock', 'E1') * 10) / 10,
    4947.6,
    '1420 + 860 + 1661.4 + 1006.2',
  );
});

test('the workbook formulas evaluate to the right numbers', () => {
  const report = inspectCalculation(FIXTURE);
  assert.equal(report.formulaCount, 2);
  assert.equal(report.calculated, 2);
  assert.deepEqual(report.cycles, []);
  assert.deepEqual(report.errors, []);

  const { sheet } = toSpreadsheet(Workbook.open(FIXTURE));
  sheet.recalculate();
  // C2 is =B2*1.17 where B2 is 1420
  assert.equal(Math.round(sheet.getByRef('Stock', 'C2') * 100) / 100, 1661.4);
  assert.equal(Math.round(sheet.getByRef('Stock', 'C3') * 100) / 100, 1006.2);
});

test('recalculating an already-correct workbook changes literally nothing', () => {
  // The fixture's cached values agree with its formulas, so a recalculation has
  // nothing to write. Byte-identical output is the strongest possible statement
  // that we only write when a value actually moved.
  const { output, calculated } = recalculateWorkbook(FIXTURE);
  assert.equal(calculated, 2, 'both formulas were evaluated');
  assert.deepEqual(output, FIXTURE, 'agreeing with the cache means writing nothing');
});

test('recalculation writes cached values and touches nothing else', () => {
  // make one cache stale, then check the repair is surgical
  const wb = Workbook.open(FIXTURE);
  wb.setCell('Stock', 'B2', 2000);
  const stale = wb.save();

  const { output } = recalculateWorkbook(stale);
  const diff = comparePackages(stale, output);
  assert.deepEqual(diff.changed.map((c) => c.name), ['xl/worksheets/sheet1.xml']);
  assert.equal(diff.removed.length, 0);
  for (const part of FRAGILE_PARTS) assert.ok(diff.identical.includes(part), part + ' disturbed');

  const sheetXml = OoxmlPackage.read(output).text('xl/worksheets/sheet1.xml');
  for (const element of FRAGILE_SHEET_ELEMENTS) {
    assert.ok(sheetXml.includes('<' + element), element + ' lost during recalculation');
  }
});

test('a recalculated cell keeps its formula, style and reference', () => {
  const { output } = recalculateWorkbook(FIXTURE);
  const sheetXml = OoxmlPackage.read(output).text('xl/worksheets/sheet1.xml');
  // the formula text survives, the style index survives, the cached value updates
  assert.match(sheetXml, /<c r="C2" s="5"><f>B2\*1\.17<\/f><v>1661\.4<\/v><\/c>/);
  assert.match(sheetXml, /<c r="C3" s="5"><f>B3\*1\.17<\/f><v>1006\.2<\/v><\/c>/);
  // untouched cells are unchanged
  assert.match(sheetXml, /<c r="B2" s="4"><v>1420<\/v><\/c>/);
});

test('editing a precedent then recalculating updates the dependent', () => {
  const wb = Workbook.open(FIXTURE);
  wb.setCell('Stock', 'B2', 2000);
  const edited = wb.save();

  const { output } = recalculateWorkbook(edited);
  const sheetXml = OoxmlPackage.read(output).text('xl/worksheets/sheet1.xml');
  assert.match(sheetXml, /<c r="C2" s="5"><f>B2\*1\.17<\/f><v>2340<\/v><\/c>/);
});

test('a workbook with no formulas comes back byte-identical', () => {
  const wb = Workbook.open(FIXTURE);
  // clear both formula cells by overwriting them with literals
  wb.setCell('Stock', 'C2', 1);
  wb.setCell('Stock', 'C3', 2);
  const literalOnly = wb.save();

  const { output, calculated } = recalculateWorkbook(literalOnly);
  assert.equal(calculated, 0);
  assert.deepEqual(output, literalOnly, 'nothing to calculate means nothing to write');
});

test('an error result is cached as an error, with the right cell type', () => {
  const wb = Workbook.open(FIXTURE);
  wb.setCell('Summary', 'A2', '=1/0');
  const withError = wb.save();

  const { output, errors } = recalculateWorkbook(withError);
  assert.equal(errors.length, 1);
  assert.equal(errors[0].error, '#DIV/0!');
  assert.equal(errors[0].ref, 'A2');

  const sheetXml = OoxmlPackage.read(output).text('xl/worksheets/sheet2.xml');
  assert.match(sheetXml, /<c r="A2"[^>]*t="e"><f>1\/0<\/f><v>#DIV\/0!<\/v><\/c>/);
});

test('a text result is cached with t="str"', () => {
  const wb = Workbook.open(FIXTURE);
  wb.setCell('Summary', 'A3', '=UPPER("done")');
  const { output } = recalculateWorkbook(wb.save());
  const sheetXml = OoxmlPackage.read(output).text('xl/worksheets/sheet2.xml');
  // the formula text stays XML-escaped inside <f>, as it must
  assert.match(sheetXml, /<c r="A3"[^>]*t="str"><f>UPPER\(&quot;done&quot;\)<\/f><v>DONE<\/v><\/c>/);
});

test('a circular reference is reported and does not hang', () => {
  const wb = Workbook.open(FIXTURE);
  wb.setCell('Summary', 'B1', '=B2');
  wb.setCell('Summary', 'B2', '=B1');
  const { cycles } = recalculateWorkbook(wb.save());
  assert.equal(cycles.length, 2);
  assert.ok(cycles.every((c) => c.startsWith('Summary!')));
});

test('a formula we cannot evaluate keeps its existing cached value', () => {
  const wb = Workbook.open(FIXTURE);
  // A function outside our library — degrade loudly, do not overwrite. This
  // used to say XLOOKUP, until XLOOKUP was implemented; CUBEVALUE needs an
  // OLAP connection and will stay outside.
  wb.setCell('Summary', 'C1', '=CUBEVALUE(1,2,3)');
  const before = wb.save();

  const { output, unparsed } = recalculateWorkbook(before);
  assert.equal(unparsed.length, 1);
  assert.match(unparsed[0].reason, /CUBEVALUE/);

  const sheetXml = OoxmlPackage.read(output).text('xl/worksheets/sheet2.xml');
  assert.ok(!sheetXml.includes('#NAME?'), 'we must not write our own failure into the file');
});

test('cross-sheet formulas calculate against the real workbook', () => {
  const wb = Workbook.open(FIXTURE);
  wb.setCell('Summary', 'A5', '=SUM(Stock!B2:B3)');
  const { output } = recalculateWorkbook(wb.save());
  const sheetXml = OoxmlPackage.read(output).text('xl/worksheets/sheet2.xml');
  assert.match(sheetXml, /<c r="A5"[^>]*><f>SUM\(Stock!B2:B3\)<\/f><v>2280<\/v><\/c>/);
});

test('ten recalculation cycles are stable and accumulate no damage', () => {
  // start from a stale cache so the first pass actually writes
  const wb = Workbook.open(FIXTURE);
  wb.setCell('Stock', 'B2', 2000);
  const start = wb.save();

  let buf = start;
  for (let i = 0; i < 10; i++) buf = recalculateWorkbook(buf).output;

  const diff = comparePackages(start, buf);
  assert.deepEqual(diff.changed.map((c) => c.name), ['xl/worksheets/sheet1.xml']);
  assert.equal(diff.removed.length, 0);

  const sheetXml = OoxmlPackage.read(buf).text('xl/worksheets/sheet1.xml');
  for (const element of FRAGILE_SHEET_ELEMENTS) assert.ok(sheetXml.includes('<' + element), element + ' eroded');
  assert.match(sheetXml, /<c r="C2" s="5"><f>B2\*1\.17<\/f><v>2340<\/v><\/c>/, 'value drifted across cycles');

  // and once it has converged, further passes are byte-identical
  assert.deepEqual(recalculateWorkbook(buf).output, buf, 'recalculation is idempotent once settled');
});

// The ERP-bound recalculation test stayed with the bindings it exercises, in
// the consumer suite: the engine here has no notion of a business database.
