/**
 * The fidelity claim, made falsifiable.
 *
 * Claim: we can edit a workbook and change nothing except the part we edited,
 * and inside that part change nothing except the cells we edited.
 *
 * If any of these fail, the "build our own" path is not viable for files that
 * go to banks, and we should say so rather than ship it.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { OoxmlPackage, Workbook, fidelityOf, comparePackages, preservedButUnsupported } from '@rutba/ooxml';
import { buildComplexWorkbook, FRAGILE_PARTS, FRAGILE_SHEET_ELEMENTS } from './fixtures/complex-workbook.js';

const FIXTURE = buildComplexWorkbook();

test('the fixture is a readable OOXML package with the parts we expect', () => {
  const pkg = OoxmlPackage.read(FIXTURE);
  assert.equal(pkg.kind(), 'sheet');
  assert.equal(pkg.mainDocument(), 'xl/workbook.xml');
  for (const part of FRAGILE_PARTS) assert.ok(pkg.has(part), 'fixture is missing ' + part);
  assert.equal(pkg.contentTypeOf('xl/media/image1.png'), 'image/png');
});

test('open and save with NO edits changes nothing at all', async () => {
  const result = await fidelityOf(FIXTURE);
  assert.equal(result.changed.length, 0, 'changed: ' + JSON.stringify(result.changed));
  assert.equal(result.removed.length, 0);
  assert.equal(result.added.length, 0);
  assert.equal(result.score, 1, 'a no-op round trip must be 100% byte-identical');
  assert.equal(result.identical.length, result.total);
});

test('a no-op round trip is byte-identical at the ARCHIVE level too', async () => {
  const result = await fidelityOf(FIXTURE);
  assert.deepEqual(result.output, FIXTURE, 'untouched entries are written back verbatim');
});

test('editing one cell rewrites exactly one part', async () => {
  const result = await fidelityOf(FIXTURE, (pkg) => {
    const wb = new Workbook(pkg);
    wb.setCell('Stock', 'B2', 999);
    wb.save();
  });

  assert.deepEqual(result.intendedEdits, ['xl/worksheets/sheet1.xml']);
  assert.deepEqual(result.unintendedChanges, [], 'nothing may change that we did not edit');
  assert.equal(result.removed.length, 0, 'no part may be dropped');
  assert.equal(result.added.length, 0);
  assert.equal(result.identical.length, result.total - 1);
});

test('every fragile part survives an edit byte-identically', async () => {
  const result = await fidelityOf(FIXTURE, (pkg) => {
    const wb = new Workbook(pkg);
    wb.setCell('Stock', 'B2', 999);
    wb.save();
  });
  for (const part of FRAGILE_PARTS) {
    assert.ok(result.identical.includes(part), part + ' was not preserved byte-identically');
  }
});

test('everything outside <sheetData> survives inside the EDITED part', async () => {
  const pkg = OoxmlPackage.read(FIXTURE);
  const wb = new Workbook(pkg);
  wb.setCell('Stock', 'B2', 999);
  const out = wb.save();

  const after = OoxmlPackage.read(out).text('xl/worksheets/sheet1.xml');
  for (const element of FRAGILE_SHEET_ELEMENTS) {
    assert.ok(after.includes('<' + element), element + ' was lost from the edited sheet');
  }
  // the specific ones that would be a compliance problem to lose
  assert.ok(after.includes('errorTitle="Invalid"'), 'data validation message lost');
  assert.ok(after.includes('<sheetProtection sheet="1"'), 'sheet protection lost');
  assert.ok(after.includes('<mergeCell ref="A20:D20"/>'), 'merged cell lost');
  assert.ok(after.includes('customWidth="1"'), 'column widths lost');
});

test('an edited cell keeps its style index, and untouched rows keep their exact XML', async () => {
  const before = OoxmlPackage.read(FIXTURE).text('xl/worksheets/sheet1.xml');
  const pkg = OoxmlPackage.read(FIXTURE);
  const wb = new Workbook(pkg);
  wb.setCell('Stock', 'B2', 999);
  const after = OoxmlPackage.read(wb.save()).text('xl/worksheets/sheet1.xml');

  // B2 had s="4" — the number format must survive a value change
  assert.match(after, /<c r="B2" s="4"><v>999<\/v><\/c>/);

  // rows 1 and 3 are byte-identical substrings of the original
  const row = (xml, r) => new RegExp('<row r="' + r + '"[\\s\\S]*?</row>').exec(xml)[0];
  assert.equal(row(after, 1), row(before, 1), 'untouched header row was rewritten');
  assert.equal(row(after, 3), row(before, 3), 'untouched data row was rewritten');
});

test('a formula cell keeps its neighbours and drops only its own stale cache', async () => {
  const pkg = OoxmlPackage.read(FIXTURE);
  const wb = new Workbook(pkg);
  wb.setCell('Stock', 'C2', '=B2*1.25');
  const after = OoxmlPackage.read(wb.save()).text('xl/worksheets/sheet1.xml');

  assert.match(after, /<c r="C2" s="5"><f>B2\*1\.25<\/f><\/c>/);
  assert.ok(!/<c r="C2"[^>]*>[\s\S]*?<v>1661\.4<\/v>/.test(after), 'stale cached value kept next to a new formula');
  assert.match(after, /<c r="C3" s="5"><f>B3\*1\.17<\/f><v>1006\.2<\/v><\/c>/, 'neighbour formula disturbed');
});

test('shared strings are read but never renumbered', async () => {
  const wb = Workbook.open(FIXTURE);
  assert.equal(wb.getCell('Stock', 'A1'), 'Item');
  assert.equal(wb.getCell('Stock', 'A2'), 'Steel bracket 40mm');
  assert.equal(wb.getCell('Stock', 'B2'), 1420, 'a numeric cell reads back as a number');
  assert.equal(wb.getCell('Stock', 'C2'), '=B2*1.17');

  wb.setCell('Stock', 'A2', 'Renamed item');
  const out = wb.save();
  const diff = comparePackages(FIXTURE, out);
  assert.ok(diff.identical.includes('xl/sharedStrings.xml'), 'shared string table was rewritten');
  assert.equal(Workbook.open(out).getCell('Stock', 'A2'), 'Renamed item');
  // the other cells still resolve through the untouched table
  assert.equal(Workbook.open(out).getCell('Stock', 'A3'), 'Steel bracket 60mm');
});

test('adding a defined name preserves the ones already there', async () => {
  const pkg = OoxmlPackage.read(FIXTURE);
  const wb = new Workbook(pkg);
  wb.setDefinedName('RUTBA_STOCK_WH1', 'Stock!$A$1:$C$3');
  const out = wb.save();

  const names = new Workbook(OoxmlPackage.read(out)).definedNames();
  const byName = Object.fromEntries(names.map((n) => [n.name, n.ref]));
  assert.equal(byName.RUTBA_STOCK_WH1, 'Stock!$A$1:$C$3');
  assert.equal(byName.EXISTING_RANGE, 'Stock!$A$1:$C$3', 'existing defined name lost');
  assert.ok(byName['_xlnm.Print_Area'], 'print area lost');

  const diff = comparePackages(FIXTURE, out);
  assert.deepEqual(diff.changed.map((c) => c.name), ['xl/workbook.xml']);
});

test('writing to a second sheet leaves the first untouched', async () => {
  const result = await fidelityOf(FIXTURE, (pkg) => {
    const wb = new Workbook(pkg);
    wb.setCell('Summary', 'B5', 'added');
    wb.save();
  });
  assert.deepEqual(result.intendedEdits, ['xl/worksheets/sheet2.xml']);
  assert.ok(result.identical.includes('xl/worksheets/sheet1.xml'));
});

test('setRange writes a block without disturbing the rest of the sheet', async () => {
  const pkg = OoxmlPackage.read(FIXTURE);
  const wb = new Workbook(pkg);
  wb.setRange('Summary', 'A1', [['SKU', 'Qty'], ['SKU-1001', 1420], ['SKU-1002', 860]]);
  const out = wb.save();

  const reread = Workbook.open(out);
  assert.equal(reread.getCell('Summary', 'A1'), 'SKU');
  assert.equal(reread.getCell('Summary', 'B3'), 860);
  const diff = comparePackages(FIXTURE, out);
  assert.deepEqual(diff.changed.map((c) => c.name), ['xl/worksheets/sheet2.xml']);
});

test('binary parts survive, including a stored (uncompressed) entry', async () => {
  const original = OoxmlPackage.read(FIXTURE);
  const image = original.read('xl/media/image1.png');
  assert.equal(original.part('xl/media/image1.png').method, 0, 'fixture should store this entry');

  const result = await fidelityOf(FIXTURE, (pkg) => {
    const wb = new Workbook(pkg);
    wb.setCell('Stock', 'B2', 1);
    wb.save();
  });
  assert.deepEqual(OoxmlPackage.read(result.output).read('xl/media/image1.png'), image);
  assert.match(image.subarray(1, 4).toString('latin1'), /PNG/);
});

test('we report honestly which preserved features we cannot render', () => {
  const features = preservedButUnsupported(FIXTURE);
  assert.deepEqual(features, ['charts', 'custom XML', 'drawings and shapes', 'embedded media', 'pivot tables']);
});

test('repeated edit cycles do not accumulate damage', async () => {
  let buf = FIXTURE;
  for (let i = 0; i < 10; i++) {
    const wb = Workbook.open(buf);
    wb.setCell('Stock', 'B2', 1000 + i);
    buf = wb.save();
  }
  const diff = comparePackages(FIXTURE, buf);
  assert.deepEqual(diff.changed.map((c) => c.name), ['xl/worksheets/sheet1.xml']);
  assert.equal(diff.removed.length, 0);
  assert.equal(Workbook.open(buf).getCell('Stock', 'B2'), 1009);

  const after = OoxmlPackage.read(buf).text('xl/worksheets/sheet1.xml');
  for (const element of FRAGILE_SHEET_ELEMENTS) {
    assert.ok(after.includes('<' + element), element + ' eroded after 10 cycles');
  }
});

test('a non-package input is refused clearly rather than corrupted', () => {
  assert.throws(() => OoxmlPackage.read(Buffer.from('not a zip at all')), /not an OOXML package/);
  const zipButNotOoxml = Buffer.concat([Buffer.from('PK'), Buffer.alloc(64)]);
  assert.throws(() => OoxmlPackage.read(zipButNotOoxml), /zip|OOXML/i);
});
