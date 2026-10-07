/**
 * Excel's binary workbooks, of every age, opened in full.
 *
 * showcase.xls and showcase-95.xls are tests/fixtures/rich/showcase.xlsx as
 * Excel itself saved it in its 97-2003 and 5.0/95 formats
 * (tools/make-binary-fixtures.ps1); the reader is judged by whether what
 * the cells show after the round trip is what they showed before. Excel
 * 2.1, 3.0 and 4.0 files, which no Excel on hand writes, are laid out from
 * their published layouts in fixtures/old-excel.js.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { readXls, xlsKind, rkNumber } from '@rutba/office-formats/msxls';
import { xlsModelToXlsx } from '@rutba/office-formats/msxls-xlsx';
import { sniff } from '@rutba/office-formats/sniff';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { SheetView } from '@rutba/sheet-view';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { buildOldExcel, refTok, areaTok, funcVar } from './fixtures/old-excel.js';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const fixture = (...p) => new Uint8Array(fs.readFileSync(path.join(HERE, 'fixtures', ...p)));
const open = (bytes) => SheetView.open(Buffer.from(bytes), { viewportWidth: 2400, viewportHeight: 2400 });
const SHEETS = ['Sales', 'Summary', 'Data types', 'Table'];

/** What every cell of the four sheets shows, original against converted: [same, differences]. */
function compare(convertedBytes) {
  const before = open(fixture('rich', 'showcase.xlsx'));
  const after = open(convertedBytes);
  let same = 0;
  const differ = [];
  for (const name of SHEETS) {
    before.selectSheet(name);
    after.selectSheet(name);
    const shown = new Map(after.render().cells.map((c) => [c.ref, c.text]));
    for (const c of before.render().cells) {
      if (!c.text) continue;
      if (shown.get(c.ref) === c.text) same += 1;
      else differ.push(`${name}!${c.ref}`);
    }
  }
  return { same, differ };
}

test('an Excel 97-2003 workbook shows every cell as the workbook it was saved from does', () => {
  const bytes = fixture('binary', 'showcase.xls');
  assert.equal(xlsKind(bytes), 'biff8');
  const book = readXls(bytes);
  assert.deepEqual(book.sheets.map((s) => s.name), ['Sales', 'Summary', 'Charts', 'Data types', 'Table']);
  const { same, differ } = compare(xlsModelToXlsx(book));
  assert.deepEqual(differ, [], 'no cell shows anything different');
  assert.ok(same > 200, `${same} cells compared`);
});

test('an Excel 97-2003 workbook keeps its formulas, names, formats, merged cells and frozen panes', () => {
  const book = readXls(fixture('binary', 'showcase.xls'));
  const sales = book.sheets[0];
  const at = (s, ref) => [...s.cells.values()].find((c) => String.fromCharCode(65 + c.col) + (c.row + 1) === ref);
  assert.equal(at(sales, 'F2').formula, 'SUM(B2:E2)');
  assert.equal(at(sales, 'F2').v, 7518, 'with the value Excel last calculated');
  assert.equal(at(sales, 'H3').formula, '_xlfn.IFERROR((F3-F2)/F2,0)', 'a newer function, as .xlsx names it');
  assert.deepEqual(book.names.find((n) => n.name === 'GrandTotal'), { name: 'GrandTotal', builtin: false, hidden: false, sheet: null, formula: 'Sales!$F$14' });
  assert.deepEqual(sales.frozen, { rows: 1, cols: 1 });

  const pkg = OoxmlPackage.read(xlsModelToXlsx(book));
  const salesXml = pkg.text('xl/worksheets/sheet1.xml');
  assert.match(salesXml, /<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"\/>/);
  assert.match(salesXml, /<c r="F2" s="\d+"><f>SUM\(B2:E2\)<\/f><v>7518<\/v><\/c>/);
  assert.match(pkg.text('xl/worksheets/sheet2.xml'), /<mergeCell ref="A1:D1"\/>/);
  assert.match(pkg.text('xl/workbook.xml'), /<definedName name="GrandTotal">Sales!\$F\$14<\/definedName>/);
  // The header cell's format: bold, filled.
  const s = Number(/<c r="A1" s="(\d+)"/.exec(salesXml)[1]);
  const styles = pkg.text('xl/styles.xml');
  const xf = [...styles.matchAll(/<xf [^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].filter((m) => !m.index || styles.lastIndexOf('<cellXfs', m.index) > styles.lastIndexOf('</cellStyleXfs>', m.index))[s]?.[0] ?? '';
  const font = [...styles.matchAll(/<font>[\s\S]*?<\/font>/g)][Number(/fontId="(\d+)"/.exec(xf)[1])][0];
  assert.match(font, /<b\/>/, 'bold');
  assert.notEqual(/fillId="(\d+)"/.exec(xf)[1], '0', 'filled');
});

test('an Excel 5.0/95 workbook opens too, as far as that format could hold it', () => {
  const bytes = fixture('binary', 'showcase-95.xls');
  assert.equal(xlsKind(bytes), 'biff5');
  const { same, differ } = compare(xlsModelToXlsx(readXls(bytes)));
  assert.ok(same > 200, `${same} cells the same`);
  // Excel 95 had no IFERROR, and no characters beyond its code page: Excel itself wrote those as #N/A and "?".
  assert.ok(differ.every((ref) => /^Sales!H\d+$/.test(ref) || ref === 'Data types!B22'), differ.join(', '));
});

for (const biff of [2, 3, 4]) {
  test(`an Excel ${biff === 2 ? '2.1' : biff + '.0'} worksheet opens: numbers, text, a truth value, formulas, fonts, formats, a width, a height, a frozen row`, () => {
    const bytes = buildOldExcel(biff, {
      fonts: [{ name: 'Arial', height: 200 }, { name: 'Arial', height: 240, bold: true }],
      formats: ['General', '0.00'],
      colWidth: { col: 0, width: 20 * 256 },
      rowHeight: { row: 0, height: 400 },
      frozen: { rows: 1, cols: 0 },
      cells: [
        { row: 0, col: 0, text: 'Item', xf: 1 },
        { row: 0, col: 1, text: 'Amount', xf: 1 },
        { row: 1, col: 0, text: 'Tea' },
        { row: 1, col: 1, value: 12.5, xf: 2 },
        { row: 2, col: 0, text: 'Cake' },
        { row: 2, col: 1, value: 30 },
        { row: 3, col: 0, text: 'Total' },
        { row: 3, col: 1, formula: [...areaTok(1, 2, 1, 1), ...funcVar(biff, 1, 4)], result: 42.5, xf: 2 },
        { row: 4, col: 1, formula: [...refTok(1, 1), ...refTok(2, 1), 0x05], result: 375 },
        { row: 5, col: 0, bool: true },
      ],
    });
    assert.equal(xlsKind(bytes), 'biff' + biff);
    assert.equal(sniff(bytes, 'old.xls').kind, 'xls');
    const book = readXls(bytes);
    const cell = (ref) => [...book.sheets[0].cells.values()].find((c) => String.fromCharCode(65 + c.col) + (c.row + 1) === ref);
    assert.equal(cell('A1').v, 'Item');
    assert.equal(cell('B2').v, 12.5);
    assert.equal(cell('B4').formula, 'SUM(B2:B3)');
    assert.equal(cell('B4').v, 42.5);
    assert.equal(cell('B5').formula, 'B2*B3');
    assert.equal(cell('A6').t, 'b');
    assert.equal(book.fonts[1].bold, true);
    assert.deepEqual(book.sheets[0].frozen, { rows: 1, cols: 0 });

    const view = open(xlsModelToXlsx(book));
    const shown = new Map(view.render().cells.map((c) => [c.ref, c.text]));
    assert.equal(shown.get('B2'), '12.50', 'in its number format');
    assert.equal(shown.get('B4'), '42.50');
    assert.equal(shown.get('B5'), '375');
    assert.equal(shown.get('A6'), 'TRUE');
    const xml = OoxmlPackage.read(xlsModelToXlsx(book)).text('xl/worksheets/sheet1.xml');
    assert.match(xml, /<col min="1" max="1" width="20" customWidth="1"\/>/);
    assert.match(xml, /<row r="1" ht="20" customHeight="1">/);
    assert.match(xml, /<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"\/>/);
  });
}

test('RK numbers: integers, and doubles cut to their top 30 bits, each perhaps a hundredth', () => {
  assert.equal(rkNumber((7 << 2) | 2), 7);
  assert.equal(rkNumber((1234 << 2) | 3), 12.34);
  assert.equal(rkNumber(0x3ff00000), 1, 'the top of the double 1.0');
});

test('the suite opens an .xls as a worksheet, read in full, and says it saves an .xlsx', () => {
  const doc = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-xls-'));
  const openFile = (name, bytes) => {
    const p = path.join(dir, name);
    fs.writeFileSync(p, bytes);
    const s = doc.open({ path: p });
    doc.close({ id: s.id });
    return s;
  };
  const xls = openFile('showcase.xls', fixture('binary', 'showcase.xls'));
  assert.equal(xls.kind, 'sheet');
  assert.equal(xls.converted.from, 'xls');
  assert.ok(!xls.converted.partial, 'not the old "could not be converted" sheet');
  const old = openFile('old.xls', buildOldExcel(4, { fonts: [{ name: 'Arial', height: 200 }, { name: 'Arial', height: 200 }], formats: ['General', '0'], cells: [{ row: 0, col: 0, text: 'Excel 4' }] }));
  assert.equal(old.kind, 'sheet');
  assert.equal(old.converted.format, 'biff4');
});
