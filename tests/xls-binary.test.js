/**
 * Excel's binary workbooks, of every age, opened in full.
 *
 * showcase.xls and showcase-95.xls are tests/fixtures/rich/showcase.xlsx as
 * Excel itself saved it in its 97-2003 and 5.0/95 formats
 * (tools/make-binary-fixtures.ps1); the reader is judged by whether what
 * the cells show after the round trip is what they showed before. Excel
 * 2.1, 3.0 and 4.0 files, which no Excel on hand writes, are laid out from
 * their published layouts in fixtures/old-excel.js. charts.xls and
 * charts-95.xls are the charts the showcase has not got, made by Excel the
 * same way.
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
import { parseChartXml } from '@rutba/drawing/ooxml';
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

test('an Excel 97-2003 workbook keeps its hyperlink, its note and its picture', () => {
  const book = readXls(fixture('binary', 'showcase.xls'));
  const sales = book.sheets[0];
  assert.deepEqual(sales.links[0], { range: { top: 18, bottom: 18, left: 0, right: 0 }, href: 'https://office.rutba.io/', location: null, display: 'Rutba Office', tooltip: 'The free suite' });
  assert.equal(sales.notes[0].author, 'Ejaz Arain');
  assert.match(sales.notes[0].text, /^Months of the year/);
  const summary = book.sheets.find((s) => s.name === 'Summary');
  assert.equal(summary.pictures.length, 1);
  assert.equal(summary.pictures[0].blip.contentType, 'image/png');
  assert.equal(summary.pictures[0].from.col, 3);

  const pkg = OoxmlPackage.read(xlsModelToXlsx(book));
  const salesXml = pkg.text('xl/worksheets/sheet1.xml');
  assert.match(salesXml, /<hyperlink [^>]*ref="A19"/);
  assert.ok(pkg.partNames().some((n) => /^xl\/comments\d+\.xml$/.test(n)), 'a comments part');
  assert.match(pkg.partNames().filter((n) => /^xl\/comments/.test(n)).map((n) => pkg.text(n)).join(''), /Months of the year/);
  assert.match(pkg.text('xl/worksheets/sheet2.xml'), /<drawing r:id="rId1"\/>/);
  assert.ok(pkg.partNames().some((n) => /^xl\/media\/image\d+\.png$/.test(n)), 'the picture\'s bytes');
});

test('an Excel 5.0/95 workbook opens too, as far as that format could hold it', () => {
  const bytes = fixture('binary', 'showcase-95.xls');
  assert.equal(xlsKind(bytes), 'biff5');
  const { same, differ } = compare(xlsModelToXlsx(readXls(bytes)));
  assert.ok(same > 200, `${same} cells the same`);
  // Excel 95 had no IFERROR, and no characters beyond its code page: Excel itself wrote those as #N/A and "?".
  assert.ok(differ.every((ref) => /^Sales!H\d+$/.test(ref) || ref === 'Data types!B22'), differ.join(', '));
});

test('an Excel 5.0/95 workbook keeps its picture, kept the way Excel 95 kept one', () => {
  const book = readXls(fixture('binary', 'showcase-95.xls'));
  const [picture] = book.sheets.find((s) => s.name === 'Summary').pictures;
  assert.equal(picture.blip.contentType, 'image/bmp');
  assert.equal(String.fromCharCode(...picture.blip.bytes.subarray(0, 2)), 'BM', 'a .bmp any reader opens');
  assert.equal(new DataView(picture.blip.bytes.buffer, picture.blip.bytes.byteOffset).getInt32(18, true), 240, 'its own width');
  assert.deepEqual([picture.from.col, picture.from.row, picture.to.col, picture.to.row], [3, 0, 6, 6], 'over the cells it covered');
  const pkg = OoxmlPackage.read(xlsModelToXlsx(book));
  assert.match(pkg.text('xl/worksheets/sheet2.xml'), /<drawing r:id="rId1"\/>/);
  assert.ok(pkg.partNames().some((n) => /^xl\/media\/image\d+\.bmp$/.test(n)), 'the picture\'s bytes');
});

test('an Excel 97-2003 workbook keeps its exact colours, its rich text, and what is drawn over what', () => {
  const pkg = OoxmlPackage.read(Buffer.from(xlsModelToXlsx(readXls(fixture('binary', 'showcase.xls')))));
  // The header's fill as Excel 2007 kept it (XFEXT), not the palette's nearest (333399).
  const styles = pkg.text('xl/styles.xml');
  const xfs = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)[1].match(/<xf [^>]*?(\/>|>[\s\S]*?<\/xf>)/g);
  const fills = /<fills[^>]*>([\s\S]*?)<\/fills>/.exec(styles)[1].match(/<fill>[\s\S]*?<\/fill>/g);
  const a1 = Number(/<c r="A1" s="(\d+)"/.exec(pkg.text('xl/worksheets/sheet1.xml'))[1]);
  assert.match(fills[Number(/fillId="(\d+)"/.exec(xfs[a1])[1])], /<fgColor rgb="FF1F4E79"\/>/);
  // Rich text: each run in its own font, and the whole of it read back.
  const types = pkg.text('xl/worksheets/sheet4.xml');
  const rich = /<c r="B18"[^>]*>[\s\S]*?<\/c>/.exec(types)[0];
  assert.equal((rich.match(/<r>/g) || []).length, 4, 'bold, plain, red, plain');
  assert.match(/<c r="B21"[^>]*>[\s\S]*?<\/c>/.exec(types)[0], /<vertAlign val="superscript"\/><\/rPr><t xml:space="preserve">2<\/t>/, 'the raised 2');
  const view = open(xlsModelToXlsx(readXls(fixture('binary', 'showcase.xls'))));
  view.selectSheet('Data types');
  assert.equal(new Map(view.render().cells.map((c) => [c.ref, c.text])).get('B18'), 'Bold start, red middle, plain end');
  // The picture added after the pie chart is drawn over it, with the box Excel needs to draw it at all.
  const summary = pkg.text('xl/drawings/drawing2.xml');
  assert.ok(summary.indexOf('<xdr:graphicFrame') < summary.indexOf('<xdr:pic>'), 'the chart first, the picture over it');
  assert.match(summary, /<xdr:pic>[\s\S]*<a:prstGeom prst="rect">/);
});

test('an Excel 97-2003 sheet keeps its shapes: presets, freeforms, gradients, a text box, a group, a turned shape, a connector', () => {
  const book = readXls(fixture('binary', 'showcase.xls'));
  const charts = book.sheets.find((s) => s.name === 'Charts');
  // The drawing Excel carried on in CONTINUE records once it grew: the picture and the shapes past that point.
  assert.equal(charts.pictures.length, 1);
  const presets = charts.shapes.map((s) => s.preset ?? 'path');
  for (const p of ['rect', 'roundRect', 'ellipse', 'chevron', 'hexagon', 'smileyFace', 'wedgeRectCallout', 'bentConnector3']) assert.ok(presets.includes(p), p);
  assert.equal(presets.filter((p) => p === 'path').length, 2, 'the star and the heart, from their own points');
  const box = charts.shapes.find((s) => s.textBox);
  assert.deepEqual(box.paragraphs.map((p) => p.runs.map((r) => r.text).join('')), ['A text box beside the picture.', 'Second line, in colour.', 'Third line, bold.']);
  assert.equal(box.paragraphs[2].runs[0].bold, true);
  assert.equal(charts.shapes.filter((s) => s.place).length, 2, 'the group\'s two ellipses, placed in their group');
  assert.equal(charts.shapes.find((s) => s.rotation).rotation, 30);

  const drawing = OoxmlPackage.read(Buffer.from(xlsModelToXlsx(book))).text('xl/drawings/drawing3.xml');
  assert.equal((drawing.match(/<xdr:sp /g) || []).length, 16);
  assert.match(drawing, /<xdr:cxnSp macro="">[\s\S]*?<a:prstGeom prst="bentConnector3">/);
  assert.match(drawing, /<a:custGeom>[\s\S]*?<a:path w="100000" h="100000">/);
  assert.match(drawing, /<a:xfrm rot="1800000">/, 'turned thirty degrees');
  assert.match(drawing, /<xdr:cNvSpPr txBox="1"\/>[\s\S]*?<a:t>Third line, bold.<\/a:t>/);
});

/** Each chart part of a package: its XML, and what the suite's chart drawing reads from it. */
const chartsOf = (bytes) => {
  const pkg = OoxmlPackage.read(Buffer.from(bytes));
  return pkg.partNames().filter((n) => /^xl\/charts\/chart\d+\.xml$/.test(n)).map((n) => ({ xml: pkg.text(n), spec: parseChartXml(pkg.text(n)) }));
};
const plotted = (s) => ({ type: s.type, stacked: s.stacked, categories: s.categories, series: s.series.map((x) => [x.name, x.values]) });

for (const [file, label] of [['showcase.xls', '97-2003'], ['showcase-95.xls', '5.0/95']]) {
  test(`an Excel ${label} workbook's charts are drawn as the workbook it was saved from draws them`, () => {
    const before = chartsOf(fixture('rich', 'showcase.xlsx'));
    const after = chartsOf(xlsModelToXlsx(readXls(fixture('binary', file))));
    assert.equal(after.length, before.length, 'every chart');
    for (const b of before) {
      const a = after.find((x) => x.spec.title === b.spec.title);
      assert.ok(a, `"${b.spec.title}"`);
      assert.deepEqual(plotted(a.spec), plotted(b.spec), `"${b.spec.title}": its kind, stacking, categories and series`);
    }
  });
}

test('an Excel 97-2003 chart keeps its place, its theme colours, its legend and its labels', () => {
  const book = readXls(fixture('binary', 'showcase.xls'));
  const pie = book.sheets.find((s) => s.name === 'Summary').charts[0];
  assert.deepEqual([pie.anchor.from.col, pie.anchor.from.row, pie.anchor.to.col, pie.anchor.to.row], [3, 1, 10, 17], 'over the cells it covered');
  const bytes = xlsModelToXlsx(book);
  assert.match(OoxmlPackage.read(bytes).text('xl/theme/theme1.xml'), /<a:accent1><a:srgbClr val="156082"\/>/, 'the workbook\'s theme, which its colours are');
  const xml = chartsOf(bytes).find((c) => c.spec.title === 'Share by region').xml;
  assert.match(xml, /<c:dPt><c:idx val="1"\/>[\s\S]*?<a:schemeClr val="accent2"\/>/, 'each slice in its theme colour, as Excel 2007 kept it');
  assert.match(xml, /<c:legendPos val="b"\/>/);
  assert.match(xml, /<c:showVal val="1"\/>[\s\S]*<c:showPercent val="1"\/>/, 'value and percentage on each slice');
  assert.match(xml, /<c:f>Summary!\$A\$20:\$D\$20<\/c:f>/, 'still plotting its cells');
});

test('the charts the showcase has not got: a doughnut, a scatter with its axes\' titles, a line on a second axis, a radar, 100% stacks, a 3-D chart sheet', () => {
  for (const file of ['charts.xls', 'charts-95.xls']) {
    const book = readXls(fixture('binary', file));
    assert.deepEqual(book.sheets.map((s) => s.name), ['Data', 'Chart3D'], `${file}: the chart sheet a sheet holding its chart`);
    const xmls = chartsOf(xlsModelToXlsx(book)).map((c) => c.xml);
    const titled = (t) => xmls.find((x) => x.includes('<a:t>' + t + '</a:t>')) ?? '';
    assert.match(titled('North, by month'), /<c:doughnutChart>[\s\S]*<c:holeSize val="50"\/>/, file);
    assert.match(titled('North, by month'), /<c:legendPos val="r"\/>/);
    assert.match(titled('Y against X'), /<c:scatterChart>[\s\S]*<c:valAx>[\s\S]*<a:t>Week<\/a:t>[\s\S]*<c:valAx>[\s\S]*<a:t>Orders<\/a:t>/, 'its axes\' titles');
    const combo = titled('North in columns, South as a line');
    assert.match(combo, /<c:barChart>[\s\S]*<c:lineChart>/);
    assert.match(combo, /<c:axPos val="r"\/>/, 'the line on a second value axis');
    assert.match(combo, /<c:catAx><c:axId val="50020"\/>[\s\S]*?<c:delete val="1"\/>/, 'the second axes set\'s categories hidden, as they were');
    assert.match(titled('Three regions round'), /<c:radarChart><c:radarStyle val="marker"\/>/);
    assert.match(titled('Each month, as shares'), /<c:grouping val="percentStacked"\/>/);
    assert.match(titled('In three dimensions'), /<c:view3D>[\s\S]*<c:bar3DChart>/);
  }
});

test('an Excel 3.0 or 4.0 picture opens, though they wrote it oddly', () => {
  for (const biff of [3, 4]) {
    const bytes = buildOldExcel(biff, {
      fonts: [{ name: 'Arial', height: 200 }, { name: 'Arial', height: 200 }], formats: ['General', '0'],
      cells: [{ row: 0, col: 0, text: 'Logo' }],
      picture: { from: { col: 1, row: 1 }, to: { col: 3, row: 4 }, width: 4, height: 3, colour: [200, 30, 60] },
    });
    const [picture] = readXls(bytes).sheets[0].pictures;
    assert.ok(picture, `Excel ${biff}.0`);
    const bmp = picture.blip.bytes;
    const dv = new DataView(bmp.buffer, bmp.byteOffset, bmp.byteLength);
    assert.deepEqual([dv.getUint32(14, true), dv.getInt32(18, true), dv.getInt32(22, true), dv.getUint16(28, true)], [40, 4, 3, 32], 'a Windows DIB header');
    const first = dv.getUint32(10, true);
    assert.deepEqual([...bmp.subarray(first, first + 3)], [60, 30, 200], 'its first pixel, the three stray bytes left out');
    assert.deepEqual([picture.from.col, picture.from.row, picture.to.col, picture.to.row], [1, 1, 3, 4]);
  }
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
