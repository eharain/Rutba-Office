// What the second review of Sheets found, held to what Excel itself answers.
// Every expected value below came from Excel, asked over COM on the release PC.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Spreadsheet, isError, dateToSerial } from '@rutba/formula';
import { buildXlsx } from '@rutba/ooxml/build';
import { Workbook } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { Deck, buildPptx } from '@rutba/presentation';

function ask(cells, formula) {
  const s = new Spreadsheet({});
  for (const [a1, v] of Object.entries(cells)) s.setByRef('Sheet1', a1, v);
  s.setByRef('Sheet1', 'J20', formula);
  s.recalculate();
  return s.getValue('Sheet1', 19, 9);
}
const errorOf = (v) => (isError(v) ? v.type : v);

const GRID = { A1: 1, A2: '=NA()', A3: 3, B1: 10, B2: 20, B3: 30, D1: 'text', D2: 'more', E1: 5, E2: 7 };

for (const [formula, answer] of [
  ['=SUMIF(A1:A3,3,B1)', 30],
  ['=SUMIF(B1:B3,">1",B1)', 60],
  ['=AVERAGEIF(B1:B3,">1",B1)', 20],
  ['=SUMIF(B1:B3,">1",B1:B2)', 60],
  ['=SUMIF(B1:B2,">1",B1:C3)', 30],
  ['=0.1+0.2=0.3', true],
  ['=1.1*3=3.3', true],
  ['=1.1*3>3.3', false],
  ['=1.1*3<3.3', false],
  ['=1.1*3<>3.3', false],
  ['=1+1E-15=1', true],
  ['=1+1E-14=1', false],
  ['=SUM(INDEX(D1:D2,1))', 0],
  ['=SUM(OFFSET(D1,0,0))', 0],
  ['=SUM(INDIRECT("D1"))', 0],
  ['=AVERAGE(INDEX(D1:E2,1,1),E1)', 5],
  ['=COUNT(INDEX(D1:E2,1,1))', 0],
  ['=SUM(CHOOSE(1,D1,E1))', 0],
  ['=SUM(IF(TRUE,D1,E1))', 0],
  ['=INDEX(B1:B3,2.7)', 20],
  ['=INDEX(B1:B3,1.1)', 10],
  ['=VLOOKUP(10,B1:C3,1.9,FALSE)', 10],
  ['=HLOOKUP(10,B1:C3,1.9,FALSE)', 10],
  ['=INDEX(B1:C3,1,1.9)', 10],
  ['=COUNTIF(H:H,"")', 1048576],
  ['=COUNTIF(H:H,"<>x")', 1048576],
  ['=COUNTIFS(H:H,"")', 1048576],
  ['=COUNTIFS(H:H,"<>x")', 1048576],
  ['=FACT(170)>1E+300', true],
  ['=FACT(171)', '#NUM!'],
  ['=FACT(1E+15)', '#NUM!'],
  ['=WORKDAY(1,1E+6)', 1399999],
  ['=WORKDAY(1,1E+7)', '#NUM!'],
  ['=WORKDAY(1,1E+12)', '#NUM!'],
  ['=COMBIN(1000,500)>1E+299', true],
  ['=COMBIN(1030,515)', '#NUM!'],
  ['=COMBIN(1E+15,5E+14)', '#NUM!'],
  ['=PERMUT(1E+15,1E+14)', '#NUM!'],
  ['=NETWORKDAYS(1,2958465)', 2113190],
  ['=NETWORKDAYS(1,3000000)', '#NUM!'],
  ['=NETWORKDAYS(1,1E+15)', '#NUM!'],
]) {
  test(`${formula} answers as Excel does`, () => {
    assert.equal(errorOf(ask(GRID, formula)), answer);
  });
}

test('a whole column with something in it counts the empty cells that are left', () => {
  assert.equal(ask({ H1: 'x', H2: 'y' }, '=COUNTIF(H:H,"<>x")'), 1048575);
  assert.equal(ask({ H1: 'x', H2: 'y' }, '=COUNTIF(H:H,"")'), 1048574);
  assert.equal(ask({ H1: 'x', I1: 'x' }, '=COUNTIFS(H:H,"",I:I,"")'), 1048575);
});

test('SUMIF follows the cells its short last range reaches, not only the one written', () => {
  const s = new Spreadsheet({});
  s.setByRef('Sheet1', 'A1', 1); s.setByRef('Sheet1', 'A2', 1);
  s.setByRef('Sheet1', 'B1', 10); s.setByRef('Sheet1', 'B2', 20);
  s.setByRef('Sheet1', 'D1', '=SUMIF(A1:A2,1,B1)');
  s.recalculate();
  assert.equal(s.getValue('Sheet1', 0, 3), 30);
  s.setByRef('Sheet1', 'B2', 200);
  s.recalculate();
  assert.equal(s.getValue('Sheet1', 0, 3), 210, 'B2 is read through B1 and has to wake the formula');
});

test('INDIRECT reading a formula placed after it gets the worked-out answer, and passes it on', () => {
  const s = new Spreadsheet({});
  s.setByRef('Sheet1', 'A1', '=INDIRECT("B1")');
  s.setByRef('Sheet1', 'D1', '=A1*2');
  s.setByRef('Sheet1', 'B1', '=C1+1');
  s.setByRef('Sheet1', 'C1', 5);
  s.recalculate();
  assert.equal(s.getValue('Sheet1', 0, 0), 6);
  assert.equal(s.getValue('Sheet1', 0, 3), 12);
  s.setByRef('Sheet1', 'C1', 7);
  s.recalculate();
  assert.equal(s.getValue('Sheet1', 0, 0), 8);
  assert.equal(s.getValue('Sheet1', 0, 3), 16);
});

test('OFFSET settles too, through a chain of them', () => {
  const s = new Spreadsheet({});
  s.setByRef('Sheet1', 'A1', '=OFFSET(A1,0,1)');
  s.setByRef('Sheet1', 'B1', '=OFFSET(B1,0,1)');
  s.setByRef('Sheet1', 'C1', '=D1*3');
  s.setByRef('Sheet1', 'D1', 4);
  s.recalculate();
  assert.equal(s.getValue('Sheet1', 0, 0), 12);
});

test('a long column of OFFSETs reading downwards settles', () => {
  const s = new Spreadsheet({});
  const rows = 40;
  for (let r = 0; r < rows; r++) s.setCell('Sheet1', r, 0, '=OFFSET(A' + (r + 1) + ',1,0)+1');
  s.setCell('Sheet1', rows, 0, 0);
  s.recalculate();
  assert.equal(s.getValue('Sheet1', 0, 0), rows);
});

test('a formula that calls something unknown inside IFERROR keeps what the file stored', () => {
  const s = new Spreadsheet({});
  s.setCell('Sheet1', 0, 0, '=IFERROR(NOSUCHFUNCTION(1),0)', { cached: 42 });
  s.recalculate();
  assert.equal(s.getValue('Sheet1', 0, 0), 42, 'IFERROR must not turn what this program cannot work out into 0');
});

test('INDIRECT finds a name kept on its own sheet', () => {
  const s = new Spreadsheet({});
  s.setByRef('Sheet1', 'B1', 9);
  s.defineName('Local', { sheet: 'Sheet1', start: { row: 0, col: 1 }, end: { row: 0, col: 1 } }, 'Sheet1');
  s.setByRef('Sheet1', 'A1', '=INDIRECT("Local")');
  s.recalculate();
  assert.equal(s.getValue('Sheet1', 0, 0), 9);
});

const book = (rows, name = 'Sales') => buildXlsx({ sheets: [{ name, rows }] });
const TABLE = [['Item', 'Qty', 'Price', 'Note'], ['Pens', 12, 1.5, 'a'], ['Ink', 3, 9, 'b'], ['Clips', 40, 0.5, 'c']];

function tabled(rows = TABLE) {
  const view = new SheetView(book(rows));
  view.select(1, 1);
  view.formatAsTable({ style: 'TableStyleMedium2', stripes: true });
  return view;
}

const FILTERED = '<autoFilter ref="A1:D4">'
  + '<filterColumn colId="0"><filters><filter val="Pens"/></filters></filterColumn>'
  + '<filterColumn colId="2"><customFilters><customFilter operator="greaterThan" val="1"/></customFilters></filterColumn>'
  + '<filterColumn colId="3" hiddenButton="1"/>'
  + '</autoFilter>';

function withTableXml(view, change) {
  const part = view.workbook.tables()[0].part;
  view.workbook.pkg.write_(part, change(view.workbook.pkg.text(part)));
  return SheetView.open(view.save());
}
const tableXml = (view) => view.workbook.pkg.text(view.workbook.tables()[0].part);
const filtered = (xml) => xml.replace(/<autoFilter\b[^>]*\/>|<autoFilter\b[\s\S]*?<\/autoFilter>/, () => FILTERED);
const colIds = (xml) => [...xml.matchAll(/<filterColumn\b[^>]*?colId="(\d+)"/g)].map((m) => Number(m[1]));

test('a column taken out of a table takes its filter whole, and the ones after it close up', () => {
  const view = withTableXml(tabled(), filtered);
  view.select(0, 2);
  view.deleteCols(2, 1);
  const xml = tableXml(view);
  assert.deepEqual(colIds(xml), [0, 2]);
  assert.doesNotMatch(xml, /customFilter/, 'what was inside the filter went with it');
  assert.match(xml, /<filters><filter val="Pens"\/><\/filters><\/filterColumn>/, 'its neighbour is whole');
  assert.equal((xml.match(/<\/filterColumn>/g) ?? []).length, 1, 'no closing tag left behind');
});

test('a column put into a table moves the filters to its right', () => {
  const view = withTableXml(tabled(), filtered);
  view.select(0, 1);
  view.insertCols(1, 1);
  assert.deepEqual(colIds(tableXml(view)), [0, 3, 4]);
});

function withSheetXml(rows, change) {
  const wb = Workbook.open(book(rows));
  const part = wb.partNameFor('Sales');
  wb.pkg.write_(part, change(wb.pkg.text(part)));
  return Workbook.open(wb.save());
}
const saved = (wb, sheet) => { const again = Workbook.open(wb.save()); return again.pkg.text(again.partNameFor(sheet)); };
const putFilter = (xml) => xml.replace('</sheetData>', () => '</sheetData>' + FILTERED);

test('the sheet\'s own filter keeps its columns when one is put in or taken out', () => {
  const taken = withSheetXml(TABLE, putFilter);
  taken.deleteCols('Sales', 2, 1);
  const xml = saved(taken, 'Sales');
  assert.deepEqual(colIds(xml), [0, 2]);
  assert.match(xml, /<autoFilter ref="A1:C4"/);
  assert.doesNotMatch(xml, /customFilter/);

  const put = withSheetXml(TABLE, putFilter);
  put.insertCols('Sales', 1, 1);
  assert.deepEqual(colIds(saved(put, 'Sales')), [0, 3, 4]);

  const atEdge = withSheetXml(TABLE, putFilter);
  atEdge.insertCols('Sales', 0, 1);
  const moved = saved(atEdge, 'Sales');
  assert.deepEqual(colIds(moved), [0, 2, 3], 'the whole filter moved, so its columns count the same');
  assert.match(moved, /<autoFilter ref="B1:E4"/);
});

test('a reference to another workbook stays put when a row goes in on a sheet of the same name', () => {
  const wb = Workbook.open(buildXlsx({ sheets: [{ name: 'Data', rows: [[1], [2], [3]] }, { name: 'Sum', rows: [['=[1]Data!B3', '=Data!B3']] }] }));
  wb.insertRows('Data', 0, 1);
  const xml = saved(wb, 'Sum');
  assert.match(xml, /<f>\[1\]Data!B3<\/f>/, 'the other workbook has not moved');
  assert.match(xml, /<f>Data!B4<\/f>/, 'this one has');
});

test('a calculated column and a total row follow the row or column that moved', () => {
  const view = withTableXml(tabled(), (xml) => xml.replace(/<tableColumn id="4"[^>]*\/>/, (m) => m.replace('/>', '><calculatedColumnFormula>SUM(F1:F4)</calculatedColumnFormula><totalsRowFormula>F1&amp;"x"</totalsRowFormula></tableColumn>')));
  view.select(0, 4);
  view.insertCols(4, 1);
  const xml = tableXml(view);
  assert.match(xml, /<calculatedColumnFormula>SUM\(G1:G4\)<\/calculatedColumnFormula>/);
  assert.match(xml, /<totalsRowFormula>G1&amp;(?:"|&quot;)x(?:"|&quot;)<\/totalsRowFormula>/, 'written back escaped');
});

test('deleting the header row of a table is refused, and leaves no undo step', () => {
  const view = tabled();
  const before = view.workbook.tables()[0].ref;
  const canUndo = view.canUndo;
  assert.throws(() => view.deleteRows(0, 2), /header row/);
  assert.equal(view.workbook.tables()[0].ref, before);
  assert.equal(view.canUndo, canUndo);
});

test('deleting the whole of a table, header and all, is still allowed', () => {
  const view = tabled();
  view.deleteRows(0, 4);
  assert.deepEqual(view.workbook.tables(), []);
});

test('deleting every data row leaves the header and one empty row, as Excel does', () => {
  const view = tabled();
  view.deleteRows(1, 3);
  assert.equal(view.workbook.tables()[0].ref, 'A1:D2');
});

test('taking a sheet away drops the list of formula cells', () => {
  const wb = Workbook.open(buildXlsx({ sheets: [{ name: 'One', rows: [['=1+1']] }, { name: 'Two', rows: [['=2+2']] }] }));
  wb.pkg.addPart('xl/calcChain.xml', '<?xml version="1.0"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/><c r="A1" i="2"/></calcChain>');
  wb.removeSheet('Two');
  assert.equal(wb.pkg.has('xl/calcChain.xml'), false);
});

test('a "this day" rule goes by the day on this computer, as TODAY() does', () => {
  // Half past midnight here is still the day before in UTC, east of Greenwich.
  const clock = new Date(2026, 8, 2, 0, 30);
  const today = Math.floor(dateToSerial(new Date(Date.UTC(2026, 8, 2))));
  const view = new SheetView(book([[today], [today - 1]]), { now: () => clock });
  view.styles.dxfs = [{ font: null, fill: { pattern: 'solid', colour: '#ffeb9c' } }];
  view.conditionals.set('Sales', [{ type: 'timePeriod', timePeriod: 'today', priority: 1, dxfId: 0, formulas: [], cfvos: [], colours: [], ranges: [{ top: 0, left: 0, bottom: 1, right: 0 }], stopIfTrue: false }]);
  const fires = (row) => Boolean(view.render().cells.find((c) => c.row === row && c.col === 0)?.style?.fill);
  assert.equal(fires(0), true);
  assert.equal(fires(1), false);
});

test('a deck made in another language keeps its slide titles current', () => {
  const deck = Deck.open(buildPptx({ slides: [{ title: 'One' }, { title: 'Two' }, { title: 'Three' }] }));
  const app = '<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    + '<Slides>2</Slides><HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Tema</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Títulos de diapositiva</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs>'
    + '<TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Office</vt:lpstr><vt:lpstr>Stale one</vt:lpstr><vt:lpstr>Stale two</vt:lpstr></vt:vector></TitlesOfParts></Properties>';
  if (deck.pkg.has('docProps/app.xml')) deck.pkg.write_('docProps/app.xml', app);
  else deck.pkg.addPart('docProps/app.xml', app);
  deck.removeSlide(2);
  const saved = Deck.open(deck.save());
  const xml = saved.pkg.text('docProps/app.xml');
  assert.doesNotMatch(xml, /Stale/);
  assert.match(xml, /<vt:lpstr>Office<\/vt:lpstr>/, 'the theme name is left');
});
