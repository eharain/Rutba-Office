// What a workbook from elsewhere holds that the engine has to respect: stored
// results it cannot work out, names that belong to one sheet, tabs with
// awkward names, and dates counted from 1904.
import test from 'node:test';
import assert from 'node:assert/strict';

import { Spreadsheet, isError } from '@rutba/formula';
import { buildXlsx } from '@rutba/ooxml/build';
import { Workbook, recalculateWorkbook } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';

/** A one-sheet workbook whose cells are the given XML. */
function withCells(sheetData, { workbookEdit = (x) => x } = {}) {
  const wb = Workbook.open(buildXlsx({ sheets: [{ name: 'S', rows: [[1]] }] }));
  const part = wb.partNameFor('S');
  wb.pkg.write_(part, wb.pkg.text(part).replace(/<sheetData>[\s\S]*<\/sheetData>|<sheetData\/>/, () => '<sheetData>' + sheetData + '</sheetData>'));
  wb.pkg.write_(wb.mainPart, workbookEdit(wb.pkg.text(wb.mainPart)));
  return wb.save();
}
const stored = (bytes, ref) => {
  const wb = Workbook.open(bytes);
  const xml = wb.pkg.text(wb.partNameFor('S'));
  return new RegExp('<c r="' + ref + '"[^>]*>[\\s\\S]*?</c>').exec(xml)[0];
};

test('a reference to another workbook keeps the result the file stored, through a recalculation and a save', () => {
  const bytes = withCells('<row r="1"><c r="A1"><f>[1]Book!B2</f><v>42</v></c><c r="B1"><f>A1+1</f><v>43</v></c></row>');
  const { output, unparsed } = recalculateWorkbook(bytes);
  assert.match(stored(output, 'A1'), /<v>42<\/v>/);
  assert.match(stored(output, 'B1'), /<v>43<\/v>/);
  assert.equal(unparsed.length, 1);
  const view = new SheetView(bytes);
  assert.equal(view.calc.getValue('S', 0, 0), 42);
  assert.match(stored(view.save(), 'A1'), /<v>42<\/v>/);
});

test('a function nothing here knows keeps the stored result, of whatever kind', () => {
  const bytes = withCells('<row r="1"><c r="A1"><f>NOTAFUNCTION(1)</f><v>42</v></c><c r="B1" t="str"><f>NOTAFUNCTION(2)</f><v>words</v></c>'
    + '<c r="C1" t="b"><f>NOTAFUNCTION(3)</f><v>1</v></c><c r="D1" t="e"><f>NOTAFUNCTION(4)</f><v>#DIV/0!</v></c></row>');
  const view = new SheetView(bytes);
  assert.equal(view.calc.getValue('S', 0, 0), 42);
  assert.equal(view.calc.getValue('S', 0, 1), 'words');
  assert.equal(view.calc.getValue('S', 0, 2), true);
  assert.equal(view.calc.getValue('S', 0, 3).type, '#DIV/0!');
  const saved = view.save();
  assert.match(stored(saved, 'A1'), /<v>42<\/v>/);
  assert.match(stored(saved, 'B1'), /<v>words<\/v>/);
});

test('with nothing stored, or typed here, an unknown function is still #NAME?', () => {
  const s = new Spreadsheet();
  s.setByRef('S', 'A1', '=NOTAFUNCTION(1)');
  s.recalculate();
  assert.equal(s.getValue('S', 0, 0).type, '#NAME?');
  const none = new SheetView(withCells('<row r="1"><c r="A1"><f>NOTAFUNCTION(1)</f></c></row>'));
  assert.equal(none.calc.getValue('S', 0, 0).type, '#NAME?');
});

test('words that look like a workbook reference are only words', () => {
  const s = new Spreadsheet();
  s.setCell('S', 0, 0, '="[1]Book"&"x"', { cached: 'old' });
  s.recalculate();
  assert.equal(s.getValue('S', 0, 0), '[1]Bookx');
});

test('a formula that can be worked out ignores what was stored', () => {
  const s = new Spreadsheet();
  s.setCell('S', 0, 0, 2);
  s.setCell('S', 0, 1, '=A1*10', { cached: 999 });
  s.recalculate();
  assert.equal(s.getValue('S', 0, 1), 20);
});

test('INDIRECT and OFFSET follow the cells they build a reference to', () => {
  const s = new Spreadsheet();
  s.setByRef('S', 'A1', 1);
  s.setByRef('S', 'B1', '=INDIRECT("A1")');
  s.setByRef('S', 'C1', '=OFFSET(A1,0,0)');
  s.setByRef('S', 'D1', '=B1+C1');
  s.recalculate();
  s.setByRef('S', 'A1', 5);
  s.recalculate();
  assert.deepEqual([s.getValue('S', 0, 1), s.getValue('S', 0, 2), s.getValue('S', 0, 3)], [5, 5, 10]);
});

const NAMES = {
  sheets: [{ name: 'One', rows: [[1, '=Rate', '=Only']] }, { name: 'Two', rows: [[2, '=Rate', '=Only', '=One!Rate']] }, { name: 'Three', rows: [[3, '=Rate']] }],
};

function namedBook() {
  const wb = Workbook.open(buildXlsx(NAMES));
  const names = '<definedNames>'
    + '<definedName name="Only" localSheetId="0">One!$A$1</definedName>'
    + '<definedName name="Rate">Three!$A$1</definedName>'
    + '<definedName name="Rate" localSheetId="0">One!$A$1</definedName>'
    + '<definedName name="Rate" localSheetId="1">Two!$A$1</definedName>'
    + '</definedNames>';
  const xml = wb.pkg.text(wb.mainPart).replace(/<definedNames>[\s\S]*?<\/definedNames>/, '');
  wb.pkg.write_(wb.mainPart, xml.replace('</sheets>', () => '</sheets>' + names));
  return wb.save();
}

test('a name belongs to its sheet, hides the workbook\'s there, and can be asked for by sheet', () => {
  const view = new SheetView(namedBook());
  const at = (sheet, col) => view.calc.getValue(sheet, 0, col);
  assert.equal(at('One', 1), 1, 'One has its own Rate');
  assert.equal(at('Two', 1), 2, 'Two has its own Rate');
  assert.equal(at('Three', 1), 3, 'Three uses the workbook\'s');
  assert.equal(at('One', 2), 1);
  assert.equal(at('Two', 3), 1, 'One!Rate from another sheet');
  assert.ok(isError(at('Two', 2)), 'a name on One is not known on Two');
});

test('a sheet-scoped name is followed when its cell changes', () => {
  const view = new SheetView(namedBook());
  view.calc.setCell('Two', 0, 0, 20);
  view.calc.recalculate();
  assert.equal(view.calc.getValue('Two', 0, 1), 20);
  assert.equal(view.calc.getValue('One', 0, 1), 1);
});

test('a tab with "!" in its name can be edited, styled and saved', () => {
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'Q1!', rows: [[1, '=A1+1']] }] }));
  view.calc.setCell('Q1!', 0, 0, 7);
  view.dirtyCells.add('Q1!!A1');
  view.styledCells.add('Q1!!A1');
  const again = Workbook.open(view.save());
  assert.deepEqual([...again.cells('Q1!')].map((c) => c.ref), ['A1', 'B1']);
});

test('1904 workbooks: the file\'s serials mean 1904 dates when shown and when worked out', () => {
  const bytes = withCells('<row r="1"><c r="A1"><v>0</v></c><c r="B1"><f>YEAR(A1)</f><v>0</v></c><c r="C1"><f>DATE(1904,1,2)</f><v>0</v></c></row>', {
    workbookEdit: (x) => (/<workbookPr\b/.test(x) ? x.replace(/<workbookPr\b/, '<workbookPr date1904="1"') : x.replace('<sheets>', '<workbookPr date1904="1"/><sheets>')),
  });
  const view = new SheetView(bytes);
  assert.equal(view.date1904, true);
  assert.equal(view.calc.getValue('S', 0, 1), 1904);
  assert.equal(view.calc.getValue('S', 0, 2), 1);
});

test('the date system goes back to 1900 for the next workbook', async () => {
  const { withDate1904, serialToParts } = await import('@rutba/formula');
  assert.equal(withDate1904(true, () => serialToParts(0).y), 1904);
  assert.equal(serialToParts(1).y, 1900);
  assert.throws(() => withDate1904(true, () => { throw new Error('x'); }));
  assert.equal(serialToParts(1).y, 1900, 'a throw does not leave it on');
});
