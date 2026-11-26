/**
 * A sheet tab given a colour, and a whole sheet copied. Before 2.28 the tab
 * menu held Insert, Delete, Rename, Move left, Move right and Hide: a tab had
 * no colour and a sheet could only be built again by hand.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';

const book = () => buildXlsx({
  sheets: [
    { name: 'Budget', rows: [['Item', 'Cost', 'Twice'], ['Ink', 2, '=B2*2'], ['Paper', 3, '=B3*2'], ['All', '=SUM(B2:B3)', '=Budget!B4*2']] },
    { name: 'Notes', rows: [['Read', '=Budget!B4']] },
  ],
});
const open = () => SheetView.open(book());
const partOf = (bytes, sheet) => {
  const wb = Workbook.open(bytes);
  return wb.pkg.text(wb.partNameFor(sheet));
};

// ---- the tab's colour -------------------------------------------------------

test('a tab takes a colour, the file holds it as Excel writes it, and it is read back', () => {
  const view = open();
  assert.deepEqual(view.tabColors(), {});
  view.setTabColor('Budget', '#00B050');
  assert.deepEqual(view.tabColors(), { Budget: '#00B050' });
  assert.deepEqual(view.render().tabColors, { Budget: '#00B050' }, 'the window is told');
  const xml = partOf(view.save(), 'Budget');
  assert.match(xml, /<worksheet\b[^>]*><sheetPr><tabColor rgb="FF00B050"\/><\/sheetPr>/, 'sheetPr first in the sheet, the colour inside it');
  assert.deepEqual(SheetView.open(view.save()).tabColors(), { Budget: '#00B050' });
});

test('a second colour replaces the first, null takes it off, and each is one undo step', () => {
  const view = open();
  view.setTabColor('Notes', 'FF0000');
  view.setTabColor('Notes', '0070C0');
  assert.deepEqual(view.tabColors(), { Notes: '#0070C0' });
  assert.equal((partOf(view.save(), 'Notes').match(/<tabColor/g) || []).length, 1);
  view.setTabColor('Notes', null);
  assert.deepEqual(view.tabColors(), {});
  assert.doesNotMatch(partOf(view.save(), 'Notes'), /tabColor/);
  view.undo();
  assert.deepEqual(view.tabColors(), { Notes: '#0070C0' });
  view.undo();
  assert.deepEqual(view.tabColors(), { Notes: '#FF0000' });
  view.undo();
  assert.deepEqual(view.tabColors(), {});
});

test('the colour goes first inside a sheetPr that is already there', () => {
  const wb = Workbook.open(book());
  const part = wb._sheetPart('Budget').part;
  part.prefix = part.prefix.replace(/(<worksheet\b[^>]*>)/, '$1<sheetPr filterMode="1"><pageSetUpPr fitToPage="1"/></sheetPr>');
  part.setTabColor('FFC000');
  assert.match(part.prefix, /<sheetPr filterMode="1"><tabColor rgb="FFFFC000"\/><pageSetUpPr fitToPage="1"\/><\/sheetPr>/);
  assert.deepEqual(part.tabColor(), { rgb: 'FFC000' });
});

test('a colour Excel wrote from the theme row is shown, and one that is not a colour is refused', () => {
  const wb = Workbook.open(book());
  const part = wb._sheetPart('Budget').part;
  part.prefix = part.prefix.replace(/(<worksheet\b[^>]*>)/, '$1<sheetPr><tabColor theme="4" tint="0.4"/></sheetPr>');
  part.dirty = true;
  assert.deepEqual(Workbook.open(wb.save()).tabColor('Budget'), { theme: 4, tint: 0.4 }, 'read from a sheet nothing has opened yet');
  const view = SheetView.open(wb.save());
  // This workbook carries no theme part, so there is no colour to work out; with one, the tint is applied.
  assert.equal(view.tabColors().Budget, undefined);
  view.setTabColor('Budget', '00B0F0');
  assert.deepEqual(view.tabColors(), { Budget: '#00B0F0' }, 'a colour picked here replaces the theme one');
  assert.throws(() => view.setTabColor('Budget', 'green'), /six hex digits/);
  assert.throws(() => view.setTabColor('Nowhere', 'FF0000'), /no such sheet/);
});

test('a workbook whose structure is locked keeps its tab colours and takes no copy', () => {
  const view = open();
  view.protectWorkbook();
  assert.throws(() => view.setTabColor('Budget', 'FF0000'));
  assert.throws(() => view.copySheet('Budget'));
  assert.deepEqual(view.sheetNames(), ['Budget', 'Notes']);
});

// ---- a sheet copied ----------------------------------------------------------

test('a copy sits straight after its sheet under Excel\'s name, with the cells, the formulas and the formats', () => {
  const view = open();
  view.select(0, 0);
  view.select(0, 2, { extend: true });
  view.setFormat({ bold: true, fill: '#FFFF00' });
  view.setColWidth(0, 160);
  view.setTabColor('Budget', '7030A0');
  const made = view.copySheet('Budget');
  assert.equal(made.name, 'Budget (2)');
  assert.deepEqual(made.left, []);
  assert.deepEqual(view.sheetNames(), ['Budget', 'Budget (2)', 'Notes']);
  assert.equal(view.activeSheet, 'Budget (2)', 'the view is on the copy');
  assert.equal(view.calc.getValue('Budget (2)', 3, 1), 5);
  assert.equal(view.calc.getValue('Budget (2)', 1, 2), 4);
  assert.equal(view.geo.colWidth(0), 160, 'the column width came along');
  const head = view.render().cells.find((c) => c.ref === 'A1');
  assert.equal(head.style?.font?.bold, true, 'and the format');
  assert.equal(view.tabColors()['Budget (2)'], '#7030A0', 'and the tab colour');

  // The copy is its own sheet: a change there leaves the first alone.
  view.setCell(1, 1, 10);
  assert.equal(view.calc.getValue('Budget (2)', 3, 1), 13);
  assert.equal(view.calc.getValue('Budget', 3, 1), 5);
  // A formula that named its own sheet reads the copy; one on another sheet still reads the first.
  assert.equal(view.calc.getValue('Budget (2)', 3, 2), 26);
  assert.equal(view.calc.getValue('Notes', 0, 1), 5);
});

test('the copy is in the file, opens again, and a second copy is (3)', () => {
  const view = open();
  view.copySheet('Budget');
  view.selectSheet('Budget');
  assert.equal(view.copySheet().name, 'Budget (3)');
  assert.deepEqual(view.sheetNames(), ['Budget', 'Budget (3)', 'Budget (2)', 'Notes']);
  const again = SheetView.open(view.save());
  assert.deepEqual(again.sheetNames(), ['Budget', 'Budget (3)', 'Budget (2)', 'Notes']);
  assert.equal(again.calc.getValue('Budget (2)', 3, 1), 5);
  assert.equal(again.calc.getValue('Budget (3)', 2, 0), 'Paper');
  const wb = Workbook.open(view.save());
  const main = wb.pkg.text(wb.mainPart);
  const ids = [...main.matchAll(/<sheet\b[^>]*\bsheetId="(\d+)"/g)].map((m) => m[1]);
  assert.equal(new Set(ids).size, 4, 'every sheet has an id of its own');
  const selected = wb.sheetNames().filter((n) => /tabSelected="1"/.test(wb.pkg.text(wb.partNameFor(n))));
  assert.ok(selected.length <= 1, 'no two tabs selected together, which Excel opens as a group');
});

test('one undo takes the copy out, and redo puts it back', () => {
  const view = open();
  view.copySheet('Budget');
  view.undo();
  assert.deepEqual(view.sheetNames(), ['Budget', 'Notes']);
  assert.equal(Workbook.open(view.save()).sheetNames().length, 2);
  view.redo();
  assert.deepEqual(view.sheetNames(), ['Budget', 'Budget (2)', 'Notes']);
  assert.equal(view.calc.getValue('Budget (2)', 3, 1), 5);
});

test('links, notes, merged cells, rules, validation, hidden rows and the print area come with the copy', () => {
  const view = open();
  view.setHyperlink({ row: 1, col: 0, href: 'https://example.com/ink' });
  view.setNote({ row: 2, col: 0, author: 'Sana', text: 'Order more' });
  view.select(5, 0);
  view.select(5, 2, { extend: true });
  view.mergeSelection();
  view.select(1, 1);
  view.select(2, 1, { extend: true });
  view.addConditionalRule({ kind: 'top10', rank: 1 });
  view.addValidationRule({ kind: 'whole', operator: 'between', values: [0, 100] });
  view.select(2, 0);
  view.hideLines('row');
  view.workbook.setDefinedName('_xlnm.Print_Area', 'Budget!$A$1:$C$4', { localSheetId: 0 });

  const made = view.copySheet('Budget');
  assert.deepEqual(made.left, []);
  const wb = Workbook.open(view.save());
  assert.deepEqual(wb.hyperlinks('Budget (2)').map((h) => [h.ref, h.href]), [['A2', 'https://example.com/ink']]);
  assert.deepEqual(wb.comments('Budget (2)').map((n) => [n.ref, n.text]), [['A3', 'Order more']]);
  assert.deepEqual(wb.hyperlinks('Budget').map((h) => h.href), ['https://example.com/ink'], 'the first sheet keeps its own');
  assert.equal(wb.comments('Budget').length, 1);
  const xml = wb.pkg.text(wb.partNameFor('Budget (2)'));
  assert.match(xml, /<mergeCell ref="A6:C6"\/>/);
  assert.match(xml, /<conditionalFormatting\b/);
  assert.match(xml, /<dataValidation\b/);
  assert.match(xml, /<row[^>]*r="3"[^>]*hidden="1"/);
  const areas = wb.definedNames().filter((d) => d.name === '_xlnm.Print_Area');
  assert.deepEqual(areas.map((d) => d.ref).sort(), ["'Budget (2)'!$A$1:$C$4", 'Budget!$A$1:$C$4']);
  assert.match(areas.find((d) => d.ref.startsWith("'")).attrsStr, /localSheetId="1"/, 'scoped to the copy, which is the second sheet');
  // Every relationship the copy names is one it has.
  const rels = new Set(wb.pkg.rels(wb.partNameFor('Budget (2)')).map((r) => r.Id));
  for (const m of xml.matchAll(/\br:id="([^"]+)"/g)) assert.ok(rels.has(m[1]), 'relationship ' + m[1] + ' exists');
});

test('a table and a chart stay on the first sheet: the copy has the cells and says what it left', () => {
  const view = open();
  view.select(0, 0);
  view.select(2, 1, { extend: true });
  view.formatAsTable();
  view.insertChart({ kind: 'column' });
  const made = view.copySheet('Budget');
  assert.deepEqual([...made.left].sort(), ['drawings', 'tables']);
  const wb = Workbook.open(view.save());
  const xml = wb.pkg.text(wb.partNameFor('Budget (2)'));
  assert.doesNotMatch(xml, /<tableParts|<drawing\b/);
  assert.equal(wb.tables().length, 1, 'one table still, on the first sheet');
  const rels = new Set(wb.pkg.rels(wb.partNameFor('Budget (2)')).map((r) => r.Id));
  for (const m of xml.matchAll(/\br:id="([^"]+)"/g)) assert.ok(rels.has(m[1]));
  assert.equal(SheetView.open(view.save()).calc.getValue('Budget (2)', 1, 0), 'Ink');
});

test('a long name is cut so the copy\'s name fits 31 characters, and a taken name is refused', () => {
  const long = 'A sheet with a very long name 1';
  const view = SheetView.open(buildXlsx({ sheets: [{ name: long, rows: [[1]] }] }));
  const made = view.copySheet();
  assert.equal(made.name.length <= 31, true);
  assert.match(made.name, / \(2\)$/);
  assert.throws(() => view.copySheet(long, long), /already a sheet/);
  assert.equal(view.copySheet(long, 'Spare').name, 'Spare');
});
