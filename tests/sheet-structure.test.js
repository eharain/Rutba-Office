/**
 * Sheet structure: resizing rows and columns, inserting and deleting them.
 *
 * A resize is a small thing with one hard requirement — the pixel width chosen
 * on screen must reopen as the SAME pixel width, which means the px→character
 * conversion has to be the exact inverse of the read. Insert and delete are
 * the opposite: small trigger, wide blast radius. Rows renumber, cell refs
 * move, merges travel, `<col>` records shift, and — above all — every formula
 * in the workbook that points at the edited axis adjusts, including the ones
 * on OTHER sheets and the defined names in workbook.xml. These tests pin the
 * Excel semantics:
 *
 *   - a `=SUM` range spanning an insertion GROWS; a delete inside it SHRINKS;
 *   - a reference INTO a deleted slice becomes #REF!;
 *   - absolute (`$A$5`) references adjust exactly like relative ones;
 *   - a delete that would split a merged cell is refused, changing nothing;
 *   - the whole edit is ONE undo step that restores the sheet XML byte-for-byte;
 *
 * — and, as always, that a workbook none of this touches round-trips
 * byte-identical.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const book = (rows, extra = {}) =>
  buildXlsx({ sheets: [{ name: 'S', rows, ...extra }] });

const open = (bytes) => new SheetView(bytes);

const sheetXml = (view, name = 'S') => {
  const part = view.workbook.partNameFor(name);
  return view.workbook.snapshotParts([part])[part];
};

// ── resizing ─────────────────────────────────────────────────────────────────

test('a column resize survives save and reopen at the same pixel width', () => {
  const view = open(book([['a', 'b', 'c']]));
  view.setColWidth(2, 120);
  assert.equal(view.geo.colWidth(2), 120, 'the grid draws the new width at once');
  assert.equal(view.isDirty, true, 'a resize is unsaved work');

  const reopened = open(view.save());
  assert.equal(reopened.geo.colWidth(2), 120, 'the width round-tripped exactly');
  assert.equal(reopened.geo.colWidth(1), reopened.geo.defaultColWidth, 'neighbours untouched');
});

test('a row resize survives save and reopen at the same pixel height', () => {
  const view = open(book([['a'], ['b']]));
  view.setRowHeight(1, 41);
  assert.equal(view.geo.rowHeight(1), 41);

  const reopened = open(view.save());
  assert.equal(reopened.geo.rowHeight(1), 41, 'the height round-tripped exactly');
  assert.equal(reopened.geo.rowHeight(0), reopened.geo.defaultRowHeight);
});

test('every plausible pixel width round-trips through the file exactly', () => {
  const view = open(book([['a']]));
  for (const px of [16, 23, 64, 97, 120, 255, 400]) {
    view.setColWidth(0, px);
    const reopened = open(view.save());
    assert.equal(reopened.geo.colWidth(0), px, px + 'px column width survives');
  }
  for (const px of [12, 20, 41, 63, 100]) {
    view.setRowHeight(0, px);
    const reopened = open(view.save());
    assert.equal(reopened.geo.rowHeight(0), px, px + 'px row height survives');
  }
});

test('a resize clamps to the minimums instead of vanishing the column', () => {
  const view = open(book([['a']]));
  view.setColWidth(0, 3);
  view.setRowHeight(0, 2);
  assert.equal(view.geo.colWidth(0), 16, 'columns stop at 16px');
  assert.equal(view.geo.rowHeight(0), 12, 'rows stop at 12px');
});

test('undoing a resize restores the sheet XML byte-for-byte', () => {
  const view = open(book([['a', 'b']]));
  const before = sheetXml(view);
  view.setColWidth(1, 200);
  assert.notEqual(sheetXml(view), before);
  view.undo();
  assert.equal(sheetXml(view), before, 'the cols record is gone without a trace');
  assert.equal(view.geo.colWidth(1), view.geo.defaultColWidth, 'and the grid agrees');
  assert.equal(view.isDirty, false, 'nothing left to save');
});

// ── inserting rows ───────────────────────────────────────────────────────────

test('inserting a row renumbers what follows and keeps values, styles and merges aligned', () => {
  const view = open(book(
    [['a', 'b'], ['x', 'y'], ['p', 'q']],
    { styles: { A2: { bold: true } }, merges: ['B2:B3'] },
  ));
  view.insertRows(1, 1);

  assert.equal(view.editValue(0, 0), 'a', 'rows above the insertion stay put');
  assert.equal(view.editValue(1, 0), '', 'the inserted row is empty');
  assert.equal(view.editValue(2, 0), 'x', 'rows below moved down by one');
  assert.equal(view.editValue(3, 1), 'q');
  assert.equal(view.styleFor(2, 0)?.font?.bold, true, 'the bold cell moved with its row');
  assert.equal(view.mergeAt(2, 1)?.ref, 'B3:B4', 'the merge moved with its rows');
});

test('a =SUM range spanning the insertion grows to keep covering it', () => {
  const view = open(book([[1], [2], [3], ['=SUM(A1:A3)']]));
  view.insertRows(1, 1);
  assert.equal(view.editValue(4, 0), '=SUM(A1:A4)', 'the range grew');
  assert.equal(view.displayValue(4, 0).text, '6', 'and still sums the same cells');

  const reopened = open(view.save());
  assert.equal(reopened.editValue(4, 0), '=SUM(A1:A4)', 'the grown range was saved');
  assert.equal(reopened.displayValue(4, 0).text, '6');
});

test('inserting several rows at once shifts by the full count', () => {
  const view = open(book([[1], ['=A1*10']]));
  view.insertRows(1, 3);
  assert.equal(view.editValue(4, 0), '=A1*10', 'the formula rode down three rows');
  assert.equal(view.displayValue(4, 0).text, '10');
});

// ── deleting rows ────────────────────────────────────────────────────────────

test('deleting inside a =SUM range shrinks it', () => {
  const view = open(book([[1], [2], [3], ['=SUM(A1:A3)']]));
  view.deleteRows(1, 1);
  assert.equal(view.editValue(2, 0), '=SUM(A1:A2)', 'the range shrank');
  assert.equal(view.displayValue(2, 0).text, '4', '1 + 3, the deleted 2 is gone');
});

test('a formula referencing a deleted row yields #REF!', () => {
  const view = open(book([[1], [2], ['=A2*2']]));
  view.deleteRows(1, 1);
  assert.equal(view.editValue(1, 0), '=#REF!*2', 'the dead reference is spelled out');
  const value = view.calc.getValue('S', 1, 0);
  assert.equal(String(value.type ?? value), '#REF!', 'and evaluates to the error');
});

test('absolute references adjust exactly like relative ones', () => {
  const view = open(book([['=$A$5'], [''], [''], [''], [7]]));
  view.insertRows(1, 1);
  assert.equal(view.editValue(0, 0), '=$A$6', 'the $ pins the ref against copying, not against inserts');
  assert.equal(view.displayValue(0, 0).text, '7');
});

test('references from another sheet adjust when the referenced sheet shifts', () => {
  const bytes = buildXlsx({
    sheets: [
      { name: 'Data', rows: [[1], [2]] },
      { name: 'Report', rows: [['=Data!A2', "='Data'!$A$2"]] },
    ],
  });
  const view = open(bytes); // Data is active
  view.insertRows(0, 1);
  assert.equal(view.calc.getInput('Report', 0, 0), '=Data!A3', 'the bare qualifier moved');
  assert.equal(view.calc.getInput('Report', 0, 1), "='Data'!$A$3", 'the quoted, absolute one too');
  assert.equal(view.calc.getValue('Report', 0, 0), 2, 'and still reads the same cell');
});

test('a sheet name with a space adjusts through its quoted references', () => {
  const bytes = buildXlsx({
    sheets: [
      { name: 'Raw Data', rows: [[1], [2], [3]] },
      { name: 'Report', rows: [["='Raw Data'!A3"]] },
    ],
  });
  const view = open(bytes); // 'Raw Data' is active
  view.deleteRows(0, 1);
  assert.equal(view.calc.getInput('Report', 0, 0), "='Raw Data'!A2", 'the quoted ref shifted up');
  assert.equal(view.calc.getValue('Report', 0, 0), 3);
});

test('a cross-sheet reference into a deleted row becomes #REF! with its qualifier kept', () => {
  const bytes = buildXlsx({
    sheets: [
      { name: 'Data', rows: [[1], [2]] },
      { name: 'Report', rows: [['=Data!A2+1']] },
    ],
  });
  const view = open(bytes);
  view.deleteRows(1, 1);
  assert.equal(view.calc.getInput('Report', 0, 0), '=Data!#REF!+1');
});

test('cell-shaped text inside string literals is left alone', () => {
  const view = open(book([[5], ['=CONCATENATE("A1 is ", A1)']]));
  view.insertRows(0, 1);
  assert.equal(view.editValue(2, 0), '=CONCATENATE("A1 is ", A2)',
    'the reference moved, the words did not');
});

test('a delete that would split a merged cell is refused and changes nothing', () => {
  const view = open(book([['a'], ['b'], ['c'], ['d']], { merges: ['A2:A4'] }));
  assert.throws(() => view.deleteRows(1, 1), /split a merged cell/);
  assert.equal(view.mergeAt(1, 0)?.ref, 'A2:A4', 'the merge is intact');
  assert.equal(view.editValue(2, 0), 'c', 'the covered value is exactly where it was');
  assert.equal(view.canUndo, false, 'no undo step was recorded for a refusal');
  assert.equal(view.isDirty, false, 'and nothing reads as unsaved');
});

test('a merge fully inside the deleted slice goes with it', () => {
  const view = open(book([['a'], ['b'], ['c'], ['d']], { merges: ['A2:A3'] }));
  view.deleteRows(1, 2);
  assert.equal((view.merges.get('S') ?? []).length, 0, 'the merge went with its rows');
  assert.equal(view.editValue(1, 0), 'd', 'the row below closed the gap');
});

test('defined names shift with the rows they point at', () => {
  const bytes = buildXlsx({
    sheets: [{ name: 'S', rows: [[1], [2], [3]] }],
    definedNames: [{ name: 'TOTAL', ref: 'S!$A$3' }],
  });
  const view = open(bytes);
  view.insertRows(0, 2);
  assert.equal(view.workbook.definedNames()[0].ref, 'S!$A$5', 'the name follows its cell');

  const reopened = open(view.save());
  assert.equal(reopened.workbook.definedNames()[0].ref, 'S!$A$5', 'and survives the file');
});

// ── columns ──────────────────────────────────────────────────────────────────

test('inserting a column re-letters cells and grows a spanning range', () => {
  const view = open(book([[1, 2, 3, '=SUM(A1:C1)']]));
  view.insertCols(1, 1);
  assert.equal(view.editValue(0, 0), '1', 'A stays');
  assert.equal(view.editValue(0, 1), '', 'the inserted column is empty');
  assert.equal(view.editValue(0, 2), '2', 'B moved to C');
  assert.equal(view.editValue(0, 4), '=SUM(A1:D1)', 'the range grew sideways');
  assert.equal(view.displayValue(0, 4).text, '6');
});

test('deleting a column shrinks ranges and kills direct references', () => {
  const view = open(book([[1, 2, 3, '=SUM(A1:C1)', '=B1']]));
  view.deleteCols(1, 1);
  assert.equal(view.editValue(0, 1), '3', 'C moved into B');
  assert.equal(view.editValue(0, 2), '=SUM(A1:B1)', 'the range shrank');
  assert.equal(view.displayValue(0, 2).text, '4');
  assert.equal(view.editValue(0, 3), '=#REF!', 'the direct reference died honestly');
});

test('a column width travels with its column through an insert', () => {
  const view = open(book([['a', 'b', 'c']]));
  view.setColWidth(2, 120);
  view.insertCols(0, 1);
  assert.equal(view.geo.colWidth(3), 120, 'the wide column is one to the right now');
  assert.equal(view.geo.colWidth(2), view.geo.defaultColWidth, 'its old slot is default');

  const reopened = open(view.save());
  assert.equal(reopened.geo.colWidth(3), 120, 'and the file says the same');
});

// ── one undo step ────────────────────────────────────────────────────────────

test('undo restores the sheet XML byte-for-byte after an insert', () => {
  const view = open(book([[1], [2], ['=SUM(A1:A2)']], { merges: ['A1:B1'] }));
  const before = sheetXml(view);
  view.insertRows(1, 1);
  assert.notEqual(sheetXml(view), before, 'the edit really rewrote the part');

  view.undo();
  assert.equal(sheetXml(view), before, 'byte-identical, merges and all');
  assert.equal(view.editValue(2, 0), '=SUM(A1:A2)', 'the formula reads as before');
  assert.equal(view.mergeAt(0, 0)?.ref, 'A1:B1');
  assert.equal(view.isDirty, false, 'nothing left to save');
});

test('redo puts the structural edit back', () => {
  const view = open(book([[1], [2], ['=SUM(A1:A2)']]));
  view.insertRows(0, 1);
  view.undo();
  view.redo();
  assert.equal(view.editValue(3, 0), '=SUM(A2:A3)', 'the shifted formula is back');
  assert.equal(view.displayValue(3, 0).text, '3');
  assert.equal(view.isDirty, true, 'and the document reads as unsaved again');
});

test('an unsaved cell edit survives a structural edit and its undo', () => {
  const view = open(book([[1], [2]]));
  view.setCell(0, 0, '9'); // pending: lives in the calc model, not the file yet
  view.insertRows(1, 1);
  assert.equal(view.displayValue(0, 0).text, '9', 'the pending edit survived the flush');

  const reopened = open(view.save());
  assert.equal(reopened.displayValue(0, 0).text, '9', 'and reached the file');

  view.undo(); // undo the insert
  assert.equal(view.displayValue(0, 0).text, '9', 'undo of the insert kept the typed value');
  assert.equal(view.editValue(1, 0), '2', 'and closed the inserted row');
  view.undo(); // undo the typing
  assert.equal(view.displayValue(0, 0).text, '1', 'a second undo reverts the typing itself');
});

// ── fidelity: a workbook none of this touches ────────────────────────────────

test('a workbook untouched by any of this still round-trips byte-identical', () => {
  // No formulas on purpose: save() has always written fresh cached values next
  // to formulas, which is its own (tested) behaviour. This pins that STRUCTURE
  // code adds no rewrite of its own.
  const bytes = book([['a', 'b'], [1, 2]], { merges: ['A1:B1'] });
  const view = open(bytes);
  const out = view.save();
  assert.deepEqual(Buffer.from(out), Buffer.from(bytes), 'no structural edit, no rewrite');
});
