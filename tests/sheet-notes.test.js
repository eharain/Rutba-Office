// A note on a worksheet cell, written as Excel keeps one.
//
// The comments part carries the words and the author; a VML drawing part
// carries the box Excel draws them in, and the sheet points at both. A note
// on a cell that has one replaces it; a removed note leaves the parts,
// emptied, so an undo has somewhere to put it back.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { parseComments, commentsXml } from '@rutba/ooxml/workbook';

const BOOK = () => buildXlsx({ sheets: [{ name: 'Sales', rows: [['Item', 'Qty'], ['Pens', 12], ['Ink', 3]] }] });

test('a note is written as Excel keeps one: comments part, VML box, the sheet pointing at both, and read back', () => {
  const view = new SheetView(BOOK());
  view.setNote({ row: 1, col: 1, author: 'Kim Lee', text: 'Check this figure.' });
  const frame = view.render();
  assert.equal(frame.cells.find((c) => c.ref === 'B2').note.text, 'Check this figure.', 'the frame shows the note at once');
  assert.equal(frame.notes, 1);

  const saved = SheetView.open(view.save());
  const sheetPart = saved.workbook.partNameFor('Sales');
  const rels = saved.pkg.rels(sheetPart);
  const commentsRel = rels.find((r) => r.Type.endsWith('/comments'));
  const vmlRel = rels.find((r) => r.Type.endsWith('/vmlDrawing'));
  assert.ok(commentsRel && vmlRel, 'the sheet has a comments relationship and a VML one');
  assert.deepEqual(saved.workbook.comments('Sales'), [{ ref: 'B2', author: 'Kim Lee', text: 'Check this figure.' }]);
  const xml = saved.pkg.text(sheetPart);
  assert.match(xml, new RegExp('<legacyDrawing r:id="' + vmlRel.Id + '"/>'), 'the sheet points at the VML drawing');
  assert.match(xml, /<worksheet[^>]*\sxmlns:r=/, 'the r prefix is declared');
  const vml = saved.pkg.text('xl/drawings/vmlDrawing1.vml');
  assert.match(vml, /<v:shapetype id="_x0000_t202"/);
  assert.match(vml, /ObjectType="Note"[\s\S]*<x:Row>1<\/x:Row>\s*<x:Column>1<\/x:Column>/, 'the box is anchored on B2');
  assert.equal(saved.pkg.contentTypes().defaults.get('vml'), 'application/vnd.openxmlformats-officedocument.vmlDrawing');
  assert.equal(saved.pkg.contentTypeOf('xl/comments1.xml'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml');
  assert.equal(saved.render().cells.find((c) => c.ref === 'B2').note.author, 'Kim Lee', 'the reopened file shows it');
});

test('a second note on the same cell replaces the first; another cell adds a box; a note comes off; undo puts it back', () => {
  const view = new SheetView(BOOK());
  view.setNote({ row: 1, col: 1, author: 'Kim Lee', text: 'First.' });
  view.setNote({ row: 1, col: 1, author: 'Kim Lee', text: 'Second.' });
  view.setNote({ row: 2, col: 0, author: 'Sam Roe', text: 'Ink is low.' });
  let notes = view.workbook.comments('Sales');
  assert.deepEqual(notes.map((n) => n.ref + ':' + n.text), ['B2:Second.', 'A3:Ink is low.']);
  const boxes = (view.pkg.text('xl/drawings/vmlDrawing1.vml').match(/<v:shape\b/g) || []).length;
  assert.equal(boxes, 2, 'one box per note, the replaced one not doubled');
  const ids = [...view.pkg.text('xl/drawings/vmlDrawing1.vml').matchAll(/id="_x0000_s(\d+)"/g)].map((m) => m[1]);
  assert.equal(new Set(ids).size, ids.length, 'every box has an id of its own');

  assert.equal(view.removeNote({ row: 1, col: 1 }).render().notes, 1);
  notes = view.workbook.comments('Sales');
  assert.deepEqual(notes.map((n) => n.ref), ['A3']);
  assert.equal((view.pkg.text('xl/drawings/vmlDrawing1.vml').match(/<v:shape\b/g) || []).length, 1, 'its box went with it');

  view.undo();
  assert.deepEqual(view.workbook.comments('Sales').map((n) => n.ref + ':' + n.text), ['B2:Second.', 'A3:Ink is low.'], 'undo puts the note back');
  assert.equal(view.render().cells.find((c) => c.ref === 'B2').note.text, 'Second.', 'and the frame reads it again');
  view.undo();
  view.undo();
  view.undo();
  assert.deepEqual(view.workbook.comments('Sales'), [], 'undone to the start, no notes');
  assert.equal(view.render().notes, 0);
  view.redo();
  assert.equal(view.render().cells.find((c) => c.ref === 'B2').note.text, 'First.');
});

test('the comments part reads and writes both ways, authors once each and the words escaped', () => {
  const xml = commentsXml([{ ref: 'A1', author: 'Kim <Lee>', text: 'Fish & chips\n"cheap"' }, { ref: 'B2', author: 'Kim <Lee>', text: 'Again' }, { ref: 'C3', author: '', text: 'Unsigned' }]);
  assert.equal((xml.match(/<author>/g) || []).length, 2, 'two authors: Kim once, and the empty one');
  assert.match(xml, /xml:space="preserve"/);
  assert.deepEqual(parseComments(xml), [
    { ref: 'A1', author: 'Kim <Lee>', text: 'Fish & chips\n"cheap"' },
    { ref: 'B2', author: 'Kim <Lee>', text: 'Again' },
    { ref: 'C3', author: '', text: 'Unsigned' },
  ]);
});
