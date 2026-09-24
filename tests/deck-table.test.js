// A table on a slide — Insert → Table.
//
// Written the way PowerPoint writes one it has just drawn (a graphic frame
// holding a:tbl, Medium Style 2 — Accent 1, banded rows), read back with
// each cell's words and its own box so the window can edit it in place, and
// its rows and columns grown or cut with the frame's own size following.
import test from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Deck, buildPptx, renderSlide } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const RICH = join(dirname(fileURLToPath(import.meta.url)), 'fixtures', 'rich');

const DECK = buildPptx({
  title: 'Tables',
  slides: [
    { layout: 'title', title: 'Cover', body: 'A deck with a table' },
    { layout: 'obj', title: 'Numbers', body: ['One'] },
  ],
});

const STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';
const wordsOf = (shape) => shape.table.rows.map((r) => r.cells.map((c) => (c.text?.paragraphs || []).map((p) => p.runs.map((run) => run.text).join('')).join('')));

test('addTable writes a graphic frame exactly as PowerPoint writes one', () => {
  const deck = Deck.open(DECK);
  const id = deck.addTable(1, { rows: 3, cols: 3 });
  const xml = OoxmlPackage.read(deck.save()).text('ppt/slides/slide2.xml');

  assert.match(xml, new RegExp(`<p:cNvPr id="${id}" name="Table ${id}"`));
  assert.match(xml, /<a:graphicFrameLocks noGrp="1"\/>/);
  assert.match(xml, /<a:graphicData uri="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/table">/);
  assert.match(xml, new RegExp(`<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${STYLE_ID.replace(/[{}]/g, '\\$&')}</a:tableStyleId></a:tblPr>`));

  const cols = [...xml.matchAll(/<a:gridCol w="(\d+)"\/>/g)].map((m) => Number(m[1]));
  assert.equal(cols.length, 3, 'three columns');
  assert.ok(cols.every((w) => w === cols[0]), 'the columns are equal widths');
  const frameW = Number(/<p:xfrm><a:off[^/]*\/><a:ext cx="(\d+)" cy="\d+"\/>/.exec(xml)[1]);
  assert.ok(Math.abs(cols.reduce((a, b) => a + b, 0) - frameW) <= cols.length, 'the columns sum to the frame width');

  assert.equal((xml.match(/<a:tr h="\d+">/g) || []).length, 3, 'three rows');
  assert.equal((xml.match(/<a:tc>/g) || []).length, 9, 'nine cells');
  assert.equal((xml.match(/<a:txBody>/g) || []).length, 9, 'every cell has a text body');
});

test('the reader lists a table shape with its rows, columns, cell text and each cell\'s own box; the renderer draws it', () => {
  const deck = Deck.open(DECK);
  deck.addTable(1, { rows: 2, cols: 3, cells: [['A', 'B', 'C'], ['1', '2', '3']] });
  const reopened = Deck.open(deck.save());
  const slide = reopened.slide(1);
  const shape = slide.shapes.find((s) => s.kind === 'table');
  assert.ok(shape, 'a table shape');
  assert.equal(shape.table.rows.length, 2);
  assert.equal(shape.table.columns.length, 3);
  assert.deepEqual(wordsOf(shape), [['A', 'B', 'C'], ['1', '2', '3']]);

  for (const row of shape.table.rows) {
    for (const cell of row.cells) {
      assert.ok(cell.box, 'a box for the cell');
      assert.ok(cell.box.w > 0 && cell.box.h > 0);
    }
  }
  // Cells tile left to right, top to bottom, with no gap between them.
  const r0 = shape.table.rows[0].cells;
  assert.ok(Math.abs(r0[0].box.x + r0[0].box.w - r0[1].box.x) < 0.5);
  assert.ok(Math.abs(r0[0].box.y + shape.table.rows[0].height - shape.table.rows[1].cells[0].box.y) < 0.5);

  const svg = renderSlide(slide);
  const cellRects = svg.match(/<rect x="[\d.]+" y="[\d.]+" width="[\d.]+" height="[\d.]+" fill="#ffffff" stroke="#c9ccd1" stroke-width="1"\/>/g) || [];
  assert.equal(cellRects.length, 6, 'six cell rectangles drawn');
  assert.match(svg, />A<\/tspan>/);
  assert.match(svg, />C<\/tspan>/);
});

test('setTableCell rewrites one cell, keeps the other eight byte-identical, and a run\'s bold survives', () => {
  const deck = Deck.open(DECK);
  const id = deck.addTable(1, { rows: 3, cols: 3, cells: [['A', 'B', 'C'], ['1', '2', '3'], ['x', 'y', 'z']] });
  const before = deck.pkg.text(deck.slideParts[1].part);
  const cellsBefore = [...before.matchAll(/<a:tc>[\s\S]*?<\/a:tc>/g)].map((m) => m[0]);
  assert.equal(cellsBefore.length, 9);

  deck.setTableCell(1, id, 1, 1, [{ runs: [{ text: 'Total', bold: true }] }]);
  const after = deck.pkg.text(deck.slideParts[1].part);
  const cellsAfter = [...after.matchAll(/<a:tc>[\s\S]*?<\/a:tc>/g)].map((m) => m[0]);
  assert.equal(cellsAfter.length, 9);
  cellsBefore.forEach((c, i) => {
    if (i === 4) { assert.notEqual(cellsAfter[i], c, 'the edited cell changed'); return; }
    assert.equal(cellsAfter[i], c, `cell ${i} should be byte-identical`);
  });
  assert.match(cellsAfter[4], /<a:t[^>]*>Total<\/a:t>/);
  assert.match(cellsAfter[4], /b="1"/);

  const reopened = Deck.open(deck.save());
  const shape = reopened.slide(1).shapes.find((s) => s.id === String(id));
  assert.deepEqual(wordsOf(shape), [['A', 'B', 'C'], ['1', 'Total', '3'], ['x', 'y', 'z']]);
  assert.equal(shape.table.rows[1].cells[1].text.paragraphs[0].runs[0].bold, true, 'the bold run survives');
});

test('a row is inserted, then removed, the frame\'s own height growing and shrinking with it', () => {
  const deck = Deck.open(DECK);
  const id = deck.addTable(1, { rows: 2, cols: 2 });
  const part = deck.slideParts[1].part;
  const frameH = () => Number(/<p:xfrm><a:off[^/]*\/><a:ext cx="\d+" cy="(\d+)"\/>/.exec(deck.pkg.text(part))[1]);
  const rowCount = () => (deck.pkg.text(part).match(/<a:tr h="\d+">/g) || []).length;
  const h0 = frameH();
  assert.equal(rowCount(), 2);

  deck.insertTableRow(1, id, 1);
  assert.equal(rowCount(), 3);
  const h1 = frameH();
  assert.ok(h1 > h0, 'the frame grew by the new row\'s height');

  deck.removeTableRow(1, id, 0);
  assert.equal(rowCount(), 2);
  assert.ok(frameH() < h1, 'the frame shrank by the removed row\'s height');

  const reopened = Deck.open(deck.save());
  assert.equal(reopened.slide(1).shapes.find((s) => s.id === String(id)).table.rows.length, 2, 'survives Deck.open(deck.save())');

  assert.throws(() => deck.removeTableRow(1, id, 50), /no row/);
});

test('a column is inserted, then removed, the columns sharing the frame\'s own width', () => {
  const deck = Deck.open(DECK);
  const id = deck.addTable(1, { rows: 2, cols: 2 });
  const part = deck.slideParts[1].part;
  const gridWidths = () => [...deck.pkg.text(part).matchAll(/<a:gridCol w="(\d+)"\/>/g)].map((m) => Number(m[1]));
  const frameW = () => Number(/<p:xfrm><a:off[^/]*\/><a:ext cx="(\d+)" cy="\d+"\/>/.exec(deck.pkg.text(part))[1]);
  const cellCount = () => (deck.pkg.text(part).match(/<a:tc>/g) || []).length;
  const w0 = frameW();

  deck.insertTableColumn(1, id, 1);
  assert.equal(gridWidths().length, 3);
  assert.equal(frameW(), w0, 'the frame\'s own width does not change on a column insert');
  assert.ok(Math.abs(gridWidths().reduce((a, b) => a + b, 0) - w0) <= 3, 'the columns still sum to the frame width');
  assert.equal(cellCount(), 6, 'a cell added to every row');

  deck.removeTableColumn(1, id, 0);
  assert.equal(gridWidths().length, 2);
  assert.equal(cellCount(), 4);

  const reopened = Deck.open(deck.save());
  const shape = reopened.slide(1).shapes.find((s) => s.id === String(id));
  assert.equal(shape.table.columns.length, 2, 'survives Deck.open(deck.save())');
  assert.equal(shape.table.rows[0].cells.length, 2);

  assert.throws(() => deck.removeTableColumn(1, id, 50), /no column/);
});

test('adding and editing a table leaves the other slide byte-identical', () => {
  const before = OoxmlPackage.read(DECK);
  const deck = Deck.open(DECK);
  const id = deck.addTable(1, { rows: 2, cols: 2 });
  deck.setTableCell(1, id, 0, 0, [{ runs: [{ text: 'Hi' }] }]);
  const after = OoxmlPackage.read(deck.save());
  assert.deepEqual(after.read('ppt/slides/slide1.xml'), before.read('ppt/slides/slide1.xml'), 'slide 1 was not touched');
  assert.notDeepEqual(after.read('ppt/slides/slide2.xml'), before.read('ppt/slides/slide2.xml'));
});

test('a bad slide or shape throws', () => {
  const deck = Deck.open(DECK);
  const id = deck.addTable(1, { rows: 2, cols: 2 });
  assert.throws(() => deck.addTable(9), /no slide/);
  assert.throws(() => deck.setTableCell(9, id, 0, 0, []), /no slide/);
  assert.throws(() => deck.setTableCell(1, 999999, 0, 0, []), /not found/);
  assert.throws(() => deck.insertTableRow(1, 999999, 0), /not found/);
  assert.throws(() => deck.removeTableRow(1, 999999, 0), /not found/);
  assert.throws(() => deck.insertTableColumn(1, 999999, 0), /not found/);
  assert.throws(() => deck.removeTableColumn(1, 999999, 0), /not found/);
});

const showcaseFile = join(RICH, 'showcase.pptx');
test('the showcase deck\'s table — written by PowerPoint itself — still reads its cells, and each now carries a box', { skip: !existsSync(showcaseFile) && 'fixture not generated' }, () => {
  const deck = Deck.open(readFileSync(showcaseFile));
  const shape = deck.slide(3).shapes.find((s) => s.kind === 'table');
  assert.ok(shape, 'the table shape on slide 4');
  const words = wordsOf(shape);
  assert.deepEqual(words[0], ['Region', 'Q1', 'Q2', 'Total']);
  assert.deepEqual(words[1], ['North', '1,200', '1,350', '2,550']);
  assert.deepEqual(words[4], ['All', '3,610', '3,980', '7,590']);
  for (const row of shape.table.rows) for (const cell of row.cells) assert.ok(cell.box && cell.box.w > 0 && cell.box.h > 0, 'a box for every cell');
});
