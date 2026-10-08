// Presentations: Table Design — which of its style's parts a table takes
// (a:tblPr's firstRow, lastRow, firstCol, lastCol, bandRow, bandCol) and its
// cells' own shading, written where PowerPoint writes them and drawn.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const tableOf = (deck) => deck.slide(0).shapes.find((s) => s.kind === 'table');

test('Table Style Options: the switches written on a:tblPr, the style and the direction kept, and read back as the table\'s flags', () => {
  const deck = Deck.open(buildPptx({ slides: [{ layout: 'blank' }] }));
  deck.addTable(0, { rows: 3, cols: 3, x: 40, y: 40, w: 400, h: 120 });
  const shape = tableOf(deck);
  assert.deepEqual(shape.table.flags, { firstRow: true, lastRow: false, firstCol: false, lastCol: false, bandRow: true, bandCol: false }, 'PowerPoint\'s for a new table');
  deck.setTableLook(0, shape.id, { firstRow: false, bandRow: false, bandCol: true, lastCol: true });
  const xml = deck.pkg.text('ppt/slides/slide1.xml');
  assert.match(xml, /<a:tblPr lastCol="1" bandCol="1"><a:tableStyleId>\{[0-9A-F-]+\}<\/a:tableStyleId><\/a:tblPr>/);
  const after = tableOf(Deck.open(deck.pkg.write())).table;
  assert.deepEqual(after.flags, { firstRow: false, lastRow: false, firstCol: false, lastCol: true, bandRow: false, bandCol: true });
  assert.notEqual(after.rows[0].cells[0].fill?.color, shape.table.rows[0].cells[0].fill?.color, 'no header look on the first row now');
});

test('Shading: one cell or every cell in a colour of its own, or none so the style shows, in a:tcPr after its lines', () => {
  const deck = Deck.open(buildPptx({ slides: [{ layout: 'blank' }] }));
  deck.addTable(0, { rows: 2, cols: 2, x: 40, y: 40, w: 400, h: 120 });
  const id = tableOf(deck).id;
  // A cell with a line of its own keeps it, the fill after it.
  const part = 'ppt/slides/slide1.xml';
  deck.pkg.write_(part, deck.pkg.text(part).replace(/<a:tcPr\/>|<a:tcPr>[\s\S]*?<\/a:tcPr>/, () => '<a:tcPr anchor="ctr"><a:lnB w="25400"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:lnB></a:tcPr>'));
  deck.setTableCellFill(0, id, [{ row: 0, col: 0 }], '#FFC000');
  assert.match(deck.pkg.text(part), /<a:tcPr anchor="ctr"><a:lnB w="25400"><a:solidFill><a:srgbClr val="00FF00"\/><\/a:solidFill><\/a:lnB><a:solidFill><a:srgbClr val="FFC000"\/><\/a:solidFill><\/a:tcPr>/);
  let table = tableOf(Deck.open(deck.pkg.write())).table;
  assert.equal(table.rows[0].cells[0].fill.color.toLowerCase(), '#ffc000');
  assert.notEqual(table.rows[1].cells[1].fill?.color?.toLowerCase(), '#ffc000', 'only that cell');
  deck.setTableCellFill(0, id, null, '0070C0');
  table = tableOf(Deck.open(deck.pkg.write())).table;
  assert.deepEqual(table.rows.flatMap((r) => r.cells.map((c) => c.fill.color.toLowerCase())), ['#0070c0', '#0070c0', '#0070c0', '#0070c0'], 'every cell');
  deck.setTableCellFill(0, id, [{ row: 1, col: 1 }], null);
  table = tableOf(Deck.open(deck.pkg.write())).table;
  assert.notEqual(table.rows[1].cells[1].fill?.color?.toLowerCase(), '#0070c0', 'its own fill gone, the style\'s showing');
  assert.throws(() => deck.setTableCellFill(0, id, null, 'blue'), /colour like/);
});

test('through the document service, the options and the shading are operations of a deck', () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-deck-table-')), 'table.pptx');
  const deck = Deck.open(buildPptx({ slides: [{ layout: 'blank' }] }));
  deck.addTable(0, { rows: 2, cols: 2, x: 40, y: 40, w: 400, h: 120 });
  fs.writeFileSync(file, deck.pkg.write());
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const s = docs.open({ path: file });
  const shapeOf = () => docs.model({ id: s.id }).slide.shapes.find((x) => x.kind === 'table');
  const id = shapeOf().id;
  docs.apply({ id: s.id, ops: [{ op: 'setTableLook', slide: 0, shape: id, flags: { firstRow: true, lastRow: true } }] });
  assert.equal(shapeOf().table.flags.lastRow, true);
  docs.apply({ id: s.id, ops: [{ op: 'setTableCellFill', slide: 0, shape: id, cells: [{ row: 1, col: 0 }], fill: 'FF0000' }] });
  assert.equal(shapeOf().table.cells[1][0].fill.toLowerCase(), '#ff0000');
});

test('Table Style Options: an option changed on its own leaves the others as the table has them', () => {
  const deck = Deck.open(buildPptx({ slides: [{ layout: 'blank' }] }));
  deck.addTable(0, { rows: 3, cols: 3, x: 40, y: 40, w: 400, h: 120 });
  const shape = tableOf(deck);
  deck.setTableLook(0, shape.id, { bandCol: true });
  deck.setTableLook(0, shape.id, { firstRow: false });
  assert.deepEqual(tableOf(Deck.open(deck.pkg.write())).table.flags, { firstRow: false, lastRow: false, firstCol: false, lastCol: false, bandRow: true, bandCol: true });
});
