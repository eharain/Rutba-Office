// A table cell's look as Word draws it, on the page and on paper alike: its
// own lines and shading, its table style's conditional parts by where it
// stands (as tblLook turns them on, in Word's order), the table's sides.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { cellLook, partsAt, tableRuled } from '@rutba/doc-view/table-look';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { renderPdf } from '@rutba/doc-view/export/pdf';

const line = (colour, style = 'single') => ({ style, widthPx: 1, colour });
const GRID4 = {
  fill: null, rowBand: 1, colBand: 1,
  look: { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true },
  parts: {
    firstRow: { fill: '#156082', bold: true, colour: '#ffffff', borders: { top: line('#156082'), left: line('#156082'), bottom: line('#156082'), right: line('#156082'), insideH: { style: 'nil', widthPx: 0, colour: null }, insideV: { style: 'nil', widthPx: 0, colour: null } } },
    lastRow: { bold: true, borders: { top: line('#156082', 'double') } },
    firstCol: { bold: true },
    band1Vert: { fill: '#c1e4f5' },
    band1Horz: { fill: '#c1e4f5' },
  },
};
const TABLE = { top: line('#45b0e1'), left: line('#45b0e1'), bottom: line('#45b0e1'), right: line('#45b0e1'), insideH: line('#45b0e1'), insideV: line('#45b0e1') };
const at = (row, column, extra = {}) => ({ row, column, span: 1, rowSpan: 1, rows: 4, columns: 3, ...extra });

test('the parts that reach a cell, in Word\'s order: bands, then columns, then the header row', () => {
  assert.deepEqual(partsAt(GRID4, at(0, 0)).map((p) => p.fill ?? (p.bold ? 'bold' : '')), ['bold', '#156082'], 'the corner: first column, then header row');
  assert.deepEqual(partsAt(GRID4, at(1, 1)).map((p) => p.fill), ['#c1e4f5'], 'the first data row is the first band');
  assert.deepEqual(partsAt(GRID4, at(2, 1)).length, 0, 'the second band has no part in this style');
  assert.deepEqual(partsAt({ ...GRID4, look: { ...GRID4.look, noVBand: false } }, at(2, 1)).map((p) => p.fill), ['#c1e4f5'], 'banded columns when tblLook turns them on');
  assert.deepEqual(partsAt({ ...GRID4, look: { ...GRID4.look, firstRow: false } }, at(0, 1)).map((p) => p.fill), ['#c1e4f5'], 'no header row: the first row is a band');
  assert.deepEqual(partsAt({ ...GRID4, look: { ...GRID4.look, lastRow: true } }, at(3, 1)).map((p) => p.bold), [true], 'a total row is not banded');
});

test('a cell\'s sides, shading and words: a header row without its inside lines, a total row\'s double top, the cell\'s own over all', () => {
  const header = cellLook({ style: GRID4, borders: TABLE, at: at(0, 1) });
  assert.deepEqual([header.sides.top.colour, header.sides.bottom.colour, header.sides.left.style, header.sides.right.style], ['#156082', '#156082', 'nil', 'nil']);
  assert.deepEqual([header.fill, header.text.bold, header.text.colour], ['#156082', true, '#ffffff']);
  const corner = cellLook({ style: GRID4, borders: TABLE, at: at(0, 0) });
  assert.equal(corner.sides.left.colour, '#156082', 'at the table\'s edge the header row\'s own side');
  const band = cellLook({ style: GRID4, borders: TABLE, at: at(1, 1) });
  assert.deepEqual([band.fill, band.text.bold, band.sides.top.colour], ['#c1e4f5', undefined, '#45b0e1']);
  const label = cellLook({ style: GRID4, borders: TABLE, at: at(2, 0) });
  assert.deepEqual([label.fill, label.text.bold], [null, true], 'the first column bold, the second band plain');
  const total = cellLook({ style: { ...GRID4, look: { ...GRID4.look, lastRow: true } }, borders: TABLE, at: at(3, 2) });
  assert.equal(total.sides.top.style, 'double');
  const own = cellLook({ style: GRID4, borders: TABLE, own: { borders: { top: line('#ff0000') }, fill: '#ffff00' }, at: at(0, 1) });
  assert.deepEqual([own.sides.top.colour, own.fill], ['#ff0000', '#ffff00']);
  assert.equal(cellLook({ borders: null, at: at(1, 1), ruled: false }).sides, null, 'a table with no lines draws none');
  assert.equal(tableRuled({ borders: null, style: GRID4 }), true, 'a style\'s part with lines rules the table');
  assert.equal(tableRuled({ borders: null, style: { parts: { band1Horz: { fill: '#eeeeee' } } } }), false);
});

test('printed, a table in Word\'s Grid Table 4 has its header row filled and its rows banded, as the page shows it', () => {
  const file = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures', 'rich', 'showcase.docx');
  const { buffer } = renderPdf(openDocx(fs.readFileSync(file)), { created: '2026-09-03T00:00:00Z' });
  const text = buffer.toString('latin1');
  const count = (s) => text.split(s).length - 1;
  assert.equal(count('0.082 0.376 0.51 rg'), 4, 'the header row\'s four cells in accent blue');
  assert.equal(count('0.757 0.894 0.961 rg'), 8, 'the two banded rows\' cells in the light tint');
});
