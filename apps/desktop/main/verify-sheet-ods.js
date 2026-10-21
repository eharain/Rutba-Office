// Worksheets: a workbook saved as .ods opens again looking as it did (owner,
// 2026-10-21: "the excel formatted contents or styled text does not reload
// same after save, close reopen"). Cells are given words in bold, italic,
// underline and strike, a face, a size and a colour, fills, borders,
// alignment and wrapping, two cells are merged, and the workbook is saved as
// .ods; opened again, every cell's look and the merge are as they were.

import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';

/** What a drawn cell looks like, as the check compares it. */
function lookOf(cell) {
  if (!cell) return null;
  const style = cell.style || {};
  const f = style.font || {};
  const edges = Object.fromEntries(['top', 'right', 'bottom', 'left'].filter((s) => style.border?.[s]).map((s) => [s, `${style.border[s].style}/${style.border[s].widthPx}/${String(style.border[s].colour || '').toUpperCase()}`]));
  return {
    bold: Boolean(f.bold), italic: Boolean(f.italic), underline: Boolean(f.underline), strike: Boolean(f.strike),
    font: f.family || null, size: f.sizePt || null, colour: String(f.colour || '').toUpperCase() || null,
    fill: String(style.fill?.colour || '').toUpperCase() || null, edges,
    align: cell.align || null, valign: cell.valign || null, wrap: Boolean(cell.wrap),
  };
}

export async function verifySheetOds(h, { dir }) {
  const { check, doc } = h;
  const xlsx = path.join(dir, 'looks.xlsx');
  const ods = path.join(dir, 'looks.ods');
  try {
    fs.writeFileSync(xlsx, buildXlsx({ sheets: [{ name: 'Looks', rows: [['Name', 'Amount', 'Note'], ['Amira', 1234.5, 'paid in full'], ['Ben', 99, 'late'], ['Total', 1333.5, '']] }] }));
    const opened = await doc.open({ path: xlsx });
    const looks = [
      [0, 0, { bold: true, italic: true, fontName: 'Georgia', fontSize: 14, fontColour: '#C00000' }],
      [0, 1, { underline: true, strike: true, fill: '#FFFF00', align: 'center', valign: 'center' }],
      [1, 2, { wrap: true, align: 'right', border: { top: { style: 'thin', colour: '#000000' }, bottom: { style: 'double', colour: '#FF0000' } } }],
      [3, 0, { bold: true, fill: '#92D050', border: { left: { style: 'medium', colour: '#0070C0' } } }],
      [5, 4, { fill: '#00B0F0' }],
    ];
    const ops = looks.flatMap(([row, col, delta]) => [{ op: 'select', row, col }, { op: 'setFormat', delta }]);
    await doc.apply({ id: opened.id, ops: [...ops, { op: 'select', row: 2, col: 2 }, { op: 'select', row: 3, col: 2, extend: true }, { op: 'merge' }, { op: 'select', row: 0, col: 0 }] });
    const cellsOf = (id) => doc.model({ id }).cells || [];
    const at = (cells, r, c) => cells.find((x) => x.row === r && x.col === c);
    const beforeCells = cellsOf(opened.id);
    const before = looks.map(([r, c]) => lookOf(at(beforeCells, r, c)));
    await doc.save({ id: opened.id, path: ods });
    doc.close({ id: opened.id });

    const again = await doc.open({ path: ods });
    const afterCells = cellsOf(again.id);
    const after = looks.map(([r, c]) => lookOf(at(afterCells, r, c)));
    const differ = looks.map(([r, c], i) => (JSON.stringify(before[i]) === JSON.stringify(after[i]) ? null : { cell: `${r},${c}`, before: before[i], after: after[i] })).filter(Boolean);
    check('sheets: a workbook saved as .ods opens again with every cell\'s words, fill, borders and alignment as they were',
      before.every(Boolean) && differ.length === 0, differ.length ? JSON.stringify(differ).slice(0, 1200) : `${looks.length} cells alike`);
    const merged = at(afterCells, 2, 2)?.merged?.ref || null;
    check('sheets: a merge saved in an .ods is a merge when it opens again', merged === 'C3:C4', String(merged));
    doc.close({ id: again.id });
  } catch (err) {
    check('sheets: the .ods round trip ran', false, err.message);
  }
}
