// DIF (.dif), VisiCalc's Data Interchange Format: a header of
// TABLE, VECTORS, TUPLES and DATA entries, then the values a tuple (a row)
// at a time, each as two lines: "type,number" and a string. Values only, no
// formulas or looks. Read into the old spreadsheet readers' sheets
// (legacy-grid.js).

import { decode1252 } from './codepage.js';
import { Grid } from './legacy-grid.js';

/** Whether text looks like DIF: it opens with TABLE, 0,1. */
export function isDif(bytes) {
  const head = decode1252(bytes.subarray(0, 32)).replace(/\r/g, '');
  return /^TABLE\n0,\d+\n/.test(head);
}

/** Read a DIF file: the old spreadsheet readers' model, one sheet named by its title. */
export function readDif(bytes, { name = 'Sheet1' } = {}) {
  const lines = decode1252(bytes).split(/\r\n|\n|\r/);
  let i = 0;
  const next = () => lines[i++] ?? '';
  const unquote = (s) => (s.startsWith('"') && s.endsWith('"') && s.length > 1 ? s.slice(1, -1).replace(/""/g, '"') : s);
  let title = '';
  // The header: an entry is a topic, then "vector,value", then a string.
  for (;;) {
    if (i >= lines.length) break;
    const topic = next().trim().toUpperCase();
    next();
    const str = next();
    if (topic === 'TABLE') title = unquote(str);
    if (topic === 'DATA') break;
  }
  const grid = new Grid(title.trim().slice(0, 31) || name);
  let row = -1;
  let col = 0;
  while (i < lines.length) {
    const [type, num] = next().split(',');
    const str = next();
    const t = Number(type);
    if (t === -1) {
      const word = str.trim().toUpperCase();
      if (word === 'BOT') { row += 1; col = 0; continue; }
      if (word === 'EOD') break;
      continue;
    }
    if (row < 0) row = 0;
    if (t === 0) {
      const ind = str.trim().toUpperCase();
      if (ind === 'V') grid.set(row, col, Number(num));
      else if (ind === 'TRUE') grid.set(row, col, true);
      else if (ind === 'FALSE') grid.set(row, col, false);
      else if (ind === 'NA') grid.set(row, col, '#N/A');
      else if (ind === 'ERROR') grid.set(row, col, '#VALUE!');
    } else if (t === 1) {
      const s = unquote(str);
      if (s !== '') grid.set(row, col, s);
    }
    col += 1;
  }
  return { sheets: [grid.toSheet()], names: [], notes: [] };
}
