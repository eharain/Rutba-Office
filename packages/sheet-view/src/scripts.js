// Automate → scripts: a workbook a script can read and change, in the shape
// of Excel's Office Scripts (`function main(workbook) { … }`), worked out
// on a snapshot of the workbook and answered as a list of edits.
//
// A script runs where it can touch nothing but the workbook it is handed —
// in a worker, with no file, no network and no window (scripts-worker.js) —
// and what it does is recorded, not done: each value, formula and format it
// sets, each sheet it adds or renames. The window then applies the edits to
// the real workbook as one undo step. What a script reads is the workbook
// as it was when the script started, with the script's own writes over it;
// a formula it writes is kept as a formula, its result worked out once the
// edits are applied.

const colName = (n) => { let s = ''; for (n += 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const colIndex = (s) => [...String(s).toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;

/** "B2", "A1:C3", "Sheet2!A1", "$A$1:$B$2", "A:A", "2:2" → { sheet, top, left, bottom, right }. */
export function parseAddress(address, defaultSheet) {
  let text = String(address ?? '').trim();
  let sheet = defaultSheet;
  const bang = text.lastIndexOf('!');
  if (bang >= 0) { sheet = text.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'"); text = text.slice(bang + 1); }
  text = text.replace(/\$/g, '');
  const cell = /^([A-Za-z]{1,3})(\d+)$/;
  const parts = text.split(':');
  const one = (p) => {
    const m = cell.exec(p);
    if (m) return { row: Number(m[2]) - 1, col: colIndex(m[1]) };
    if (/^[A-Za-z]{1,3}$/.test(p)) return { row: null, col: colIndex(p) };
    if (/^\d+$/.test(p)) return { row: Number(p) - 1, col: null };
    throw new Error(`"${address}" is not a cell or a range`);
  };
  const a = one(parts[0]);
  const b = parts[1] ? one(parts[1]) : a;
  const top = Math.min(a.row ?? 0, b.row ?? 0);
  const bottom = a.row === null || b.row === null ? 1048575 : Math.max(a.row, b.row);
  const left = Math.min(a.col ?? 0, b.col ?? 0);
  const right = a.col === null || b.col === null ? 16383 : Math.max(a.col, b.col);
  return { sheet, top, left, bottom, right };
}
export const addressOf = (r) => `${colName(r.left)}${r.top + 1}${r.bottom !== r.top || r.right !== r.left ? `:${colName(r.right)}${r.bottom + 1}` : ''}`;

const LIMIT = 100000; // the cells one call may set: a script is not a way to fill a million rows by accident

/**
 * A workbook for a script, over `snapshot` — `{ activeSheet, selection,
 * sheets: [{ name, cells: { 'r,c': { v, f } } }] }` — and the edits it makes.
 * Answers `{ workbook, edits }`.
 */
export function scriptWorkbook(snapshot) {
  const edits = [];
  const sheets = new Map((snapshot.sheets || []).map((s) => [s.name, { name: s.name, cells: new Map(Object.entries(s.cells || {})) }]));
  let active = snapshot.activeSheet || [...sheets.keys()][0];
  const sheetOf = (name) => {
    const s = sheets.get(name);
    if (!s) throw new Error(`There is no worksheet called "${name}"`);
    return s;
  };
  const used = (s) => {
    let top = Infinity, left = Infinity, bottom = -1, right = -1;
    for (const [key, c] of s.cells) {
      if (c.v === null && !c.f) continue;
      const [r, col] = key.split(',').map(Number);
      top = Math.min(top, r); left = Math.min(left, col); bottom = Math.max(bottom, r); right = Math.max(right, col);
    }
    return bottom < 0 ? null : { top, left, bottom, right };
  };
  const value = (s, r, c) => { const x = s.cells.get(`${r},${c}`); return x ? x.v ?? '' : ''; };
  const formula = (s, r, c) => { const x = s.cells.get(`${r},${c}`); return x ? x.f || (x.v ?? '') : ''; };
  const write = (s, r, c, input) => {
    const isFormula = typeof input === 'string' && input.startsWith('=');
    s.cells.set(`${r},${c}`, isFormula ? { v: null, f: input } : { v: input === '' ? null : input, f: null });
    edits.push({ kind: isFormula ? 'formula' : 'value', sheet: s.name, row: r, col: c, value: input });
  };
  const size = (box) => (box.bottom - box.top + 1) * (box.right - box.left + 1);
  const clampUsed = (s, box) => {
    // A whole column or row is read as far as the sheet is used.
    if (box.bottom < 1048575 && box.right < 16383) return box;
    const u = used(s) || { top: 0, left: 0, bottom: 0, right: 0 };
    return { ...box, bottom: Math.min(box.bottom, u.bottom), right: Math.min(box.right, u.right) };
  };

  function makeRange(s, box) {
    const rows = () => clampUsed(s, box).bottom - box.top + 1;
    const cols = () => clampUsed(s, box).right - box.left + 1;
    const grid = (fn) => { const b = clampUsed(s, box); const out = []; for (let r = b.top; r <= b.bottom; r++) { const row = []; for (let c = b.left; c <= b.right; c++) row.push(fn(r, c)); out.push(row); } return out; };
    const format = (patch) => edits.push({ kind: 'format', sheet: s.name, ...box, format: patch });
    const range = {
      getAddress: () => `${s.name}!${addressOf(box)}`,
      getRowCount: rows,
      getColumnCount: cols,
      getRowIndex: () => box.top,
      getColumnIndex: () => box.left,
      getWorksheet: () => makeSheet(s.name),
      getCell: (r, c) => makeRange(s, { top: box.top + r, left: box.left + c, bottom: box.top + r, right: box.left + c }),
      getOffsetRange: (dr, dc) => makeRange(s, { top: box.top + dr, left: box.left + dc, bottom: box.bottom + dr, right: box.right + dc }),
      getResizedRange: (dr, dc) => makeRange(s, { ...box, bottom: box.bottom + dr, right: box.right + dc }),
      getValue: () => value(s, box.top, box.left),
      getValues: () => grid((r, c) => value(s, r, c)),
      getFormula: () => formula(s, box.top, box.left),
      getFormulas: () => grid((r, c) => formula(s, r, c)),
      getText: () => String(value(s, box.top, box.left)),
      setValue: (v) => { if (size(box) > LIMIT) throw new Error('That is more cells than a script may set at once'); for (let r = box.top; r <= box.bottom; r++) for (let c = box.left; c <= box.right; c++) write(s, r, c, v ?? ''); },
      setValues: (m) => {
        if (!Array.isArray(m) || !m.every(Array.isArray)) throw new Error('setValues takes rows of values');
        if (m.length !== box.bottom - box.top + 1 || m.some((row) => row.length !== box.right - box.left + 1)) throw new Error(`setValues needs ${box.bottom - box.top + 1} rows of ${box.right - box.left + 1} values for ${addressOf(box)}`);
        m.forEach((row, i) => row.forEach((v, j) => write(s, box.top + i, box.left + j, v ?? '')));
      },
      setFormula: (f) => range.setValue(String(f).startsWith('=') ? f : `=${f}`),
      setFormulas: (m) => range.setValues(m.map((row) => row.map((f) => (String(f).startsWith('=') ? f : `=${f}`)))),
      // Walks the cells that exist, not the box: a whole-sheet range is
      // seventeen billion addresses, and the script ran out its time on them.
      clear: () => { for (const key of [...s.cells.keys()]) { const [r, c] = key.split(',').map(Number); if (r >= box.top && r <= box.bottom && c >= box.left && c <= box.right) s.cells.delete(key); } edits.push({ kind: 'clear', sheet: s.name, ...box }); },
      setNumberFormat: (code) => format({ numberFormat: String(code) }),
      setNumberFormatLocal: (code) => format({ numberFormat: String(code) }),
      select: () => { edits.push({ kind: 'select', sheet: s.name, ...box }); },
      // merge(across) merges the range, or each of its rows when `across`; unmerge() undoes any in it.
      merge: (across = false) => {
        if (across) for (let r = box.top; r <= box.bottom; r++) edits.push({ kind: 'merge', sheet: s.name, ...box, top: r, bottom: r });
        else edits.push({ kind: 'merge', sheet: s.name, ...box });
        for (let r = box.top; r <= box.bottom; r++) for (let c = box.left; c <= box.right; c++) if ((across ? c !== box.left : r !== box.top || c !== box.left)) s.cells.delete(`${r},${c}`);
      },
      unmerge: () => { edits.push({ kind: 'unmerge', sheet: s.name, ...box }); },
      getFormat: () => ({
        getFill: () => ({ setColor: (c) => format({ fill: String(c) }), clear: () => format({ fill: null }) }),
        getFont: () => ({
          setBold: (b) => format({ bold: Boolean(b) }),
          setItalic: (b) => format({ italic: Boolean(b) }),
          setColor: (c) => format({ color: String(c) }),
          setSize: (n) => format({ fontSize: Number(n) }),
          setName: (n) => format({ fontFamily: String(n) }),
          setUnderline: (u) => format({ underline: u && u !== 'None' }),
          setStrikethrough: (b) => format({ strike: Boolean(b) }),
        }),
        setHorizontalAlignment: (a) => format({ align: String(a).toLowerCase().replace('centeracrossselection', 'center') }),
        setVerticalAlignment: (a) => format({ valign: String(a).toLowerCase() === 'justify' ? 'justify' : String(a).toLowerCase() }),
        setWrapText: (b) => format({ wrap: Boolean(b) }),
        setColumnWidth: (w) => edits.push({ kind: 'colWidth', sheet: s.name, ...box, width: Number(w) }),
        setRowHeight: (h) => edits.push({ kind: 'rowHeight', sheet: s.name, ...box, height: Number(h) }),
        autofitColumns: () => edits.push({ kind: 'autofit', sheet: s.name, ...box }),
      }),
    };
    return Object.freeze(range);
  }

  function makeSheet(name) {
    const sheet = {
      getName: () => name,
      getRange: (address = 'A1') => { const box = parseAddress(address, name); return makeRange(sheetOf(box.sheet), box); },
      getCell: (r, c) => makeRange(sheetOf(name), { top: r, left: c, bottom: r, right: c }),
      getUsedRange: () => { const u = used(sheetOf(name)); return u ? makeRange(sheetOf(name), u) : undefined; },
      activate: () => { active = name; edits.push({ kind: 'activate', sheet: name }); },
      setName: (to) => {
        const next = String(to).trim();
        if (!next || sheets.has(next)) throw new Error(`"${to}" cannot be a worksheet's name here`);
        const s = sheetOf(name);
        sheets.delete(name); s.name = next; sheets.set(next, s);
        edits.push({ kind: 'rename', sheet: name, to: next });
        name = next;
      },
    };
    return Object.freeze(sheet);
  }

  const workbook = Object.freeze({
    getActiveWorksheet: () => makeSheet(active),
    getWorksheet: (name) => (sheets.has(name) ? makeSheet(name) : undefined),
    getWorksheets: () => [...sheets.keys()].map(makeSheet),
    addWorksheet: (name) => {
      let n = name ? String(name).trim() : '';
      if (!n) { let k = sheets.size + 1; while (sheets.has(`Sheet${k}`)) k += 1; n = `Sheet${k}`; }
      if (sheets.has(n)) throw new Error(`There is already a worksheet called "${n}"`);
      sheets.set(n, { name: n, cells: new Map() });
      edits.push({ kind: 'addSheet', sheet: n });
      return makeSheet(n);
    },
    getSelectedRange: () => { const sel = snapshot.selection || { sheet: active, top: 0, left: 0, bottom: 0, right: 0 }; return makeRange(sheetOf(sel.sheet || active), sel); },
    getName: () => snapshot.name || 'Workbook',
  });
  return { workbook, edits };
}

/**
 * A script run on a snapshot: `{ edits, logs, error }`. `code` is the
 * script's text, which declares `function main(workbook)`.
 */
export function runScript(code, snapshot, { timeLimitMs = 10000 } = {}) {
  const logs = [];
  const say = (...args) => logs.push(args.map((a) => (typeof a === 'string' ? a : (() => { try { return JSON.stringify(a); } catch { return String(a); } })())).join(' '));
  const console = Object.freeze({ log: say, info: say, warn: say, error: say });
  const { workbook, edits } = scriptWorkbook(snapshot);
  const started = Date.now();
  try {
    // eslint-disable-next-line no-new-func
    const fn = new Function('workbook', 'console', `"use strict";\n${code}\n;if (typeof main !== 'function') throw new Error('A script needs a function called main(workbook)');\nreturn main(workbook);`);
    const out = fn(workbook, console);
    if (out && typeof out.then === 'function') throw new Error('main must not be async here: it runs on a snapshot, all at once');
    if (Date.now() - started > timeLimitMs) throw new Error('The script ran too long');
    return { edits, logs, error: null };
  } catch (err) {
    return { edits: [], logs, error: String(err?.message || err) };
  }
}
