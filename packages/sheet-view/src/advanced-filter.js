/**
 * Data → Advanced: a list filtered by a criteria range, in place or copied
 * to another place, with Excel's semantics throughout.
 *
 * The criteria range is a header row of field names over rows of
 * conditions. Conditions on one row must all hold (AND); any one row
 * holding is enough (OR). A condition is a value — text matches a field
 * that BEGINS with it, a number one that equals it — or a comparison
 * (`>`, `>=`, `<`, `<=`, `=`, `<>`) with `*` and `?` as wildcards and `~`
 * to take one literally; `=` alone is a blank field and `<>` alone a
 * filled one. A formula under a heading that names no field is a
 * condition of its own, written for the list's first row and read for
 * every row the way a formula filled down reads. A blank criteria row
 * lets every row through.
 *
 * In place, the rows that fail are hidden and the sheet says it is
 * filtered (`sheetPr filterMode`); copied, the rows that pass are written
 * under the headings at the destination. Either way the workbook keeps
 * the names Excel keeps: `_xlnm._FilterDatabase` (the list, hidden),
 * `_xlnm.Criteria`, and `_xlnm.Extract` for a copy.
 */
import { calculate, shiftFormula, isError } from '@rutba/formula';
import { ref, colName } from './selection.js';

/** A1 text — with or without `$` and a sheet in front — as a range. */
export function parseRange(text) {
  const raw = String(text ?? '').trim().replace(/^'?.*?'?!/, '').replace(/\$/g, '').toUpperCase();
  const m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(raw);
  if (!m) return null;
  const col = (letters) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const a = { row: Number(m[2]) - 1, col: col(m[1]) };
  const b = m[3] ? { row: Number(m[4]) - 1, col: col(m[3]) } : a;
  return {
    top: Math.min(a.row, b.row), bottom: Math.max(a.row, b.row),
    left: Math.min(a.col, b.col), right: Math.max(a.col, b.col),
  };
}

const absRef = (sheet, r) => {
  const q = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : "'" + sheet.replace(/'/g, "''") + "'";
  const cell = (row, col) => '$' + colName(col) + '$' + (row + 1);
  return q + '!' + cell(r.top, r.left) + ':' + cell(r.bottom, r.right);
};

/** A wildcard pattern as a whole-value regex: `*` any run, `?` one, `~` the next literally. */
function wildcard(pattern, { prefix = false } = {}) {
  let out = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) { out += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); continue; }
    if (ch === '*') out += '[\\s\\S]*';
    else if (ch === '?') out += '[\\s\\S]';
    else out += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp('^' + out + (prefix ? '' : '$'), 'i');
}

/**
 * One criteria cell as a test of a value. `raw` is what the cell holds:
 * a number, a boolean, or text (a comparison, or a value to begin with).
 * Returns null for a blank cell — no condition.
 */
export function condition(raw) {
  if (raw === null || raw === undefined || raw === '') return null;
  if (typeof raw === 'number') return (v) => typeof v === 'number' && v === raw;
  if (typeof raw === 'boolean') return (v) => v === raw;
  const text = String(raw);
  const m = /^(>=|<=|<>|>|<|=)?([\s\S]*)$/.exec(text);
  const op = m[1] || null;
  const operand = m[2];
  const blank = (v) => v === null || v === undefined || v === '';
  if (op === '=' && operand === '') return blank;
  if (op === '<>' && operand === '') return (v) => !blank(v);
  const n = operand.trim() !== '' && Number.isFinite(Number(operand)) ? Number(operand) : null;
  const lower = (v) => String(v).toLowerCase();
  if (!op) {
    // A bare number matches that number; bare text, a field beginning with it.
    if (n !== null) return (v) => (typeof v === 'number' ? v === n : lower(v) === operand.toLowerCase());
    const re = wildcard(operand, { prefix: true });
    return (v) => !blank(v) && typeof v === 'string' && re.test(v);
  }
  if (op === '=' || op === '<>') {
    const re = wildcard(operand);
    const eq = n !== null
      ? (v) => (typeof v === 'number' ? v === n : !blank(v) && re.test(String(v)))
      : (v) => !blank(v) && typeof v !== 'number' && re.test(String(v));
    return op === '=' ? eq : (v) => !eq(v);
  }
  // An ordering compares numbers with numbers and text with text, as Excel's does.
  const cmp = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
  const holds = (c) => (op === '>' ? c > 0 : op === '>=' ? c >= 0 : op === '<' ? c < 0 : c <= 0);
  if (n !== null) return (v) => typeof v === 'number' && holds(cmp(v, n));
  return (v) => typeof v === 'string' && v !== '' && holds(cmp(v.toLowerCase(), operand.toLowerCase()));
}

/** The list's values, a row at a time: what a condition reads. */
const valueAt = (view, row, col) => {
  const v = view.calc.getValue(view.activeSheet, row, col);
  return isError(v) ? null : v;
};

/**
 * Read a criteria range against a list: per criteria row, the tests of
 * that row, each a field test (`col`, `test`) or a formula test.
 */
function readCriteria(view, list, criteria) {
  const sheet = view.activeSheet;
  const heads = new Map();
  for (let c = list.left; c <= list.right; c++) {
    const name = String(view.displayValue(list.top, c).text ?? '').trim().toLowerCase();
    if (name && !heads.has(name)) heads.set(name, c);
  }
  const rows = [];
  for (let r = criteria.top + 1; r <= criteria.bottom; r++) {
    const tests = [];
    for (let c = criteria.left; c <= criteria.right; c++) {
      const input = String(view.calc.getInput(sheet, r, c) ?? '');
      const head = String(view.displayValue(criteria.top, c).text ?? '').trim().toLowerCase();
      if (input.startsWith('=') && !heads.has(head)) {
        // A computed condition: written for the list's first data row.
        tests.push({ formula: input.slice(1), at: { row: r, col: c } });
        continue;
      }
      if (!head) continue;
      const col = heads.get(head);
      if (col === undefined) throw new Error('The criteria heading "' + String(view.displayValue(criteria.top, c).text).trim() + '" names no field of the list');
      const raw = input.startsWith('=') ? valueAt(view, r, c) : view.calc.getValue(sheet, r, c);
      const test = condition(raw);
      if (test) tests.push({ col, test });
    }
    rows.push(tests);
  }
  if (!rows.length) throw new Error('The criteria range needs its headings and at least one row of conditions under them');
  return rows;
}

/** The data rows of the list that pass, and in the order they sit. */
function matchingRows(view, list, criteriaRows) {
  const sheet = view.activeSheet;
  const first = list.top + 1;
  const resolver = view.calc.resolver();
  const out = [];
  for (let row = first; row <= list.bottom; row++) {
    const pass = criteriaRows.some((tests) => tests.every((t) => {
      if (t.formula) {
        const shifted = shiftFormula(t.formula, row - first, 0);
        const v = calculate('=' + shifted, resolver, { sheet, row: t.at.row, col: t.at.col });
        return v === true || (typeof v === 'number' && v !== 0);
      }
      return t.test(valueAt(view, row, t.col));
    }));
    if (pass) out.push(row);
  }
  return out;
}

/** Drop rows that repeat an earlier one across the given columns, letter case aside. */
function unique(view, rows, cols) {
  const seen = new Set();
  return rows.filter((row) => {
    const key = JSON.stringify(cols.map((c) => String(view.displayValue(row, c).text ?? '').trim().toLowerCase()));
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

const overlaps = (a, b) => a.left <= b.right && b.left <= a.right && a.top <= b.bottom && b.top <= a.bottom;

/**
 * Data → Advanced. `list`, `criteria` and `copyTo` are A1 ranges on the
 * active sheet (the list defaults to the block round the cell).
 *
 * @returns {{ action: string, matched: number, total: number, range: string }}
 */
export function advancedFilter(view, { list: listText, criteria: criteriaText, action = 'filter', copyTo = null, unique: onlyUnique = false } = {}) {
  const sheet = view.activeSheet;
  const active = view.selection.active;
  const list = listText ? parseRange(listText) : view._currentRegion(active.row, active.col);
  if (!list || list.bottom <= list.top) throw new Error('The list range needs a header row and at least one row under it');
  const criteria = parseRange(criteriaText);
  if (!criteria || criteria.bottom <= criteria.top) throw new Error('The criteria range needs its headings and at least one row of conditions — e.g. F1:G3');
  if (overlaps(list, criteria)) throw new Error('The criteria range cannot sit inside the list');
  const criteriaRows = readCriteria(view, list, criteria);
  let rows = matchingRows(view, list, criteriaRows);
  const total = list.bottom - list.top;
  const index = view.workbook.sheetNames().indexOf(sheet);
  const names = (extra) => {
    view.workbook.setDefinedName('_xlnm._FilterDatabase', absRef(sheet, list), { localSheetId: index, hidden: 1 });
    view.workbook.setDefinedName('_xlnm.Criteria', absRef(sheet, criteria), { localSheetId: index });
    if (extra) view.workbook.setDefinedName('_xlnm.Extract', absRef(sheet, extra), { localSheetId: index });
    view._syncNames();
  };

  if (action !== 'copy') {
    const cols = [];
    for (let c = list.left; c <= list.right; c++) cols.push(c);
    if (onlyUnique) rows = unique(view, rows, cols);
    const keep = new Set(rows);
    const partName = view.workbook.partNameFor(sheet);
    view._edit('advanced filter', null, [], () => {
      const { part } = view.workbook._sheetPart(sheet);
      const geo = view.geo;
      for (let row = list.top + 1; row <= list.bottom; row++) {
        const hide = !keep.has(row);
        if (hide !== geo.hiddenRows.has(row)) {
          part.setRowOutline(row, { hidden: hide });
          if (hide) geo.hiddenRows.add(row);
          else geo.hiddenRows.delete(row);
        }
      }
      part.setSheetPrAttr('filterMode', rows.length < total ? '1' : null);
      names(null);
      view._structuralDirty = true;
    }, { parts: [partName, view.workbook.mainPart] });
    return { action: 'filter', matched: rows.length, total, range: ref(list.top, list.left) + ':' + ref(list.bottom, list.right) };
  }

  // Copy: to the headings named at the destination, or every column.
  const to = parseRange(copyTo);
  if (!to) throw new Error('Say where to copy the rows — a cell such as I1');
  const heads = new Map();
  for (let c = list.left; c <= list.right; c++) {
    const name = String(view.displayValue(list.top, c).text ?? '').trim().toLowerCase();
    if (name && !heads.has(name)) heads.set(name, c);
  }
  const named = [];
  for (let c = to.left; c <= to.right; c++) {
    const col = heads.get(String(view.displayValue(to.top, c).text ?? '').trim().toLowerCase());
    if (col !== undefined) named.push(col);
  }
  // A row of headings picks the columns; one cell is where every column starts.
  const byName = to.right > to.left && named.length === to.right - to.left + 1;
  const cols = byName ? named : Array.from({ length: list.right - list.left + 1 }, (_, i) => list.left + i);
  if (onlyUnique) rows = unique(view, rows, cols);
  // What is written: the headings (unless they are there already) and the
  // rows; and, as Excel does, the columns below are cleared to the bottom of
  // what the sheet uses, so an earlier, longer extract leaves nothing behind.
  const width = cols.length;
  const bottom = Math.max(to.top + rows.length, view.bounds.maxRow);
  const area = { top: to.top, left: to.left, bottom, right: to.left + width - 1 };
  if (overlaps(area, list) || overlaps(area, criteria)) throw new Error('The rows copied would land on the list or the criteria — choose a place beside or below them');
  const cells = [];
  for (let r = area.top; r <= area.bottom; r++) for (let c = area.left; c <= area.right; c++) cells.push({ row: r, col: c });
  // Headings typed at the destination stay as typed; otherwise the list's go there.
  const header = byName
    ? cols.map((_, k) => view.calc.getInput(sheet, to.top, to.left + k) ?? '')
    : cols.map((c) => view.calc.getInput(sheet, list.top, c) ?? '');
  const lines = rows.map((row) => cols.map((c) => {
    const input = view.calc.getInput(sheet, row, c) ?? '';
    // A formula is copied as the value it shows, as Excel's extract is.
    if (String(input).startsWith('=')) {
      const v = view.calc.getValue(sheet, row, c);
      return isError(v) ? view.displayValue(row, c).text : v;
    }
    return input;
  }));
  view._edit('advanced filter', null, cells, () => {
    for (let r = area.top; r <= area.bottom; r++) {
      for (let k = 0; k < width; k++) {
        const value = r === area.top ? header[k] : lines[r - area.top - 1]?.[k] ?? '';
        const current = view.calc.getInput(sheet, r, area.left + k) ?? '';
        if (String(current) === String(value ?? '')) continue;
        view._setCell(r, area.left + k, value === '' || value === null || value === undefined ? null : value);
      }
    }
    names({ top: to.top, left: to.left, bottom: to.top, right: to.left + width - 1 });
    view.selection.collapseTo(to.top, to.left);
    view.selection.extendTo(to.top + rows.length, to.left + width - 1);
  }, { parts: [view.workbook.mainPart] });
  return { action: 'copy', matched: rows.length, total, range: ref(to.top, to.left) + ':' + ref(to.top + rows.length, to.left + width - 1) };
}

/**
 * Data → Clear: the rows an advanced filter hid come back, and the sheet
 * stops saying it is filtered. The names stay, as Excel keeps them, so
 * the dialog offers the same ranges next time.
 */
export function clearAdvancedFilter(view) {
  const sheet = view.activeSheet;
  const index = view.workbook.sheetNames().indexOf(sheet);
  const db = view.workbook.definedNames().find((n) => n.name === '_xlnm._FilterDatabase' && new RegExp('localSheetId="' + index + '"').test(n.attrsStr));
  const list = db ? parseRange(db.ref) : null;
  const geo = view.geo;
  let hidden = 0;
  if (list) for (let r = list.top + 1; r <= list.bottom; r++) if (geo.hiddenRows.has(r)) hidden += 1;
  if (!list || !hidden) throw new Error('No advanced filter is hiding rows on this sheet');
  const partName = view.workbook.partNameFor(sheet);
  view._edit('clear filter', null, [], () => {
    const { part } = view.workbook._sheetPart(sheet);
    for (let r = list.top + 1; r <= list.bottom; r++) {
      if (!geo.hiddenRows.has(r)) continue;
      part.setRowOutline(r, { hidden: false });
      geo.hiddenRows.delete(r);
    }
    part.setSheetPrAttr('filterMode', null);
    view._structuralDirty = true;
  }, { parts: [partName] });
  return { shown: hidden };
}

/** The ranges an earlier Advanced filter left in the names, for the dialog to offer again. */
export function filterNames(view) {
  const index = view.workbook.sheetNames().indexOf(view.activeSheet);
  const out = {};
  for (const n of view.workbook.definedNames()) {
    if (!new RegExp('localSheetId="' + index + '"').test(n.attrsStr)) continue;
    const range = parseRange(n.ref);
    if (!range) continue;
    const text = ref(range.top, range.left) + ':' + ref(range.bottom, range.right);
    if (n.name === '_xlnm._FilterDatabase') out.list = text;
    if (n.name === '_xlnm.Criteria') out.criteria = text;
    if (n.name === '_xlnm.Extract') out.extract = range.right > range.left ? ref(range.top, range.left) + ':' + ref(range.bottom, range.right) : ref(range.top, range.left);
  }
  return out;
}
