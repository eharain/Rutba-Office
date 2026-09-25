/**
 * Data → Consolidate: several ranges — on this sheet or others — summarised
 * into one at the active cell, by position or by their labels, with one of
 * Excel's eleven functions.
 *
 *   - By position (no labels): the first cell of every range goes into the
 *     first cell of the result, and so on; the result is as large as the
 *     largest range.
 *   - By label: with "Top row" the ranges' first rows name their columns,
 *     with "Left column" their first columns name their rows; like labels
 *     meet whatever their place, in the order they are first met, matched
 *     without regard to capitals, and the labels are written round the result.
 *   - "Create links to source data" writes formulas instead of values, as
 *     Excel does: for every row of the result, a detail row per range that
 *     has it, linking to its cells (the source's sheet named beside them),
 *     hidden one level down an outline, and the row itself summing them with
 *     the function — Excel's own layout, folded to level 1, so the outline's
 *     2 opens the detail. Excel refuses links into the sheet a source is on,
 *     and so does this.
 *
 * What was consolidated is kept where Excel keeps it: `<dataConsolidate>`
 * in the destination sheet, so the dialog opens on the same references.
 */
import { ref, colName } from './selection.js';
import { structuralEdit } from './outline.js';

/** The functions, by the name the dialog and `<dataConsolidate function>` use. */
export const CONSOLIDATE_FUNCTIONS = {
  sum: { label: 'Sum', formula: 'SUM', file: 'sum' },
  count: { label: 'Count', formula: 'COUNTA', file: 'count' },
  average: { label: 'Average', formula: 'AVERAGE', file: 'average' },
  max: { label: 'Max', formula: 'MAX', file: 'max' },
  min: { label: 'Min', formula: 'MIN', file: 'min' },
  product: { label: 'Product', formula: 'PRODUCT', file: 'product' },
  countNumbers: { label: 'Count Numbers', formula: 'COUNT', file: 'countNums' },
  stdDev: { label: 'StdDev', formula: 'STDEV', file: 'stdDev' },
  stdDevp: { label: 'StdDevp', formula: 'STDEVP', file: 'stdDevp' },
  var: { label: 'Var', formula: 'VAR', file: 'var' },
  varp: { label: 'Varp', formula: 'VARP', file: 'varp' },
};

/** The same arithmetic in the engine's own hands, for the values Consolidate writes. */
function aggregate(fn, values) {
  const nums = values.filter((v) => typeof v === 'number' && Number.isFinite(v));
  const sum = nums.reduce((a, b) => a + b, 0);
  const mean = nums.length ? sum / nums.length : 0;
  const ss = nums.reduce((a, b) => a + (b - mean) ** 2, 0);
  switch (fn) {
    case 'sum': return sum;
    case 'count': return values.filter((v) => v !== '' && v !== null && v !== undefined).length;
    case 'countNumbers': return nums.length;
    case 'average': return nums.length ? mean : null;
    case 'max': return nums.length ? Math.max(...nums) : 0;
    case 'min': return nums.length ? Math.min(...nums) : 0;
    case 'product': return nums.length ? nums.reduce((a, b) => a * b, 1) : 0;
    case 'stdDev': return nums.length > 1 ? Math.sqrt(ss / (nums.length - 1)) : null;
    case 'stdDevp': return nums.length ? Math.sqrt(ss / nums.length) : null;
    case 'var': return nums.length > 1 ? ss / (nums.length - 1) : null;
    case 'varp': return nums.length ? ss / nums.length : null;
    default: return null;
  }
}

/** A sheet name as a formula writes it: quoted when it has to be. */
export function quoteSheet(name) {
  return /[^A-Za-z0-9_]/.test(name) || /^\d/.test(name) ? "'" + String(name).replace(/'/g, "''") + "'" : name;
}

const colIndex = (letters) => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0) - 1;

/**
 * A reference as the dialog lists it — `Sheet1!$A$1:$C$5`, `'Q1 sales'!B2:D9`
 * or, on the active sheet, `A1:C5` — to its sheet and box. Refused in words
 * when it is not one.
 */
export function parseConsolidateRef(text, view) {
  const s = String(text ?? '').trim().replace(/^=/, '');
  const m = /^(?:'((?:[^']|'')+)'|([^'!]+))?!?\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(s.includes('!') ? s : '!' + s);
  if (!m) throw new Error('"' + text + '" is not a reference to a range, such as Sheet1!$A$1:$C$5.');
  const sheet = m[1] ? m[1].replace(/''/g, "'") : m[2] ? m[2] : view.activeSheet;
  if (!view.sheetNames().includes(sheet)) throw new Error('There is no sheet called "' + sheet + '".');
  const r1 = Number(m[4]) - 1;
  const c1 = colIndex(m[3]);
  const r2 = m[6] ? Number(m[6]) - 1 : r1;
  const c2 = m[5] ? colIndex(m[5]) : c1;
  return { sheet, top: Math.min(r1, r2), left: Math.min(c1, c2), bottom: Math.max(r1, r2), right: Math.max(c1, c2) };
}

/** The reference as Excel lists it: sheet, dollars. */
export function consolidateRefText(box) {
  const a = '$' + colName(box.left) + '$' + (box.top + 1);
  const b = '$' + colName(box.right) + '$' + (box.bottom + 1);
  return quoteSheet(box.sheet) + '!' + (a === b ? a : a + ':' + b);
}

const key = (v) => String(v ?? '').trim().toLowerCase();

/**
 * Consolidate at the active cell. `refs` are references as above; `fn` one of
 * CONSOLIDATE_FUNCTIONS; `topRow` / `leftCol` use those labels; `links`
 * writes formulas and the outline. One undo step. Answers the result's
 * reference and its size.
 */
export function consolidate(view, { fn = 'sum', refs = [], topRow = false, leftCol = false, links = false } = {}) {
  const f = CONSOLIDATE_FUNCTIONS[fn];
  if (!f) throw new Error('Consolidate has no function "' + fn + '".');
  const boxes = refs.map((r) => parseConsolidateRef(r, view));
  if (!boxes.length) throw new Error('Add at least one reference to consolidate.');
  const dest = view.activeSheet;
  const at = { row: view.selection.active.row, col: view.selection.active.col };
  if (links && boxes.some((b) => b.sheet === dest)) {
    throw new Error('Links cannot be created when the source and destination areas are on the same sheet.');
  }
  if (view.protection().sheet) throw new Error('This sheet is protected — unprotect it before consolidating into it.');

  // Every source read once: its labels and its cells, by what they are keyed on.
  const sources = boxes.map((b) => {
    const value = (r, c) => {
      const v = view.calc.getValue(b.sheet, r, c);
      return v && typeof v === 'object' && v.type ? '' : v;
    };
    const dataTop = b.top + (topRow ? 1 : 0);
    const dataLeft = b.left + (leftCol ? 1 : 0);
    const rows = [];
    for (let r = dataTop; r <= b.bottom; r++) rows.push({ key: leftCol ? key(value(r, b.left)) : String(r - dataTop), label: leftCol ? value(r, b.left) : null, at: r });
    const cols = [];
    for (let c = dataLeft; c <= b.right; c++) cols.push({ key: topRow ? key(value(b.top, c)) : String(c - dataLeft), label: topRow ? value(b.top, c) : null, at: c });
    return { box: b, rows: rows.filter((x) => !leftCol || x.key !== ''), cols: cols.filter((x) => !topRow || x.key !== ''), value };
  });
  // The result's rows and columns: every key, in the order first met.
  const order = (list) => {
    const seen = new Map();
    for (const item of list) if (!seen.has(item.key)) seen.set(item.key, item.label);
    return [...seen].map(([k, label]) => ({ key: k, label }));
  };
  const outRows = order(sources.flatMap((s) => s.rows));
  const outCols = order(sources.flatMap((s) => s.cols));
  if (!position(outRows, outCols)) throw new Error('The references hold nothing to consolidate.');

  // The result must not overlap a range it reads, on the same sheet (Excel's rule).
  const labelCol = leftCol ? 1 : 0;
  const nameCol = links ? 1 : 0;
  const width = labelCol + nameCol + outCols.length;
  const detailCount = links ? outRows.reduce((n, row) => n + sources.filter((s) => s.rows.some((x) => x.key === row.key)).length, 0) : 0;
  const height = (topRow ? 1 : 0) + outRows.length + detailCount;
  const box = { top: at.row, left: at.col, bottom: at.row + height - 1, right: at.col + width - 1 };
  for (const b of boxes) {
    if (b.sheet === dest && !(b.bottom < box.top || b.top > box.bottom || b.right < box.left || b.left > box.right)) {
      throw new Error('Source references overlap destination area.');
    }
  }

  const A1 = (r, c) => colName(c) + (r + 1);
  return structuralEdit(view, 'consolidate', () => {
    const wb = view.workbook;
    const { part } = wb._sheetPart(dest);
    const put = (r, c, v) => wb.setCell(dest, A1(r, c), v === null || v === undefined ? '' : v);
    // Whatever the result will cover is cleared first, the outline with it.
    for (let r = box.top; r <= box.bottom; r++) {
      for (let c = box.left; c <= box.right; c++) put(r, c, '');
      part.setRowOutline(r, { level: 0, hidden: false, collapsed: false });
    }
    let row = box.top;
    const firstData = box.left + labelCol + nameCol;
    if (topRow) {
      outCols.forEach((col, j) => put(row, firstData + j, col.label ?? ''));
      row += 1;
    }
    let deepest = 0;
    for (const out of outRows) {
      if (links) {
        // A detail row per source that has this row, linking to its cells.
        const first = row;
        for (const s of sources) {
          const src = s.rows.find((x) => x.key === out.key);
          if (!src) continue;
          if (nameCol) put(row, box.left + labelCol, s.box.sheet);
          outCols.forEach((col, j) => {
            const c = s.cols.find((x) => x.key === col.key);
            if (c) put(row, firstData + j, '=' + quoteSheet(s.box.sheet) + '!$' + colName(c.at) + '$' + (src.at + 1));
          });
          part.setRowOutline(row, { level: 1, hidden: true });
          row += 1;
        }
        deepest = 1;
        if (leftCol) put(row, box.left, out.label ?? '');
        outCols.forEach((col, j) => {
          if (row > first) put(row, firstData + j, '=' + f.formula + '(' + A1(first, firstData + j) + ':' + A1(row - 1, firstData + j) + ')');
        });
        part.setRowOutline(row, { collapsed: true });
        row += 1;
      } else {
        if (leftCol) put(row, box.left, out.label ?? '');
        outCols.forEach((col, j) => {
          const vals = [];
          for (const s of sources) {
            const src = s.rows.find((x) => x.key === out.key);
            const c = s.cols.find((x) => x.key === col.key);
            if (src && c) vals.push(s.value(src.at, c.at));
          }
          if (!vals.some((v) => v !== '' && v !== null && v !== undefined)) return;
          put(row, firstData + j, aggregate(fn, vals));
        });
        row += 1;
      }
    }
    if (deepest) part.setOutlineProps({ rowLevels: deepest, summaryBelow: true });
    // Kept as Excel keeps it, for the dialog to open on next time.
    part.setTailElement('dataConsolidate', '<dataConsolidate' + (fn !== 'sum' ? ' function="' + f.file + '"' : '')
      + (leftCol ? ' leftLabels="1"' : '') + (topRow ? ' topLabels="1"' : '') + (links ? ' link="1"' : '')
      + '><dataRefs count="' + boxes.length + '">'
      + boxes.map((b) => '<dataRef ref="' + ref(b.top, b.left) + (b.top === b.bottom && b.left === b.right ? '' : ':' + ref(b.bottom, b.right)) + '" sheet="' + escapeXml(b.sheet) + '"/>').join('')
      + '</dataRefs></dataConsolidate>');
    view.selection.collapseTo(box.top, box.left);
    return { ref: ref(box.top, box.left) + ':' + ref(box.bottom, box.right), rows: outRows.length, cols: outCols.length };
  });
}

/** Is there anything at all to write? */
function position(rows, cols) {
  return rows.length > 0 && cols.length > 0;
}

/**
 * What the sheet kept from its last consolidation — the dialog's starting
 * point, as Excel's does: `{ fn, topRow, leftCol, links, refs }`, or null.
 */
export function lastConsolidation(view) {
  const el = view.workbook._sheetPart(view.activeSheet).part.tailElement('dataConsolidate');
  if (!el) return null;
  const a = (name) => new RegExp('<dataConsolidate\\b[^>]*\\b' + name + '="([^"]*)"').exec(el)?.[1];
  const fileFn = a('function') || 'sum';
  const fn = Object.entries(CONSOLIDATE_FUNCTIONS).find(([, v]) => v.file === fileFn)?.[0] ?? 'sum';
  const on = (v) => v === '1' || v === 'true';
  const refs = [...el.matchAll(/<dataRef\b([^>]*)\/?>/g)].map((m) => {
    const at = /\bref="([^"]+)"/.exec(m[1])?.[1];
    const sheet = /\bsheet="([^"]*)"/.exec(m[1])?.[1];
    if (!at) return null;
    const [p, q] = at.split(':');
    const dollar = (x) => x.replace(/([A-Z]+)(\d+)/i, '$$$1$$$2');
    return (sheet ? quoteSheet(unescapeXml(sheet)) + '!' : '') + dollar(p) + (q ? ':' + dollar(q) : '');
  }).filter(Boolean);
  return { fn, topRow: on(a('topLabels')), leftCol: on(a('leftLabels')), links: on(a('link')), refs };
}

function escapeXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function unescapeXml(s) {
  return String(s).replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}
