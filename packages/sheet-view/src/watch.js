/**
 * Formulas → Watch Window: the small, engine-free bookkeeping behind the
 * watched-cell list. A watch is just `{ sheet, row, col }` — the workbook it
 * points at is read fresh every time `resolveWatches` is asked for the pane's
 * rows, the same way Error Checking rereads `errorCells()` on every frame
 * rather than caching a value that could go stale.
 *
 * None of `addWatches` or `removeWatch` touches a workbook at all, which is
 * what keeps them cheap to test and cheap to call from a ribbon handler.
 */
import { ref } from './selection.js';
import { formatValue } from './numfmt.js';

const keyOf = (w) => `${w.sheet}!${ref(w.row, w.col)}`;

/**
 * Add every ref in `refs` to `list`, in order, skipping one already on it
 * (same sheet, row and column) — so pressing Add Watch again on a cell
 * already watched, or on a range that overlaps one, never doubles a row.
 * Neither argument is mutated; a new array comes back.
 *
 * @param {Array<{sheet:string,row:number,col:number}>} list
 * @param {Array<{sheet:string,row:number,col:number}>} refs
 */
export function addWatches(list, refs) {
  const out = list.slice();
  const seen = new Set(out.map(keyOf));
  for (const r of refs) {
    const w = { sheet: r.sheet, row: r.row, col: r.col };
    const k = keyOf(w);
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(w);
  }
  return out;
}

/**
 * Drop one watch (same sheet, row and column) from the list. Every other
 * watch keeps its place; a ref that is not on the list leaves it unchanged.
 */
export function removeWatch(list, ref) {
  return list.filter((w) => !(w.sheet === ref.sheet && w.row === ref.row && w.col === ref.col));
}

/**
 * Turn each watch into a row for the pane — its reference text, its live
 * value and its formula (null for a plain value) — read straight off the
 * workbook's calc model, the way the grid itself reads a cell: nothing here
 * is cached, so a row is only ever as stale as the workbook's last
 * recalculation. A watch on a sheet since removed from the workbook is left
 * out rather than thrown over, the same as a stale cell reference anywhere
 * else in the engine.
 *
 * @param {Array<{sheet:string,row:number,col:number}>} list
 * @param {{sheetNames: () => string[], calc: {cell: Function, getValue: Function}}} workbook
 */
export function resolveWatches(list, workbook) {
  const names = new Set(workbook.sheetNames());
  const out = [];
  for (const w of list) {
    if (!names.has(w.sheet)) continue;
    const cell = workbook.calc.cell(w.sheet, w.row, w.col);
    const value = workbook.calc.getValue(w.sheet, w.row, w.col);
    out.push({
      sheet: w.sheet,
      row: w.row,
      col: w.col,
      ref: ref(w.row, w.col),
      value: formatValue(value, 'General').text,
      formula: cell?.formula ?? null,
    });
  }
  return out;
}
