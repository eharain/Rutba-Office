/**
 * Data → Outline: Group, Ungroup, Show and Hide Detail, the level buttons,
 * Clear Outline, and Subtotal.
 *
 * An outline in a workbook is three attributes on rows and columns and two
 * on the sheet: `outlineLevel` (1..7) on each grouped row or column,
 * `hidden` on the ones folded away, `collapsed` on the summary row whose
 * group is folded; `sheetFormatPr` carries the deepest level of each axis
 * and `sheetPr/outlinePr` says whether summaries sit below (and to the
 * right of) their detail, which is Excel's default, or above it. Everything
 * here writes exactly those, into the part, and mirrors them into the
 * geometry in place — the offsets already skip hidden rows, so a folded
 * group costs a frame nothing it did not cost before.
 *
 * The functions take the view; the view's own methods are one line each.
 */
import { ref, colName } from './selection.js';

export const MAX_OUTLINE_LEVEL = 7;

/**
 * Subtotal's functions, as Excel numbers them for SUBTOTAL and names them
 * on the rows it adds. "Count" is COUNTA — every entry counted — and
 * "Count Numbers" is COUNT, which is how Excel's dialog maps them.
 */
export const SUBTOTAL_FUNCTIONS = {
  sum: { number: 9, label: 'Total', name: 'Sum' },
  count: { number: 3, label: 'Count', name: 'Count' },
  average: { number: 1, label: 'Average', name: 'Average' },
  max: { number: 4, label: 'Max', name: 'Max' },
  min: { number: 5, label: 'Min', name: 'Min' },
  product: { number: 6, label: 'Product', name: 'Product' },
  countNumbers: { number: 2, label: 'Count', name: 'Count Numbers' },
  stdDev: { number: 7, label: 'StdDev', name: 'StdDev' },
  stdDevp: { number: 8, label: 'StdDevp', name: 'StdDevp' },
  var: { number: 10, label: 'Var', name: 'Var' },
  varp: { number: 11, label: 'Varp', name: 'Varp' },
};

/** A cell's formula reads as a subtotal: what Remove All takes away, and what Replace replaces. */
const SUBTOTAL_FORMULA = /^=.*\bSUBTOTAL\s*\(/i;

/** One axis of the outline, over the part and the geometry at once. */
function axisOf(view, axis) {
  const geo = view.geo;
  const { part } = view.workbook._sheetPart(view.activeSheet);
  const row = axis === 'row';
  const levels = row ? geo.rowLevels : geo.colLevels;
  const hidden = row ? geo.hiddenRows : geo.hiddenCols;
  const collapsed = row ? geo.collapsedRows : geo.collapsedCols;
  return {
    geo,
    part,
    levels,
    hidden,
    collapsed,
    level: (i) => levels.get(i) ?? 0,
    /** Write one index's attributes to the file and the geometry together. */
    write(i, spec) {
      if (row) part.setRowOutline(i, spec);
      else part.setColOutline(i, spec);
      if (spec.level !== undefined) {
        if (spec.level > 0) levels.set(i, spec.level);
        else if (levels.has(i)) levels.delete(i);
      }
      if (spec.hidden !== undefined) {
        if (spec.hidden && !hidden.has(i)) hidden.add(i);
        else if (!spec.hidden && hidden.has(i)) hidden.delete(i);
      }
      if (spec.collapsed !== undefined) {
        if (spec.collapsed && !collapsed.has(i)) collapsed.add(i);
        else if (!spec.collapsed && collapsed.has(i)) collapsed.delete(i);
      }
    },
    groups: () => geo.outlineGroups(axis),
    /** The deepest level, written where Excel keeps it. */
    settleDepth() {
      part.setOutlineProps(row ? { rowLevels: geo.outlineDepth('row') } : { colLevels: geo.outlineDepth('col') });
    },
  };
}

/** Run one outline change as one undo step on the active sheet's part. */
function outlineEdit(view, label, fn) {
  const partName = view.workbook.partNameFor(view.activeSheet);
  return view._edit(label, null, [], () => {
    const out = fn();
    view._structuralDirty = true;
    return out;
  }, { parts: [partName] });
}

/**
 * Which axis a selection means: whole rows are rows, whole columns are
 * columns ("whole" as far as the sheet is used — a heading click selects
 * that much), and anything else is neither, which is when Excel asks.
 */
export function selectionAxis(view) {
  const r = view.selection.range;
  const { maxRow, maxCol } = view.bounds;
  const wholeRows = r.left === 0 && r.right >= maxCol;
  const wholeCols = r.top === 0 && r.bottom >= maxRow;
  if (wholeRows && !wholeCols) return 'row';
  if (wholeCols && !wholeRows) return 'col';
  return null;
}

/**
 * Data → Group: the rows (or columns) one level deeper. Seven levels is
 * Excel's limit and the file's; an eighth is refused before anything moves.
 */
export function group(view, axis, from, to) {
  const lo = Math.min(from, to);
  const hi = Math.max(from, to);
  const a = axisOf(view, axis);
  for (let i = lo; i <= hi; i++) {
    if (a.level(i) >= MAX_OUTLINE_LEVEL) throw new Error('An outline goes seven levels deep at most');
  }
  return outlineEdit(view, 'group', () => {
    for (let i = lo; i <= hi; i++) a.write(i, { level: a.level(i) + 1 });
    a.settleDepth();
    return { axis, from: lo, to: hi, levels: a.geo.outlineDepth(axis) };
  });
}

/**
 * Data → Ungroup: the rows one level shallower. Rows that were folded
 * away and now belong to no folded group come back into view, and a
 * summary left with no group loses its `collapsed` mark.
 */
export function ungroup(view, axis, from, to) {
  const lo = Math.min(from, to);
  const hi = Math.max(from, to);
  const a = axisOf(view, axis);
  let any = false;
  for (let i = lo; i <= hi; i++) if (a.level(i) > 0) { any = true; break; }
  if (!any) throw new Error(axis === 'row' ? 'Those rows are not in a group' : 'Those columns are not in a group');
  return outlineEdit(view, 'ungroup', () => {
    for (let i = lo; i <= hi; i++) if (a.level(i) > 0) a.write(i, { level: a.level(i) - 1 });
    settle(a, lo - 1, hi + 1);
    a.settleDepth();
    return { axis, from: lo, to: hi, levels: a.geo.outlineDepth(axis) };
  });
}

/**
 * After levels change: an index hidden in the range that no folded group
 * holds any more is shown, and a `collapsed` mark on an index that is no
 * longer a folded group's summary is taken off.
 */
function settle(a, lo, hi) {
  const groups = a.groups();
  const held = inside(groups.filter((g) => a.collapsed.has(g.summary)));
  const summaries = new Set(groups.map((g) => g.summary));
  for (let i = Math.max(0, lo); i <= hi; i++) {
    if (a.hidden.has(i) && !held(i)) a.write(i, { hidden: false });
    if (a.collapsed.has(i) && !summaries.has(i)) a.write(i, { collapsed: false });
  }
}

/**
 * Whether an index lies in any of some groups, asked in rising order of
 * index: the groups are walked once with a pointer, not searched per row —
 * six hundred groups over sixty thousand rows was thirty-six million tests.
 */
function inside(groups) {
  const spans = groups.map((g) => [g.start, g.end]).sort((x, y) => x[0] - y[0]);
  let k = 0;
  let reach = -1;
  return (i) => {
    while (k < spans.length && spans[k][0] <= i) { reach = Math.max(reach, spans[k][1]); k += 1; }
    return i <= reach;
  };
}

/** Fold one group: its rows hidden, its summary marked. */
function fold(a, g) {
  for (let i = g.start; i <= g.end; i++) if (!a.hidden.has(i)) a.write(i, { hidden: true });
  if (g.summary >= 0) a.write(g.summary, { collapsed: true });
}

/** Unfold one group: its rows shown, except those of a folded group inside it. */
function unfold(a, g) {
  const keep = inside(a.groups().filter((x) => x.level > g.level && x.start >= g.start && x.end <= g.end && a.collapsed.has(x.summary)));
  for (let i = g.start; i <= g.end; i++) {
    if (!keep(i) && a.hidden.has(i)) a.write(i, { hidden: false });
  }
  if (g.summary >= 0 && a.collapsed.has(g.summary)) a.write(g.summary, { collapsed: false });
}

/** The + or − beside a group: fold it when it is open, open it when it is folded. */
export function toggleGroup(view, axis, level, start) {
  const a = axisOf(view, axis);
  const g = a.groups().find((x) => x.level === level && x.start === start);
  if (!g) throw new Error('There is no group of level ' + level + ' there');
  return outlineEdit(view, g.collapsed ? 'show detail' : 'hide detail', () => {
    if (g.collapsed) unfold(a, g);
    else fold(a, g);
    return { collapsed: !g.collapsed };
  });
}

/**
 * The level buttons, 1 2 3: level n shows every row above level n and
 * folds every group of level n or deeper. The last button shows all.
 */
export function showLevel(view, axis, n) {
  const a = axisOf(view, axis);
  const depth = a.geo.outlineDepth(axis);
  if (!depth) throw new Error(axis === 'row' ? 'No rows are grouped on this sheet' : 'No columns are grouped on this sheet');
  const level = Math.max(1, Math.min(depth + 1, Math.trunc(Number(n) || 1)));
  return outlineEdit(view, 'outline level', () => {
    const groups = a.groups();
    for (const [i, lv] of [...a.levels]) {
      const hide = lv >= level;
      if (hide !== a.hidden.has(i)) a.write(i, { hidden: hide });
    }
    for (const g of groups) {
      if (g.summary < 0) continue;
      const fold = g.level >= level;
      if (fold !== a.collapsed.has(g.summary)) a.write(g.summary, { collapsed: fold });
    }
    return { level };
  });
}

/**
 * Data → Show Detail / Hide Detail, at the active cell: the group it sits
 * in (Hide), or the folded group whose summary it is (Show). Rows first,
 * then columns, as Excel looks.
 */
export function detail(view, show) {
  const { row, col } = view.selection.active;
  for (const [axis, at] of [['row', row], ['col', col]]) {
    const a = axisOf(view, axis);
    const groups = a.groups();
    let g = null;
    if (show) {
      const folded = groups.filter((x) => x.collapsed && (x.summary === at || (at >= x.start && at <= x.end)));
      g = folded.sort((x, y) => x.level - y.level)[0] ?? null;
    } else {
      const inside = groups.filter((x) => !x.collapsed && at >= x.start && at <= x.end);
      const summarised = groups.filter((x) => !x.collapsed && x.summary === at);
      g = (inside.length ? inside : summarised).sort((x, y) => y.level - x.level)[0] ?? null;
    }
    if (g) {
      return outlineEdit(view, show ? 'show detail' : 'hide detail', () => {
        if (show) unfold(a, g);
        else fold(a, g);
        return { axis, level: g.level, start: g.start };
      });
    }
  }
  throw new Error(show ? 'There is no folded group here to show' : 'There is no group here to hide');
}

/**
 * Data → Ungroup → Clear Outline: every level on the sheet goes, and the
 * rows and columns folded away by it come back. (Excel leaves those hidden
 * and tells you to unhide them by hand; showing them is the kinder reading
 * of "clear", and nothing else hides a row inside a group.)
 */
export function clearOutline(view) {
  const rows = axisOf(view, 'row');
  const cols = axisOf(view, 'col');
  if (!rows.levels.size && !cols.levels.size) throw new Error('This sheet has no outline to clear');
  return outlineEdit(view, 'clear outline', () => {
    for (const a of [rows, cols]) {
      for (const g of a.groups()) {
        for (let i = g.start; i <= g.end; i++) if (a.hidden.has(i)) a.write(i, { hidden: false });
      }
      for (const i of [...a.collapsed]) a.write(i, { collapsed: false });
      for (const i of [...a.levels.keys()]) a.write(i, { level: 0 });
    }
    rows.part.setOutlineProps({ rowLevels: 0, colLevels: 0 });
    return true;
  });
}

/**
 * What a frame carries of the outline: per axis, the depth, the side the
 * summaries sit on, and the groups near the viewport with their pixel
 * extents — the window draws the brackets and buttons from these and
 * needs no lookups of its own. Null on an axis nothing is grouped on.
 */
export function outlineFrame(geo, vp, frozen) {
  const one = (axis) => {
    const groups = geo.outlineGroups(axis);
    if (!groups.length) return null;
    const row = axis === 'row';
    const first = row ? vp.firstRow : vp.firstCol;
    const last = row ? vp.lastRow : vp.lastCol;
    const pinned = row ? frozen.rows : frozen.cols;
    const offset = (i) => (row ? geo.rowOffset(i) : geo.colOffset(i));
    const size = (i) => (row ? geo.rowHeight(i) : geo.colWidth(i));
    let depth = 0;
    const near = [];
    for (const g of groups) {
      if (g.level > depth) depth = g.level;
      const lo = Math.min(g.start, g.summary);
      const hi = Math.max(g.end, g.summary);
      const seen = (hi >= first - 1 && lo <= last + 1) || (pinned && lo < pinned);
      if (!seen) continue;
      near.push({
        level: g.level,
        start: g.start,
        end: g.end,
        summary: g.summary,
        collapsed: g.collapsed,
        from: offset(g.start),
        to: offset(g.end) + size(g.end),
        // The summary's own place, for its button; none where the summary
        // would sit above the first row, or is folded into a group itself.
        at: g.summary >= 0 && size(g.summary) > 0 ? offset(g.summary) : null,
        size: g.summary >= 0 ? size(g.summary) : 0,
      });
    }
    return { levels: depth, below: row ? geo.summaryBelow : geo.summaryRight, groups: near };
  };
  const rows = one('row');
  const cols = one('col');
  return rows || cols ? { rows, cols } : null;
}

// ── Subtotal ──────────────────────────────────────────────────────────────

/** The list the Subtotal and Remove All act on: the selection, or the block round the cell. */
function listRange(view) {
  const sel = view.selection.range;
  const single = sel.top === sel.bottom && sel.left === sel.right;
  const r = single ? view._currentRegion(sel.top, sel.left) : { ...sel };
  if (r.bottom <= r.top) throw new Error('Subtotal needs a list with a header row and at least one row of data under it');
  return r;
}

/**
 * How many of a sorted list lie before a row (or at it, with `inclusive`)
 * — a binary search, since it is asked once per row of a long list.
 */
function countBelow(sorted, row, inclusive) {
  let lo = 0;
  let hi = sorted.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (inclusive ? sorted[mid] <= row : sorted[mid] < row) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/** Rows of the list that hold a SUBTOTAL formula in any of its columns. */
function subtotalRows(view, r) {
  const out = [];
  for (let row = r.top + 1; row <= r.bottom; row++) {
    for (let c = r.left; c <= r.right; c++) {
      if (SUBTOTAL_FORMULA.test(String(view.calc.getInput(view.activeSheet, row, c) ?? ''))) { out.push(row); break; }
    }
  }
  return out;
}

/** The list's columns by header, for the dialog. */
export function listFields(view) {
  const sel = view.selection.range;
  const single = sel.top === sel.bottom && sel.left === sel.right;
  const r = single ? view._currentRegion(sel.top, sel.left) : { ...sel };
  const columns = [];
  for (let c = r.left; c <= r.right; c++) {
    columns.push({ col: c, name: String(view.displayValue(r.top, c).text ?? '').trim() || 'Column ' + colName(c) });
  }
  return {
    ref: ref(r.top, r.left) + ':' + ref(r.bottom, r.right),
    top: r.top, left: r.left, bottom: r.bottom, right: r.right,
    columns,
    subtotals: r.bottom > r.top ? subtotalRows(view, r).length : 0,
  };
}

/**
 * The flush, the transform and the rebuild of a structural edit, as one
 * undo step. The rows move, so every sheet part and the names travel.
 */
export function structuralEdit(view, label, fn) {
  const parts = [...new Set([...view.workbook.sheets().map((s) => s.part), view.workbook.mainPart])];
  return view._edit(label, null, [], () => {
    view._flushPendingEdits();
    view.dirtyCells.clear();
    view.styledCells.clear();
    const out = fn();
    view._structuralDirty = true;
    view._rebuildDerivedState();
    return out;
  }, { parts, structural: true });
}

/** Every outline mark taken off the rows of a band, and the rows shown. */
function stripRows(part, top, bottom) {
  for (const row of part.rows) {
    if (row.index < top || row.index > bottom) continue;
    part.setRowOutline(row.index, { level: 0, hidden: false, collapsed: false });
  }
}

/** The manual page breaks a sheet carries, as 0-based rows that start a page. */
function readBreaks(part) {
  const el = part.tailElement('rowBreaks');
  if (!el) return [];
  return [...el.matchAll(/<brk\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1])).filter((n) => n > 0);
}

function writeBreaks(part, ids) {
  const list = [...new Set(ids)].filter((n) => n > 0).sort((a, b) => a - b);
  part.setTailElement('rowBreaks', list.length
    ? '<rowBreaks count="' + list.length + '" manualBreakCount="' + list.length + '">'
      + list.map((id) => '<brk id="' + id + '" max="16383" man="1"/>').join('') + '</rowBreaks>'
    : null);
}

/**
 * Data → Subtotal: at each change in one column, a row with SUBTOTAL
 * formulas under the chosen columns and a bold "<value> Total" label; a
 * Grand Total at the end (the top, with summaries above); the three-level
 * outline round them. Replace takes the list's earlier subtotals away
 * first; without it the new groups nest inside the old ones. Page breaks
 * between groups are Excel's manual breaks.
 *
 * @param {{ groupBy: number, fn?: string, columns: number[], replace?: boolean, pageBreaks?: boolean, summaryBelow?: boolean }} spec
 */
export function subtotal(view, {
  groupBy, fn = 'sum', columns = [], replace = true, pageBreaks = false, summaryBelow = true,
} = {}) {
  const f = SUBTOTAL_FUNCTIONS[fn];
  if (!f) throw new Error('Subtotal has no function "' + fn + '"');
  const r = listRange(view);
  if (!(groupBy >= r.left && groupBy <= r.right)) throw new Error('The column to group by is not in the list');
  const cols = [...new Set(columns.map(Number))].filter((c) => c >= r.left && c <= r.right).sort((a, b) => a - b);
  if (!cols.length) throw new Error('Choose at least one column to add a subtotal to');

  const sheet = view.activeSheet;
  const old = subtotalRows(view, r);
  const oldSet = new Set(old);
  const nest = !replace && old.length > 0;
  // Old subtotal rows go when replacing; kept, they are walls a new group
  // cannot cross.
  const removed = nest ? [] : old;
  const after = (row) => row - countBelow(removed, row, false);

  // The groups, in the rows as they stand after the old subtotals go.
  const key = (row) => String(view.displayValue(row, groupBy).text ?? '').trim().toLowerCase();
  const groups = [];
  let current = null;
  for (let row = r.top + 1; row <= r.bottom; row++) {
    if (oldSet.has(row)) { current = null; continue; }
    const k = key(row);
    if (current && current.key === k) { current.last = after(row); continue; }
    current = { key: k, label: String(view.displayValue(row, groupBy).text ?? '').trim(), first: after(row), last: after(row), source: row };
    groups.push(current);
  }
  if (!groups.length) throw new Error('There are no rows of data to subtotal');
  const bottomAfter = r.bottom - removed.length;
  // Where the grand row goes: past the list, or — nesting under Excel's own
  // Grand row — right above the first grand row already there.
  const isGrand = (row) => {
    for (let c = r.left; c <= r.right; c++) if (/^grand\b/i.test(String(view.displayValue(row, c).text ?? '').trim())) return true;
    return false;
  };
  const grandRows = nest ? old.filter(isGrand) : [];
  const grandPoint = summaryBelow
    ? (grandRows.length ? after(grandRows[0]) : bottomAfter + 1)
    : r.top + 1;
  const inserts = [];
  if (!summaryBelow) inserts.push({ point: grandPoint, grand: true });
  for (const g of groups) inserts.push({ point: summaryBelow ? g.last + 1 : g.first, group: g });
  if (summaryBelow) inserts.push({ point: grandPoint, grand: true });
  // Stable by point: the order above is the order rows land at one point.
  const order = inserts.map((x, i) => ({ ...x, i })).sort((a, b) => a.point - b.point || a.i - b.i);
  order.forEach((x, k) => { x.at = x.point + k; });
  const points = order.map((x) => x.point);
  const shifted = (row) => row + countBelow(points, row, true);

  // What the data's own cells look like, read before anything moves.
  const styleOf = (row, col) => view._styleIndexAt(sheet, row, col);

  return structuralEdit(view, 'subtotal', () => {
    const { part } = view.workbook._sheetPart(sheet);
    if (removed.length) view.workbook.deleteRowsAt(sheet, removed);
    if (!nest) stripRows(part, r.top, bottomAfter);
    view.workbook.insertRowsAt(sheet, points, { materialize: true });
    // The levels the rows carry now, read before any is changed: nesting
    // puts every one a level deeper.
    const levelBefore = new Map();
    if (nest) for (const row of part.rows) levelBefore.set(row.index, Number(/\boutlineLevel="(\d+)"/.exec(row.attrsStr)?.[1] ?? 0));
    const bold = new Map();
    const boldOf = (base) => {
      const k = base ?? -1;
      if (!bold.has(k)) bold.set(k, view._boldIndex(base));
      return bold.get(k);
    };
    const lastGroup = order.filter((y) => !y.grand).at(-1);

    const A1 = (row, col) => colName(col) + (row + 1);
    const grandTop = summaryBelow ? r.top + 1 : order.find((x) => x.grand).at + 1;
    const lastRow = shifted(bottomAfter);
    const grandBottom = summaryBelow ? order.find((x) => x.grand).at - 1 : lastRow;
    const levelOf = (row) => Number(levelBefore.get(row) ?? 0);
    const breaks = [];
    for (const x of order) {
      const at = x.at;
      const label = x.grand ? 'Grand ' + f.label : (x.group.label || '(blank)') + ' ' + f.label;
      const labelBase = x.grand ? null : styleOf(x.group.source, groupBy);
      view.workbook.setCell(sheet, A1(at, groupBy), label);
      view.workbook.setCellStyle(sheet, A1(at, groupBy), boldOf(labelBase));
      for (const c of cols) {
        if (c === groupBy) continue;
        const top = x.grand ? grandTop : shifted(x.group.first);
        const bottom = x.grand ? grandBottom : shifted(x.group.last);
        view.workbook.setCell(sheet, A1(at, c), '=SUBTOTAL(' + f.number + ',' + A1(top, c) + ':' + A1(bottom, c) + ')');
        const look = styleOf(x.grand ? groups[0].source : x.group.source, c);
        if (look !== null && look !== undefined) view.workbook.setCellStyle(sheet, A1(at, c), look);
      }
      if (x.grand) continue;
      // The outline: the group's rows one level under its subtotal row.
      const base = nest ? levelOf(shifted(x.group.first)) : 1;
      for (let row = shifted(x.group.first); row <= shifted(x.group.last); row++) {
        part.setRowOutline(row, { level: (nest ? levelOf(row) : 1) + 1 });
      }
      part.setRowOutline(at, { level: base });
      if (pageBreaks && x !== lastGroup) breaks.push(summaryBelow ? at + 1 : shifted(x.group.last) + 1);
    }
    if (breaks.length) writeBreaks(part, [...readBreaks(part), ...breaks]);
    let depth = 0;
    for (const row of part.rows) {
      const lv = Number(/\boutlineLevel="(\d+)"/.exec(row.attrsStr)?.[1] ?? 0);
      if (lv > depth) depth = lv;
    }
    part.setOutlineProps({ rowLevels: depth, summaryBelow });
    const bottom = bottomAfter + order.length;
    // The cursor on the list's first cell. Excel leaves the whole list
    // selected; on a long list that makes every later frame and toolbar
    // read walk every cell of it.
    view.selection.collapseTo(r.top, r.left);
    return { groups: groups.length, rows: order.length, range: ref(r.top, r.left) + ':' + ref(bottom, r.right), levels: depth };
  });
}

/**
 * Subtotal → Remove All: every row of the list holding a SUBTOTAL formula
 * goes, the outline over the list goes with them, and so do the page
 * breaks inside it — the list as it was before the first Subtotal.
 */
export function removeSubtotals(view) {
  const r = listRange(view);
  const sheet = view.activeSheet;
  const rows = subtotalRows(view, r);
  if (!rows.length) throw new Error('This list has no subtotals to remove');
  return structuralEdit(view, 'remove subtotals', () => {
    const { part } = view.workbook._sheetPart(sheet);
    view.workbook.deleteRowsAt(sheet, rows);
    const bottom = r.bottom - rows.length;
    stripRows(part, r.top, bottom);
    const breaks = readBreaks(part);
    if (breaks.length) writeBreaks(part, breaks.filter((id) => id <= r.top || id > bottom));
    let depth = 0;
    for (const row of part.rows) {
      const lv = Number(/\boutlineLevel="(\d+)"/.exec(row.attrsStr)?.[1] ?? 0);
      if (lv > depth) depth = lv;
    }
    part.setOutlineProps({ rowLevels: depth, summaryBelow: true });
    // The cursor on the list's first cell. Excel leaves the whole list
    // selected; on a long list that makes every later frame and toolbar
    // read walk every cell of it.
    view.selection.collapseTo(r.top, r.left);
    return { removed: rows.length, range: ref(r.top, r.left) + ':' + ref(bottom, r.right) };
  });
}
