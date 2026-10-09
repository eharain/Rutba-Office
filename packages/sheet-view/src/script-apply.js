// Automate → scripts, in the workbook: the snapshot a script reads, and its
// edits applied as one undo step.

import { applyFormat } from './styles-write.js';
import { pixelsToCharWidth } from './geometry.js';

/** Points, as Office Scripts measure a width or height, to the grid's pixels. */
const ptToPx = (pt) => Math.max(1, Math.round((Number(pt) || 0) * (96 / 72)));

/** A cell's A1 name, and a range's box from its A1 name. */
const a1 = (row, col) => { let s = ''; for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s + (row + 1); };
const cellOf = (text) => {
  const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(String(text).trim());
  if (!m) return null;
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]) - 1, col: col - 1 };
};
const boxOf = (refText) => {
  const [p, q = p] = String(refText).split(':');
  const a = cellOf(p);
  const b = cellOf(q);
  return a && b ? { top: Math.min(a.row, b.row), bottom: Math.max(a.row, b.row), left: Math.min(a.col, b.col), right: Math.max(a.col, b.col) } : null;
};
/** A sheet's merged regions, by their A1 ranges. */
const mergesOf = (view, sheet) => view.workbook._sheetPart(sheet).part.mergeRefs();

const NAMED = { black: '#000000', white: '#FFFFFF', red: '#FF0000', green: '#00B050', blue: '#0070C0', yellow: '#FFFF00', orange: '#FFC000', purple: '#7030A0', grey: '#808080', gray: '#808080' };
const colour = (c) => {
  if (c == null) return null;
  const s = String(c).trim();
  if (NAMED[s.toLowerCase()]) return NAMED[s.toLowerCase()];
  const m = /^#?([0-9a-f]{6})$/i.exec(s);
  if (!m) throw new Error(`"${c}" is not a colour a script can set: a name or #RRGGBB`);
  return `#${m[1].toUpperCase()}`;
};

/** The workbook as a script reads it: each sheet's filled cells, values and formulas. */
export function scriptSnapshot(view, { limit = 200000 } = {}) {
  let count = 0;
  const sheets = view.sheetNames().map((name) => {
    const cells = {};
    const { maxRow, maxCol } = view.calc.usedBounds(name);
    for (let row = 0; row <= maxRow; row++) {
      for (let col = 0; col <= maxCol; col++) {
        const input = view.calc.getInput(name, row, col);
        const v = view.calc.getValue(name, row, col);
        if ((input === '' || input == null) && (v === null || v === undefined || v === '')) continue;
        if (++count > limit) throw new Error('This workbook is too large for a script to read at once');
        const f = typeof input === 'string' && input.startsWith('=') ? input : null;
        cells[`${row},${col}`] = { v: v === undefined || (v && typeof v === 'object') ? null : v, f };
      }
    }
    return { name, cells };
  });
  const s = view.selection;
  return { name: view.fileName || 'Workbook', activeSheet: view.activeSheet, selection: { sheet: view.activeSheet, top: s.top, left: s.left, bottom: s.bottom, right: s.right }, sheets };
}

/** Every part an edit of the whole workbook may change, for its one undo step. */
function partsOf(view) {
  const out = [...view.workbook.sheets().map((s) => s.part), view.workbook.mainPart, 'xl/_rels/workbook.xml.rels', '[Content_Types].xml', 'xl/styles.xml'];
  return [...new Set(out)].filter((p) => view.pkg.has(p));
}

/**
 * A script's edits applied to the workbook as one undo step: values and
 * formulas on any sheet, formats, sheets added and renamed; then the sheet
 * and the range it activated or selected last. Answers how many cells it set.
 */
export function applyScriptEdits(view, edits = []) {
  if (!edits.length) return 0;
  // The edits come back from a worker running someone's script: a cell
  // beyond Excel's grid would be written into the file, which Excel then
  // calls damaged, so it is refused before anything is changed.
  for (const e of edits) {
    if ((e.kind === 'value' || e.kind === 'formula') && !(Number.isInteger(e.row) && e.row >= 0 && e.row <= 1048575 && Number.isInteger(e.col) && e.col >= 0 && e.col <= 16383)) {
      throw new Error('A script cannot set a cell outside the worksheet (A1 to XFD1048576)');
    }
  }
  view._structureGate();
  let cells = 0;
  const FORMAT_KEYS = { bold: 'bold', italic: 'italic', underline: 'underline', strike: 'strike', color: 'fontColour', fill: 'fill', fontSize: 'fontSize', fontFamily: 'fontName', align: 'align', valign: 'valign', wrap: 'wrap', numberFormat: 'numberFormat' };
  let show = null;
  view._edit('script', null, [], () => {
    view._flushForStructure();
    let xml = null;
    for (const e of edits) {
      switch (e.kind) {
        case 'addSheet':
          // What the script set so far goes into the parts first: the rebuild reads them.
          view._flushForStructure();
          view.workbook.addSheet(e.sheet);
          view._rebuildDerivedState();
          break;
        case 'rename':
          view._flushForStructure();
          view.workbook.renameSheet(e.sheet, e.to);
          if (view.activeSheet === e.sheet) view.activeSheet = e.to;
          view._rebuildDerivedState();
          break;
        case 'value':
        case 'formula':
          view._setCellOn(e.sheet, e.row, e.col, e.value);
          cells += 1;
          break;
        case 'clear':
          // Only as far as the sheet is used: a whole-sheet range walked
          // 1.6 billion empty cells and held the window for minutes.
          const used = view.calc.usedBounds(e.sheet);
          for (let r = e.top; r <= Math.min(e.bottom, used.maxRow); r++) {
            for (let c = e.left; c <= Math.min(e.right, used.maxCol); c++) if (view.calc.getInput(e.sheet, r, c) !== '') { view._setCellOn(e.sheet, r, c, ''); cells += 1; }
          }
          break;
        case 'format': {
          view._ensureStylesPart();
          const resolved = {};
          for (const [k, v] of Object.entries(e.format || {})) {
            if (!FORMAT_KEYS[k]) continue;
            resolved[FORMAT_KEYS[k]] = k === 'color' || k === 'fill' ? (v == null ? null : colour(v)) : v;
          }
          xml = view.pkg.text('xl/styles.xml');
          const bottom = Math.min(e.bottom, view.calc.usedBounds(e.sheet).maxRow);
          const right = Math.min(e.right, view.calc.usedBounds(e.sheet).maxCol);
          for (let r = e.top; r <= Math.max(e.top, bottom); r++) {
            for (let c = e.left; c <= Math.max(e.left, right); c++) {
              const base = view._styleIndexAt(e.sheet, r, c);
              const out = applyFormat(xml, base, resolved);
              xml = out.xml;
              if (out.index !== base) view._setStyleIndex(e.sheet, r, c, out.index);
            }
          }
          view.pkg.write_('xl/styles.xml', xml);
          break;
        }
        // getFormat().setColumnWidth / setRowHeight, in points as Office Scripts take them,
        // and autofitColumns: the widest text in each column, as the ribbon's AutoFit makes it.
        // These were offered to a script and refused when it ran.
        case 'colWidth':
        case 'autofit': {
          const used = view.calc.usedBounds(e.sheet);
          for (let c = e.left; c <= Math.min(e.right, Math.max(e.left, used.maxCol)); c++) {
            let px;
            if (e.kind === 'colWidth') px = ptToPx(e.width);
            else {
              let widest = 0;
              for (let r = 0; r <= used.maxRow; r++) {
                const v = view.calc.getValue(e.sheet, r, c);
                if (v != null && v !== '') widest = Math.max(widest, String(typeof v === 'object' ? v.value ?? '' : v).length);
              }
              px = Math.max(40, Math.min(600, Math.round(widest * 7.2 + 14)));
            }
            view.workbook.setColWidthChars(e.sheet, c, pixelsToCharWidth(px));
            if (e.sheet === view.activeSheet) view.geo.colWidths.set(c, px);
          }
          view._structuralDirty = true;
          break;
        }
        case 'rowHeight':
          for (let r = e.top; r <= e.bottom; r++) {
            view.workbook.setRowHeightPoints(e.sheet, r, Number(e.height) || 15);
            if (e.sheet === view.activeSheet) view.geo.rowHeights.set(r, ptToPx(e.height));
          }
          view._structuralDirty = true;
          break;
        // merge() and unmerge(): the cells but the first emptied, as the ribbon's Merge does.
        case 'merge': {
          if (e.top === e.bottom && e.left === e.right) break;
          // A merge inside the range gives way to it; one only partly inside is refused, as the ribbon refuses it.
          for (const m of mergesOf(view, e.sheet)) {
            const box = boxOf(m);
            if (!box || box.right < e.left || box.left > e.right || box.bottom < e.top || box.top > e.bottom) continue;
            if (box.top < e.top || box.bottom > e.bottom || box.left < e.left || box.right > e.right) throw new Error('A script cannot merge across part of a merged region');
            view.workbook.removeMerge(e.sheet, m);
          }
          view.workbook.addMerge(e.sheet, `${a1(e.top, e.left)}:${a1(e.bottom, e.right)}`);
          for (let r = e.top; r <= e.bottom; r++) {
            for (let c = e.left; c <= e.right; c++) if ((r !== e.top || c !== e.left) && view.calc.getInput(e.sheet, r, c) !== '') { view._setCellOn(e.sheet, r, c, ''); cells += 1; }
          }
          view._structuralDirty = true;
          break;
        }
        case 'unmerge':
          for (const m of mergesOf(view, e.sheet)) {
            const box = boxOf(m);
            if (box && !(box.right < e.left || box.left > e.right || box.bottom < e.top || box.top > e.bottom)) view.workbook.removeMerge(e.sheet, m);
          }
          view._structuralDirty = true;
          break;
        case 'activate':
          show = { sheet: e.sheet, range: null };
          break;
        case 'select':
          show = { sheet: e.sheet, range: e };
          break;
        default:
          throw new Error(`"${e.kind}" is not an edit a script makes`);
      }
    }
    view._flushForStructure();
    view._rebuildDerivedState();
    return cells;
  }, { parts: partsOf(view), tracksNewParts: true, structural: true, sheetGate: false });
  if (show && view.sheetNames().includes(show.sheet)) {
    view.selectSheet(show.sheet);
    if (show.range) { view.select(show.range.top, show.range.left); view.selection.extendTo(show.range.bottom, show.range.right); }
  }
  return cells;
}
