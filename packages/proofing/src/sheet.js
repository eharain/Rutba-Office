// Worksheets: what the accessibility checker reads, the fixes it writes, and
// the cells the spelling pass walks — against a SheetView (sheet-view) over
// the Workbook (ooxml/workbook.js).

import { OoxmlPackage } from '@rutba/ooxml/package';
import { findElements, readAltProps, writeAltProps } from './alt-text.js';
import { normaliseColour } from './colour.js';

const colName = (c) => {
  let s = '';
  let n = c + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
};
const refOf = (row, col) => `${colName(col)}${row + 1}`;
const parseRef = (ref) => {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(String(ref).trim().toUpperCase());
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]) - 1, col: col - 1 };
};

/** The drawing part a sheet's pictures and charts live in, or null. */
function drawingPartOf(view, sheet) {
  const sheetPart = view.workbook.partNameFor(sheet);
  if (!sheetPart) return null;
  const rel = view.pkg.rels(sheetPart).find((r) => String(r.Type).endsWith('/drawing'));
  if (!rel) return null;
  const part = OoxmlPackage.resolveTarget(sheetPart, rel.Target);
  return view.pkg.has(part) ? part : null;
}

const ANCHOR_RE = /<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>[\s\S]*?<\/xdr:\1>/g;

/** A sheet's anchored objects in drawing order: kind, where it sits, and its alt text. */
export function sheetDrawings(view, sheet) {
  const part = drawingPartOf(view, sheet);
  if (!part) return [];
  const xml = view.pkg.text(part);
  const out = [];
  let i = 0;
  for (const m of xml.matchAll(ANCHOR_RE)) {
    const a = m[0];
    const kind = /<xdr:grpSp\b/.test(a) ? 'group' : /<xdr:pic\b/.test(a) ? 'picture' : /<c:chart\b/.test(a) ? 'chart' : /<xdr:cxnSp\b/.test(a) ? 'connector' : /<xdr:sp\b/.test(a) ? 'shape' : 'object';
    const cNvPr = findElements(a, 'xdr:cNvPr')[0];
    const from = /<xdr:from>\s*<xdr:col>(\d+)<\/xdr:col>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(a);
    const hasText = /<a:t>[^<]*\S[^<]*<\/a:t>/.test(a);
    out.push({ anchor: i++, part, kind, hasText, row: from ? Number(from[2]) : 0, col: from ? Number(from[1]) : 0, ...(cNvPr ? readAltProps(cNvPr.xml) : { name: '', descr: '', decorative: false }) });
  }
  return out;
}

const KIND_LABEL = { picture: 'Picture', chart: 'Chart', group: 'Group', shape: 'Shape', object: 'Object' };
const shortText = (s, n = 40) => {
  const t = String(s ?? '').replace(/\s+/g, ' ').trim();
  return t.length > n ? `${t.slice(0, n - 1)}…` : t;
};

/** Every filled cell of a sheet, row by row: `{ row, col, input }`. */
export function sheetCells(view, sheet, { limit = 200000 } = {}) {
  const cells = view.calc.sheets.get(sheet);
  const out = [];
  if (view.calc.lazy?.has?.(sheet)) {
    const provider = view.calc.lazy.get(sheet);
    const b = provider.bounds?.() || { maxRow: -1, maxCol: -1 };
    for (let row = 0; row <= b.maxRow && out.length < limit; row++) {
      for (let col = 0; col <= b.maxCol; col++) {
        const input = view.calc.getInput(sheet, row, col);
        if (input !== '' && input != null) out.push({ row, col, input });
      }
    }
    return out;
  }
  for (const c of cells?.values() || []) {
    if (c.tombstone || c.input === '' || c.input == null) continue;
    out.push({ row: c.row, col: c.col, input: c.input });
    if (out.length >= limit) break;
  }
  return out.sort((a, b) => a.row - b.row || a.col - b.col);
}

/** The workbook, described for the rules in accessibility.js. */
export function describeSheet(view) {
  const model = { app: 'sheets', objects: [], tables: [], merges: [], texts: [], links: [], sheets: [] };
  const names = view.workbook.sheetNames();
  for (const sheet of names) {
    model.sheets.push({ name: sheet, where: { sheet } });

    const counters = {};
    for (const d of sheetDrawings(view, sheet)) {
      if (d.kind === 'connector' || (d.kind === 'shape' && d.hasText)) continue;
      counters[d.kind] = (counters[d.kind] || 0) + 1;
      model.objects.push({
        key: `${sheet}:${d.anchor}`, kind: d.kind,
        label: `${d.name || `${KIND_LABEL[d.kind]} ${counters[d.kind]}`} (${sheet})`,
        alt: d.descr, decorative: d.decorative,
        where: { sheet, row: d.row, col: d.col, drawing: d.anchor },
        target: { sheet, anchor: d.anchor },
      });
    }

    for (const m of view.merges.get(sheet) || []) {
      model.merges.push({ key: `${sheet}!${m.ref}`, label: `${sheet}!${m.ref}`, where: { sheet, row: m.top, col: m.left }, target: { sheet, row: m.top, col: m.left, ref: m.ref } });
    }

    // Text in a colour of its own against the cell's fill (or the white of the sheet).
    for (const c of sheetCells(view, sheet, { limit: 20000 })) {
      const index = view._styleIndexAt(sheet, c.row, c.col);
      const style = index == null ? null : view.styles.byStyleIndex[index];
      const fg = normaliseColour(style?.font?.colour);
      if (!fg) continue;
      const bg = normaliseColour(style?.fill?.colour) || '#FFFFFF';
      const ref = refOf(c.row, c.col);
      model.texts.push({ key: `${sheet}!${ref}`, label: `${sheet}!${ref}`, where: { sheet, row: c.row, col: c.col }, fg, bg, sizePt: style?.font?.sizePt || 11, bold: Boolean(style?.font?.bold), target: { sheet, row: c.row, col: c.col } });
    }

    for (const h of view.workbook.hyperlinks(sheet)) {
      const at = parseRef(String(h.ref).split(':')[0]);
      if (!at) continue;
      const shown = view.calc.getValue(sheet, at.row, at.col);
      const text = h.display || (shown == null ? '' : String(shown?.text ?? shown));
      model.links.push({ key: `${sheet}!${h.ref}`, text, url: h.href || h.location || '', label: `${sheet}!${h.ref}: "${shortText(text)}"`, where: { sheet, row: at.row, col: at.col } });
    }
  }

  for (const t of view.workbook.tables()) {
    const [a] = String(t.ref).split(':');
    const at = parseRef(a) || { row: 0, col: 0 };
    model.tables.push({
      key: t.part, label: `${t.name} (${t.sheet}!${t.ref})`,
      hasHeader: t.headerRowCount !== 0, merged: false,
      where: { sheet: t.sheet, row: at.row, col: at.col },
      target: { part: t.part },
      headerFix: 'Use the first row as the header',
    });
  }
  return model;
}

/* ── the fixes ─────────────────────────────────────────────────────────── */

/** Alt text on the `anchor`-th object of a sheet's drawing part. One undo step. */
export function setSheetAltText(view, { sheet, anchor, descr = '', decorative = false }) {
  const part = drawingPartOf(view, sheet);
  if (!part) throw new Error('that sheet has no pictures or charts');
  return view._edit('alt text', null, [], () => {
    const xml = view.pkg.text(part);
    const anchors = [...xml.matchAll(ANCHOR_RE)];
    const hit = anchors[Number(anchor)];
    if (!hit) throw new Error('that object is no longer on the sheet');
    const cNvPr = findElements(hit[0], 'xdr:cNvPr')[0];
    if (!cNvPr) throw new Error('that object has no properties to describe');
    const next = hit[0].slice(0, cNvPr.start) + writeAltProps(cNvPr.xml, { descr, decorative }) + hit[0].slice(cNvPr.end);
    view.pkg.write_(part, xml.slice(0, hit.index) + next + xml.slice(hit.index + hit[0].length));
    view._structuralDirty = true;
    return view;
  }, { parts: [part] });
}

const escAttr = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/**
 * Excel's "Use First Row as Header" for a table saved without one: the
 * table's first row becomes its header — each column named from the cell,
 * made unique, "Column1" for an empty one — and the filter buttons with it.
 */
export function setSheetTableHeader(view, { part }) {
  if (!view.pkg.has(part)) throw new Error('that table is no longer in the workbook');
  const info = view.workbook.tables().find((t) => t.part === part);
  if (!info) throw new Error('that table is no longer in the workbook');
  const [a, b] = String(info.ref).split(':');
  const from = parseRef(a);
  const to = parseRef(b || a);
  const parts = [...new Set([...view.workbook.sheets().map((s) => s.part), view.workbook.mainPart, part])];
  return view._edit('table header', null, [], () => {
    view._flushPendingEdits();
    const used = new Set();
    const names = [];
    for (let col = from.col; col <= to.col; col++) {
      const value = view.calc.getValue(info.sheet, from.row, col);
      let name = String(value?.text ?? value ?? '').trim() || `Column${col - from.col + 1}`;
      let n = 2;
      const base = name;
      while (used.has(name.toLowerCase())) name = `${base}${n++}`;
      used.add(name.toLowerCase());
      names.push(name);
    }
    let xml = view.pkg.text(part);
    xml = xml.replace(/(<table\b[^>]*?)\sheaderRowCount="0"/, '$1');
    let i = 0;
    xml = xml.replace(/<tableColumn\b([^>]*?)\sname="[^"]*"/g, (m, pre) => `<tableColumn${pre} name="${escAttr(names[i++] ?? `Column${i}`)}"`);
    if (!/<autoFilter\b/.test(xml)) xml = xml.replace(/(<table\b[^>]*>)/, `$1<autoFilter ref="${info.ref}"/>`);
    view.pkg.write_(part, xml);
    view.dirtyCells.clear();
    view.styledCells.clear();
    view._structuralDirty = true;
    view._rebuildDerivedState();
    return view;
  }, { parts, structural: true });
}

/* ── the words, for the spelling pass ──────────────────────────────────── */

/**
 * The text cells of the sheets asked for, row by row — words, not numbers,
 * dates or formulas (Excel does not check a formula's text either).
 */
export function sheetSegments(view, { sheets = [view.activeSheet] } = {}) {
  const out = [];
  for (const sheet of sheets) {
    for (const c of sheetCells(view, sheet)) {
      if (typeof c.input !== 'string' || c.input.startsWith('=') || !/\p{L}/u.test(c.input)) continue;
      out.push({ key: `c:${sheet}:${c.row}:${c.col}`, text: c.input.startsWith("'") ? c.input.slice(1) : c.input, prefix: c.input.startsWith("'") ? 1 : 0, where: { sheet, row: c.row, col: c.col, ref: refOf(c.row, c.col) } });
    }
  }
  return out;
}

export function sheetStart(view) {
  const a = view.selection?.active || { row: 0, col: 0 };
  return { key: `c:${view.activeSheet}:${a.row}:${a.col}`, offset: 0, sheet: view.activeSheet, row: a.row, col: a.col };
}

/** Replace spans of cells' text, one undo step per sheet touched. */
export function replaceSheetText(view, edits) {
  const bySheet = new Map();
  for (const e of edits) {
    const [, sheet, row, col] = /^c:(.*):(\d+):(\d+)$/.exec(e.key) || [];
    if (sheet == null) continue;
    if (!bySheet.has(sheet)) bySheet.set(sheet, new Map());
    const cells = bySheet.get(sheet);
    const k = `${row}:${col}`;
    if (!cells.has(k)) cells.set(k, { row: Number(row), col: Number(col), edits: [] });
    cells.get(k).edits.push(e);
  }
  const back = view.activeSheet;
  let changed = 0;
  for (const [sheet, cells] of bySheet) {
    if (view.activeSheet !== sheet) view.selectSheet(sheet);
    const list = [...cells.values()];
    view._edit('spelling', null, list.map(({ row, col }) => ({ row, col })), () => {
      for (const { row, col, edits: es } of list) {
        const input = String(view.calc.getInput(sheet, row, col) ?? '');
        const prefix = input.startsWith("'") ? 1 : 0;
        let next = input;
        for (const e of [...es].sort((x, y) => y.from - x.from)) {
          next = next.slice(0, prefix + e.from) + e.text + next.slice(prefix + e.to);
        }
        if (next !== input) {
          view._setCell(row, col, next);
          changed += 1;
        }
      }
      return view;
    });
  }
  if (view.activeSheet !== back) view.selectSheet(back);
  return changed;
}
