// OpenDocument — .odt, .ods, .odp.
//
// Read here, written in odf-write.js. A file that came in as ODF goes out as
// ODF holding what this suite models — the text, structure, tables, values,
// formulas and slide content below. The preserving trick that makes our OOXML
// safe (rewrite only what you touched) does not transfer, because ODF keeps
// one document in one part rather than spreading it across many, so what is
// not modelled is not kept; the reader and the writer agree on the subset.
//
// What we do promise: the text, the structure, the tables, the sheet values and
// the slide content all arrive intact, which is what someone opening a
// colleague's .ods actually needs.

import { readZip } from '@rutba/ooxml/zip';
import { parse, all, kids, textOf, first } from './xml.js';
import { formulaFromOdf } from './odf-write.js';
import { enhancedFigures } from './odf-geometry.js';

const dec = new TextDecoder('utf-8');

function entriesOf(bytes) {
  const buf = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes);
  const { entries } = readZip(buf);
  const map = new Map();
  for (const e of entries) map.set(e.name, e);
  return map;
}

function textPart(map, name) {
  const e = map.get(name);
  return e ? dec.decode(e.data) : '';
}

/** ODF hides the flavour in mimetype; the extension is only a hint. */
export function odfFlavour(bytes) {
  const map = entriesOf(bytes);
  const mime = textPart(map, 'mimetype').trim();
  if (mime.includes('opendocument.text')) return 'odt';
  if (mime.includes('opendocument.spreadsheet')) return 'ods';
  if (mime.includes('opendocument.presentation')) return 'odp';
  if (mime.includes('opendocument.graphics')) return 'odp';
  const content = textPart(map, 'content.xml');
  if (content.includes('<office:spreadsheet')) return 'ods';
  if (content.includes('<office:presentation')) return 'odp';
  if (content.includes('<office:text')) return 'odt';
  return null;
}

function repeatOf(attrs, key) {
  const n = Number(attrs[key] || 1);
  // A row repeated a million times is how ODF writes "the rest of the sheet is
  // empty". Materialising that is how a viewer dies on open.
  return Number.isFinite(n) ? Math.min(Math.max(1, n), 65536) : 1;
}

/** Inline text of a paragraph, with tabs, line breaks and repeated spaces. */
function inlineText(node) {
  let out = '';
  const visit = (n) => {
    for (const c of n.children) {
      if (typeof c === 'string') {
        out += c;
        continue;
      }
      if (c.name === 'text:s') out += ' '.repeat(Number(c.attrs['text:c'] || 1));
      else if (c.name === 'text:tab') out += '\t';
      else if (c.name === 'text:line-break') out += '\n';
      else visit(c);
    }
  };
  visit(node);
  return out;
}

function styleName(node) {
  return node.attrs['text:style-name'] || node.attrs['table:style-name'] || '';
}

/** A heading level, when the element is a heading. */
function outlineLevel(node) {
  const l = Number(node.attrs['text:outline-level'] || 0);
  return Number.isFinite(l) && l > 0 ? Math.min(l, 9) : 0;
}

function readTextBody(body) {
  /** @type {Array<object>} */
  const blocks = [];
  const visit = (node) => {
    for (const c of kids(node)) {
      if (c.name === 'text:h') {
        blocks.push({ type: 'heading', level: outlineLevel(c) || 1, text: inlineText(c), style: styleName(c) });
      } else if (c.name === 'text:p') {
        blocks.push({ type: 'paragraph', text: inlineText(c), style: styleName(c) });
      } else if (c.name === 'text:list') {
        const items = [];
        for (const li of all(c, 'text:list-item')) {
          const t = kids(li, 'text:p').map(inlineText).join('\n');
          items.push(t);
        }
        blocks.push({ type: 'list', ordered: /number|ordered/i.test(styleName(c)), items });
      } else if (c.name === 'table:table') {
        blocks.push(readTable(c));
      } else if (c.name === 'text:section' || c.name === 'office:text') {
        visit(c);
      } else if (c.name === 'draw:frame') {
        const img = first(c, 'draw:image');
        if (img) blocks.push({ type: 'image', href: img.attrs['xlink:href'] || '', width: c.attrs['svg:width'], height: c.attrs['svg:height'] });
        else blocks.push({ type: 'paragraph', text: inlineText(c), style: styleName(c) });
      }
    }
  };
  visit(body);
  return blocks;
}

/** What a sheet draws: frames (charts, pictures, words), shapes, lines, groups. */
const DRAWN = new Set(['draw:frame', 'draw:custom-shape', 'draw:rect', 'draw:ellipse', 'draw:circle', 'draw:line', 'draw:connector', 'draw:g']);

function readTable(table, rowHeights = new Map()) {
  const rows = [];
  const hiddenRows = [];
  // The drawings anchored in its cells, with the cell, and those on the sheet's page itself.
  const drawings = [];
  for (const shapes of kids(table, 'table:shapes')) for (const node of kids(shapes).filter((k) => DRAWN.has(k.name))) drawings.push({ row: null, col: null, node });
  const heights = {};
  for (const r of all(table, 'table:table-row')) {
    const rowRepeat = repeatOf(r.attrs, 'table:number-rows-repeated');
    const height = rowHeights.get(r.attrs['table:style-name']);
    if (height) for (let i = 0; i < rowRepeat && rows.length + i < 4096; i++) heights[rows.length + i] = height;
    // A row collapsed or filtered out is hidden; so is each it repeats.
    const rowHidden = r.attrs['table:visibility'] === 'collapse' || r.attrs['table:visibility'] === 'filter';
    const cells = [];
    // Empty cells not yet written out: a run of them is only built when something comes after it.
    let pending = 0;
    let runs = [];
    for (const c of kids(r).filter((k) => k.name === 'table:table-cell' || k.name === 'table:covered-table-cell')) {
      const repeat = repeatOf(c.attrs, 'table:number-columns-repeated');
      for (const node of kids(c).filter((k) => DRAWN.has(k.name))) drawings.push({ row: rows.length, col: cells.length + pending, node });
      const text = kids(c, 'text:p').map(inlineText).join('\n');
      const cell = {
        text,
        type: c.attrs['office:value-type'] || (text ? 'string' : null),
        value: c.attrs['office:value'] ?? c.attrs['office:date-value'] ?? c.attrs['office:time-value'] ?? c.attrs['office:boolean-value'] ?? null,
        // `of:=SUM([.A1:.B2];3)` in the file is `=SUM(A1:B2,3)` to the engine.
        formula: c.attrs['table:formula'] ? formulaFromOdf(c.attrs['table:formula']) : null,
        colspan: Number(c.attrs['table:number-columns-spanned'] || 1),
        rowspan: Number(c.attrs['table:number-rows-spanned'] || 1),
        // The automatic cell style, which names the data style the number wears.
        style: c.attrs['table:style-name'] || null,
        covered: c.name === 'table:covered-table-cell',
      };
      if (!cell.text && cell.value == null && !cell.formula && !cell.covered && cell.colspan === 1 && cell.rowspan === 1) {
        pending += repeat;
        runs.push([cell, repeat]);
        continue;
      }
      // Each run as it was, its own style kept.
      for (const [empty, n] of runs) for (let i = 0; i < n; i++) cells.push({ ...empty });
      pending = 0;
      runs = [];
      for (let i = 0; i < repeat; i++) cells.push(i === 0 ? cell : { ...cell });
    }
    // Trim the trailing run of empty cells ODF writes to pad the row.
    while (cells.length && !cells[cells.length - 1].text && cells[cells.length - 1].value == null) cells.pop();
    for (let i = 0; i < rowRepeat; i++) {
      if (rowHidden && rows.length < 200000) hiddenRows.push(rows.length);
      rows.push(cells.map((c) => ({ ...c })));
    }
    if (rows.length > 200000) break;
  }
  while (rows.length && rows[rows.length - 1].length === 0) rows.pop();
  return { type: 'table', name: table.attrs['table:name'] || '', rows, hiddenRows: hiddenRows.filter((r) => r < rows.length), drawings, heights };
}

const colName = (i) => {
  let s = '';
  for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(64 + ((n - 1) % 26) + 1) + s;
  return s;
};

/**
 * The sheets, each with the ranges its spanned cells merge and the number
 * format each styled cell wears — `formats` maps an automatic cell style's
 * name to a format code, see readDataStyles.
 */
function readSheets(body, formats = new Map(), columnWidths = new Map(), { rowHeights = new Map(), styles = () => ({}) } = {}) {
  return all(body, 'table:table').map((t) => {
    const table = readTable(t, rowHeights);
    // The columns, in order: which are hidden, and each one's width (px) from its style.
    const hiddenCols = [];
    const widths = {};
    let col = 0;
    for (const c of all(t, 'table:table-column')) {
      const repeat = repeatOf(c.attrs, 'table:number-columns-repeated');
      const width = columnWidths.get(c.attrs['table:style-name']);
      for (let i = 0; i < repeat && col < 16384; i++, col++) {
        if (c.attrs['table:visibility'] === 'collapse' || c.attrs['table:visibility'] === 'filter') hiddenCols.push(col);
        if (width && col < 1024) widths[col] = width;
      }
    }
    const merges = [];
    const cellFormats = {};
    table.rows.forEach((row, r) => {
      row.forEach((cell, c) => {
        if (cell.covered) return;
        if (cell.colspan > 1 || cell.rowspan > 1) merges.push(`${colName(c)}${r + 1}:${colName(c + cell.colspan - 1)}${r + cell.rowspan}`);
        const fmt = cell.style && formats.get(cell.style);
        if (fmt) cellFormats[`${colName(c)}${r + 1}`] = fmt;
      });
    });
    // Its drawings, each with the cell it is anchored in (null on the page) and its offset from there.
    const drawings = [];
    for (const { row, col, node } of table.drawings) {
      for (const d of readDrawings({ children: [node] }, styles)) drawings.push({ ...d, anchor: row == null ? null : { row, col } });
    }
    return { name: table.name, rows: table.rows, merges, formats: cellFormats, hiddenRows: table.hiddenRows, hiddenCols: hiddenCols.filter((c) => c < 1024), widths, heights: table.heights, drawings };
  });
}

/**
 * Named ranges: `<table:named-range table:name="GrandTotal"
 * table:cell-range-address="Sales.$F$14"/>` is `GrandTotal = Sales!$F$14`
 * to the engine, and a formula spells it `$$GrandTotal` (formulaFromOdf).
 */
function readNames(body) {
  const out = [];
  for (const n of all(body, 'table:named-range')) {
    const name = n.attrs['table:name'];
    const addr = n.attrs['table:cell-range-address'] || '';
    if (!name || !addr) continue;
    const parts = addr.split(':').map((p) => {
      const dot = p.lastIndexOf('.');
      return { sheet: dot > 0 ? p.slice(0, dot).replace(/^\$/, '') : '', cell: p.slice(dot + 1) };
    });
    const sheet = parts[0].sheet;
    const quoted = sheet && /[^A-Za-z0-9_]/.test(sheet.replace(/^'|'$/g, '')) ? `'${sheet.replace(/^'|'$/g, '')}'` : sheet;
    out.push({ name, ref: (quoted ? `${quoted}!` : '') + parts.map((p) => p.cell).join(':') });
  }
  return out;
}

/**
 * ODF number formats as the format codes the grid speaks. A data style is a
 * small tree — `<number:number number:decimal-places="2" number:grouping="true"/>`
 * between `<number:text>` pieces, a currency symbol, date parts — and the
 * common ones map onto Excel's codes exactly; anything stranger falls back
 * to the closest plain number. Answers a map of automatic CELL style name →
 * format code, since that is what a cell names.
 */
function readDataStyles(roots) {
  const codes = new Map();
  const escText = (s) => (s ? (/^[\s%/:\-.,]+$/.test(s) ? s : `"${s.replace(/"/g, '""')}"`) : '');
  const numberCode = (n) => {
    if (!n) return 'General';
    const dec = Number(n.attrs['number:decimal-places'] ?? n.attrs['number:min-decimal-places'] ?? 0);
    const grouping = n.attrs['number:grouping'] === 'true';
    const minInt = Number(n.attrs['number:min-integer-digits'] ?? 1);
    let code = (grouping ? '#,##' : '') + (minInt ? '0' : '#');
    if (dec > 0) code += '.' + '0'.repeat(dec);
    return code;
  };
  const dateCode = (style) => {
    let code = '';
    for (const k of kids(style)) {
      const long = k.attrs['number:style'] === 'long';
      if (k.name === 'number:day') code += long ? 'dd' : 'd';
      else if (k.name === 'number:month') code += k.attrs['number:textual'] === 'true' ? (long ? 'mmmm' : 'mmm') : long ? 'mm' : 'm';
      else if (k.name === 'number:year') code += long ? 'yyyy' : 'yy';
      else if (k.name === 'number:day-of-week') code += long ? 'dddd' : 'ddd';
      else if (k.name === 'number:hours') code += long ? 'hh' : 'h';
      else if (k.name === 'number:minutes') code += long ? 'mm' : 'm';
      else if (k.name === 'number:seconds') code += long ? 'ss' : 's';
      else if (k.name === 'number:am-pm') code += ' AM/PM';
      else if (k.name === 'number:text') code += escText(textOf(k));
    }
    return code || 'General';
  };
  for (const root of roots.filter(Boolean)) {
    for (const style of all(root, 'number:number-style')) {
      const n = first(style, 'number:number');
      const sci = first(style, 'number:scientific-number');
      const frac = first(style, 'number:fraction');
      const texts = kids(style, 'number:text').map(textOf);
      let code;
      if (sci) code = numberCode(sci) + 'E+' + '0'.repeat(Number(sci.attrs['number:min-exponent-digits'] || 2));
      else if (frac) code = `# ${'?'.repeat(Number(frac.attrs['number:min-numerator-digits'] || 1))}/${'?'.repeat(Number(frac.attrs['number:min-denominator-digits'] || 1))}`;
      else code = numberCode(n);
      const suffix = texts.length && !n ? escText(texts[texts.length - 1]) : '';
      codes.set(style.attrs['style:name'], code + suffix);
    }
    for (const style of all(root, 'number:percentage-style')) codes.set(style.attrs['style:name'], numberCode(first(style, 'number:number')) + '%');
    for (const style of all(root, 'number:currency-style')) {
      const sym = first(style, 'number:currency-symbol');
      const n = first(style, 'number:number');
      const symbol = sym ? textOf(sym) : '';
      const before = sym && n && kids(style).indexOf(sym) < kids(style).indexOf(n);
      const body = numberCode(n);
      codes.set(style.attrs['style:name'], before ? `"${symbol}"${body}` : `${body}"${symbol}"`);
    }
    for (const style of all(root, 'number:date-style')) codes.set(style.attrs['style:name'], dateCode(style));
    for (const style of all(root, 'number:time-style')) codes.set(style.attrs['style:name'], dateCode(style));
    for (const style of all(root, 'number:boolean-style')) codes.set(style.attrs['style:name'], 'General');
  }
  // Cell styles name their data style; the map answers for the cell style.
  const byCell = new Map();
  for (const root of roots.filter(Boolean)) {
    for (const st of all(root, 'style:style')) {
      if (st.attrs['style:family'] !== 'table-cell') continue;
      const data = st.attrs['style:data-style-name'];
      if (data && codes.has(data) && codes.get(data) !== 'General') byCell.set(st.attrs['style:name'], codes.get(data));
    }
  }
  return byCell;
}

/** An ODF length — "2.5cm", "12mm", "1in", "72pt", "96px" — in CSS pixels; null for none. */
export function lengthPx(value) {
  const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|pc|px)?$/.exec(String(value ?? '').trim());
  if (!m) return null;
  return Number(m[1]) * ({ cm: 96 / 2.54, mm: 96 / 25.4, in: 96, pt: 96 / 72, pc: 16, px: 1 }[m[2] || 'px']);
}

/** An ODF angle as degrees: a bare number is tenths of a degree (ODF 1.2), else its unit says. */
export function angleDeg(value) {
  const m = /^\s*(-?[\d.]+)\s*(deg|rad|grad)?\s*$/.exec(String(value ?? ''));
  if (!m) return 0;
  const n = Number(m[1]);
  return m[2] === 'deg' ? n : m[2] === 'rad' ? (n * 180) / Math.PI : m[2] === 'grad' ? n * 0.9 : n / 10;
}

const pct = (v, fallback) => {
  const m = /^\s*(-?[\d.]+)%\s*$/.exec(String(v ?? ''));
  return m ? Number(m[1]) / 100 : fallback;
};

/** A colour made darker by an intensity (0–1), as ODF's start- and end-intensity scale a gradient's colours. */
function intensify(hex, by) {
  const m = /^#?([0-9a-f]{6})$/i.exec(String(hex || ''));
  if (!m || by >= 1) return m ? '#' + m[1].toUpperCase() : '#000000';
  const n = parseInt(m[1], 16);
  const ch = (shift) => Math.round(((n >> shift) & 255) * Math.max(0, by)).toString(16).padStart(2, '0');
  return ('#' + ch(16) + ch(8) + ch(0)).toUpperCase();
}

/**
 * The named gradients (`<draw:gradient>` in the styles), as a drawing
 * fills with one: its style — linear, axial, radial, ellipsoid, square,
 * rectangular — its two colours at their intensities, its angle in degrees
 * (ODF's: 0 runs top to bottom, counter-clockwise from there), how much of
 * it is border, and the centre a radial one runs out from.
 */
function readGradients(roots) {
  const out = new Map();
  for (const root of roots.filter(Boolean)) {
    for (const g of all(root, 'draw:gradient')) {
      const a = g.attrs;
      out.set(a['draw:name'], {
        style: a['draw:style'] || 'linear',
        start: intensify(a['draw:start-color'] || '#000000', pct(a['draw:start-intensity'], 1)),
        end: intensify(a['draw:end-color'] || '#FFFFFF', pct(a['draw:end-intensity'], 1)),
        angle: angleDeg(a['draw:angle']),
        border: Math.max(0, Math.min(0.99, pct(a['draw:border'], 0))),
        cx: pct(a['draw:cx'], 0.5),
        cy: pct(a['draw:cy'], 0.5),
      });
    }
  }
  return out;
}

/**
 * The graphic styles a drawing names — its fill, its outline — from the
 * automatic styles and the named ones, a style's parent filling in what it
 * does not say; a drawing page's background with them, and the gradient a
 * fill names, read whole.
 */
function readGraphicStyles(roots) {
  const raw = new Map();
  const gradients = readGradients(roots);
  for (const root of roots.filter(Boolean)) {
    for (const st of all(root, 'style:style')) {
      const family = st.attrs['style:family'];
      if (family !== 'graphic' && family !== 'presentation' && family !== 'drawing-page') continue;
      const g = first(st, family === 'drawing-page' ? 'style:drawing-page-properties' : 'style:graphic-properties');
      const a = g ? g.attrs : {};
      raw.set(st.attrs['style:name'], {
        parent: st.attrs['style:parent-style-name'] || null,
        fill: a['draw:fill'] ?? null,
        fillColor: a['draw:fill-color'] ?? null,
        gradient: a['draw:fill-gradient-name'] ? gradients.get(a['draw:fill-gradient-name']) || null : null,
        stroke: a['draw:stroke'] ?? null,
        strokeColor: a['svg:stroke-color'] ?? null,
        strokeWidth: a['svg:stroke-width'] ?? null,
      });
    }
  }
  const resolve = (name, depth = 0) => {
    const own = raw.get(name);
    if (!own) return {};
    const up = own.parent && depth < 8 ? resolve(own.parent, depth + 1) : {};
    const out = { ...up };
    for (const k of ['fill', 'fillColor', 'gradient', 'stroke', 'strokeColor', 'strokeWidth']) if (own[k] !== null) out[k] = own[k];
    return out;
  };
  return (name) => resolve(name);
}

/** The page size a presentation is laid out on, from its page layout — null when it says none. */
function readPageSize(stylesRoot) {
  if (!stylesRoot) return null;
  for (const props of all(stylesRoot, 'style:page-layout-properties')) {
    const width = lengthPx(props.attrs['fo:page-width']);
    const height = lengthPx(props.attrs['fo:page-height']);
    if (width && height) return { width, height };
  }
  return null;
}

/**
 * A drawing's box, in pixels, from its own svg:x, svg:y, svg:width and
 * svg:height — or, for one placed by draw:transform (a turned drawing),
 * where the transform takes its centre, and the turn in degrees clockwise.
 * ODF applies the transform's steps as they are written, left to right,
 * and turns counter-clockwise.
 */
const boxOf = (node) => {
  const w = lengthPx(node.attrs['svg:width']) ?? 0;
  const h = lengthPx(node.attrs['svg:height']) ?? 0;
  const transform = node.attrs['draw:transform'];
  if (!transform) return { x: lengthPx(node.attrs['svg:x']) ?? 0, y: lengthPx(node.attrs['svg:y']) ?? 0, w, h };
  let p = [(lengthPx(node.attrs['svg:x']) ?? 0) + w / 2, (lengthPx(node.attrs['svg:y']) ?? 0) + h / 2];
  let turn = 0;
  for (const m of String(transform).matchAll(/(translate|rotate|scale)\s*\(([^)]*)\)/g)) {
    const args = m[2].trim().split(/[\s,]+/);
    if (m[1] === 'translate') p = [p[0] + (lengthPx(args[0]) ?? 0), p[1] + (lengthPx(args[1] ?? '0') ?? 0)];
    else if (m[1] === 'rotate') {
      const a = Number(args[0]) || 0;
      p = [p[0] * Math.cos(a) + p[1] * Math.sin(a), -p[0] * Math.sin(a) + p[1] * Math.cos(a)];
      turn += a;
    } else {
      const sx = Number(args[0]) || 1;
      const sy = Number(args[1] ?? args[0]) || 1;
      p = [p[0] * sx, p[1] * sy];
    }
  }
  const rotation = ((((-turn * 180) / Math.PI) % 360) + 360) % 360;
  return { x: p[0] - w / 2, y: p[1] - h / 2, w, h, ...(Math.abs(rotation) > 0.01 && Math.abs(rotation - 360) > 0.01 ? { rotation: Math.round(rotation * 100) / 100 } : {}) };
};

/** The paragraphs a text box or shape holds, lists' items among them. */
const paragraphsOf = (node) => {
  const out = [];
  const walk = (n) => {
    for (const c of n.children || []) {
      if (typeof c === 'string') continue;
      if (c.name === 'text:p' || c.name === 'text:h') out.push(inlineText(c));
      else if (c.name === 'text:list' || c.name === 'text:list-item' || c.name === 'text:list-header') walk(c);
    }
  };
  walk(node);
  return out;
};

/**
 * A fill as the style says it: 'none', a gradient ({ gradient }), a hex,
 * or null for the default.
 */
const fillOf = (st) => (st.fill === 'none' ? 'none' : st.fill === 'gradient' && st.gradient ? { gradient: st.gradient } : st.fillColor || null);

/** A shape's fill and outline, as the style it names says: a hex, a gradient, 'none', or null for the default. */
const lookOf = (node, styles) => {
  const st = styles(node.attrs['draw:style-name'] || node.attrs['presentation:style-name']);
  return {
    fill: fillOf(st),
    line: st.stroke === 'none' ? 'none' : st.strokeColor || null,
    lineWidth: lengthPx(st.strokeWidth),
  };
};

/**
 * One page's drawings, in drawing order: frames holding a picture, words
 * or a table; custom shapes, rectangles and ellipses with their fill,
 * outline and words; lines; a group's members in its place.
 */
function readDrawings(page, styles) {
  const shapes = [];
  const visit = (parent) => {
    for (const c of parent.children || []) {
      if (typeof c === 'string') continue;
      if (c.name === 'draw:g') { visit(c); continue; }
      if (c.name === 'draw:frame') {
        const img = first(c, 'draw:image');
        const box = first(c, 'draw:text-box');
        const table = first(c, 'table:table');
        const name = c.attrs['draw:name'] || null;
        const object = first(c, 'draw:object');
        if (object) {
          shapes.push({ type: 'object', name, href: object.attrs['xlink:href'] || '', ...boxOf(c) });
        } else if (table) {
          const rows = all(table, 'table:table-row').map((row) => kids(row, 'table:table-cell').map((cell) => paragraphsOf(cell).join('\n')));
          shapes.push({ type: 'table', name, rows, ...boxOf(c) });
        } else if (img) {
          shapes.push({ type: 'image', name, href: img.attrs['xlink:href'] || '', ...boxOf(c) });
        } else if (box) {
          shapes.push({ type: 'text', name, role: c.attrs['presentation:class'] || null, paragraphs: paragraphsOf(box).filter((p, i, list) => p || list.length === 1), ...boxOf(c) });
        }
        continue;
      }
      if (c.name === 'draw:custom-shape' || c.name === 'draw:rect' || c.name === 'draw:ellipse' || c.name === 'draw:circle') {
        const enhanced = c.name === 'draw:custom-shape' ? first(c, 'draw:enhanced-geometry') : null;
        const geometry = c.name === 'draw:custom-shape' ? (enhanced?.attrs['draw:type'] || 'rectangle') : c.name === 'draw:rect' ? 'rectangle' : 'ellipse';
        const box = boxOf(c);
        // A custom shape's own outline, worked out from its path and equations.
        const figures = enhanced ? enhancedFigures(enhanced, { width: box.w, height: box.h }) : null;
        shapes.push({ type: 'shape', name: c.attrs['draw:name'] || null, geometry, ...box, ...lookOf(c, styles), paragraphs: paragraphsOf(c), ...(figures ? { figures } : {}) });
        continue;
      }
      // A connector is drawn as the line between its ends.
      if (c.name === 'draw:line' || c.name === 'draw:connector') {
        const x1 = lengthPx(c.attrs['svg:x1']) ?? 0;
        const y1 = lengthPx(c.attrs['svg:y1']) ?? 0;
        const x2 = lengthPx(c.attrs['svg:x2']) ?? 0;
        const y2 = lengthPx(c.attrs['svg:y2']) ?? 0;
        shapes.push({ type: 'line', name: c.attrs['draw:name'] || null, x1, y1, x2, y2, ...lookOf(c, styles) });
      }
    }
  };
  visit(page);
  return shapes;
}

/** Each row style's height, in pixels. */
function readRowHeights(roots) {
  const out = new Map();
  for (const root of roots.filter(Boolean)) {
    for (const st of all(root, 'style:style')) {
      if (st.attrs['style:family'] !== 'table-row') continue;
      const height = lengthPx(first(st, 'style:table-row-properties')?.attrs['style:row-height']);
      if (height) out.set(st.attrs['style:name'], height);
    }
  }
  return out;
}

/**
 * A cell range as ODF writes one — "Sales.$B$2:.$B$13", "'My data'.A1" —
 * as { sheet, top, left, bottom, right } (0-based); null when it is not one.
 */
export function odfRange(text) {
  // The first of a list of ranges, a quoted sheet name's spaces and colons its own.
  const head = /^(?:'(?:[^']|'')*'|[^\s'])+/.exec(String(text || '').trim())?.[0] || '';
  const parts = head.split(/:(?=(?:[^']*'[^']*')*[^']*$)/);
  const one = (p, sheet) => {
    const m = /^\$?(?:'((?:[^']|'')*)'|([^.']*))\.\$?([A-Za-z]+)\$?(\d+)$/.exec(p);
    if (!m) return null;
    const name = m[1] != null ? m[1].replace(/''/g, "'") : m[2];
    const col = m[3].toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    return { sheet: name || sheet, row: Number(m[4]) - 1, col };
  };
  const a = one(parts[0], null);
  if (!a) return null;
  const b = parts[1] ? one(parts[1], a.sheet) : a;
  if (!b) return null;
  return { sheet: a.sheet, top: Math.min(a.row, b.row), left: Math.min(a.col, b.col), bottom: Math.max(a.row, b.row), right: Math.max(a.col, b.col) };
}

/** The A1 reference a range is in a workbook: Sales!$B$2:$B$13, the sheet quoted when it must be. */
export function rangeRef(r) {
  const name = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(r.sheet || '') ? r.sheet : `'${String(r.sheet || '').replace(/'/g, "''")}'`;
  const cell = (row, col) => `$${colName(col)}$${row + 1}`;
  return r.top === r.bottom && r.left === r.right ? `${name}!${cell(r.top, r.left)}` : `${name}!${cell(r.top, r.left)}:${cell(r.bottom, r.right)}`;
}

/** The cells of a range, row by row then column by column, from `cellAt(sheet, row, col)`. */
const cellsOf = (r, cellAt) => {
  const out = [];
  for (let row = r.top; row <= r.bottom && out.length < 4096; row++) for (let col = r.left; col <= r.right && out.length < 4096; col++) out.push(cellAt(r.sheet, row, col));
  return out;
};

/**
 * An embedded chart — "Object 1/content.xml" — as a chart a workbook or a
 * deck writes: its kind (bar charts upright are columns), its title, its
 * categories and each series' name and values, with the ranges they come
 * from. `cellAt(sheet, row, col)` answers a cell, from the sheets the chart
 * plots or from the chart's own local table; null when the part is no chart.
 */
function readChartObject(map, href, cellAt) {
  const dir = String(href || '').replace(/^\.\//, '').replace(/\/$/, '');
  const xml = textPart(map, `${dir}/content.xml`);
  if (!xml) return null;
  const root = parse(xml);
  const chart = first(root, 'chart:chart');
  if (!chart) return null;
  const props = new Map(all(root, 'style:style').map((s) => [s.attrs['style:name'], first(s, 'style:chart-properties')?.attrs || {}]));
  const plot = first(chart, 'chart:plot-area');
  const plotProps = props.get(plot?.attrs['chart:style-name']) || {};
  const cls = String(chart.attrs['chart:class'] || '').replace(/^chart:/, '');
  const kind = cls === 'bar' ? (plotProps['chart:vertical'] === 'true' ? 'bar' : 'column')
    : cls === 'line' || cls === 'radar' || cls === 'filled-radar' ? 'line'
      : cls === 'area' ? 'area' : cls === 'circle' ? 'pie' : cls === 'ring' ? 'doughnut' : cls === 'scatter' ? 'scatter' : 'column';
  const title = first(chart, 'chart:title');
  // The chart's own table of data, for a chart with no sheet behind it.
  const local = first(root, 'table:table');
  const localRows = local ? readTable(local).rows : [];
  const localName = local?.attrs['table:name'] || 'local-table';
  const lookup = (sheet, row, col) => (sheet === localName || !cellAt ? localRows[row]?.[col] ?? null : cellAt(sheet, row, col));
  const numberOf = (cell) => {
    const n = Number(cell?.value ?? cell?.text);
    return cell && (cell.value != null || /^-?[\d.]+$/.test(String(cell.text || '').trim())) && Number.isFinite(n) ? n : null;
  };
  const textOfCell = (cell) => (cell ? String(cell.text ?? cell.value ?? '') : '');
  const catRange = odfRange(first(plot, 'chart:categories')?.attrs['table:cell-range-address']);
  const series = all(plot || chart, 'chart:series').map((s, i) => {
    const values = odfRange(s.attrs['chart:values-cell-range-address']);
    const label = odfRange(s.attrs['chart:label-cell-address']);
    return {
      name: label ? textOfCell(lookup(label.sheet, label.top, label.left)) : `Series ${i + 1}`,
      nameRef: label && label.sheet !== localName ? rangeRef(label) : null,
      ref: values && values.sheet !== localName ? rangeRef(values) : null,
      values: values ? cellsOf(values, lookup).map(numberOf) : [],
    };
  }).filter((s) => s.values.length);
  if (!series.length) return null;
  return {
    kind,
    // Stacked, or stacked to a hundred, as its plot area's style says.
    ...(plotProps['chart:percentage'] === 'true' ? { grouping: 'percentStacked' } : plotProps['chart:stacked'] === 'true' ? { grouping: 'stacked' } : {}),
    title: title ? all(title, 'text:p').map(textOf).join(' ').trim() || null : null,
    categories: catRange ? { ref: catRange.sheet !== localName ? rangeRef(catRange) : null, values: cellsOf(catRange, lookup).map(textOfCell) } : null,
    series,
  };
}

/** Each column style's width, in pixels. */
function readColumnWidths(roots) {
  const out = new Map();
  for (const root of roots.filter(Boolean)) {
    for (const st of all(root, 'style:style')) {
      if (st.attrs['style:family'] !== 'table-column') continue;
      const width = lengthPx(first(st, 'style:table-column-properties')?.attrs['style:column-width']);
      if (width) out.set(st.attrs['style:name'], width);
    }
  }
  return out;
}

/**
 * The panes each sheet keeps frozen, from settings.xml: a split mode of 2
 * is a freeze, its position the columns (horizontally) or rows
 * (vertically) held. `{ rows, cols }` by sheet name.
 */
function readFrozenPanes(xml) {
  const out = new Map();
  if (!xml) return out;
  const root = parse(xml);
  for (const map of all(root, 'config:config-item-map-named')) {
    if (map.attrs['config:name'] !== 'Tables') continue;
    for (const entry of kids(map, 'config:config-item-map-entry')) {
      const items = new Map(kids(entry, 'config:config-item').map((i) => [i.attrs['config:name'], textOf(i)]));
      const cols = items.get('HorizontalSplitMode') === '2' ? Number(items.get('HorizontalSplitPosition') || 0) : 0;
      const rows = items.get('VerticalSplitMode') === '2' ? Number(items.get('VerticalSplitPosition') || 0) : 0;
      if (rows > 0 || cols > 0) out.set(entry.attrs['config:name'], { rows, cols });
    }
  }
  return out;
}

/**
 * The slides: each page's drawings and notes, and its background — its own
 * page style's fill, else its master page's — as a colour or a gradient.
 */
function readSlides(body, styles = () => ({}), masters = new Map()) {
  const backgroundOf = (styleName) => {
    if (!styleName) return null;
    const st = styles(styleName);
    if (st.fill === 'solid' && st.fillColor) return { colour: st.fillColor };
    if (st.fill === 'gradient' && st.gradient) return { gradient: st.gradient };
    return null;
  };
  return all(body, 'draw:page').map((page, i) => {
    const notes = first(page, 'presentation:notes');
    const background = backgroundOf(page.attrs['draw:style-name']) || backgroundOf(masters.get(page.attrs['draw:master-page-name']));
    return {
      index: i,
      name: page.attrs['draw:name'] || `Slide ${i + 1}`,
      shapes: readDrawings(page, styles),
      notes: notes ? all(notes, 'text:p').map(textOf).join('\n') : '',
      ...(background ? { background } : {}),
    };
  });
}

/** Each master page's drawing-page style, by the master's name. */
function readMasterPages(stylesRoot) {
  const out = new Map();
  if (!stylesRoot) return out;
  for (const m of all(stylesRoot, 'style:master-page')) if (m.attrs['draw:style-name']) out.set(m.attrs['style:name'], m.attrs['draw:style-name']);
  return out;
}

function readMeta(map) {
  const xml = textPart(map, 'meta.xml');
  if (!xml) return {};
  const root = parse(xml);
  const pick = (name) => textOf(first(root, name)) || undefined;
  return {
    title: pick('dc:title'),
    subject: pick('dc:subject'),
    creator: pick('meta:initial-creator') || pick('dc:creator'),
    created: pick('meta:creation-date'),
    modified: pick('dc:date'),
    generator: pick('meta:generator'),
    keywords: pick('meta:keyword'),
  };
}

/**
 * Read an ODF file into the neutral model the apps render.
 * @param {Uint8Array|Buffer} bytes
 * @returns {{ flavour: 'odt'|'ods'|'odp', meta: object, blocks?: object[], sheets?: object[], slides?: object[], images: Map<string, Uint8Array> }}
 */
export function readOdf(bytes) {
  const map = entriesOf(bytes);
  const flavour = odfFlavour(bytes);
  if (!flavour) throw new Error('not an OpenDocument file');
  const content = textPart(map, 'content.xml');
  if (!content) throw new Error('OpenDocument file has no content.xml');
  const root = parse(content);
  const body = first(root, 'office:body') || root;

  const images = new Map();
  for (const [name, entry] of map) {
    if (name.startsWith('Pictures/') || name.startsWith('media/')) {
      images.set(name, new Uint8Array(entry.data));
    }
  }

  const out = { flavour, meta: readMeta(map), images };
  if (flavour === 'odt') out.blocks = readTextBody(first(body, 'office:text') || body);
  else if (flavour === 'ods') {
    // Data styles live in content.xml's automatic styles or in styles.xml;
    // both are read, and a cell's style name resolves through either.
    const stylesXml = textPart(map, 'styles.xml');
    const formats = readDataStyles([root, stylesXml ? parse(stylesXml) : null]);
    const stylesRoot = stylesXml ? parse(stylesXml) : null;
    out.sheets = readSheets(body, formats, readColumnWidths([root, stylesRoot]), { rowHeights: readRowHeights([root, stylesRoot]), styles: readGraphicStyles([stylesRoot, root]) });
    out.names = readNames(body);
    // Each sheet's charts, read from their own parts with the cells they plot.
    const byName = new Map(out.sheets.map((s) => [s.name, s]));
    const cellAt = (sheet, row, col) => byName.get(sheet)?.rows[row]?.[col] ?? null;
    for (const sheet of out.sheets) {
      sheet.drawings = sheet.drawings.flatMap((d) => {
        if (d.type !== 'object') return [d];
        const chart = readChartObject(map, d.href, cellAt);
        return chart ? [{ ...d, type: 'chart', chart }] : [];
      });
    }
    // Frozen panes live in settings.xml, by sheet.
    const frozen = readFrozenPanes(textPart(map, 'settings.xml'));
    for (const sheet of out.sheets) if (frozen.has(sheet.name)) sheet.frozen = frozen.get(sheet.name);
  } else {
    // The drawings' styles, from content.xml's own and styles.xml's named
    // ones, and the page every position is measured on.
    const stylesXml = textPart(map, 'styles.xml');
    const stylesRoot = stylesXml ? parse(stylesXml) : null;
    out.slides = readSlides(body, readGraphicStyles([stylesRoot, root]), readMasterPages(stylesRoot));
    // A slide's charts, each from its own part and its own table of data.
    for (const slide of out.slides) {
      slide.shapes = slide.shapes.flatMap((d) => {
        if (d.type !== 'object') return [d];
        const chart = readChartObject(map, d.href, null);
        return chart ? [{ ...d, type: 'chart', chart }] : [];
      });
    }
    out.size = readPageSize(stylesRoot);
  }
  return out;
}

export default { readOdf, odfFlavour };
