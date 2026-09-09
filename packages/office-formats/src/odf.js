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

function readTable(table) {
  const rows = [];
  for (const r of all(table, 'table:table-row')) {
    const rowRepeat = repeatOf(r.attrs, 'table:number-rows-repeated');
    const cells = [];
    for (const c of kids(r).filter((k) => k.name === 'table:table-cell' || k.name === 'table:covered-table-cell')) {
      const repeat = repeatOf(c.attrs, 'table:number-columns-repeated');
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
      for (let i = 0; i < repeat; i++) cells.push(i === 0 ? cell : { ...cell });
    }
    // Trim the trailing run of empty cells ODF writes to pad the row.
    while (cells.length && !cells[cells.length - 1].text && cells[cells.length - 1].value == null) cells.pop();
    for (let i = 0; i < rowRepeat; i++) rows.push(cells.map((c) => ({ ...c })));
    if (rows.length > 200000) break;
  }
  while (rows.length && rows[rows.length - 1].length === 0) rows.pop();
  return { type: 'table', name: table.attrs['table:name'] || '', rows };
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
function readSheets(body, formats = new Map()) {
  return all(body, 'table:table').map((t) => {
    const table = readTable(t);
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
    return { name: table.name, rows: table.rows, merges, formats: cellFormats };
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

function readSlides(body) {
  return all(body, 'draw:page').map((page, i) => {
    const shapes = [];
    for (const frame of all(page, 'draw:frame')) {
      const img = first(frame, 'draw:image');
      const box = first(frame, 'draw:text-box');
      if (img) {
        shapes.push({ type: 'image', href: img.attrs['xlink:href'] || '', x: frame.attrs['svg:x'], y: frame.attrs['svg:y'], w: frame.attrs['svg:width'], h: frame.attrs['svg:height'] });
      } else if (box) {
        const paras = kids(box, 'text:p').map(inlineText).concat(kids(box, 'text:list').flatMap((l) => all(l, 'text:p').map(inlineText)));
        shapes.push({ type: 'text', paragraphs: paras.filter(Boolean), x: frame.attrs['svg:x'], y: frame.attrs['svg:y'], w: frame.attrs['svg:width'], h: frame.attrs['svg:height'] });
      }
    }
    const notes = first(page, 'presentation:notes');
    return {
      index: i,
      name: page.attrs['draw:name'] || `Slide ${i + 1}`,
      shapes,
      notes: notes ? all(notes, 'text:p').map(textOf).join('\n') : '',
    };
  });
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
    out.sheets = readSheets(body, formats);
    out.names = readNames(body);
  } else out.slides = readSlides(body);
  return out;
}

export default { readOdf, odfFlavour };
