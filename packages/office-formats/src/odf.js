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
        value: c.attrs['office:value'] ?? c.attrs['office:date-value'] ?? c.attrs['office:boolean-value'] ?? null,
        // `of:=SUM([.A1:.B2];3)` in the file is `=SUM(A1:B2,3)` to the engine.
        formula: c.attrs['table:formula'] ? formulaFromOdf(c.attrs['table:formula']) : null,
        colspan: Number(c.attrs['table:number-columns-spanned'] || 1),
        rowspan: Number(c.attrs['table:number-rows-spanned'] || 1),
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

function readSheets(body) {
  return all(body, 'table:table').map((t) => {
    const table = readTable(t);
    return { name: table.name, rows: table.rows };
  });
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
  else if (flavour === 'ods') out.sheets = readSheets(body);
  else out.slides = readSlides(body);
  return out;
}

export default { readOdf, odfFlavour };
