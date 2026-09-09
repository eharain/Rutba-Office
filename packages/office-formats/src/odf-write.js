// OpenDocument, written — .odt, .ods, .odp.
//
// The installer tells Windows this suite is the editor of .odt, .ods and
// .odp, and for a year Ctrl+S on one refused. The reader (odf.js) takes the
// text, the structure, the tables, the sheet values and formulas and the slide
// content; this writer puts the same back, so a file that came in as
// OpenDocument goes out as OpenDocument, holding what this suite models.
//
// What is written is what the suite models, no more: a document's paragraphs,
// headings, lists and tables with their run formatting; a workbook's values,
// formulas and the value types a cell needs to stay a date or a percentage; a
// deck's text boxes and pictures, in place. LibreOffice opens all three.
//
// The archive follows the rule the format has and readers check: `mimetype`
// is the first entry and is stored uncompressed, so the first bytes of the
// file name its type.

import { writeZip, ZipEntry } from '@rutba/ooxml/zip';

const NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ' +
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" ' +
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ' +
  'xmlns:xlink="http://www.w3.org/1999/xlink" ' +
  'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" ' +
  'xmlns:dc="http://purl.org/dc/elements/1.1/" ' +
  'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" ' +
  'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" ' +
  'office:version="1.3"';

const MIME = {
  odt: 'application/vnd.oasis.opendocument.text',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
  odp: 'application/vnd.oasis.opendocument.presentation',
};

const esc = (s) =>
  String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

/** Pixels (at 96 to the inch) as the centimetres ODF measures in. */
const cm = (px) => `${(Number(px || 0) / 96 * 2.54).toFixed(3)}cm`;

/**
 * Text inside a paragraph: tabs, line breaks and runs of spaces have their own
 * elements, because a reader collapses whitespace in character data.
 */
function textXml(text) {
  let out = '';
  const s = String(text ?? '');
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '\t') out += '<text:tab/>';
    else if (ch === '\n') out += '<text:line-break/>';
    else if (ch === ' ' && s[i + 1] === ' ') {
      let n = 0;
      while (s[i + n] === ' ') n += 1;
      out += ' ' + (n > 1 ? `<text:s text:c="${n - 1}"/>` : '');
      i += n - 1;
    } else out += esc(ch);
  }
  return out;
}

/** A table of automatic styles, each written once and referred to by name. */
function styleTable(prefix) {
  const styles = [];
  const names = new Map();
  return {
    name(key, body) {
      if (!body) return null;
      if (names.has(key)) return names.get(key);
      const name = `${prefix}${styles.length + 1}`;
      names.set(key, name);
      styles.push(body(name));
      return name;
    },
    xml() {
      return styles.join('');
    },
  };
}

function runStyle(run, texts, fonts) {
  const props = [];
  if (run.bold) props.push('fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"');
  if (run.italic) props.push('fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"');
  if (run.underline) props.push('style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"');
  if (run.strike) props.push('style:text-line-through-style="solid"');
  if (run.size) props.push(`fo:font-size="${Number(run.size)}pt" style:font-size-asian="${Number(run.size)}pt" style:font-size-complex="${Number(run.size)}pt"`);
  const colour = run.colour || run.color;
  if (colour && /^#?[0-9a-f]{6}$/i.test(colour)) props.push(`fo:color="#${colour.replace('#', '').toLowerCase()}"`);
  const font = run.font || run.fontName;
  if (font) {
    fonts.add(font);
    props.push(`style:font-name="${esc(font)}"`);
  }
  if (!props.length) return null;
  const key = props.join(' ');
  return texts.name(key, (name) => `<style:style style:name="${name}" style:family="text"><style:text-properties ${key}/></style:style>`);
}

function runsXml(runs, texts, fonts) {
  return (runs || [])
    .map((run) => {
      const inner = textXml(run.text);
      const name = runStyle(run, texts, fonts);
      return name ? `<text:span text:style-name="${name}">${inner}</text:span>` : inner;
    })
    .join('');
}

function paragraphStyle(block, paras) {
  const props = [];
  const align = { center: 'center', right: 'end', justify: 'justify', left: 'start' }[block.align];
  if (align && align !== 'start') props.push(`fo:text-align="${align}"`);
  if (block.level && !block.list) props.push(`fo:margin-left="${(Number(block.level) * 0.635).toFixed(3)}cm"`);
  if (!props.length) return null;
  const key = props.join(' ');
  return paras.name(key, (name) => `<style:style style:name="${name}" style:family="paragraph"><style:paragraph-properties ${key}/></style:style>`);
}

/** The text of a block as runs, from whichever shape the frame hands over. */
function runsOf(block) {
  if (Array.isArray(block.runs) && block.runs.length) return block.runs;
  if (typeof block.text === 'string') return [{ text: block.text }];
  return [];
}

function listKind(block) {
  const list = block.list;
  if (!list) return null;
  if (typeof list === 'string') return /num|dec|order/i.test(list) ? 'number' : 'bullet';
  return /num|dec|order/i.test(String(list.kind || list.type || list.format || '')) ? 'number' : 'bullet';
}

/**
 * A document: headings, paragraphs, lists and tables from the frame the
 * editor renders — the same blocks the RTF writer takes.
 */
export function writeOdt({ blocks = [], title = '' } = {}) {
  const texts = styleTable('T');
  const paras = styleTable('P');
  const fonts = new Set();
  const body = [];
  let open = null; // an open list: { kind, items: [] }

  const flushList = () => {
    if (!open) return;
    const style = open.kind === 'number' ? 'LNumber' : 'LBullet';
    body.push(`<text:list text:style-name="${style}">${open.items.map((p) => `<text:list-item>${p}</text:list-item>`).join('')}</text:list>`);
    open = null;
  };

  const paragraph = (block, tag = 'text:p', extra = '') => {
    const name = paragraphStyle(block, paras);
    const attr = name ? ` text:style-name="${name}"` : '';
    return `<${tag}${attr}${extra}>${runsXml(runsOf(block), texts, fonts)}</${tag}>`;
  };

  for (const block of blocks) {
    if (!block) continue;
    if (block.type === 'table') {
      flushList();
      const rows = block.rows || [];
      const cols = Math.max(1, ...rows.map((r) => (r.cells || r).length));
      const rowXml = rows
        .map((row) => {
          const cells = row.cells || row;
          return `<table:table-row>${cells
            .map((cell) => {
              const inner = Array.isArray(cell?.blocks)
                ? cell.blocks.map((b) => paragraph(b)).join('')
                : `<text:p>${runsXml(runsOf(cell || {}), texts, fonts)}</text:p>`;
              return `<table:table-cell office:value-type="string">${inner}</table:table-cell>`;
            })
            .join('')}</table:table-row>`;
        })
        .join('');
      body.push(`<table:table table:name="${esc(block.name || `Table${body.length + 1}`)}"><table:table-column table:number-columns-repeated="${cols}"/>${rowXml}</table:table>`);
      continue;
    }
    const kind = listKind(block);
    if (kind) {
      if (open && open.kind !== kind) flushList();
      if (!open) open = { kind, items: [] };
      open.items.push(paragraph(block));
      continue;
    }
    flushList();
    if (block.type === 'heading') {
      const level = Math.min(9, Math.max(1, Number(block.level) || 1));
      body.push(paragraph(block, 'text:h', ` text:outline-level="${level}"`));
    } else if (block.type === 'pageBreak') {
      body.push('<text:p text:style-name="PBreak"/>');
    } else {
      body.push(paragraph(block));
    }
  }
  flushList();

  const content =
    `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${NS}>` +
    fontDecls(fonts) +
    `<office:automatic-styles>${texts.xml()}${paras.xml()}` +
    '<style:style style:name="PBreak" style:family="paragraph"><style:paragraph-properties fo:break-before="page"/></style:style>' +
    '<text:list-style style:name="LBullet"><text:list-level-style-bullet text:level="1" text:bullet-char="•"><style:list-level-properties text:space-before="0.635cm" text:min-label-width="0.635cm"/></text:list-level-style-bullet></text:list-style>' +
    '<text:list-style style:name="LNumber"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."><style:list-level-properties text:space-before="0.635cm" text:min-label-width="0.635cm"/></text:list-level-style-number></text:list-style>' +
    `</office:automatic-styles><office:body><office:text>${body.join('')}</office:text></office:body></office:document-content>`;

  return archive('odt', { content, styles: textStyles(), title });
}

function fontDecls(fonts) {
  if (!fonts.size) return '';
  return `<office:font-face-decls>${[...fonts].map((f) => `<style:font-face style:name="${esc(f)}" svg:font-family="${esc(f)}"/>`).join('')}</office:font-face-decls>`;
}

function textStyles() {
  const heading = (n, size) =>
    `<style:style style:name="Heading_20_${n}" style:display-name="Heading ${n}" style:family="paragraph" style:parent-style-name="Standard" style:default-outline-level="${n}"><style:paragraph-properties fo:margin-top="0.4cm" fo:margin-bottom="0.2cm" fo:keep-with-next="always"/><style:text-properties fo:font-size="${size}pt" fo:font-weight="bold"/></style:style>`;
  return (
    `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${NS}><office:styles>` +
    '<style:default-style style:family="paragraph"><style:text-properties style:font-name="Calibri" fo:font-size="11pt"/></style:default-style>' +
    '<style:style style:name="Standard" style:family="paragraph" style:class="text"/>' +
    heading(1, 20) + heading(2, 16) + heading(3, 14) +
    '</office:styles><office:font-face-decls><style:font-face style:name="Calibri" svg:font-family="Calibri"/></office:font-face-decls>' +
    '<office:automatic-styles><style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm"/></style:page-layout></office:automatic-styles>' +
    '<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="PM1"/></office:master-styles></office:document-styles>'
  );
}

/* ── formulas: A1 in the engine, [.A1] in the file ─────────────────────── */

const REF = /(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?(\$?[A-Z]{1,3}\$?[0-9]+(?::\$?[A-Z]{1,3}\$?[0-9]+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?[0-9]+:\$?[0-9]+)(?![A-Za-z0-9_(])/g;

/** Whole columns and rows, which ODF spells out. */
function expandRange(range) {
  const cols = /^(\$?[A-Z]{1,3}):(\$?[A-Z]{1,3})$/.exec(range);
  if (cols) return `${cols[1]}1:${cols[2]}1048576`;
  const rows = /^(\$?[0-9]+):(\$?[0-9]+)$/.exec(range);
  if (rows) return `A${rows[1]}:XFD${rows[2]}`;
  return range;
}

/**
 * `=SUM(Sheet2!A1:B2, 3)` becomes `of:=SUM([$Sheet2.A1:.B2]; 3)`: every
 * reference in brackets with a dot before its cell, the sheet with a dollar,
 * and arguments parted by semicolons. Strings pass through untouched.
 */
export function formulaToOdf(formula) {
  const src = String(formula || '');
  if (!src.startsWith('=')) return null;
  let out = '';
  let i = 1;
  while (i < src.length) {
    const ch = src[i];
    if (ch === '"') {
      const end = src.indexOf('"', i + 1);
      const stop = end < 0 ? src.length : end + 1;
      out += src.slice(i, stop);
      i = stop;
      continue;
    }
    if (ch === ',') {
      out += ';';
      i += 1;
      continue;
    }
    REF.lastIndex = i;
    const m = REF.exec(src);
    if (m && m.index === i) {
      const sheet = m[1] ? m[1].replace(/''/g, "'") : m[2];
      const range = expandRange(m[3]);
      const [from, to] = range.split(':');
      const prefix = sheet ? `$${/[^A-Za-z0-9_]/.test(sheet) ? `'${sheet.replace(/'/g, "''")}'` : sheet}` : '';
      out += `[${prefix}.${from}${to ? `:.${to}` : ''}]`;
      i += m[0].length;
      continue;
    }
    out += ch;
    i += 1;
  }
  return `of:=${out}`;
}

/** The way back: `[$Sheet2.A1:.B2]` to `Sheet2!A1:B2`, semicolons to commas. */
export function formulaFromOdf(formula) {
  let src = String(formula || '').replace(/^(?:of|oooc):/, '');
  if (!src.startsWith('=')) src = `=${src}`;
  let out = '';
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (ch === '"') {
      const end = src.indexOf('"', i + 1);
      const stop = end < 0 ? src.length : end + 1;
      out += src.slice(i, stop);
      i = stop;
      continue;
    }
    if (ch === ';') {
      out += ',';
      i += 1;
      continue;
    }
    // A named range is `$$GrandTotal` in the file and `GrandTotal` to the engine.
    if (ch === '$' && src[i + 1] === '$') {
      i += 2;
      continue;
    }
    if (ch === '[') {
      const end = src.indexOf(']', i);
      if (end > i) {
        const inner = src.slice(i + 1, end);
        const parts = inner.split(':').map((p) => {
          const dot = p.lastIndexOf('.');
          const sheet = dot > 0 ? p.slice(0, dot).replace(/^\$/, '') : '';
          const cell = p.slice(dot + 1);
          return { sheet, cell };
        });
        const sheet = parts[0].sheet;
        const prefix = sheet ? `${/[^A-Za-z0-9_]/.test(sheet.replace(/^'|'$/g, '')) ? sheet : sheet.replace(/^'|'$/g, '')}!` : '';
        out += prefix + parts.map((p) => p.cell).join(':');
        i = end + 1;
        continue;
      }
    }
    out += ch;
    i += 1;
  }
  return out;
}

/* ── a workbook ────────────────────────────────────────────────────────── */

const EPOCH = Date.UTC(1899, 11, 30);

function isoDate(serial) {
  const d = new Date(EPOCH + Math.round(Number(serial) * 86400000));
  if (!Number.isFinite(d.getTime())) return null;
  return d.toISOString().slice(0, Number(serial) % 1 ? 19 : 10);
}

/** What kind of value a cell holds, from its value and the format it wears. */
function valueAttrs(cell) {
  const v = cell.value;
  const fmt = String(cell.format || '');
  if (typeof v === 'boolean') return `office:value-type="boolean" office:boolean-value="${v}"`;
  if (typeof v === 'number' && Number.isFinite(v)) {
    if (/[dmyh]/i.test(fmt.replace(/"[^"]*"/g, '')) && !/General/i.test(fmt)) {
      const iso = isoDate(v);
      if (iso) return `office:value-type="date" office:date-value="${iso}"`;
    }
    if (/%/.test(fmt)) return `office:value-type="percentage" office:value="${v}"`;
    return `office:value-type="float" office:value="${v}"`;
  }
  if (v === null || v === undefined || v === '') return cell.formula ? 'office:value-type="string"' : '';
  return 'office:value-type="string"';
}

/**
 * A workbook: each sheet's cells with their values, the type a value needs
 * to stay what it is, and the formulas in the file's own spelling.
 *
 * `sheets` is `[{ name, rows: [[cell]] }]`, a cell being
 * `{ value, text, formula, format }` — the calculated value, what the grid
 * shows, the formula as typed (with its `=`) and the number format code.
 */
export function writeOds({ sheets = [], title = '' } = {}) {
  const tables = sheets
    .map((sheet, si) => {
      const rows = sheet.rows || [];
      const cols = Math.max(1, ...rows.map((r) => r.length));
      const rowXml = rows
        .map((row) => {
          let xml = '';
          let gap = 0;
          const flush = () => {
            if (gap) xml += gap === 1 ? '<table:table-cell/>' : `<table:table-cell table:number-columns-repeated="${gap}"/>`;
            gap = 0;
          };
          for (const cell of row) {
            const attrs = cell ? valueAttrs(cell) : '';
            const formula = cell?.formula ? formulaToOdf(cell.formula) : null;
            if (!attrs && !formula) {
              gap += 1;
              continue;
            }
            flush();
            const text = cell.text ?? (cell.value === null || cell.value === undefined ? '' : String(cell.value));
            xml += `<table:table-cell ${attrs}${formula ? ` table:formula="${esc(formula)}"` : ''}>${text === '' ? '' : `<text:p>${textXml(text)}</text:p>`}</table:table-cell>`;
          }
          return `<table:table-row>${xml || '<table:table-cell/>'}</table:table-row>`;
        })
        .join('');
      return `<table:table table:name="${esc(sheet.name || `Sheet${si + 1}`)}"><table:table-column table:number-columns-repeated="${cols}"/>${rowXml || '<table:table-row><table:table-cell/></table:table-row>'}</table:table>`;
    })
    .join('');

  const content =
    `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${NS}><office:automatic-styles/>` +
    `<office:body><office:spreadsheet>${tables}</office:spreadsheet></office:body></office:document-content>`;
  const styles = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${NS}><office:styles><style:default-style style:family="table-cell"><style:text-properties style:font-name="Calibri" fo:font-size="11pt"/></style:default-style></office:styles><office:font-face-decls><style:font-face style:name="Calibri" svg:font-family="Calibri"/></office:font-face-decls></office:document-styles>`;
  return archive('ods', { content, styles, title });
}

/* ── a deck ────────────────────────────────────────────────────────────── */

/**
 * A deck: each slide's text boxes and pictures where they sit, and the
 * speaker's notes.
 *
 * `slides` is `[{ name, shapes, notes }]`; a shape is
 * `{ kind: 'text', paragraphs: [{ runs }|string], x, y, w, h }` or
 * `{ kind: 'picture', data, contentType, x, y, w, h }`, positions in pixels
 * at 96 to the inch, as the deck measures. `size` is `{ width, height }` in
 * the same pixels.
 */
export function writeOdp({ slides = [], size = { width: 1280, height: 720 }, title = '' } = {}) {
  const texts = styleTable('T');
  const fonts = new Set();
  const pictures = [];
  const pageXml = slides
    .map((slide, i) => {
      const frames = (slide.shapes || [])
        .map((shape, j) => {
          const box = `svg:x="${cm(shape.x)}" svg:y="${cm(shape.y)}" svg:width="${cm(shape.w)}" svg:height="${cm(shape.h)}"`;
          if (shape.kind === 'picture' && shape.data) {
            const ext = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp', 'image/bmp': 'bmp', 'image/svg+xml': 'svg' }[shape.contentType] || 'bin';
            const name = `Pictures/image${pictures.length + 1}.${ext}`;
            pictures.push({ name, data: shape.data, contentType: shape.contentType || 'application/octet-stream' });
            return `<draw:frame draw:name="Picture ${j + 1}" draw:layer="layout" ${box}><draw:image xlink:href="${name}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
          }
          if (shape.kind !== 'text') return '';
          const paras = (shape.paragraphs || [])
            .map((p) => `<text:p>${typeof p === 'string' ? textXml(p) : runsXml(p.runs && p.runs.length ? p.runs : [{ text: p.plain ?? p.text ?? '' }], texts, fonts)}</text:p>`)
            .join('');
          const cls = shape.placeholder === 'title' || shape.placeholder === 'ctrTitle' ? ' presentation:class="title"' : '';
          return `<draw:frame draw:name="${esc(shape.name || `Text ${j + 1}`)}" draw:layer="layout"${cls} ${box}><draw:text-box>${paras || '<text:p/>'}</draw:text-box></draw:frame>`;
        })
        .join('');
      const notes = slide.notes
        ? `<presentation:notes><draw:frame presentation:class="notes" draw:layer="layout" svg:x="2cm" svg:y="12cm" svg:width="17cm" svg:height="12cm"><draw:text-box>${String(slide.notes).split('\n').map((line) => `<text:p>${textXml(line)}</text:p>`).join('')}</draw:text-box></draw:frame></presentation:notes>`
        : '';
      return `<draw:page draw:name="${esc(slide.name || `Slide ${i + 1}`)}" draw:master-page-name="Default">${frames}${notes}</draw:page>`;
    })
    .join('');

  const content =
    `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${NS}>` +
    fontDecls(fonts) +
    `<office:automatic-styles>${texts.xml()}</office:automatic-styles>` +
    `<office:body><office:presentation>${pageXml}</office:presentation></office:body></office:document-content>`;
  const styles =
    `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${NS}><office:styles/>` +
    `<office:automatic-styles><style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="${cm(size.width)}" fo:page-height="${cm(size.height)}" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm" style:print-orientation="landscape"/></style:page-layout>` +
    '<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties draw:background-size="border" draw:fill="none"/></style:style></office:automatic-styles>' +
    '<office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1" draw:style-name="dp1"/></office:master-styles></office:document-styles>';
  return archive('odp', { content, styles, title, pictures });
}

/* ── the archive ───────────────────────────────────────────────────────── */

function metaXml(title) {
  return `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-meta ${NS}><office:meta><meta:generator>Rutba Office</meta:generator>${title ? `<dc:title>${esc(title)}</dc:title>` : ''}<dc:date>${new Date().toISOString()}</dc:date></office:meta></office:document-meta>`;
}

function entry(name, data, { stored = false } = {}) {
  const e = new ZipEntry({
    name,
    method: stored ? 0 : 8,
    crc: 0,
    compressedSize: 0,
    uncompressedSize: 0,
    compressed: Buffer.alloc(0),
    flags: 0,
    // A fixed timestamp, so the same document written twice is the same bytes.
    dosTime: 0,
    dosDate: 0x2821,
    externalAttrs: 0,
    comment: Buffer.alloc(0),
  });
  e.data = Buffer.isBuffer(data) ? data : Buffer.from(String(data), 'utf8');
  return e;
}

function archive(kind, { content, styles, title, pictures = [] }) {
  const mime = MIME[kind];
  const manifest =
    '<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">' +
    `<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${mime}"/>` +
    '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>' +
    pictures.map((p) => `<manifest:file-entry manifest:full-path="${esc(p.name)}" manifest:media-type="${esc(p.contentType)}"/>`).join('') +
    '</manifest:manifest>';
  const entries = [
    entry('mimetype', mime, { stored: true }),
    entry('META-INF/manifest.xml', manifest),
    entry('content.xml', content),
    entry('styles.xml', styles),
    entry('meta.xml', metaXml(title)),
    ...pictures.map((p) => entry(p.name, p.data)),
  ];
  return writeZip(entries);
}

export default { writeOdt, writeOds, writeOdp, formulaToOdf, formulaFromOdf };
