// An OpenDocument text document, as the .docx the Word engine edits.
//
// odt.js reads the document whole; this writes it as WordprocessingML: each
// paragraph with its heading style, alignment, indents, spacing and page
// break, each run with its look, links as links, tabs and line breaks;
// lists as numbering — an abstract list per list style, an instance per
// list so each starts again as it did, bullets with their characters and
// numbers in their formats; tables with their grid, spans (gridSpan, and
// vMerge where a cell reaches down), cell backgrounds and borders; pictures
// inline at their size, each a part of its own; the page's size and margins.

import { buildDocx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
  + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
const PICTURE_TYPES = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff' };

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const clean = (s) => String(s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, '');
/** Pixels (96 to the inch) as twentieths of a point. */
const twips = (px) => Math.round((Number(px) || 0) * 15);
const EMU = 9525;

/** ODF's number formats as Word's. */
const NUMFMT = { 1: 'decimal', a: 'lowerLetter', A: 'upperLetter', i: 'lowerRoman', I: 'upperRoman', '': 'none' };

/** Register a file extension's content type, once. */
function ensureDefaultType(pkg, ext, type) {
  const name = '[Content_Types].xml';
  const xml = pkg.text(name);
  if (new RegExp(`Extension="${ext}"`, 'i').test(xml)) return;
  pkg.write_(name, xml.replace('<Override', () => `<Default Extension="${ext}" ContentType="${type}"/><Override`));
}

/** A run's properties, in the schema's order. */
function rPrXml(r) {
  const bits = [];
  if (r.font) bits.push(`<w:rFonts w:ascii="${esc(r.font)}" w:hAnsi="${esc(r.font)}" w:cs="${esc(r.font)}"/>`);
  if (r.bold) bits.push('<w:b/>');
  if (r.italic) bits.push('<w:i/>');
  if (r.strike) bits.push('<w:strike/>');
  if (r.color) bits.push(`<w:color w:val="${r.color.replace('#', '')}"/>`);
  if (r.size) bits.push(`<w:sz w:val="${Math.round(r.size * 2)}"/><w:szCs w:val="${Math.round(r.size * 2)}"/>`);
  if (r.highlight) bits.push(`<w:shd w:val="clear" w:color="auto" w:fill="${r.highlight.replace('#', '')}"/>`);
  if (r.underline) bits.push('<w:u w:val="single"/>');
  if (r.vertical) bits.push(`<w:vertAlign w:val="${r.vertical}"/>`);
  return bits.length ? `<w:rPr>${bits.join('')}</w:rPr>` : '';
}

/**
 * Write an .odt read by readOdt as a .docx: returns the file's bytes.
 * @param {{ blocks: object[], lists: Map, page: object|null, images: Map<string, Uint8Array> }} doc
 */
export function odtToDocx(doc) {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [] }));
  const part = 'word/document.xml';
  const images = doc.images || new Map();
  const media = new Map();
  let mediaN = 0;
  let drawingId = 0;
  const links = new Map();

  /** A picture's part and relationship, once per file in the archive. */
  const imageRel = (href) => {
    const key = String(href || '').replace(/^\.\//, '');
    if (media.has(key)) return media.get(key);
    const bytes = images.get(key);
    const ext = key.split('.').pop().toLowerCase();
    if (!bytes || !PICTURE_TYPES[ext]) return null;
    const name = `media/image${++mediaN}.${ext === 'jpeg' ? 'jpg' : ext}`;
    pkg.addPart(`word/${name}`, Buffer.from(bytes), null);
    ensureDefaultType(pkg, ext === 'jpeg' ? 'jpg' : ext, PICTURE_TYPES[ext]);
    const rId = pkg.addRelationshipTo(part, REL + 'image', name);
    media.set(key, rId);
    return rId;
  };
  const linkRel = (url) => {
    if (!links.has(url)) links.set(url, pkg.addRelationshipTo(part, REL + 'hyperlink', url, { external: true }));
    return links.get(url);
  };

  // Lists: an abstract list per list style, an instance per list (each starting again).
  const abstractOf = new Map();
  const abstracts = [];
  const nums = [];
  const numOf = new Map();
  const listNum = (list) => {
    const key = `${list.style}|${list.id}`;
    if (numOf.has(key)) return numOf.get(key);
    if (!abstractOf.has(list.style)) {
      abstractOf.set(list.style, abstracts.length);
      abstracts.push(doc.lists?.get(list.style) || []);
    }
    const id = nums.length + 1;
    nums.push({ id, abstract: abstractOf.get(list.style) });
    numOf.set(key, id);
    return id;
  };

  const runXml = (r) => {
    if (r.image) {
      const rId = imageRel(r.image.href);
      if (!rId) return '';
      const cx = Math.max(1, Math.round((r.image.width || 96) * EMU));
      const cy = Math.max(1, Math.round((r.image.height || 96) * EMU));
      const id = ++drawingId;
      const name = esc(r.image.name || `Picture ${id}`);
      return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="${name}"/>`
        + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
        + `<pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>`
        + `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
        + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
        + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    }
    const rPr = rPrXml(r);
    if (r.tab) return `<w:r>${rPr}<w:tab/></w:r>`;
    if (r.br) return `<w:r>${rPr}<w:br/></w:r>`;
    return `<w:r>${rPr}<w:t xml:space="preserve">${esc(clean(r.text ?? ''))}</w:t></w:r>`;
  };

  const paragraphXml = (p) => {
    const pPr = [];
    if (p.heading) pPr.push(`<w:pStyle w:val="Heading${Math.min(9, p.heading)}"/>`);
    if (p.pageBreakBefore) pPr.push('<w:pageBreakBefore/>');
    if (p.list) pPr.push(`<w:numPr><w:ilvl w:val="${p.list.level}"/><w:numId w:val="${listNum(p.list)}"/></w:numPr>`);
    if (p.rtl) pPr.push('<w:bidi/>');
    if (p.spaceBefore != null || p.spaceAfter != null) pPr.push(`<w:spacing${p.spaceBefore != null ? ` w:before="${twips(p.spaceBefore)}"` : ''}${p.spaceAfter != null ? ` w:after="${twips(p.spaceAfter)}"` : ''}/>`);
    // A list item's indents are its list level's; the paragraph's own beside them only when it states them.
    if (!p.list && (p.indentLeft || p.indentRight || p.indentFirst)) {
      const first = p.indentFirst || 0;
      pPr.push(`<w:ind w:left="${twips(p.indentLeft || 0)}" w:right="${twips(p.indentRight || 0)}"${first < 0 ? ` w:hanging="${twips(-first)}"` : first > 0 ? ` w:firstLine="${twips(first)}"` : ''}/>`);
    }
    if (p.align && p.align !== 'left') pPr.push(`<w:jc w:val="${p.align}"/>`);
    // Consecutive runs that share a link go inside one hyperlink.
    const out = [];
    let open = null;
    for (const r of p.runs || []) {
      const url = r.link || null;
      if (url !== open) {
        if (open) out.push('</w:hyperlink>');
        if (url) out.push(`<w:hyperlink r:id="${linkRel(url)}">`);
        open = url;
      }
      out.push(runXml(url ? { underline: true, color: '#0563C1', ...r } : r));
    }
    if (open) out.push('</w:hyperlink>');
    return `<w:p>${pPr.length ? `<w:pPr>${pPr.join('')}</w:pPr>` : ''}${out.join('')}</w:p>`;
  };

  const tableXml = (t) => {
    const cols = t.columns.length;
    const known = t.columns.filter(Boolean);
    const fallback = known.length ? known.reduce((a, b) => a + b, 0) / known.length : 624 / Math.max(1, cols);
    const widths = t.columns.map((w) => twips(w || fallback));
    const borders = t.rows.some((r) => r.some((c) => c?.border));
    // Where each cell reaches: a span down leaves continuing cells in the rows below.
    const below = new Map();
    const rows = t.rows.map((row, ri) => {
      // Every place in the grid has its element in ODF — a cell, or a covered cell where a span reaches.
      const cells = [];
      const width = (col, span) => widths.slice(col, col + span).reduce((x, y) => x + y, 0);
      const shade = (fill) => (fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill.replace('#', '')}"/>` : '');
      for (let col = 0; col < cols;) {
        const cell = row[col];
        const down = below.get(`${ri}:${col}`);
        if (down && (!cell || cell.covered)) {
          cells.push(`<w:tc><w:tcPr><w:tcW w:w="${width(col, down.span)}" w:type="dxa"/>${down.span > 1 ? `<w:gridSpan w:val="${down.span}"/>` : ''}<w:vMerge/>${shade(down.fill)}</w:tcPr><w:p/></w:tc>`);
          col += down.span;
          continue;
        }
        if (!cell || cell.covered) {
          if (!cell) cells.push(`<w:tc><w:tcPr><w:tcW w:w="${widths[col]}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`);
          col++;
          continue;
        }
        const span = Math.max(1, Math.min(cols - col, cell.colspan || 1));
        for (let k = 1; k < (cell.rowspan || 1); k++) below.set(`${ri + k}:${col}`, { span, fill: cell.fill });
        const inner = (cell.blocks || []).map(blockXml).join('');
        const body = !inner ? '<w:p/>' : /<\/w:tbl>$/.test(inner) ? inner + '<w:p/>' : inner;
        cells.push(`<w:tc><w:tcPr><w:tcW w:w="${width(col, span)}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}${(cell.rowspan || 1) > 1 ? '<w:vMerge w:val="restart"/>' : ''}${shade(cell.fill)}</w:tcPr>${body}</w:tc>`);
        col += span;
      }
      return `<w:tr>${cells.join('')}</w:tr>`;
    });
    const edge = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`;
    return `<w:tbl><w:tblPr><w:tblW w:w="${widths.reduce((a, b) => a + b, 0)}" w:type="dxa"/>`
      + (borders ? `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(edge).join('')}</w:tblBorders>` : '')
      + `<w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows.join('')}</w:tbl>`;
  };

  function blockXml(b) {
    return b.type === 'table' ? tableXml(b) : paragraphXml(b);
  }

  // A table cannot end the body: Word wants a paragraph after it.
  let body = (doc.blocks || []).map(blockXml).join('');
  if (!body || /<\/w:tbl>$/.test(body)) body += '<w:p/>';
  const page = doc.page;
  const sectPr = page && page.width && page.height
    ? `<w:sectPr><w:pgSz w:w="${twips(page.width)}" w:h="${twips(page.height)}"${page.landscape ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${twips(page.top ?? 96)}" w:right="${twips(page.right ?? 96)}" w:bottom="${twips(page.bottom ?? 96)}" w:left="${twips(page.left ?? 96)}" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`
    : '';
  pkg.write_(part, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${sectPr}</w:body></w:document>`);

  // Headings past the third, which the built styles have not got: each a step smaller.
  const used = new Set([...(doc.blocks || [])].flatMap(function headings(b) { return b.type === 'table' ? b.rows.flat().flatMap((c) => (c?.blocks || []).flatMap(headings)) : b.heading ? [b.heading] : []; }));
  const extra = [...used].filter((n) => n > 3).map((n) => `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="60"/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:b/><w:sz w:val="${Math.max(20, 28 - (n - 3) * 2)}"/></w:rPr></w:style>`);
  if (extra.length) pkg.write_('word/styles.xml', pkg.text('word/styles.xml').replace('</w:styles>', () => extra.join('') + '</w:styles>'));

  if (nums.length) {
    const levelXml = (lv, i) => {
      if (!lv) return `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
      const left = twips(lv.indent || 48 * (i + 1));
      const hanging = twips(lv.hanging || 24);
      if (lv.kind === 'bullet') {
        return `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${esc(lv.char)}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="${hanging}"/></w:pPr></w:lvl>`;
      }
      // ODF shows `display` levels' numbers, this one last: %1.%2 and so on.
      const shown = Array.from({ length: Math.max(1, Math.min(lv.display || 1, i + 1)) }, (_, k) => `%${i + 2 - Math.max(1, Math.min(lv.display || 1, i + 1)) + k}`).join('.');
      return `<w:lvl w:ilvl="${i}"><w:start w:val="${lv.start || 1}"/><w:numFmt w:val="${NUMFMT[lv.format] ?? 'decimal'}"/><w:lvlText w:val="${esc(lv.prefix + shown + lv.suffix)}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="${hanging}"/></w:pPr></w:lvl>`;
    };
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}>`
      + abstracts.map((levels, a) => `<w:abstractNum w:abstractNumId="${a}"><w:multiLevelType w:val="hybridMultilevel"/>${Array.from({ length: 9 }, (_, i) => levelXml(levels[i], i)).join('')}</w:abstractNum>`).join('')
      + nums.map((n) => `<w:num w:numId="${n.id}"><w:abstractNumId w:val="${n.abstract}"/>${Array.from({ length: 9 }, (_, i) => `<w:lvlOverride w:ilvl="${i}"><w:startOverride w:val="${abstracts[n.abstract][i]?.start || 1}"/></w:lvlOverride>`).join('')}</w:num>`).join('')
      + '</w:numbering>';
    pkg.addPart('word/numbering.xml', xml, CT + 'numbering+xml');
    pkg.addRelationshipTo(part, REL + 'numbering', 'numbering.xml');
  }
  return pkg.write();
}
