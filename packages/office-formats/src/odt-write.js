// An OpenDocument text written whole from the document model odt.js reads
// an .odt into (and docx-read.js reads a .docx into).
//
// writeOdt in odf-write.js writes the editor's plain frame — paragraphs,
// headings and flat lists. This writes what the model holds: headings at
// their outline levels; paragraphs in their alignment, indents, spacing,
// page breaks and direction; runs in their looks with tabs, breaks, links
// and pictures (stored in Pictures/ and listed in the manifest); lists in
// their own list styles, nested by level and numbering on where a list
// carries on after something else; tables with their column widths, spans
// (covered cells where a span reaches), shading and borders; the page's
// size, orientation and margins; and the default look of the words.

import { ODF_NAMESPACES as NS, odfArchive, odfTextXml as textXml } from './odf-write.js';

const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const cm = (px) => `${(Number(px || 0) / 96 * 2.54).toFixed(3)}cm`;
/** A name a style can carry: letters, digits and underscores, starting with a letter. */
const ncname = (s) => { const n = String(s || 'L').replace(/[^A-Za-z0-9_]/g, '_'); return /^[A-Za-z]/.test(n) ? n : `L${n}`; };

const MEDIA = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf', webp: 'image/webp' };
const HEADING_SIZES = [20, 16, 14, 13, 12, 12, 11, 11, 11];

/** A run's look as ODF text properties. */
function textProps(look, fonts) {
  const p = [];
  if (look.bold) p.push('fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"');
  else if (look.bold === false) p.push('fo:font-weight="normal"');
  if (look.italic) p.push('fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"');
  else if (look.italic === false) p.push('fo:font-style="normal"');
  if (look.underline) p.push('style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"');
  if (look.strike) p.push('style:text-line-through-style="solid"');
  if (look.color) p.push(`fo:color="${esc(look.color)}"`);
  if (look.size) p.push(`fo:font-size="${look.size}pt" style:font-size-asian="${look.size}pt" style:font-size-complex="${look.size}pt"`);
  if (look.font) { fonts.add(look.font); p.push(`style:font-name="${esc(look.font)}"`); }
  if (look.highlight) p.push(`fo:background-color="${esc(look.highlight)}"`);
  if (look.vertical === 'superscript') p.push('style:text-position="super 58%"');
  else if (look.vertical === 'subscript') p.push('style:text-position="sub 58%"');
  return p.join(' ');
}

/** A table of automatic styles, each written once and named by its first use. */
function styleTable(prefix) {
  const names = new Map();
  const xml = [];
  return {
    name(key, body) {
      if (names.has(key)) return names.get(key);
      const name = `${prefix}${names.size + 1}`;
      names.set(key, name);
      xml.push(body(name));
      return name;
    },
    xml: () => xml.join(''),
  };
}

/** One list style: each level a bullet or a number, placed by its indent and hang. */
function listStyleXml(name, levels) {
  const out = [];
  for (let i = 0; i < 9; i++) {
    const lv = levels[i] || levels[levels.length - 1] || { kind: 'bullet', char: '•', indent: 48 + 24 * i, hanging: 24 };
    const indent = lv.indent || 48 + 24 * i;
    const hanging = lv.hanging ?? 24;
    const placed = `<style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="${cm(indent)}" fo:text-indent="${cm(-hanging)}" fo:margin-left="${cm(indent)}"/></style:list-level-properties>`;
    if (lv.kind === 'number') {
      out.push(`<text:list-level-style-number text:level="${i + 1}"${lv.prefix ? ` style:num-prefix="${esc(lv.prefix)}"` : ''} style:num-suffix="${esc(lv.suffix ?? '.')}" style:num-format="${esc(lv.format || '1')}"${lv.start && lv.start !== 1 ? ` text:start-value="${lv.start}"` : ''}${lv.display > 1 ? ` text:display-levels="${Math.min(lv.display, i + 1)}"` : ''}>${placed}</text:list-level-style-number>`);
    } else {
      out.push(`<text:list-level-style-bullet text:level="${i + 1}" text:bullet-char="${esc(lv.char || '•')}">${placed}</text:list-level-style-bullet>`);
    }
  }
  return `<text:list-style style:name="${name}">${out.join('')}</text:list-style>`;
}

/**
 * The model as .odt bytes: { blocks, lists, page, images, title, defaults }
 * in odt.js's shapes.
 */
export function writeOdtDocument(doc = {}) {
  const fonts = new Set();
  const texts = styleTable('T');
  const paras = styleTable('P');
  const tables = styleTable('Tbl');
  const cells = styleTable('Cell');
  const listNames = new Map(); // the model's list style → the written one
  const pictures = new Map(); // archive name → { name, data, contentType }
  const used = new Set(); // frame names already given
  let tableCount = 0;
  let frameCount = 0;

  const listStyle = (style) => {
    if (!listNames.has(style)) listNames.set(style, ncname(`L_${style}`));
    return listNames.get(style);
  };

  const textStyle = (look) => {
    const props = textProps(look, fonts);
    return props ? texts.name(props, (name) => `<style:style style:name="${name}" style:family="text"><style:text-properties ${props}/></style:style>`) : null;
  };

  const paraStyle = (block) => {
    const parent = block.heading ? `Heading_20_${block.heading}` : 'Standard';
    const p = [];
    const align = { left: 'start', center: 'center', right: 'end', both: 'justify' }[block.align];
    if (align) p.push(`fo:text-align="${align}"`);
    if (!block.list) {
      if (block.indentLeft) p.push(`fo:margin-left="${cm(block.indentLeft)}"`);
      if (block.indentFirst) p.push(`fo:text-indent="${cm(block.indentFirst)}"`);
    }
    if (block.indentRight) p.push(`fo:margin-right="${cm(block.indentRight)}"`);
    if (block.spaceBefore != null) p.push(`fo:margin-top="${cm(block.spaceBefore)}"`);
    if (block.spaceAfter != null) p.push(`fo:margin-bottom="${cm(block.spaceAfter)}"`);
    if (block.pageBreakBefore) p.push('fo:break-before="page"');
    if (block.rtl) p.push('style:writing-mode="rl-tb"');
    if (!p.length) return parent;
    const key = `${parent}|${p.join(' ')}`;
    return paras.name(key, (name) => `<style:style style:name="${name}" style:family="paragraph" style:parent-style-name="${parent}"><style:paragraph-properties ${p.join(' ')}/></style:style>`);
  };

  const frameXml = (image) => {
    const data = doc.images?.get?.(image.href);
    if (!data) return '';
    if (!pictures.has(image.href)) {
      const ext = (/\.([a-z0-9]+)$/i.exec(image.href)?.[1] || 'png').toLowerCase();
      pictures.set(image.href, { name: image.href, data: Buffer.from(data), contentType: MEDIA[ext] || 'application/octet-stream' });
    }
    let name = image.name || `Image${++frameCount}`;
    while (used.has(name)) name = `Image${++frameCount}`;
    used.add(name);
    return `<draw:frame draw:name="${esc(name)}" text:anchor-type="as-char" svg:width="${cm(image.width)}" svg:height="${cm(image.height)}" draw:z-index="0"><draw:image xlink:href="${esc(image.href)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
  };

  /** A paragraph's runs: spans for their looks, links round the runs that share one. */
  const runsXml = (runs) => {
    let out = '';
    let link = null;
    let inLink = '';
    const flush = () => {
      if (link) out += `<text:a xlink:type="simple" xlink:href="${esc(link)}">${inLink}</text:a>`;
      link = null;
      inLink = '';
    };
    for (const run of runs || []) {
      let piece;
      if (run.image) piece = frameXml(run.image);
      else {
        const look = Object.fromEntries(Object.entries(run).filter(([k]) => !['text', 'tab', 'br', 'link', 'image'].includes(k)));
        const inner = run.tab ? '<text:tab/>' : run.br ? '<text:line-break/>' : textXml(run.text);
        const name = textStyle(look);
        piece = name ? `<text:span text:style-name="${name}">${inner}</text:span>` : inner;
      }
      if ((run.link || null) !== link) flush();
      if (run.link) { link = run.link; inLink += piece; } else out += piece;
    }
    flush();
    return out;
  };

  const paragraphXml = (block) => {
    const style = paraStyle(block);
    const body = runsXml(block.runs);
    return block.heading
      ? `<text:h text:style-name="${style}" text:outline-level="${block.heading}">${body}</text:h>`
      : `<text:p text:style-name="${style}">${body}</text:p>`;
  };

  const tableXml = (table) => {
    tableCount += 1;
    const name = `Table${tableCount}`;
    const widths = table.columns || [];
    const total = widths.every((w) => w > 0) ? widths.reduce((a, b) => a + b, 0) : 0;
    const tStyle = tables.name(`${name}`, (n) => `<style:style style:name="${n}" style:family="table"><style:table-properties${total ? ` style:width="${cm(total)}"` : ''} table:align="left"/></style:style>`);
    const colXml = widths.map((w, i) => {
      if (!w) return '<table:table-column/>';
      const cn = `${tStyle}.C${i + 1}`;
      tables.name(cn, () => `<style:style style:name="${cn}" style:family="table-column"><style:table-column-properties style:column-width="${cm(w)}"/></style:style>`);
      return `<table:table-column table:style-name="${cn}"/>`;
    }).join('');
    const rowXml = (table.rows || []).map((row) => `<table:table-row>${row.map((cell) => {
      if (!cell || cell.covered) return '<table:covered-table-cell/>';
      const props = [cell.fill ? `fo:background-color="${esc(cell.fill)}"` : '', cell.border ? 'fo:border="0.5pt solid #000000"' : 'fo:border="none"', 'fo:padding="0.097cm"'].filter(Boolean).join(' ');
      const cs = cells.name(props, (n) => `<style:style style:name="${n}" style:family="table-cell"><style:table-cell-properties ${props}/></style:style>`);
      const span = `${cell.colspan > 1 ? ` table:number-columns-spanned="${cell.colspan}"` : ''}${cell.rowspan > 1 ? ` table:number-rows-spanned="${cell.rowspan}"` : ''}`;
      const inner = blocksXml(cell.blocks || []);
      return `<table:table-cell table:style-name="${cs}" office:value-type="string"${span}>${inner || '<text:p text:style-name="Standard"/>'}</table:table-cell>`;
    }).join('')}</table:table-row>`).join('');
    return `<table:table table:name="${name}" table:style-name="${tStyle}">${colXml}${rowXml}</table:table>`;
  };

  /** A run of one list's items as nested lists, numbering on where the list carried on after something else. */
  const listXml = (items, style, carriesOn) => {
    let out = `<text:list text:style-name="${listStyle(style)}"${carriesOn ? ' text:continue-numbering="true"' : ''}>`;
    let depth = 0;
    let open = false; // an item open at the current depth
    for (const { level, xml } of items) {
      while (depth > level) { out += `${open ? '</text:list-item>' : ''}</text:list>`; depth -= 1; open = true; }
      if (depth < level) {
        while (depth < level) { out += `${open ? '' : '<text:list-item>'}<text:list>`; depth += 1; open = false; }
      } else if (open) out += '</text:list-item>';
      out += `<text:list-item>${xml}`;
      open = true;
    }
    if (open) out += '</text:list-item>';
    while (depth > 0) { out += '</text:list></text:list-item>'; depth -= 1; }
    return `${out}</text:list>`;
  };

  const seenLists = new Set();
  function blocksXml(blocks) {
    let out = '';
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.type === 'table') { out += tableXml(b); continue; }
      if (b.list && b.list.style && doc.lists?.get?.(b.list.style)) {
        const items = [];
        const { id, style } = b.list;
        while (i < blocks.length && blocks[i].type !== 'table' && blocks[i].list?.id === id) {
          items.push({ level: Math.max(0, Math.min(8, blocks[i].list.level || 0)), xml: paragraphXml(blocks[i]) });
          i += 1;
        }
        i -= 1;
        out += listXml(items, style, seenLists.has(id));
        seenLists.add(id);
        continue;
      }
      out += paragraphXml(b);
    }
    return out;
  }

  const body = blocksXml(doc.blocks || []);
  const listStyles = [...listNames].map(([style, name]) => listStyleXml(name, doc.lists.get(style) || [])).join('');
  const defaultProps = textProps(doc.defaults?.run || {}, fonts);
  const fontDecls = `<office:font-face-decls>${[...fonts].map((f) => `<style:font-face style:name="${esc(f)}" svg:font-family="${esc(/\s/.test(f) ? `'${f}'` : f)}"/>`).join('')}</office:font-face-decls>`;

  const content = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${NS}>${fontDecls}`
    + `<office:automatic-styles>${tables.xml()}${cells.xml()}${paras.xml()}${texts.xml()}${listStyles}</office:automatic-styles>`
    + `<office:body><office:text>${body || '<text:p text:style-name="Standard"/>'}</office:text></office:body></office:document-content>`;

  const page = doc.page || {};
  const heading = (n) => `<style:style style:name="Heading_20_${n}" style:display-name="Heading ${n}" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Standard" style:default-outline-level="${n}" style:class="text"><style:paragraph-properties fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:keep-with-next="always"/><style:text-properties fo:font-size="${HEADING_SIZES[n - 1]}pt" fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"/></style:style>`;
  const styles = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${NS}>${fontDecls}<office:styles>`
    + `<style:default-style style:family="paragraph">${defaultProps ? `<style:text-properties ${defaultProps}/>` : ''}</style:default-style>`
    + '<style:style style:name="Standard" style:family="paragraph" style:class="text"/>'
    + Array.from({ length: 9 }, (_, i) => heading(i + 1)).join('')
    + '</office:styles><office:automatic-styles>'
    + `<style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="${cm(page.width || 816)}" fo:page-height="${cm(page.height || 1056)}" style:print-orientation="${page.landscape ? 'landscape' : 'portrait'}" fo:margin-top="${cm(page.top ?? 96)}" fo:margin-bottom="${cm(page.bottom ?? 96)}" fo:margin-left="${cm(page.left ?? 96)}" fo:margin-right="${cm(page.right ?? 96)}"/></style:page-layout>`
    + '</office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="PM1"/></office:master-styles></office:document-styles>';

  return odfArchive('odt', { content, styles, title: doc.title || '', pictures: [...pictures.values()] });
}
