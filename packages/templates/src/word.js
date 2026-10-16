// The gallery's documents, written through the suite's own document engine.
//
// Each template starts as the document File > New makes (buildDocx), and is
// then given what a person would give it in Documents: a theme from Design,
// a style sheet whose headings follow that theme, lists from Home, a page
// set up from Layout, a footer, a table of contents from References. The
// body is written here as WordprocessingML, paragraph by paragraph, because
// that is what a template IS; everything else goes through the engine.

import { buildDocx } from '@rutba/ooxml/build';
import { OoxmlBackend } from '@rutba/doc-view/backends/ooxml';
import { designedThemePart, themeById } from '@rutba/office-formats/themes';

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/* ---- the style sheet ---------------------------------------------------- */

const fontsOf = (slot) => `<w:rFonts w:asciiTheme="${slot}HAnsi" w:hAnsiTheme="${slot}HAnsi" w:eastAsiaTheme="${slot}EastAsia" w:cstheme="${slot}Bidi"/>`;
const themeColour = (slot, hex, shade = '') => `<w:color w:val="${hex}" w:themeColor="${slot}"${shade ? ` w:themeShade="${shade}"` : ''}/>`;

/**
 * The paragraph styles every template shares: Normal in the theme's body
 * face, Title, Subtitle, three headings in its heading face and accent
 * colour (outline levels set, so a table of contents and the navigation
 * pane find them), Quote, No Spacing, Caption, the band styles, and a plain
 * table grid. `palette` is the theme's own, for the colours a style writes
 * as a value beside its theme slot.
 */
function stylesXml(palette) {
  const p = palette;
  const style = (id, name, body, { type = 'paragraph', basedOn = 'Normal', next = null, ui = null, def = false } = {}) =>
    `<w:style w:type="${type}"${def ? ' w:default="1"' : ''} w:styleId="${id}"><w:name w:val="${name}"/>` +
    (basedOn ? `<w:basedOn w:val="${basedOn}"/>` : '') + (next ? `<w:next w:val="${next}"/>` : '') +
    (ui != null ? `<w:uiPriority w:val="${ui}"/><w:qFormat/>` : '') + body + '</w:style>';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:docDefaults><w:rPrDefault><w:rPr>' + fontsOf('minor') + '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-GB"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    style('Normal', 'Normal', '', { basedOn: null, ui: 0, def: true }) +
    style('Title', 'Title', '<w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr>' +
      `<w:rPr>${fontsOf('major')}${themeColour('text2', p.dk2)}<w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>`, { next: 'Normal', ui: 10 }) +
    style('Subtitle', 'Subtitle', '<w:pPr><w:spacing w:after="240"/></w:pPr>' +
      `<w:rPr>${themeColour('accent1', p.accent1)}<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>`, { next: 'Normal', ui: 11 }) +
    style('Heading1', 'heading 1', '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>' +
      `<w:rPr>${fontsOf('major')}${themeColour('accent1', p.accent1)}<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>`, { next: 'Normal', ui: 9 }) +
    style('Heading2', 'heading 2', '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr>' +
      `<w:rPr>${fontsOf('major')}${themeColour('text2', p.dk2)}<w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>`, { next: 'Normal', ui: 9 }) +
    style('Heading3', 'heading 3', '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="40"/><w:outlineLvl w:val="2"/></w:pPr>' +
      `<w:rPr>${fontsOf('major')}${themeColour('accent1', p.accent1)}<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>`, { next: 'Normal', ui: 9 }) +
    style('Quote', 'Quote', '<w:pPr><w:spacing w:before="200" w:after="160"/><w:ind w:left="864" w:right="864"/></w:pPr>' +
      `<w:rPr><w:i/>${themeColour('accent1', p.accent1)}</w:rPr>`, { next: 'Normal', ui: 29 }) +
    style('NoSpacing', 'No Spacing', '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>', { ui: 1 }) +
    style('Caption', 'caption', '<w:pPr><w:spacing w:after="200"/></w:pPr>' +
      `<w:rPr><w:i/>${themeColour('text2', p.dk2)}<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>`, { next: 'Normal', ui: 35 }) +
    style('Header', 'header', '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="18"/></w:rPr>') +
    style('Footer', 'footer', '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
      `<w:rPr>${themeColour('text2', p.dk2)}<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>`) +
    style('TableGrid', 'Table Grid', '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
      '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>' +
      '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>' +
      '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders>' +
      '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>', { type: 'table', basedOn: null, ui: 39 }) +
    '</w:styles>';
}

/* ---- the body ------------------------------------------------------------ */

/**
 * One run. `o`: bold, italic, caps, size (points), colour ('RRGGBB'),
 * font, spacing (twentieths of a point between letters). A tab in the text
 * is a tab, a line feed a line break.
 */
export function run(text, o = {}) {
  const rPr = [
    o.font ? `<w:rFonts w:ascii="${esc(o.font)}" w:hAnsi="${esc(o.font)}"/>` : '',
    o.bold ? '<w:b/>' : '',
    o.italic ? '<w:i/>' : '',
    o.caps ? '<w:caps/>' : '',
    o.colour ? `<w:color w:val="${o.colour}"/>` : '',
    o.spacing ? `<w:spacing w:val="${o.spacing}"/>` : '',
    o.size ? `<w:sz w:val="${Math.round(o.size * 2)}"/><w:szCs w:val="${Math.round(o.size * 2)}"/>` : '',
  ].join('');
  const parts = String(text).split(/(\t|\n)/).filter((s) => s !== '');
  const body = parts.map((s) => (s === '\t' ? '<w:tab/>' : s === '\n' ? '<w:br/>' : `<w:t xml:space="preserve">${esc(s)}</w:t>`)).join('');
  return `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${body}</w:r>`;
}

/** A PAGE field, as Word writes one, its result `1` until it is updated. */
const pageField = (o = {}) => {
  const rPr = o.colour || o.size ? `<w:rPr>${o.colour ? `<w:color w:val="${o.colour}"/>` : ''}${o.size ? `<w:sz w:val="${o.size * 2}"/>` : ''}</w:rPr>` : '';
  return `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rPr}<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>` +
    `<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${rPr}<w:t>1</w:t></w:r><w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`;
};

/**
 * One paragraph. `content` is text or an array of runs (strings, or
 * [text, runOptions]). `o`: style, align, before, after, line (240 = single),
 * keepNext, pageBreakBefore, tabs [[pos, 'left'|'right'|'center', leader?]],
 * fill ('RRGGBB'), rule ({ side, colour, size, space }), indent { left,
 * right, hanging }, and a run look for plain text.
 */
export function para(content = '', o = {}) {
  const pPr = [
    o.style ? `<w:pStyle w:val="${o.style}"/>` : '',
    o.keepNext ? '<w:keepNext/>' : '',
    o.pageBreakBefore ? '<w:pageBreakBefore/>' : '',
    o.rule ? `<w:pBdr><w:${o.rule.side || 'bottom'} w:val="single" w:sz="${o.rule.size || 8}" w:space="${o.rule.space ?? 4}" w:color="${o.rule.colour}"/></w:pBdr>` : '',
    o.fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${o.fill}"/>` : '',
    o.tabs ? `<w:tabs>${o.tabs.map(([pos, val, leader]) => `<w:tab w:val="${val}"${leader ? ` w:leader="${leader}"` : ''} w:pos="${pos}"/>`).join('')}</w:tabs>` : '',
    o.before != null || o.after != null || o.line != null
      ? `<w:spacing${o.before != null ? ` w:before="${o.before}"` : ''}${o.after != null ? ` w:after="${o.after}"` : ''}${o.line != null ? ` w:line="${o.line}" w:lineRule="auto"` : ''}/>`
      : '',
    o.indent ? `<w:ind${o.indent.left != null ? ` w:left="${o.indent.left}"` : ''}${o.indent.right != null ? ` w:right="${o.indent.right}"` : ''}${o.indent.hanging != null ? ` w:hanging="${o.indent.hanging}"` : ''}/>` : '',
    o.align ? `<w:jc w:val="${{ left: 'left', center: 'center', right: 'right', justify: 'both' }[o.align]}"/>` : '',
  ].join('');
  const runs = typeof content === 'string'
    ? (content ? run(content, o.run || {}) : '')
    : content.map((r) => (typeof r === 'string' ? (r.startsWith('<w:') ? r : run(r, o.run || {})) : run(r[0], { ...(o.run || {}), ...(r[1] || {}) }))).join('');
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
}

/**
 * A table. `rows` are arrays of cells; a cell is text, an array of
 * paragraphs' XML, or { content, fill, bold, align, colour }. `o`: widths
 * (twips), borders ('grid' | 'rows' | 'none'), header (shade the first row
 * in `headerFill`, its words white and bold), align per column.
 */
export function table(rows, o = {}) {
  const widths = o.widths;
  const total = widths.reduce((a, b) => a + b, 0);
  const line = (side, colour = 'BFBFBF', sz = 4) => `<w:${side} w:val="single" w:sz="${sz}" w:space="0" w:color="${colour}"/>`;
  const nil = (side) => `<w:${side} w:val="nil"/>`;
  const borders = o.borders === 'none'
    ? ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(nil).join('')
    : o.borders === 'rows'
      ? line('top', o.ruleColour || 'BFBFBF', 8) + nil('left') + line('bottom', o.ruleColour || 'BFBFBF', 8) + nil('right') + line('insideH', 'D9D9D9') + nil('insideV')
      : ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => line(s)).join('');
  const tblPr = `<w:tblPr>${o.borders === 'none' ? '' : '<w:tblStyle w:val="TableGrid"/>'}<w:tblW w:w="${total}" w:type="dxa"/>` +
    `<w:tblBorders>${borders}</w:tblBorders><w:tblLayout w:type="fixed"/>` +
    `<w:tblCellMar><w:top w:w="${o.padding ?? 60}" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="${o.padding ?? 60}" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>` +
    '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>';
  const grid = `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
  const spanOf = (cell) => (cell && typeof cell === 'object' && !Array.isArray(cell) && cell.span) || 1;
  const cellXml = (cell, c, r, col) => {
    const spec = cell && typeof cell === 'object' && !Array.isArray(cell) ? cell : { content: cell };
    const width = widths.slice(col, col + spanOf(cell)).reduce((a, b) => a + b, 0);
    const isHeader = o.header && r === 0;
    const fill = spec.fill || (isHeader ? o.headerFill : null);
    const align = spec.align || o.align?.[c] || null;
    const look = { bold: spec.bold || isHeader, colour: spec.colour || (isHeader ? 'FFFFFF' : null), size: spec.size || o.size || null };
    const content = Array.isArray(spec.content)
      ? spec.content.join('')
      : para(String(spec.content ?? ''), { style: 'NoSpacing', align, run: look });
    return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${spec.span ? `<w:gridSpan w:val="${spec.span}"/>` : ''}` +
      `${fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : ''}<w:vAlign w:val="${spec.valign || 'center'}"/></w:tcPr>${content}</w:tc>`;
  };
  const rowXml = (cells, r) => `<w:tr>${o.header && r === 0 ? '<w:trPr><w:tblHeader/></w:trPr>' : (o.rowHeight ? `<w:trPr><w:trHeight w:val="${o.rowHeight}"/></w:trPr>` : '')}${cells.map((cell, c) => cellXml(cell, c, r, cells.slice(0, c).reduce((n, x) => n + spanOf(x), 0))).join('')}</w:tr>`;
  return `<w:tbl>${tblPr}${grid}${rows.map(rowXml).join('')}</w:tbl>`;
}

/** How many edit-space paragraphs a piece of body XML holds (a table's cells count each of theirs). */
const paragraphsIn = (xml) => (xml.match(/<w:p\b/g) || []).length;

/** A4, with margins in twips. */
const sectPr = ({ top = 1440, right = 1440, bottom = 1440, left = 1440 } = {}) =>
  `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="${top}" w:right="${right}" w:bottom="${bottom}" w:left="${left}" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/></w:sectPr>`;

/**
 * A document from body pieces: File > New's blank document, given the
 * theme, the style sheet, the body and the page, then whatever `finish`
 * does through the engine (lists, bands, a table of contents). Each piece
 * is the XML of a paragraph or a table; `lists` names the pieces that are
 * bulleted or numbered, by their place in `pieces`.
 */
export function compose({ theme, pieces, lists = {}, margins, footer = null, finish = null }) {
  const backend = OoxmlBackend.open(buildDocx({ paragraphs: [{ text: '' }], styles: true }));
  const doc = backend.doc;
  const t = themeById(theme);
  // Design: the theme, and the style sheet whose headings follow it.
  doc.setThemePart(designedThemePart(null, { theme }));
  doc.pkg.write_('word/styles.xml', Buffer.from(stylesXml(t.palette), 'utf8'));
  // The body, piece by piece, and where each piece's first paragraph sits.
  const starts = [];
  const prose = [];
  let at = 0;
  let top = 0;
  for (const piece of pieces) {
    starts.push(at);
    prose.push(top);
    at += paragraphsIn(piece);
    if (piece.startsWith('<w:p>') || piece.startsWith('<w:p ')) top += 1;
  }
  const { prefix, suffix } = doc._body();
  doc.xml = prefix + pieces.join('') + sectPr(margins) + suffix;
  doc.dirty = true;
  // Home > Bullets and Numbering, through the engine, so the numbering part
  // is the one Documents writes.
  for (const [kind, indexes] of Object.entries(lists)) {
    for (const i of indexes) backend.setParagraphList(starts[i], kind);
  }
  if (footer) doc._addBand('footer', footer);
  if (finish) finish({ backend, doc, starts, prose });
  return Buffer.from(doc.save());
}

/** The theme's colours, as hex without the mark. */
const paletteOf = (theme) => themeById(theme).palette;

/* ---- the templates ------------------------------------------------------- */

function letter() {
  const theme = 'rutba';
  const p = paletteOf(theme);
  const pieces = [
    para('Your Name', { style: 'Title', after: 40, run: { size: 26 } }),
    para([['Street address, Town, Postcode', { colour: p.dk2 }], ['\t'], ['phone  |  you@example.com', { colour: p.dk2 }]], {
      style: 'NoSpacing', tabs: [[9026, 'right']], rule: { colour: p.accent1, size: 12, space: 6 }, run: { size: 9 },
    }),
    para('', { after: 360 }),
    para('[Date]', { after: 360 }),
    para('[Recipient name]', { style: 'NoSpacing' }),
    para('[Title]', { style: 'NoSpacing' }),
    para('[Company]', { style: 'NoSpacing' }),
    para('[Street address]', { style: 'NoSpacing' }),
    para('[Town, Postcode]', { after: 360 }),
    para([['Subject: ', { bold: true }], ['[What this letter is about]', { bold: true }]], { after: 240 }),
    para('Dear [Recipient name],'),
    para('Begin with the reason you are writing, in a sentence or two. A reader who knows why the letter has come reads the rest of it with that in mind.'),
    para('Use the paragraphs that follow for the details: what has happened, what you are asking for or offering, and any dates, amounts or references the reader will need. Keep to one subject a paragraph.'),
    para('End by saying what happens next and how the reader can reach you. Thank them for their time.'),
    para('Yours sincerely,', { before: 240, after: 720 }),
    para('Your Name', { style: 'NoSpacing', run: { bold: true } }),
    para('Your job title', { style: 'NoSpacing', run: { colour: p.dk2 } }),
  ];
  const footer = para([['Your Company Ltd  |  Registered in England and Wales, number 00000000  |  example.com', { colour: p.dk2 }]], { style: 'Footer', align: 'center' });
  return compose({ theme, pieces, footer, margins: { top: 1134, bottom: 1134, left: 1440, right: 1440 } });
}

function cv() {
  const theme = 'paperkite';
  const p = paletteOf(theme);
  const role = (title, where, when, lines) => [
    para([[title, { bold: true }], [`, ${where}`]], { style: 'Heading3', tabs: [[9026, 'right']], after: 0 }),
    para(when, { style: 'NoSpacing', after: 80, run: { italic: true, colour: p.dk2, size: 10 } }),
    ...lines.map((l) => para(l, { after: 40 })),
  ];
  const pieces = [
    para('Your Name', { style: 'Title', after: 0 }),
    para('Job title you are looking for', { style: 'Subtitle', after: 80 }),
    para('Town, Country  |  you@example.com  |  07000 000000  |  example.com/yourname', { style: 'NoSpacing', rule: { colour: p.accent1, size: 12, space: 8 }, run: { size: 9.5, colour: p.dk2 } }),
    para('Profile', { style: 'Heading1' }),
    para('Two or three sentences about who you are at work: what you are good at, the kind of work you have done and the kind you are looking for. Write it for the job you want, not the last one you had.'),
    para('Experience', { style: 'Heading1' }),
    ...role('Senior Project Coordinator', 'Example Company', '2022 to present  |  London', [
      'What you were responsible for, in one line.',
      'An achievement with a number in it: reduced turnaround from ten days to four.',
      'A second achievement, or something you built, led or improved.',
    ]),
    ...role('Project Coordinator', 'Another Company', '2018 to 2022  |  Manchester', [
      'What you were responsible for, in one line.',
      'What you are proudest of from this job.',
    ]),
    para('Education', { style: 'Heading1' }),
    ...role('BSc Business Management', 'University name', '2015 to 2018', ['Your grade, and a project or prize worth a mention.']),
    para('Skills', { style: 'Heading1' }),
    para('Planning and scheduling', { after: 40 }),
    para('Budgets and reporting', { after: 40 }),
    para('Spreadsheets, documents and presentations', { after: 40 }),
    para('Languages: English (native), French (working)', { after: 40 }),
  ];
  // The lines under each role and the skills are bullets.
  const bullets = pieces.map((x, i) => i).filter((i) => {
    const x = pieces[i];
    return !/w:pStyle/.test(x) && i > 4;
  });
  return compose({ theme, pieces, lists: { bullet: bullets }, margins: { top: 1080, bottom: 1080, left: 1260, right: 1260 } });
}

function report() {
  const theme = 'harbour';
  const p = paletteOf(theme);
  const pieces = [
    // The cover: a band in the theme's colour, the title in it, and who and when below.
    para('', { after: 2400 }),
    table([[{
      fill: p.accent1,
      content: [
        para('ANNUAL REPORT', { before: 240, after: 120, run: { colour: 'FFFFFF', size: 11, bold: true, spacing: 40 } }),
        para('Report Title', { style: 'Title', after: 60, run: { colour: 'FFFFFF', size: 40 } }),
        para('A subtitle that says what the report covers', { style: 'Subtitle', after: 240, run: { colour: 'FFFFFF' } }),
      ],
    }]], { widths: [9026], borders: 'none', padding: 120 }),
    para('', { after: 3600 }),
    para('Prepared by [Your name]', { style: 'NoSpacing', run: { bold: true, colour: p.dk2 } }),
    para('[Organisation]', { style: 'NoSpacing', run: { colour: p.dk2 } }),
    para('[Date]', { style: 'NoSpacing', run: { colour: p.dk2 } }),
    // Contents, on a page of its own; the table is inserted after this heading.
    para('Contents', { style: 'Title', pageBreakBefore: true, run: { size: 22 }, after: 240 }),
    // The sections.
    para('Introduction', { style: 'Heading1', pageBreakBefore: true }),
    para('Say what the report is about, who it is for and what question it answers. A reader should know from this section alone whether the rest is worth reading.'),
    para('Background', { style: 'Heading2' }),
    para('What the reader needs to know first: how things stood before, and why the question came up now.'),
    para('Scope', { style: 'Heading2' }),
    para('What the report covers, and what it leaves out on purpose.'),
    para('Findings', { style: 'Heading1', pageBreakBefore: true }),
    para('Set out what was found, most important first. Keep each finding to a paragraph, with the evidence for it.'),
    para('Results', { style: 'Heading2' }),
    table([
      ['Measure', 'Last year', 'This year', 'Change'],
      ['Customers', '1,240', '1,515', '+22%'],
      ['Orders', '3,980', '4,610', '+16%'],
      ['Average order', '£42.10', '£44.80', '+6%'],
      ['Returns', '3.1%', '2.4%', '-0.7 pts'],
    ], { widths: [3626, 1800, 1800, 1800], header: true, headerFill: p.accent1, borders: 'rows', align: [null, 'right', 'right', 'right'] }),
    para('Table 1: the year in four measures.', { style: 'Caption', before: 120 }),
    para('What the results mean', { style: 'Heading2' }),
    para('Explain what the figures show and what they do not. A table answers "how much"; this paragraph answers "so what".'),
    para('Recommendations', { style: 'Heading1', pageBreakBefore: true }),
    para('Give each recommendation as an action someone can take, with who should take it and by when.'),
    para('Next steps', { style: 'Heading2' }),
    para('Agree the recommendations and who owns each.', { after: 40 }),
    para('Set a date to review progress.', { after: 40 }),
    para('Share the report with everyone it affects.', { after: 40 }),
  ];
  const lastThree = [pieces.length - 3, pieces.length - 2, pieces.length - 1];
  const footer = para([['Report Title  |  Page ', { colour: p.dk2 }], pageField({ colour: p.dk2 })], { style: 'Footer', align: 'right' });
  // Each first-level heading starts a page: the cover is page 1 and the
  // contents page 2, so the sections begin on 3, 4 and 5, and the second-
  // level headings sit on their section's page.
  const pages = ['3', '3', '3', '4', '4', '4', '5', '5'];
  return compose({
    theme, pieces, lists: { number: lastThree }, footer,
    finish: ({ doc, prose }) => doc.insertTableOfContents({ at: prose[6], levels: 2, pages }),
  });
}

function invoiceDoc() {
  const theme = 'basalt';
  const p = paletteOf(theme);
  const money = (n) => '£' + n.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const items = [
    ['Design consultation (hours)', 6, 65],
    ['Website pages', 5, 180],
    ['Photography, half day', 1, 350],
    ['Hosting, twelve months', 1, 120],
  ];
  const subtotal = items.reduce((s, [, q, u]) => s + q * u, 0);
  const vat = Math.round(subtotal * 0.2 * 100) / 100;
  const right = (text, o = {}) => ({ content: text, align: 'right', ...o });
  const pieces = [
    table([[
      { content: [para('Your Company Ltd', { style: 'NoSpacing', run: { bold: true, size: 14 } }), para('Street address, Town, Postcode', { style: 'NoSpacing', run: { size: 9, colour: p.dk2 } }), para('hello@example.com  |  07000 000000', { style: 'NoSpacing', run: { size: 9, colour: p.dk2 } })] },
      { content: [para('INVOICE', { style: 'Title', align: 'right', after: 0, run: { colour: p.accent1, spacing: 40 } })], valign: 'top' },
    ]], { widths: [5513, 3513], borders: 'none' }),
    para('', { after: 240, rule: { colour: p.accent1, size: 12, space: 4 } }),
    table([
      [{ content: 'BILL TO', bold: true, colour: p.accent1, size: 9 }, { content: 'INVOICE NUMBER', bold: true, colour: p.accent1, size: 9, align: 'right' }, { content: 'DATE', bold: true, colour: p.accent1, size: 9, align: 'right' }],
      ['[Customer name]', right('INV-0001'), right('[Date]')],
      ['[Company]', { content: 'PAYMENT DUE', bold: true, colour: p.accent1, size: 9, align: 'right' }, ''],
      ['[Street address, Town, Postcode]', right('Within 30 days'), ''],
    ], { widths: [4626, 2200, 2200], borders: 'none', padding: 20 }),
    para('', { after: 240 }),
    table([
      ['Description', 'Quantity', 'Unit price', 'Amount'],
      ...items.map(([d, q, u]) => [d, right(String(q)), right(money(u)), right(money(q * u))]),
      ['', '', right('Subtotal', { bold: true }), right(money(subtotal))],
      ['', '', right('VAT at 20%', { bold: true }), right(money(vat))],
      ['', '', right('Total due', { bold: true, fill: p.lt2 }), right(money(subtotal + vat), { bold: true, fill: p.lt2 })],
    ], { widths: [4426, 1400, 1600, 1600], header: true, headerFill: p.accent1, borders: 'rows' }),
    para('', { after: 240 }),
    para('How to pay', { style: 'Heading2' }),
    para('Bank transfer to Your Company Ltd, sort code 00-00-00, account 00000000. Please give the invoice number as the reference.', { after: 80 }),
    para('Thank you for your business.', { before: 240, run: { italic: true, colour: p.dk2 } }),
  ];
  const footer = para([['Your Company Ltd  |  Company number 00000000  |  VAT number GB000 0000 00', { colour: p.dk2 }]], { style: 'Footer', align: 'center' });
  return compose({ theme, pieces, footer, margins: { top: 1134, bottom: 1134, left: 1440, right: 1440 } });
}

function minutes() {
  const theme = 'studio';
  const p = paletteOf(theme);
  const label = (text) => ({ content: text, bold: true, colour: p.dk2, fill: p.lt2 });
  const pieces = [
    para('Meeting Minutes', { style: 'Title' }),
    para('[Name of the group or project]', { style: 'Subtitle' }),
    table([
      [label('Date'), '[Date]', label('Time'), '10:00 to 11:00'],
      [label('Place'), '[Room, or the link to join]', label('Chair'), '[Name]'],
      [label('Present'), { content: '[Names of those who came]', span: 3 }],
      [label('Apologies'), { content: '[Names of those who could not come]', span: 3 }],
    ], { widths: [1500, 3013, 1500, 3013], borders: 'grid' }),
    para('Agenda', { style: 'Heading1' }),
    para('Welcome and apologies', { after: 40 }),
    para('Minutes of the last meeting', { after: 40 }),
    para('Progress report', { after: 40 }),
    para('Budget for the next quarter', { after: 40 }),
    para('Any other business', { after: 160 }),
    para('1. Minutes of the last meeting', { style: 'Heading2' }),
    para('The minutes of the meeting on [date] were agreed as a true record.'),
    para('2. Progress report', { style: 'Heading2' }),
    para('[Name] reported on progress since the last meeting. Summarise what was said in a few sentences, and the questions that were asked.'),
    para('Decision: the group agreed to [what was decided].', { run: { bold: true } }),
    para('3. Budget for the next quarter', { style: 'Heading2' }),
    para('Summarise the discussion. Record decisions as they were agreed, in the words the meeting used.'),
    para('Actions', { style: 'Heading1' }),
    table([
      ['Action', 'Who', 'By when'],
      ['Send the revised budget to everyone', '[Name]', '[Date]'],
      ['Book the room for the next meeting', '[Name]', '[Date]'],
      ['Ask the supplier for a second quote', '[Name]', '[Date]'],
    ], { widths: [5426, 1800, 1800], header: true, headerFill: p.accent1, borders: 'rows' }),
    para('Next meeting: [date and time], [place].', { before: 240, run: { bold: true } }),
  ];
  // The agenda's five items are numbered.
  return compose({ theme, pieces, lists: { number: [4, 5, 6, 7, 8] } });
}

function newsletter() {
  const theme = 'orchard';
  const p = paletteOf(theme);
  const story = (heading, lines) => [para(heading, { style: 'Heading2', before: 0 }), ...lines.map((l) => para(l, { run: { size: 10 } }))];
  const pieces = [
    // The masthead: the name across the page in the theme's colour.
    para('THE NEWSLETTER', { style: 'Title', fill: p.accent1, align: 'center', before: 0, after: 0, run: { colour: 'FFFFFF', size: 36, spacing: 60 } }),
    para('News, dates and stories from [your organisation]', { fill: p.accent1, align: 'center', after: 0, run: { colour: 'FFFFFF', italic: true } }),
    para([['Issue 1', { bold: true }], '\t', ['[Month and year]', {}]], { tabs: [[9026, 'right']], before: 120, after: 240, rule: { colour: p.accent1, size: 8, space: 4 }, run: { colour: p.dk2, size: 9 } }),
    // Two columns of stories, side by side.
    table([[
      {
        valign: 'top',
        content: [
          ...story('Our lead story', [
            'Start with the story most readers will want: what happened, who it matters to and what comes next.',
            'Keep paragraphs short. A newsletter is read in a few minutes, often on a phone.',
          ]),
          para('"A good quote from someone in the story brings it to life."', { style: 'Quote', indent: { left: 240, right: 240 } }),
          ...story('Welcome to new members', ['Introduce the people who have joined since the last issue, with a line about each.']),
          ...story('Spotlight', [
            'A longer piece about one person, project or place: what they do, why it matters and what readers can learn from it.',
            'End with how to find out more, or how to take part.',
          ]),
        ],
      },
      {
        valign: 'top',
        content: [
          ...story('News in brief', [
            'A short item: one or two sentences.',
            'Another short item, with a date or a link to find out more.',
          ]),
          ...story('From the team', ['A word from someone readers know: what the team has been working on and what is planned.']),
          para('DATES FOR THE DIARY', { fill: p.lt2, before: 120, after: 0, indent: { left: 120 }, run: { bold: true, colour: p.accent1, size: 10, spacing: 20 } }),
          para('[Date]  Open day', { fill: p.lt2, after: 0, indent: { left: 120 }, run: { size: 10 } }),
          para('[Date]  Annual meeting', { fill: p.lt2, after: 0, indent: { left: 120 }, run: { size: 10 } }),
          para('[Date]  Summer party', { fill: p.lt2, after: 120, indent: { left: 120 }, run: { size: 10 } }),
          ...story('Get involved', [
            'Tell readers how they can help: volunteer, come along, or send in a story for the next issue.',
            'Write to hello@example.com or call 07000 000000.',
          ]),
        ],
      },
    ]], { widths: [4513, 4513], borders: 'none', padding: 0 }),
  ];
  const footer = para([['[Your organisation]  |  example.com  |  hello@example.com', { colour: p.dk2 }]], { style: 'Footer', align: 'center' });
  return compose({ theme, pieces, footer, margins: { top: 1080, bottom: 1080, left: 1440, right: 1440 } });
}

/** The documents, by template id. */
export const DOCUMENTS = {
  'word-letter': letter,
  'word-cv': cv,
  'word-report': report,
  'word-invoice': invoiceDoc,
  'word-minutes': minutes,
  'word-newsletter': newsletter,
};
