/**
 * Mailings → Envelopes and Labels — the stock, and the XML Word builds from it.
 *
 * Plain data and string builders, no Node module: the window lists the same
 * envelope sizes and label products the document is built from. An envelope
 * is a section of its own at the front of the letter, on the envelope's own
 * paper, turned landscape the way it feeds: the return address as ordinary
 * paragraphs in the corner, the delivery address in a frame placed where
 * Word's Envelope Address style places it. A sheet of labels is a table the
 * size of the sheet — a column per label and a narrow one for each gap
 * between, rows of the label's exact height, no borders — with the sheet's
 * top and side margins as the page's, which is how Word lays one out.
 */

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const mmTw = (mm) => Math.round((Number(mm) * 1440) / 25.4);
const inTw = (inches) => Math.round(Number(inches) * 1440);

/**
 * Envelope sizes as Word lists them, width by height as the envelope lies
 * (landscape), with the paper number Windows gives each (`w:pgSz w:code`).
 */
export const ENVELOPE_SIZES = [
  { id: 'DL', label: 'DL — 110 × 220 mm', w: mmTw(220), h: mmTw(110), code: 27 },
  { id: 'C5', label: 'C5 — 162 × 229 mm', w: mmTw(229), h: mmTw(162), code: 28 },
  { id: 'C6', label: 'C6 — 114 × 162 mm', w: mmTw(162), h: mmTw(114), code: 31 },
  { id: 'C4', label: 'C4 — 229 × 324 mm', w: mmTw(324), h: mmTw(229), code: 30 },
  { id: 'B5', label: 'B5 — 176 × 250 mm', w: mmTw(250), h: mmTw(176), code: 34 },
  { id: 'No10', label: 'Size 10 — 4⅛ × 9½ in', w: inTw(9.5), h: inTw(4.125), code: 20 },
  { id: 'No9', label: 'Size 9 — 3⅞ × 8⅞ in', w: inTw(8.875), h: inTw(3.875), code: 19 },
  { id: 'Monarch', label: 'Monarch — 3⅞ × 7½ in', w: inTw(7.5), h: inTw(3.875), code: 37 },
  { id: 'Size634', label: 'Size 6¾ — 3⅝ × 6½ in', w: inTw(6.5), h: inTw(3.625), code: 38 },
];

/**
 * Label products, table-driven: the page, the sheet's top and side margins,
 * the pitch (label edge to next label edge, across and down), the label's
 * own size, and how many across and down — the numbers Avery prints on
 * every pack, in millimetres.
 */
const A4 = { w: 210, h: 297, name: 'A4' };
const LETTER = { w: 215.9, h: 279.4, name: 'Letter' };
const inch = (n) => n * 25.4;
export const LABEL_PRODUCTS = [
  { id: 'L7160', vendor: 'Avery A4/A5', name: 'L7160', kind: 'Address', page: A4, top: 15.15, side: 7.25, pitchX: 66.04, pitchY: 38.1, w: 63.5, h: 38.1, cols: 3, rows: 7 },
  { id: 'L7159', vendor: 'Avery A4/A5', name: 'L7159', kind: 'Address', page: A4, top: 12.9, side: 7.25, pitchX: 66.04, pitchY: 33.9, w: 63.5, h: 33.9, cols: 3, rows: 8 },
  { id: 'L7161', vendor: 'Avery A4/A5', name: 'L7161', kind: 'Address', page: A4, top: 8.8, side: 7.25, pitchX: 66.04, pitchY: 46.6, w: 63.5, h: 46.6, cols: 3, rows: 6 },
  { id: 'L7162', vendor: 'Avery A4/A5', name: 'L7162', kind: 'Address', page: A4, top: 12.9, side: 4.65, pitchX: 101.6, pitchY: 33.9, w: 99.1, h: 33.9, cols: 2, rows: 8 },
  { id: 'L7163', vendor: 'Avery A4/A5', name: 'L7163', kind: 'Address', page: A4, top: 15.15, side: 4.65, pitchX: 101.6, pitchY: 38.1, w: 99.1, h: 38.1, cols: 2, rows: 7 },
  { id: 'L7165', vendor: 'Avery A4/A5', name: 'L7165', kind: 'Parcel', page: A4, top: 13.1, side: 4.65, pitchX: 101.6, pitchY: 67.7, w: 99.1, h: 67.7, cols: 2, rows: 4 },
  { id: 'L7173', vendor: 'Avery A4/A5', name: 'L7173', kind: 'Shipping', page: A4, top: 6, side: 4.65, pitchX: 101.6, pitchY: 57, w: 99.1, h: 57, cols: 2, rows: 5 },
  { id: '3422', vendor: 'Avery Zweckform', name: '3422', kind: 'Multipurpose', page: A4, top: 8.5, side: 0, pitchX: 70, pitchY: 35, w: 70, h: 35, cols: 3, rows: 8 },
  { id: '3475', vendor: 'Avery Zweckform', name: '3475', kind: 'Multipurpose', page: A4, top: 4.5, side: 0, pitchX: 70, pitchY: 36, w: 70, h: 36, cols: 3, rows: 8 },
  { id: '5160', vendor: 'Avery US Letter', name: '5160', kind: 'Address', page: LETTER, top: inch(0.5), side: inch(0.1875), pitchX: inch(2.75), pitchY: inch(1), w: inch(2.625), h: inch(1), cols: 3, rows: 10 },
  { id: '5161', vendor: 'Avery US Letter', name: '5161', kind: 'Address', page: LETTER, top: inch(0.5), side: inch(0.15625), pitchX: inch(4.1875), pitchY: inch(1), w: inch(4), h: inch(1), cols: 2, rows: 10 },
  { id: '5162', vendor: 'Avery US Letter', name: '5162', kind: 'Address', page: LETTER, top: inch(0.83333), side: inch(0.15625), pitchX: inch(4.1875), pitchY: inch(1.33333), w: inch(4), h: inch(1.33333), cols: 2, rows: 7 },
  { id: '5163', vendor: 'Avery US Letter', name: '5163', kind: 'Shipping', page: LETTER, top: inch(0.5), side: inch(0.15625), pitchX: inch(4.1875), pitchY: inch(2), w: inch(4), h: inch(2), cols: 2, rows: 5 },
  { id: '5164', vendor: 'Avery US Letter', name: '5164', kind: 'Shipping', page: LETTER, top: inch(0.5), side: inch(0.15625), pitchX: inch(4.1875), pitchY: inch(3.33333), w: inch(4), h: inch(3.33333), cols: 2, rows: 3 },
];

export const labelProduct = (id) => LABEL_PRODUCTS.find((p) => p.id === id) || LABEL_PRODUCTS[0];

/** How a product reads in a list: "L7160 Address — 63.5 × 38.1 mm, 21 on an A4 sheet". */
export function describeLabel(p) {
  return `${p.name} ${p.kind} — ${round1(p.w)} × ${round1(p.h)} mm, ${p.cols * p.rows} on ${p.page.name === 'A4' ? 'an A4' : 'a Letter'} sheet`;
}
const round1 = (n) => Math.round(n * 10) / 10;

/* ── envelopes ─────────────────────────────────────────────────────────── */

/** Word's two envelope styles, for a styles part that has neither. */
export const ENVELOPE_STYLES_XML =
  '<w:style w:type="paragraph" w:styleId="EnvelopeAddress"><w:name w:val="envelope address"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>'
  + '<w:pPr><w:framePr w:w="7920" w:h="1980" w:hRule="exact" w:hSpace="180" w:wrap="auto" w:hAnchor="page" w:xAlign="center" w:yAnchor="page" w:yAlign="bottom"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="2880"/></w:pPr>'
  + '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="EnvelopeReturn"><w:name w:val="envelope return"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>'
  + '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
  + '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>';

/** An envelope's margins, as Word sets them: a quarter inch at the top, four-tenths at the sides. */
const ENVELOPE_MARGINS = { top: 360, right: 576, bottom: 720, left: 576, header: 720, footer: 720, gutter: 0 };

/** Where the delivery address frame goes on an envelope of this size — Word's Auto: centred across, at the foot. */
export function deliveryFrame(size) {
  const w = 7920;
  const h = 1980;
  return { w, h, x: Math.max(0, Math.round((size.w - w) / 2)), y: Math.max(0, size.h - h), indent: 2880 };
}

/** A run's look from a font choice: `{ font, sizePt }`. */
function fontRPr(font) {
  if (!font || (!font.name && !font.sizePt)) return '';
  const bits = [];
  if (font.name) bits.push('<w:rFonts w:ascii="' + esc(font.name) + '" w:hAnsi="' + esc(font.name) + '" w:cs="' + esc(font.name) + '"/>');
  if (font.sizePt) bits.push('<w:sz w:val="' + Math.round(font.sizePt * 2) + '"/><w:szCs w:val="' + Math.round(font.sizePt * 2) + '"/>');
  return '<w:rPr>' + bits.join('') + '</w:rPr>';
}

/** Lines as one paragraph's runs, a line break between each — the way a label and an envelope hold an address. */
function linesXml(lines, rPr) {
  return (lines || []).map((l, i) => (i ? '<w:r>' + rPr + '<w:br/></w:r>' : '') + (l ? '<w:r>' + rPr + '<w:t xml:space="preserve">' + esc(l) + '</w:t></w:r>' : '')).join('');
}

/** An envelope section's properties: its paper, landscape, Word's envelope margins, page numbers starting again after it. */
export function envelopeSectPr(size) {
  const m = ENVELOPE_MARGINS;
  return '<w:sectPr><w:pgSz w:w="' + size.w + '" w:h="' + size.h + '" w:orient="landscape" w:code="' + size.code + '"/>'
    + '<w:pgMar w:top="' + m.top + '" w:right="' + m.right + '" w:bottom="' + m.bottom + '" w:left="' + m.left + '" w:header="' + m.header + '" w:footer="' + m.footer + '" w:gutter="0"/>'
    + '<w:pgNumType w:start="0"/><w:cols w:space="720"/></w:sectPr>';
}

/**
 * Add to Document: the envelope's paragraphs, its section ending in the last
 * — the return address (unless omitted) in the Envelope Return style at the
 * top left, the delivery address in the Envelope Address style's frame.
 * `delivery` and `returnAddress` are arrays of lines.
 */
export function envelopeXml({ size = ENVELOPE_SIZES[0], delivery = [], returnAddress = null, deliveryFont = null, returnFont = null } = {}) {
  const f = deliveryFrame(size);
  const ret = returnAddress && returnAddress.some((l) => String(l).trim())
    ? '<w:p><w:pPr><w:pStyle w:val="EnvelopeReturn"/></w:pPr>' + linesXml(returnAddress, fontRPr(returnFont)) + '</w:p>'
    : '';
  const del = '<w:p><w:pPr><w:pStyle w:val="EnvelopeAddress"/>'
    + '<w:framePr w:w="' + f.w + '" w:h="' + f.h + '" w:hRule="exact" w:hSpace="180" w:wrap="auto" w:hAnchor="page" w:vAnchor="page" w:x="' + f.x + '" w:y="' + f.y + '"/>'
    + '</w:pPr>' + linesXml(delivery, fontRPr(deliveryFont)) + '</w:p>';
  return ret + del + '<w:p><w:pPr>' + envelopeSectPr(size) + '</w:pPr></w:p>';
}

/**
 * Start Mail Merge → Envelopes: the whole document is one envelope — the
 * return address, and the delivery address frame left empty for an Address
 * Block — on the envelope's paper, as the document's own section.
 */
export function envelopeDocumentXml({ size = ENVELOPE_SIZES[0], delivery = [], returnAddress = null, returnFont = null, deliveryFont = null } = {}) {
  // Envelopes → Print hands the addresses in; a merge leaves the frame empty.
  const xml = envelopeXml({ size, delivery, returnAddress, returnFont, deliveryFont });
  const body = xml.replace(/<w:p><w:pPr><w:sectPr>[\s\S]*<\/w:sectPr><\/w:pPr><\/w:p>$/, '');
  return { body, sectPr: envelopeSectPr(size).replace('<w:pgNumType w:start="0"/>', '') };
}

/* ── labels ────────────────────────────────────────────────────────────── */

const LABEL_P_PR = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="95" w:right="95"/></w:pPr>';

/**
 * A sheet of labels as Word builds one: a table exactly the sheet's width,
 * a column per label and a narrow one for each gap between, rows of the
 * label's height held exact, no borders, the words centred down each label;
 * the page is the sheet's paper with its top and side margins, and a last
 * paragraph too small to spill onto a second page.
 *
 * `cell(row, col)` answers each label's inner XML (its paragraphs).
 */
export function labelSheetXml(product, cell) {
  const p = product;
  const labelW = mmTw(p.w);
  const labelH = mmTw(p.h);
  const gapX = mmTw(p.pitchX - p.w);
  const gapY = mmTw(p.pitchY - p.h);
  const grid = [];
  for (let c = 0; c < p.cols; c++) {
    grid.push(labelW);
    if (c < p.cols - 1 && gapX > 10) grid.push(gapX);
  }
  const total = grid.reduce((a, b) => a + b, 0);
  const none = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => '<w:' + s + ' w:val="nil"/>').join('');
  let xml = '<w:tbl><w:tblPr><w:tblW w:w="' + total + '" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
    + '<w:tblBorders>' + none + '</w:tblBorders>'
    + '<w:tblCellMar><w:left w:w="15" w:type="dxa"/><w:right w:w="15" w:type="dxa"/></w:tblCellMar>'
    + '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr>'
    + '<w:tblGrid>' + grid.map((w) => '<w:gridCol w:w="' + w + '"/>').join('') + '</w:tblGrid>';
  const spacerCell = (w) => '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/></w:tcPr><w:p>' + LABEL_P_PR + '</w:p></w:tc>';
  for (let r = 0; r < p.rows; r++) {
    xml += '<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="' + labelH + '" w:hRule="exact"/></w:trPr>';
    for (let c = 0; c < p.cols; c++) {
      xml += '<w:tc><w:tcPr><w:tcW w:w="' + labelW + '" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>' + (cell(r, c) || '<w:p>' + LABEL_P_PR + '</w:p>') + '</w:tc>';
      if (c < p.cols - 1 && gapX > 10) xml += spacerCell(gapX);
    }
    xml += '</w:tr>';
    if (r < p.rows - 1 && gapY > 10) {
      xml += '<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="' + gapY + '" w:hRule="exact"/></w:trPr>' + grid.map((w) => spacerCell(w)).join('') + '</w:tr>';
    }
  }
  xml += '</w:tbl>';
  // A document cannot end in a table; this paragraph is one point high, so it
  // fits in what the sheet leaves at its foot instead of making a second page.
  xml += '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';
  return { body: xml, sectPr: labelSectPr(p), total };
}

/** A label sheet's section: the sheet's paper, its top and side margins, nothing at the foot. */
export function labelSectPr(p) {
  const pageW = mmTw(p.page.w);
  const pageH = mmTw(p.page.h);
  const top = mmTw(p.top);
  const side = mmTw(p.side);
  const used = mmTw(p.pitchX * (p.cols - 1) + p.w);
  const right = Math.max(0, pageW - side - used);
  return '<w:sectPr><w:pgSz w:w="' + pageW + '" w:h="' + pageH + '"/>'
    + '<w:pgMar w:top="' + top + '" w:right="' + right + '" w:bottom="0" w:left="' + side + '" w:header="0" w:footer="0" w:gutter="0"/>'
    + '<w:cols w:space="720"/></w:sectPr>';
}

/** One label's paragraph of address lines. */
export function labelParagraph(lines, font = null) {
  return '<w:p>' + LABEL_P_PR + linesXml(lines, fontRPr(font)) + '</w:p>';
}

/** Word's «Next Record» field, as the runs it writes. */
export const NEXT_FIELD_RUNS = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NEXT </w:instrText></w:r>'
  + '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t xml:space="preserve">«Next Record»</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';

/** A label paragraph that starts with Word's «Next Record» field — every label but the first on a merge sheet. */
export function nextRecordParagraph(rest = '') {
  return '<w:p>' + LABEL_P_PR + NEXT_FIELD_RUNS + rest + '</w:p>';
}

/** Lines typed in a box, as address lines: trimmed, the empty ones gone. */
export function addressLines(text) {
  return String(text ?? '').split(/\r\n|\r|\n/).map((l) => l.trim()).filter(Boolean);
}
