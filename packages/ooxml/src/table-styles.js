// Word's built-in table styles, as a document carries one once a table takes it.
//
// Word writes a table style into word/styles.xml only when something uses it,
// so a document that never had a Grid Table 4 has no definition to point at.
// These are written the way Word writes them — the whole table's lines, then
// each conditional part (header and total rows, first and last columns, the
// bands) with its shading, its lines and its words — in the document's own
// theme colours: an accent at full strength for a header, tinted for the
// lines and the bands as Word tints a theme colour (in HSL, towards white).

/** The gallery: each style's id, its name as Word shows it, and the accent it is drawn in. */
export const TABLE_STYLES = [
  { id: 'TableGrid', name: 'Table Grid', family: 'plain' },
  { id: 'PlainTable1', name: 'Plain Table 1', family: 'plain' },
  { id: 'GridTable1Light', name: 'Grid Table 1 Light', family: 'plain' },
  ...[null, 1, 2, 3, 4, 5, 6].map((n) => ({ id: n ? `GridTable4-Accent${n}` : 'GridTable4', name: n ? `Grid Table 4 – Accent ${n}` : 'Grid Table 4', family: 'grid4', accent: n })),
  ...[null, 1, 2, 3, 4, 5, 6].map((n) => ({ id: n ? `ListTable4-Accent${n}` : 'ListTable4', name: n ? `List Table 4 – Accent ${n}` : 'List Table 4', family: 'list4', accent: n })),
];

/** A colour lightened as Word tints a theme colour: its lightness moved towards white, keeping only `keep` of the way. */
export function tint(hex, keep) {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  const l2 = l * keep + (1 - keep);
  const hue = (p, q, t) => {
    let u = t;
    if (u < 0) u += 1;
    if (u > 1) u -= 1;
    if (u < 1 / 6) return p + (q - p) * 6 * u;
    if (u < 1 / 2) return q;
    if (u < 2 / 3) return p + (q - p) * (2 / 3 - u) * 6;
    return p;
  };
  let out;
  if (s === 0) out = [l2, l2, l2];
  else {
    const q = l2 < 0.5 ? l2 * (1 + s) : l2 + s - l2 * s;
    const p = 2 * l2 - q;
    out = [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)];
  }
  return out.map((v) => Math.floor(Math.max(0, Math.min(1, v)) * 255 + 1e-6).toString(16).padStart(2, '0')).join('').toUpperCase();
}

const side = (name, colour, { val = 'single', sz = 4 } = {}) => `<w:${name} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${colour}"/>`;
const nil = (name) => `<w:${name} w:val="nil"/>`;
const borders = (tag, parts) => `<w:${tag}>${parts.join('')}</w:${tag}>`;
const part = (type, { b = false, colour = null, fill = null, tc = null } = {}) =>
  `<w:tblStylePr w:type="${type}">${b || colour ? `<w:rPr>${b ? '<w:b/><w:bCs/>' : ''}${colour ? `<w:color w:val="${colour}"/>` : ''}</w:rPr>` : ''}<w:tblPr/>${tc || fill ? `<w:tcPr>${tc || ''}${fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : ''}</w:tcPr>` : ''}</w:tblStylePr>`;

/**
 * A style's definition, in the theme's colours ({ accent1…accent6, dk1, … } as
 * six-digit hex), or null for an id this does not know.
 */
export function tableStyleXml(id, colours = {}) {
  const def = TABLE_STYLES.find((s) => s.id === id);
  if (!def) return null;
  const head = (bands) => `<w:style w:type="table" w:styleId="${id}"><w:name w:val="${def.name.replace('–', '').replace(/\s+/g, ' ')}"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="${def.family === 'plain' ? 39 : 49}"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`
    + `<w:tblPr>${bands ? '<w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/>' : ''}`;
  if (id === 'TableGrid') {
    return `${head(false)}${borders('tblBorders', ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((n) => side(n, 'auto')))}</w:tblPr></w:style>`;
  }
  if (id === 'PlainTable1') {
    const line = 'BFBFBF';
    return `${head(true)}${borders('tblBorders', ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((n) => side(n, line)))}</w:tblPr>`
      + part('firstRow', { b: true }) + part('lastRow', { b: true, tc: borders('tcBorders', [side('top', line, { val: 'double' })]) }) + part('firstCol', { b: true }) + part('lastCol', { b: true })
      + part('band1Vert', { fill: 'F2F2F2' }) + part('band1Horz', { fill: 'F2F2F2' }) + '</w:style>';
  }
  if (id === 'GridTable1Light') {
    const line = '999999';
    const strong = '666666';
    return `${head(true)}${borders('tblBorders', ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((n) => side(n, line)))}</w:tblPr>`
      + part('firstRow', { b: true, tc: borders('tcBorders', [side('bottom', strong, { sz: 12 })]) }) + part('lastRow', { b: true, tc: borders('tcBorders', [side('top', strong, { val: 'double', sz: 2 })]) })
      + part('firstCol', { b: true }) + part('lastCol', { b: true }) + '</w:style>';
  }
  // Grid Table 4 and List Table 4: an accent (or the text colour) at full
  // strength for the header row, at 60% for the lines, at 20% for the bands.
  const strong = def.accent ? (colours[`accent${def.accent}`] || '4472C4').toUpperCase() : '000000';
  const lines = def.accent ? tint(strong, 0.6) : '666666';
  const band = def.accent ? tint(strong, 0.2) : 'CCCCCC';
  const around = ['top', 'left', 'bottom', 'right'];
  if (def.family === 'grid4') {
    return `${head(true)}${borders('tblBorders', [...around, 'insideH', 'insideV'].map((n) => side(n, lines)))}</w:tblPr>`
      + part('firstRow', { b: true, colour: 'FFFFFF', fill: strong, tc: borders('tcBorders', [...around.map((n) => side(n, strong)), nil('insideH'), nil('insideV')]) })
      + part('lastRow', { b: true, tc: borders('tcBorders', [side('top', strong, { val: 'double' })]) })
      + part('firstCol', { b: true }) + part('lastCol', { b: true })
      + part('band1Vert', { fill: band }) + part('band1Horz', { fill: band }) + '</w:style>';
  }
  return `${head(true)}${borders('tblBorders', [...around, 'insideH'].map((n) => side(n, lines)))}</w:tblPr>`
    + part('firstRow', { b: true, colour: 'FFFFFF', fill: strong, tc: borders('tcBorders', [...around.map((n) => side(n, strong)), nil('insideH')]) })
    + part('lastRow', { b: true, tc: borders('tcBorders', [side('top', lines, { val: 'double' })]) })
    + part('firstCol', { b: true }) + part('lastCol', { b: true })
    + part('band1Vert', { fill: band }) + part('band1Horz', { fill: band }) + '</w:style>';
}
