/**
 * Cell appearance — fonts, fills, borders and alignment.
 *
 * Until now a workbook rendered as monospaced-looking plain text with the right
 * numbers in it. Every real one a customer sends has a shaded header row, ruled
 * columns and a bold total, and their absence is the first thing a user notices:
 * the figures are correct and the file still looks like it was damaged.
 *
 * Three things in here are easy to get subtly wrong, and all three are the kind
 * of wrong that only shows up on a customer's file:
 *
 * 1. THEME COLOUR INDICES ARE NOT THE THEME'S DOCUMENT ORDER. `<a:clrScheme>`
 *    lists dk1, lt1, dk2, lt2, accent1..6 — but index 0 is lt1 and index 1 is
 *    dk1. The first two are swapped. Get this wrong and every themed header comes
 *    out black on a white sheet, or white on white.
 * 2. TINT IS APPLIED IN LUMINANCE, not by mixing with white. `accent1` with
 *    tint -0.5 is a genuinely darker accent1, not a grey blend, and eyeballing
 *    the difference on a pale fill is impossible.
 * 3. `patternType="none"` MEANS NO FILL. Fill 0 is always none and fill 1 is
 *    always gray125 by convention; a file that says fill 1 wants that pattern,
 *    not a solid.
 *
 * Everything here is READ. Nothing writes styles.xml — an edited cell keeps its
 * existing style index, which is what makes editing a formatted workbook safe.
 */

/** ARGB or RGB hex -> CSS. Alpha comes first in OOXML and is usually FF. */
function hexColour(rgb) {
  const value = String(rgb ?? '').replace(/^#/, '');
  if (/^[0-9a-fA-F]{8}$/.test(value)) {
    const alpha = parseInt(value.slice(0, 2), 16);
    const body = '#' + value.slice(2).toLowerCase();
    return alpha === 255 ? body : body + Math.round(alpha).toString(16).padStart(2, '0');
  }
  if (/^[0-9a-fA-F]{6}$/.test(value)) return '#' + value.toLowerCase();
  return null;
}

/**
 * The legacy 56-colour indexed palette.
 *
 * Superseded by themes twenty years ago and still all over real files, because
 * anything saved by an older tool — or by an export from an accounting system —
 * uses it. 64 and 65 are the system foreground and background, which have no
 * fixed value; they resolve to the caller's ink rather than to a guess.
 */
export const INDEXED_COLOURS = [
  '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
  '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#c0c0c0', '#808080',
  '#9999ff', '#993366', '#ffffcc', '#ccffff', '#660066', '#ff8080', '#0066cc', '#ccccff',
  '#000080', '#ff00ff', '#ffff00', '#00ffff', '#800080', '#800000', '#008080', '#0000ff',
  '#00ccff', '#ccffff', '#ccffcc', '#ffff99', '#99ccff', '#ff99cc', '#cc99ff', '#ffcc99',
  '#3366ff', '#33cccc', '#99cc00', '#ffcc00', '#ff9900', '#ff6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
];

/**
 * Theme slot order, which is NOT the order `<a:clrScheme>` lists them in.
 *
 * ECMA-376 numbers these by the *UI* order, where background comes first. The
 * scheme element lists dk1 first. Swapping the first two pairs is the whole
 * difference, and it is the single most common way themed colours come out wrong.
 */
export const THEME_SLOTS = [
  'lt1', 'dk1', 'lt2', 'dk2',
  'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
  'hlink', 'folHlink',
];

/** Read `xl/theme/theme1.xml` into the twelve slots, in INDEX order. */
export function readTheme(pkg) {
  const part = pkg.has('xl/theme/theme1.xml') ? 'xl/theme/theme1.xml' : null;
  if (!part) return [];
  const xml = pkg.text(part);
  const scheme = /<a:clrScheme\b[^>]*>([\s\S]*?)<\/a:clrScheme>/.exec(xml);
  if (!scheme) return [];

  const byName = {};
  const re = /<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)\b[^>]*>([\s\S]*?)<\/a:\1>/g;
  let m;
  while ((m = re.exec(scheme[1]))) {
    const inner = m[2];
    const srgb = /<a:srgbClr\b[^>]*val="([0-9a-fA-F]{6})"/.exec(inner);
    // sysClr carries lastClr — the colour the authoring system resolved it to,
    // which is the only value we can honour without a Windows theme.
    const sys = /<a:sysClr\b[^>]*lastClr="([0-9a-fA-F]{6})"/.exec(inner);
    byName[m[1]] = srgb ? '#' + srgb[1].toLowerCase() : (sys ? '#' + sys[1].toLowerCase() : null);
  }
  return THEME_SLOTS.map((name) => byName[name] ?? null);
}

/** #rrggbb -> {h, s, l} with l in 0..1. */
function toHsl(hex) {
  const r = parseInt(hex.slice(1, 3), 16) / 255;
  const g = parseInt(hex.slice(3, 5), 16) / 255;
  const b = parseInt(hex.slice(5, 7), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h;
  if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
  else if (max === g) h = ((b - r) / d + 2) / 6;
  else h = ((r - g) / d + 4) / 6;
  return { h, s, l };
}

function toHex({ h, s, l }) {
  const f = (n) => {
    const k = (n + h * 12) % 12;
    const a = s * Math.min(l, 1 - l);
    const v = l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1));
    return Math.round(Math.max(0, Math.min(1, v)) * 255).toString(16).padStart(2, '0');
  };
  return '#' + f(0) + f(8) + f(4);
}

/**
 * Apply a tint the way ECMA-376 defines it: in LUMINANCE, not by mixing.
 *
 * Negative darkens toward black, positive lightens toward white, and the curve
 * is multiplicative rather than linear — which is why `accent1` at tint 0.8
 * still reads as that hue rather than as pale grey.
 */
export function applyTint(hex, tint) {
  const t = Number(tint);
  if (!hex || !Number.isFinite(t) || t === 0) return hex;
  const hsl = toHsl(hex);
  const l = t < 0 ? hsl.l * (1 + t) : hsl.l * (1 - t) + t;
  return toHex({ ...hsl, l: Math.max(0, Math.min(1, l)) });
}

const attrsOf = (tag) => {
  const out = {};
  const re = /([\w:.-]+)\s*=\s*"([^"]*)"/g;
  let m;
  while ((m = re.exec(tag))) out[m[1]] = m[2];
  return out;
};

/**
 * Resolve one colour element — `rgb`, `theme` + `tint`, or `indexed`.
 * Returns null for "unspecified" and for the two system colours, which have no
 * fixed value and must fall back to the caller's own ink.
 */
export function readColourElement(tag, theme = []) {
  if (!tag) return null;
  const a = attrsOf(tag);
  if (a.rgb) return applyTint(hexColour(a.rgb), a.tint);
  if (a.theme !== undefined) {
    const base = theme[Number(a.theme)] ?? null;
    return base ? applyTint(base, a.tint) : null;
  }
  if (a.indexed !== undefined) {
    const i = Number(a.indexed);
    // 64 is the system foreground and 65 the system background; both are
    // "whatever the reader's ink is", so we say nothing rather than guess.
    if (i === 64 || i === 65) return null;
    return INDEXED_COLOURS[i] ?? null;
  }
  return null;
}

const first = (xml, name) => {
  const m = new RegExp('<' + name + '\\b[^>]*?(?:/>|>[\\s\\S]*?</' + name + '>)').exec(xml);
  return m ? m[0] : null;
};
const children = (xml, name) => {
  const re = new RegExp('<' + name + '\\b[^>]*?(?:/>|>[\\s\\S]*?</' + name + '>)', 'g');
  return [...String(xml).matchAll(re)].map((m) => m[0]);
};
/** A `<w:b/>`-style flag: present means on unless it says val="0". */
const flag = (xml, name) => {
  const el = first(xml, name);
  if (!el) return false;
  const v = attrsOf(el).val;
  return v === undefined || !['0', 'false'].includes(String(v).toLowerCase());
};

/** OOXML border style -> CSS width and style. `thin` is 1px, and that is it. */
export const BORDER_WIDTHS = {
  hair: [1, 'solid'], thin: [1, 'solid'], medium: [2, 'solid'], thick: [3, 'solid'],
  dotted: [1, 'dotted'], dashed: [1, 'dashed'], double: [3, 'double'],
  dashDot: [1, 'dashed'], dashDotDot: [1, 'dashed'], slantDashDot: [1, 'dashed'],
  mediumDashed: [2, 'dashed'], mediumDashDot: [2, 'dashed'], mediumDashDotDot: [2, 'dashed'],
};

const SIDES = ['left', 'right', 'top', 'bottom'];

/**
 * Read every style table, indexed by the `s=` attribute a cell carries.
 *
 * @returns {{
 *   byStyleIndex: Array<object>,      full style per cellXfs index
 *   numberFormats: string[],          format code per index (the old shape)
 *   theme: Array<string|null>,
 *   custom: Record<number, string>,
 * }}
 */
export function readStyles(pkg, { BUILTIN_FORMATS = {} } = {}) {
  const empty = { byStyleIndex: [], numberFormats: [], theme: [], custom: {}, dxfs: [] };
  if (!pkg.has('xl/styles.xml')) return empty;

  const xml = pkg.text('xl/styles.xml');
  const theme = readTheme(pkg);

  const custom = {};
  for (const m of xml.matchAll(/<numFmt\b([^>]*)\/>/g)) {
    const id = /numFmtId="(\d+)"/.exec(m[1]);
    const code = /formatCode="([^"]*)"/.exec(m[1]);
    if (id && code) custom[Number(id[1])] = unescXml(code[1]);
  }

  // ---- fonts ----
  const fontsBlock = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(xml);
  const fonts = (fontsBlock ? children(fontsBlock[1], 'font') : []).map((f) => {
    const size = first(f, 'sz');
    const name = first(f, 'name') ?? first(f, 'rFont');
    return {
      bold: flag(f, 'b'),
      italic: flag(f, 'i'),
      underline: Boolean(first(f, 'u')),
      strike: flag(f, 'strike'),
      sizePt: size ? Number(attrsOf(size).val) || null : null,
      family: name ? attrsOf(name).val ?? null : null,
      colour: readColourElement(first(f, 'color'), theme),
    };
  });

  // ---- fills ----
  const fillsBlock = /<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(xml);
  const fills = (fillsBlock ? children(fillsBlock[1], 'fill') : []).map((f) => {
    const pattern = first(f, 'patternFill');
    if (!pattern) return { pattern: 'none', colour: null };
    const type = attrsOf(pattern).patternType ?? 'none';
    if (type === 'none') return { pattern: 'none', colour: null };
    // In a pattern fill the FOREGROUND colour is the one that shows for a solid
    // fill — bgColor is the gaps, and for `solid` there are none.
    const fg = readColourElement(first(pattern, 'fgColor'), theme);
    const bg = readColourElement(first(pattern, 'bgColor'), theme);
    return { pattern: type, colour: type === 'solid' ? fg : (fg ?? bg), background: bg };
  });

  // ---- borders ----
  const bordersBlock = /<borders\b[^>]*>([\s\S]*?)<\/borders>/.exec(xml);
  const borders = (bordersBlock ? children(bordersBlock[1], 'border') : []).map((b) => {
    const out = {};
    let any = false;
    for (const side of SIDES) {
      const el = first(b, side);
      if (!el) continue;
      const style = attrsOf(el).style;
      if (!style || style === 'none') continue;
      const [width, css] = BORDER_WIDTHS[style] ?? [1, 'solid'];
      out[side] = { style: css, widthPx: width, colour: readColourElement(first(el, 'color'), theme) };
      any = true;
    }
    return any ? out : null;
  });

  // ---- cellXfs: what a cell's `s=` actually points at ----
  const cellXfsBlock = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
  const byStyleIndex = [];
  const numberFormats = [];
  if (cellXfsBlock) {
    for (const xf of children(cellXfsBlock[1], 'xf')) {
      const a = attrsOf(xf.split('>')[0] + '>');
      const numFmtId = Number(a.numFmtId ?? 0);
      const alignEl = first(xf, 'alignment');
      const align = alignEl ? attrsOf(alignEl) : {};

      // `applyFont="0"` means "ignore the fontId and inherit" — a distinction
      // that matters, because fontId 0 is a real font and not a null.
      const applied = (name, id) => (a['apply' + name] === '0' ? null : id);

      const fontId = applied('Font', Number(a.fontId ?? 0));
      const fillId = applied('Fill', Number(a.fillId ?? 0));
      const borderId = applied('Border', Number(a.borderId ?? 0));

      const protection = first(xf, 'protection');
      numberFormats.push(custom[numFmtId] ?? BUILTIN_FORMATS[numFmtId] ?? 'General');
      byStyleIndex.push({
        numFmt: custom[numFmtId] ?? BUILTIN_FORMATS[numFmtId] ?? 'General',
        font: fontId === null ? null : fonts[fontId] ?? null,
        fill: fillId === null ? null : (fills[fillId]?.pattern === 'none' ? null : fills[fillId] ?? null),
        border: borderId === null ? null : borders[borderId] ?? null,
        align: {
          horizontal: align.horizontal ?? null,
          vertical: align.vertical ?? null,
          wrap: align.wrapText === '1',
          indent: Number(align.indent ?? 0) || 0,
          rotation: Number(align.textRotation ?? 0) || 0,
        },
        // The spec's default is LOCKED — protection only bites when the
        // sheet turns it on, which is why nobody notices until it does.
        locked: protection ? attrsOf(protection).locked !== '0' : true,
      });
    }
  }
  // ---- dxfs: the DIFFERENTIAL formats conditional formatting points at ----
  // A dxf says only what CHANGES — a red font, a rose fill — and one more
  // spec inversion lives here: in a dxf's patternFill the solid colour
  // usually rides bgColor, the opposite of the fills table above.
  const dxfsBlock = /<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/.exec(xml);
  const dxfs = (dxfsBlock ? children(dxfsBlock[1], 'dxf') : []).map((d) => {
    const font = first(d, 'font');
    const pattern = first(d, 'fill') ? first(first(d, 'fill'), 'patternFill') : null;
    const bg = pattern ? readColourElement(first(pattern, 'bgColor'), theme) : null;
    const fg = pattern ? readColourElement(first(pattern, 'fgColor'), theme) : null;
    return {
      font: font ? {
        bold: flag(font, 'b'),
        italic: flag(font, 'i'),
        strike: flag(font, 'strike'),
        colour: readColourElement(first(font, 'color'), theme),
      } : null,
      fill: bg ?? fg ? { pattern: 'solid', colour: bg ?? fg } : null,
    };
  });

  return { byStyleIndex, numberFormats, theme, custom, dxfs };
}

/**
 * Merged ranges declared by a sheet.
 *
 * Returned as ranges rather than resolved into a grid, for the same reason a
 * document table reports its merges: only the renderer knows how it lays cells
 * out, and a resolved grid would be an invented one.
 */
export function readMergedCells(sheetXml) {
  const block = /<mergeCells\b[^>]*>([\s\S]*?)<\/mergeCells>/.exec(String(sheetXml));
  if (!block) return [];
  const out = [];
  for (const m of block[1].matchAll(/<mergeCell\b[^>]*ref="([A-Z]+\d+):([A-Z]+\d+)"/g)) {
    const from = cellRefToRowCol(m[1]);
    const to = cellRefToRowCol(m[2]);
    if (!from || !to) continue;
    out.push({
      top: Math.min(from.row, to.row),
      left: Math.min(from.col, to.col),
      bottom: Math.max(from.row, to.row),
      right: Math.max(from.col, to.col),
      ref: m[1] + ':' + m[2],
    });
  }
  return out;
}

function cellRefToRowCol(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(String(ref));
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]) - 1, col: col - 1 };
}

const unescXml = (s) => String(s)
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
  .replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/**
 * The `<dataValidations>` block of a sheet: the rules a workbook's author put
 * on what a cell may hold.
 *
 * Read for ENFORCEMENT, which is why the shape keeps the author's own words —
 * `errorTitle`/`error` are what the person who broke the rule should read, and
 * `prompt` is what they should have read first. `sqref` can name several
 * ranges separated by spaces; each becomes a rect. Two spec traps worth
 * naming: `showDropDown="1"` means SUPPRESS the in-cell dropdown (the
 * attribute predates the UI it now inverts), and `allowBlank` defaults to
 * false in the schema although Excel writes it explicitly almost always.
 * Extended rules living in `<extLst>` (x14) are preserved but not read here.
 */
export function readDataValidations(sheetXml) {
  const block = /<dataValidations\b[^>]*>([\s\S]*?)<\/dataValidations>/.exec(String(sheetXml));
  if (!block) return [];
  const out = [];
  for (const m of block[1].matchAll(/<dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/dataValidation>)/g)) {
    const attrs = {};
    for (const a of m[1].matchAll(/([A-Za-z0-9:]+)="([^"]*)"/g)) attrs[a[1]] = unescXml(a[2]);
    const body = m[2] ?? '';
    const f1 = /<formula1>([\s\S]*?)<\/formula1>/.exec(body);
    const f2 = /<formula2>([\s\S]*?)<\/formula2>/.exec(body);
    const ranges = [];
    for (const refText of String(attrs.sqref ?? '').split(/\s+/).filter(Boolean)) {
      const [a, b] = refText.split(':');
      const from = cellRefToRowCol(a);
      const to = b ? cellRefToRowCol(b) : from;
      if (!from || !to) continue;
      ranges.push({
        top: Math.min(from.row, to.row), left: Math.min(from.col, to.col),
        bottom: Math.max(from.row, to.row), right: Math.max(from.col, to.col),
      });
    }
    if (!ranges.length) continue;
    out.push({
      type: attrs.type ?? 'none',
      operator: attrs.operator ?? 'between',
      allowBlank: attrs.allowBlank === '1' || attrs.allowBlank === 'true',
      suppressDropdown: attrs.showDropDown === '1' || attrs.showDropDown === 'true',
      errorStyle: attrs.errorStyle ?? 'stop',
      errorTitle: attrs.errorTitle ?? null,
      error: attrs.error ?? null,
      promptTitle: attrs.promptTitle ?? null,
      prompt: attrs.prompt ?? null,
      formula1: f1 ? unescXml(f1[1]) : null,
      formula2: f2 ? unescXml(f2[1]) : null,
      sqref: attrs.sqref,
      ranges,
    });
  }
  return out;
}

/**
 * The `<conditionalFormatting>` rules of a sheet, flattened and sorted by
 * priority (lower number wins first, as the spec orders them).
 *
 * Each rule keeps its whole vocabulary: `dxfId` into the style table's dxfs,
 * the comparison operator and its `<formula>`s, a colour scale's stops, a
 * data bar's bounds and colour, and the text/rank/average attributes the
 * simpler types ride on. Colours resolve through the theme like every other
 * colour in the file.
 */
export function readConditionalFormatting(sheetXml, theme = []) {
  const out = [];
  for (const block of String(sheetXml).matchAll(/<conditionalFormatting\b([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g)) {
    const sqref = attrsOf(block[1]).sqref ?? '';
    const ranges = [];
    for (const refText of sqref.split(/\s+/).filter(Boolean)) {
      const [a, b] = refText.split(':');
      const from = cellRefToRowCol(a);
      const to = b ? cellRefToRowCol(b) : from;
      if (!from || !to) continue;
      ranges.push({
        top: Math.min(from.row, to.row), left: Math.min(from.col, to.col),
        bottom: Math.max(from.row, to.row), right: Math.max(from.col, to.col),
      });
    }
    if (!ranges.length) continue;

    for (const r of block[2].matchAll(/<cfRule\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cfRule>)/g)) {
      const a = attrsOf(r[1]);
      const body = r[2] ?? '';
      const cfvos = children(body, 'cfvo').map((el) => {
        const at = attrsOf(el);
        return { type: at.type, val: at.val !== undefined ? unescXml(at.val) : null };
      });
      const colours = children(body, 'color').map((el) => readColourElement(el, theme));
      // The iconSet element's own attributes — which set, whether the order
      // is reversed, whether the number still shows beside the icon.
      const iconEl = /<iconSet\b([^>]*?)>/.exec(body);
      const iconAttrs = iconEl ? attrsOf(iconEl[1]) : null;
      out.push({
        type: a.type,
        iconSet: iconAttrs ? (iconAttrs.iconSet ?? '3TrafficLights1') : null,
        iconReverse: iconAttrs ? iconAttrs.reverse === '1' : false,
        iconShowValue: iconAttrs ? iconAttrs.showValue !== '0' : true,
        timePeriod: a.timePeriod ?? null,
        priority: Number(a.priority ?? 0),
        dxfId: a.dxfId !== undefined ? Number(a.dxfId) : null,
        operator: a.operator ?? null,
        text: a.text !== undefined ? unescXml(a.text) : null,
        rank: Number(a.rank ?? 10) || 10,
        percent: a.percent === '1',
        bottom: a.bottom === '1',
        aboveAverage: a.aboveAverage !== '0',
        equalAverage: a.equalAverage === '1',
        stopIfTrue: a.stopIfTrue === '1',
        formulas: [...body.matchAll(/<formula>([\s\S]*?)<\/formula>/g)].map((f) => unescXml(f[1])),
        cfvos,
        colours,
        ranges,
        sqref,
      });
    }
  }
  return out.sort((x, y) => x.priority - y.priority);
}

export { hexColour, cellRefToRowCol, unescXml };
