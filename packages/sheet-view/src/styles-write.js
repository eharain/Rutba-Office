/**
 * Writing cell formatting into `styles.xml`, by APPENDING and never rewriting.
 *
 * THE RULE THIS FILE EXISTS TO KEEP. Every `<xf>` in `cellXfs` is addressed by
 * its position, and every cell in the workbook points at one by index. Editing
 * an existing entry would silently restyle every other cell that shares it —
 * bold one heading and watch a hundred unrelated cells go bold. Inserting one
 * would be worse: every index after the insertion point now means something
 * else, across a file we may not fully understand.
 *
 * So the only safe move is to append. A style that already exists is reused, a
 * new one goes on the end, and nothing that was there before changes meaning.
 * That is also what Excel itself does, which is why files edited this way stay
 * ordinary rather than becoming ours.
 *
 * The same rule covers `<fonts>`: a new weight means a new font entry, not an
 * edit to the one the cell happened to be using. And `<numFmts>`, where a new
 * custom format code takes the next id rather than overwriting an existing one.
 */
import { BUILTIN_FORMATS } from './numfmt.js';

/**
 * A standard format code back to its reserved built-in id, so setting `0.00`
 * references numFmtId 2 rather than inventing a custom entry for something every
 * consumer already agrees on. First id wins where a code is not unique.
 */
const BUILTIN_NUMFMT_ID = (() => {
  const out = {};
  for (const [id, code] of Object.entries(BUILTIN_FORMATS)) {
    if (!(code in out)) out[code] = Number(id);
  }
  return out;
})();

const attrsOf = (tag) => {
  const out = {};
  for (const m of String(tag).matchAll(/([A-Za-z:]+)\s*=\s*"([^"]*)"/g)) out[m[1]] = m[2];
  return out;
};

const esc = (s) => String(s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const unesc = (s) => String(s)
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
  .replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** The children of a container, as raw XML strings. */
function children(xml, name) {
  const out = [];
  const re = new RegExp('<' + name + '\\b([^>]*?)(/>|>)', 'g');
  let m;
  while ((m = re.exec(xml))) {
    if (m[2] === '/>') { out.push(m[0]); continue; }
    const close = '</' + name + '>';
    const end = xml.indexOf(close, m.index);
    if (end === -1) { out.push(m[0]); continue; }
    out.push(xml.slice(m.index, end + close.length));
    re.lastIndex = end + close.length;
  }
  return out;
}

/** A container's inner XML and the span it occupies, or null when absent. */
function block(xml, name) {
  const open = new RegExp('<' + name + '\\b([^>]*)>', 'g');
  const m = open.exec(xml);
  if (!m) return null;
  const close = '</' + name + '>';
  const end = xml.indexOf(close, m.index);
  if (end === -1) return null;
  return {
    attrs: m[1],
    inner: xml.slice(m.index + m[0].length, end),
    start: m.index,
    end: end + close.length,
    openTag: m[0],
  };
}

/** Replace a container's inner XML and keep its `count` attribute honest. */
function replaceBlock(xml, name, inner, count) {
  const b = block(xml, name);
  if (!b) return null;
  const openTag = b.openTag.replace(/\bcount="\d+"/, 'count="' + count + '"');
  return xml.slice(0, b.start) + openTag + inner + '</' + name + '>' + xml.slice(b.end);
}

// ── fonts ───────────────────────────────────────────────────────────────────

/**
 * Read one `<font>` back into the shape a delta is expressed in.
 *
 * Deliberately partial: only the properties this module can set are read, so a
 * font carrying anything else is matched by its XML rather than by these
 * fields. That is what stops "make it bold" from quietly dropping a colour or a
 * character set nobody modelled.
 */
function fontFacts(xml) {
  const has = (tag) => {
    const m = new RegExp('<' + tag + '\\b([^>]*?)/?>', 'i').exec(xml);
    if (!m) return false;
    const val = attrsOf(m[1]).val;
    return val === undefined || val === '1' || val === 'true';
  };
  return { bold: has('b'), italic: has('i'), underline: has('u'), strike: has('strike') };
}

/**
 * The same font with a toggle applied, as XML.
 *
 * Everything not being toggled is carried across verbatim — the size, the
 * family, the colour, the scheme, and anything else the file declared that we
 * have never heard of. Only the toggled element is added or removed.
 */
function fontWith(xml, delta) {
  let inner = /^<font\b[^>]*\/>$/.test(xml)
    ? ''
    : xml.replace(/^<font\b[^>]*>/, '').replace(/<\/font>$/, '');

  for (const [key, tag] of [['bold', 'b'], ['italic', 'i'], ['underline', 'u'], ['strike', 'strike']]) {
    if (delta[key] === undefined) continue;
    const re = new RegExp('<' + tag + '\\b[^>]*?/?>(?:</' + tag + '>)?', 'i');
    inner = inner.replace(re, '');
    // Order within <font> is schema-defined and b/i/u/strike all precede sz and
    // name, so prepending is both correct and simpler than finding the slot.
    if (delta[key]) inner = '<' + tag + '/>' + inner;
  }
  // Value elements, which are set rather than toggled. Replaced in place when
  // present so the schema's element order survives, appended otherwise.
  const setValue = (tag, attrs) => {
    const re = new RegExp('<' + tag + '\\b[^>]*?/?>(?:</' + tag + '>)?', 'i');
    const el = attrs === null ? '' : '<' + tag + ' ' + attrs + '/>';
    if (re.test(inner)) inner = inner.replace(re, el);
    else if (el) inner += el;
  };
  if (delta.fontSize !== undefined) {
    setValue('sz', delta.fontSize === null ? null : 'val="' + Number(delta.fontSize) + '"');
  }
  if (delta.fontColour !== undefined) {
    setValue('color', delta.fontColour === null ? null : 'rgb="' + argb(delta.fontColour) + '"');
  }
  if (delta.fontName !== undefined) {
    // name in a workbook font, rFont in a rich-text run. A font element
    // that already uses rFont keeps using it.
    const tag = /<rFont\b/i.test(inner) ? 'rFont' : 'name';
    setValue(tag, delta.fontName === null ? null : 'val="' + esc(delta.fontName) + '"');
  }

  return '<font>' + inner + '</font>';
}

/**
 * A colour as OOXML wants it: eight hex digits, alpha first.
 *
 * Accepts `#RRGGBB`, `RRGGBB` or `FFRRGGBB`. Anything else is refused rather
 * than guessed, because a malformed colour in styles.xml is a file Excel
 * declines to open — a much worse outcome than a rejected click.
 */
export function argb(value) {
  const hex = String(value).replace(/^#/, '').toUpperCase();
  if (/^[0-9A-F]{8}$/.test(hex)) return hex;
  if (/^[0-9A-F]{6}$/.test(hex)) return 'FF' + hex;
  throw new Error('not a colour: ' + value);
}

/**
 * Find an identical entry in a table, or append one.
 *
 * The same append-never-rewrite rule as `cellXfs`, for the same reason: fills
 * and borders are addressed by index too, so editing one restyles every cell
 * that shares it.
 */
function ensureEntry(xml, container, childName, desiredXml) {
  const b = block(xml, container);
  if (!b) return null;
  const items = children(b.inner, childName);
  const found = items.findIndex((it) => it === desiredXml);
  if (found >= 0) return { xml, index: found };
  items.push(desiredXml);
  const out = replaceBlock(xml, container, items.join(''), items.length);
  return out ? { xml: out, index: items.length - 1 } : null;
}

/** A solid fill, or the "no fill" the schema keeps at index 0. */
function fillXml(colour) {
  if (colour === null) return '<fill><patternFill patternType="none"/></fill>';
  // fgColor, not bgColor. In a solid pattern the FOREGROUND colour is the one
  // that shows — setting bgColor produces a cell that is still white and a bug
  // report that says "the colour did nothing".
  return '<fill><patternFill patternType="solid"><fgColor rgb="' + argb(colour)
    + '"/><bgColor indexed="64"/></patternFill></fill>';
}

/** The four edges, in the order the schema declares them. */
const EDGE_ORDER = ['left', 'right', 'top', 'bottom'];

/**
 * A border element from a per-edge map.
 *
 * `<diagonal/>` is emitted empty because the schema requires the element even
 * when nothing is drawn on it.
 */
function borderXml(edges) {
  let inner = '';
  for (const side of EDGE_ORDER) {
    const e = edges[side];
    if (!e || !e.style || e.style === 'none') { inner += '<' + side + '/>'; continue; }
    inner += '<' + side + ' style="' + esc(e.style) + '">'
      + '<color rgb="' + argb(e.colour || '#000000') + '"/>'
      + '</' + side + '>';
  }
  return '<border>' + inner + '<diagonal/></border>';
}

/** Read a border element back into the per-edge map a delta is written in. */
function borderEdges(xml) {
  const out = {};
  for (const side of EDGE_ORDER) {
    const m = new RegExp('<' + side + '\\b([^>]*?)(?:/>|>([\\s\\S]*?)</' + side + '>)').exec(xml || '');
    if (!m) continue;
    const style = attrsOf(m[1]).style;
    if (!style || style === 'none') continue;
    const colourEl = /<color\b([^>]*?)\/?>/.exec(m[2] || '');
    out[side] = { style, colour: colourEl ? (attrsOf(colourEl[1]).rgb || null) : null };
  }
  return out;
}

// ── number formats ───────────────────────────────────────────────────────────

/** The custom `<numFmt>` entries as an id -> code map. */
function customNumFmts(xml) {
  const out = {};
  const b = block(xml, 'numFmts');
  if (!b) return out;
  for (const m of b.inner.matchAll(/<numFmt\b([^>]*)\/>/g)) {
    const id = /numFmtId="(\d+)"/.exec(m[1]);
    const code = /formatCode="([^"]*)"/.exec(m[1]);
    if (id && code) out[Number(id[1])] = unesc(code[1]);
  }
  return out;
}

/** The format code a numFmtId resolves to: custom, then built-in, then General. */
function numFmtCodeOf(xml, numFmtId) {
  if (numFmtId === 0) return 'General';
  return customNumFmts(xml)[numFmtId] ?? BUILTIN_FORMATS[numFmtId] ?? 'General';
}

/**
 * Ensure a custom format code has a `<numFmts>` entry, and return its id.
 *
 * The same append-never-rewrite discipline as the rest of the file: an
 * identical code already present is reused, a new one takes the next id above
 * the current maximum (custom ids start at 164), and the block is created —
 * before `<fonts>`, which is where the schema orders it — when the file has none.
 */
function ensureNumFmt(xml, code) {
  const encoded = esc(code);
  const b = block(xml, 'numFmts');
  if (b) {
    const items = children(b.inner, 'numFmt');
    let maxId = 163;
    for (const it of items) {
      const idM = /numFmtId="(\d+)"/.exec(it);
      const codeM = /formatCode="([^"]*)"/.exec(it);
      if (codeM && unesc(codeM[1]) === code && idM) return { xml, id: Number(idM[1]) };
      if (idM && Number(idM[1]) > maxId) maxId = Number(idM[1]);
    }
    const id = maxId + 1;
    items.push('<numFmt numFmtId="' + id + '" formatCode="' + encoded + '"/>');
    const out = replaceBlock(xml, 'numFmts', items.join(''), items.length);
    return { xml: out ?? xml, id };
  }
  // No block yet. A self-closing empty one is replaced in place; otherwise a new
  // block is inserted before `<fonts>`, the schema's slot for numFmts.
  const entry = '<numFmt numFmtId="164" formatCode="' + encoded + '"/>';
  const blockXml = '<numFmts count="1">' + entry + '</numFmts>';
  const selfClosing = /<numFmts\b[^>]*\/>/.exec(xml);
  if (selfClosing) {
    const out = xml.slice(0, selfClosing.index) + blockXml
      + xml.slice(selfClosing.index + selfClosing[0].length);
    return { xml: out, id: 164 };
  }
  const fontsAt = xml.indexOf('<fonts');
  if (fontsAt >= 0) return { xml: xml.slice(0, fontsAt) + blockXml + xml.slice(fontsAt), id: 164 };
  const openTag = /<styleSheet\b[^>]*>/.exec(xml);
  if (openTag) {
    const p = openTag.index + openTag[0].length;
    return { xml: xml.slice(0, p) + blockXml + xml.slice(p), id: 164 };
  }
  return { xml: xml + blockXml, id: 164 };
}

// ── the public operation ────────────────────────────────────────────────────

/**
 * Apply a formatting delta to one style index.
 *
 * @param {string} xml         the whole of styles.xml
 * @param {number|null} base   the cell's current style index, or null for none
 * @param {object} delta       { bold?, italic?, underline?, strike?, align? }
 * @returns {{ xml: string, index: number, changed: boolean }}
 *
 * `align` is a horizontal alignment string, or null to clear it back to the
 * format's own default.
 */
export function applyFormat(xml, base, delta) {
  const cellXfs = block(xml, 'cellXfs');
  if (!cellXfs) return { xml, index: base ?? 0, changed: false };

  const xfs = children(cellXfs.inner, 'xf');
  const baseIndex = Number.isInteger(base) && xfs[base] ? base : 0;
  const baseXf = xfs[baseIndex] ?? '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>';
  const baseAttrs = attrsOf(baseXf.split('>')[0]);

  let nextXml = xml;
  let fontId = Number(baseAttrs.fontId ?? 0);

  // ---- the font half ----
  const fontDelta = {};
  for (const k of ['bold', 'italic', 'underline', 'strike', 'fontName', 'fontSize', 'fontColour']) {
    if (delta[k] !== undefined) fontDelta[k] = delta[k];
  }
  if (Object.keys(fontDelta).length) {
    const fontsBlock = block(nextXml, 'fonts');
    if (!fontsBlock) return { xml, index: baseIndex, changed: false };
    const fonts = children(fontsBlock.inner, 'font');
    const current = fonts[fontId] ?? '<font/>';

    // Already as asked? Then the font does not change, and if the alignment is
    // not moving either the whole operation is a no-op.
    const facts = fontFacts(current);
    const wanted = { ...facts, ...fontDelta };
    const valueChange = fontDelta.fontName !== undefined
      || fontDelta.fontSize !== undefined || fontDelta.fontColour !== undefined;
    if (valueChange || facts.bold !== wanted.bold || facts.italic !== wanted.italic
        || facts.underline !== wanted.underline || facts.strike !== wanted.strike) {
      const desired = fontWith(current, fontDelta);
      const existing = fonts.findIndex((f) => f === desired);
      if (existing >= 0) {
        fontId = existing;
      } else {
        fonts.push(desired);
        fontId = fonts.length - 1;
        nextXml = replaceBlock(nextXml, 'fonts', fonts.join(''), fonts.length) ?? nextXml;
      }
    }
  }

  // ---- the fill half ----
  let fillId = Number(baseAttrs.fillId ?? 0);
  if (delta.fill !== undefined) {
    const made = ensureEntry(nextXml, 'fills', 'fill', fillXml(delta.fill));
    if (made) { nextXml = made.xml; fillId = made.index; }
  }

  // ---- the border half ----
  let borderId = Number(baseAttrs.borderId ?? 0);
  if (delta.border !== undefined) {
    const bordersBlock = block(nextXml, 'borders');
    const existing = bordersBlock
      ? children(bordersBlock.inner, 'border')[borderId] ?? ''
      : '';
    // A delta names the edges it changes; the rest of the border is kept, so
    // adding a bottom rule to a cell that already has a left one keeps both.
    const edges = { ...borderEdges(existing) };
    for (const side of EDGE_ORDER) {
      if (delta.border[side] === undefined) continue;
      if (delta.border[side] === null) delete edges[side];
      else edges[side] = delta.border[side];
    }
    const made = ensureEntry(nextXml, 'borders', 'border', borderXml(edges));
    if (made) { nextXml = made.xml; borderId = made.index; }
  }

  // ---- the number-format half ----
  // A code that matches a built-in references that reserved id; anything else is
  // ensured as a `<numFmts>` entry. `null` (and 'General') clear back to id 0.
  let numFmtId = Number(baseAttrs.numFmtId ?? 0);
  if (delta.numberFormat !== undefined) {
    const code = delta.numberFormat === null ? 'General' : String(delta.numberFormat);
    if (code in BUILTIN_NUMFMT_ID) {
      numFmtId = BUILTIN_NUMFMT_ID[code];
    } else {
      const made = ensureNumFmt(nextXml, code);
      nextXml = made.xml;
      numFmtId = made.id;
    }
  }

  // ---- the alignment half ----
  const alignEl = /<alignment\b([^>]*?)\/?>/.exec(baseXf);
  const alignAttrs = alignEl ? attrsOf(alignEl[1]) : {};
  if (delta.align !== undefined) {
    if (delta.align === null) delete alignAttrs.horizontal;
    else alignAttrs.horizontal = delta.align;
  }
  // Vertical alignment uses the OOXML vocabulary verbatim (top/center/bottom);
  // null clears the attr back to the format's own default, exactly as horizontal
  // does. Every other alignment attr the file declared is carried across.
  if (delta.valign !== undefined) {
    if (delta.valign === null) delete alignAttrs.vertical;
    else alignAttrs.vertical = delta.valign;
  }
  // wrapText is a flag: on writes wrapText="1", off removes it. A cell that does
  // not wrap carries no attr rather than wrapText="0".
  if (delta.wrap !== undefined) {
    if (delta.wrap) alignAttrs.wrapText = '1';
    else delete alignAttrs.wrapText;
  }
  const alignTouched = delta.align !== undefined || delta.valign !== undefined || delta.wrap !== undefined;
  const alignPairs = Object.entries(alignAttrs).filter(([, v]) => v !== undefined && v !== '');
  const alignXml = alignPairs.length
    ? '<alignment ' + alignPairs.map(([k, v]) => k + '="' + esc(v) + '"').join(' ') + '/>'
    : '';

  // ---- the protection half ----
  // The locked flag lives in a <protection> child, and the DEFAULT is locked:
  // unlocking writes locked="0", re-locking removes the attr. Carrying the
  // child through is not optional even when the delta says nothing about it —
  // the rebuild below re-emits the xf's children, and before this half
  // existed, ANY format edit silently dropped an unlocked cell's protection.
  const protEl = /<protection\b([^>]*?)\/?>/.exec(baseXf);
  const protAttrs = protEl ? attrsOf(protEl[1]) : {};
  if (delta.locked !== undefined) {
    if (delta.locked === false) protAttrs.locked = '0';
    else delete protAttrs.locked;
  }
  const protPairs = Object.entries(protAttrs).filter(([, v]) => v !== undefined && v !== '');
  const protXml = protPairs.length
    ? '<protection ' + protPairs.map(([k, v]) => k + '="' + esc(v) + '"').join(' ') + '/>'
    : '';

  // ---- the xf ----
  const attrs = { ...baseAttrs, fontId: String(fontId), fillId: String(fillId), borderId: String(borderId) };
  // applyNumberFormat="0" means "inherit the format, ignore numFmtId" — the same
  // trap as applyFont: setting a format means applying it. Clearing to General
  // (id 0) says so explicitly rather than pretending the format was inherited.
  if (delta.numberFormat !== undefined) {
    attrs.numFmtId = String(numFmtId);
    attrs.applyNumberFormat = numFmtId === 0 ? '0' : '1';
  }
  // applyFill/applyBorder="0" means "inherit, ignore the id", which would throw
  // away the entry we just chose — the same trap as applyFont.
  if (delta.fill !== undefined) attrs.applyFill = '1';
  if (delta.border !== undefined) attrs.applyBorder = '1';
  // applyFont="0" means "inherit the font, ignore fontId" — which would throw
  // away the weight we just chose. Setting a font means applying it.
  if (Object.keys(fontDelta).length) attrs.applyFont = '1';
  if (alignTouched) attrs.applyAlignment = alignPairs.length ? '1' : '0';
  if (delta.locked !== undefined) attrs.applyProtection = protPairs.length ? '1' : '0';

  const attrXml = Object.entries(attrs)
    .filter(([, v]) => v !== undefined)
    .map(([k, v]) => k + '="' + esc(v) + '"')
    .join(' ');
  // Child order is the schema's: alignment, then protection.
  const inner = alignXml + protXml;
  const desiredXf = inner
    ? '<xf ' + attrXml + '>' + inner + '</xf>'
    : '<xf ' + attrXml + '/>';

  if (desiredXf === baseXf) return { xml, index: baseIndex, changed: false };

  const reuse = xfs.findIndex((x) => x === desiredXf);
  if (reuse >= 0) {
    const refreshed = nextXml === xml ? xml : nextXml;
    return { xml: refreshed, index: reuse, changed: reuse !== baseIndex };
  }

  // The cellXfs block has to be re-read: appending a font above may have moved it.
  const freshXfsBlock = block(nextXml, 'cellXfs');
  const freshXfs = children(freshXfsBlock.inner, 'xf');
  freshXfs.push(desiredXf);
  const out = replaceBlock(nextXml, 'cellXfs', freshXfs.join(''), freshXfs.length);
  return { xml: out ?? nextXml, index: freshXfs.length - 1, changed: true };
}

/** Whether a style index already reads as bold / italic / underlined. */
/**
 * The dxf a conditional rule points at — reuse an identical one, else append.
 *
 * A dxf says only what CHANGES, and its solid fill rides `bgColor` (the
 * spec's inversion — the opposite of the fills table). The block rebuild
 * keeps every existing entry byte-verbatim and only appends, so the indices
 * every other rule holds stay true.
 */
export function ensureDxf(xml, { fontColour = null, fill = null, bold = false } = {}) {
  const dxfXml = '<dxf>'
    + (fontColour || bold
      ? '<font>' + (bold ? '<b/>' : '')
        + (fontColour ? '<color rgb="' + argb(fontColour) + '"/>' : '') + '</font>'
      : '')
    + (fill ? '<fill><patternFill><bgColor rgb="' + argb(fill) + '"/></patternFill></fill>' : '')
    + '</dxf>';
  const block = /<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/.exec(xml);
  if (block) {
    const items = block[1].match(/<dxf>[\s\S]*?<\/dxf>|<dxf\/>/g) ?? [];
    const at = items.indexOf(dxfXml);
    if (at >= 0) return { xml, index: at };
    const rebuilt = '<dxfs count="' + (items.length + 1) + '">' + items.join('') + dxfXml + '</dxfs>';
    return { xml: xml.replace(block[0], rebuilt), index: items.length };
  }
  // No dxfs block yet — per schema it sits after cellStyles, before
  // tableStyles / colors / extLst / the end.
  const anchor = /<tableStyles\b|<colors\b|<extLst\b|<\/styleSheet>/.exec(xml);
  const el = '<dxfs count="1">' + dxfXml + '</dxfs>';
  return { xml: xml.slice(0, anchor.index) + el + xml.slice(anchor.index), index: 0 };
}

export function formatOf(xml, index) {
  const cellXfs = block(xml, 'cellXfs');
  const fontsBlock = block(xml, 'fonts');
  const none = {
    bold: false, italic: false, underline: false, strike: false, align: null,
    valign: null, wrap: false,
    fontName: null, fontSize: null, fontColour: null, fill: null, border: {},
    numberFormat: 'General',
    locked: true,
  };
  if (!cellXfs || !fontsBlock) return none;
  const xf = children(cellXfs.inner, 'xf')[Number.isInteger(index) ? index : 0];
  if (!xf) return none;
  const a = attrsOf(xf.split('>')[0]);
  const font = children(fontsBlock.inner, 'font')[Number(a.fontId ?? 0)];
  const alignEl = /<alignment\b([^>]*?)\/?>/.exec(xf);
  const valueOf = (tag) => {
    const m = new RegExp('<' + tag + '\\b([^>]*?)/?>', 'i').exec(font || '');
    return m ? (attrsOf(m[1]).val ?? null) : null;
  };
  const colourEl = /<color\b([^>]*?)\/?>/.exec(font || '');
  const fillsBlock = block(xml, 'fills');
  const fill = fillsBlock ? children(fillsBlock.inner, 'fill')[Number(a.fillId ?? 0)] : null;
  const fg = fill ? /<fgColor\b([^>]*?)\/?>/.exec(fill) : null;
  const bordersBlock = block(xml, 'borders');
  const border = bordersBlock ? children(bordersBlock.inner, 'border')[Number(a.borderId ?? 0)] : null;

  return {
    ...(a.applyFont === '0' || !font ? none : fontFacts(font)),
    align: alignEl ? (attrsOf(alignEl[1]).horizontal ?? null) : null,
    valign: alignEl ? (attrsOf(alignEl[1]).vertical ?? null) : null,
    wrap: alignEl ? attrsOf(alignEl[1]).wrapText === '1' : false,
    fontName: valueOf('name') ?? valueOf('rFont'),
    fontSize: valueOf('sz') === null ? null : Number(valueOf('sz')),
    fontColour: colourEl ? (attrsOf(colourEl[1]).rgb ?? null) : null,
    fill: fg ? (attrsOf(fg[1]).rgb ?? null) : null,
    border: border ? borderEdges(border) : {},
    // The number format is a property of the xf itself, not gated by applyFont,
    // so it is read straight from numFmtId regardless of the rest of the style.
    numberFormat: numFmtCodeOf(xml, Number(a.numFmtId ?? 0)),
    // Locked unless the protection child explicitly says otherwise.
    locked: !/<protection\b[^>]*\blocked="(?:0|false)"/.test(xf),
  };
}
