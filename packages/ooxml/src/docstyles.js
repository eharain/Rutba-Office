/**
 * Paragraph styles and numbering, resolved from the document's own parts.
 *
 * Until now a style was a NAME the view mapped to a hardcoded look: `Heading1`
 * was 22px because our builder writes it that way. A document using `Heading2`,
 * or a house style called `LetterBody`, rendered as body text — and pagination
 * then measured it at the wrong size, so this was a correctness bug wearing a
 * cosmetic one's clothes. The same for lists: a numbered list rendered as a
 * bullet regardless of what `numbering.xml` says, because nobody read it.
 *
 * Both parsers resolve to PLAIN DATA in CSS pixels, because the consumer is the
 * format-neutral editor and it must not learn what a half-point or a twip is.
 * The chain (`basedOn` up to `docDefaults`) is flattened here, once, rather than
 * walked at render time.
 */
import { attrs } from './package.js';

/** Half-points -> CSS px (Word stores font size doubled: sz=28 is 14pt). */
const halfPointsToPx = (hp) => (Number(hp) / 2) * (96 / 72);
/** Twips -> CSS px. */
const twipsToPx = (tw) => Number(tw) * (96 / 1440);

const first = (xml, name) => {
  const m = new RegExp('<' + name + '\\b[^>]*?(?:/>|>[\\s\\S]*?</' + name + '>)').exec(String(xml ?? ''));
  return m ? m[0] : null;
};
const val = (xml, name, attribute = 'w:val') => {
  const el = first(xml, name);
  return el ? attrs(el)[attribute] ?? null : null;
};
/** `<w:b/>` present means on unless it says val="0"; absent means "unset". */
const toggle = (xml, name) => {
  const el = first(xml, name);
  if (!el) return undefined;
  const v = attrs(el)['w:val'];
  return v === undefined || !['0', 'false', 'none'].includes(String(v).toLowerCase());
};

/** The look one `rPr`/`pPr` pair asks for. Unset fields stay undefined so the
 *  chain can fill them — undefined and false mean different things here. */
function readProps(styleXml) {
  const rPr = first(styleXml, 'w:rPr');
  const pPr = first(styleXml, 'w:pPr');
  const sz = val(rPr, 'w:sz');
  const spacing = first(pPr, 'w:spacing');
  const ind = first(pPr, 'w:ind');
  const fonts = first(rPr, 'w:rFonts');
  const out = {
    bold: toggle(rPr, 'w:b'),
    italic: toggle(rPr, 'w:i'),
    underline: (() => {
      const u = first(rPr, 'w:u');
      if (!u) return undefined;
      const v = attrs(u)['w:val'];
      return v !== 'none' && v !== '0';
    })(),
    caps: toggle(rPr, 'w:caps'),
    smallCaps: toggle(rPr, 'w:smallCaps'),
    sizePx: sz !== null ? halfPointsToPx(sz) : undefined,
    colour: (() => {
      const c = val(rPr, 'w:color');
      // "auto" is a colour, not the absence of one: a run that says auto on a
      // heading whose style says blue is black in Word, and was blue here.
      if (c === 'auto') return '#000000';
      return c ? '#' + c.toLowerCase() : undefined;

    })(),
    // A font is named outright (`w:ascii`) or by theme slot (`w:asciiTheme`);
    // the slot is kept as a marker and resolved against the theme once every
    // chain is flattened — see `readParagraphStyles`.
    fontName: (() => {
      if (!fonts) return undefined;
      const a = attrs(fonts);
      if (a['w:ascii']) return a['w:ascii'];
      if (a['w:hAnsi']) return a['w:hAnsi'];
      const slot = a['w:asciiTheme'] || a['w:hAnsiTheme'];
      if (slot) return slot.startsWith('major') ? THEME_MAJOR : THEME_MINOR;
      return undefined;
    })(),
    align: val(pPr, 'w:jc') ?? undefined,
    spaceBeforePx: spacing ? (attrs(spacing)['w:before'] !== undefined ? twipsToPx(attrs(spacing)['w:before']) : undefined) : undefined,
    spaceAfterPx: spacing ? (attrs(spacing)['w:after'] !== undefined ? twipsToPx(attrs(spacing)['w:after']) : undefined) : undefined,
    // Line spacing as a multiplier when the rule is auto (240 = single), the
    // only form a stylesheet can express without knowing the font.
    lineFactor: (() => {
      if (!spacing) return undefined;
      const a = attrs(spacing);
      if (a['w:line'] === undefined || (a['w:lineRule'] ?? 'auto') !== 'auto') return undefined;
      return Number(a['w:line']) / 240;
    })(),
    // An exact or at-least line: pixels, the way a cover title is set
    // ("exactly 60 pt") so its lines sit tight whatever the font says.
    lineExactPx: (() => {
      if (!spacing) return undefined;
      const a = attrs(spacing);
      if (a['w:line'] === undefined || (a['w:lineRule'] ?? 'auto') === 'auto') return undefined;
      return twipsToPx(a['w:line']);
    })(),
    indentPx: ind ? (attrs(ind)['w:left'] !== undefined ? twipsToPx(attrs(ind)['w:left']) : undefined) : undefined,
    rightPx: ind ? (attrs(ind)['w:right'] !== undefined ? twipsToPx(attrs(ind)['w:right']) : undefined) : undefined,
    hangingPx: ind ? (attrs(ind)['w:hanging'] !== undefined ? twipsToPx(attrs(ind)['w:hanging']) : undefined) : undefined,
    firstLinePx: ind ? (attrs(ind)['w:firstLine'] !== undefined ? twipsToPx(attrs(ind)['w:firstLine']) : undefined) : undefined,
    keepNext: toggle(pPr, 'w:keepNext'),
    contextualSpacing: toggle(pPr, 'w:contextualSpacing'),

    // A style's borders — the box around a cover title — same shape as the
    // paragraph's own.
    borders: (() => {
      const pBdr = first(pPr, 'w:pBdr');
      if (!pBdr) return undefined;
      const out = {};
      for (const m of pBdr.matchAll(/<w:(top|left|bottom|right|between|bar)\b([^>]*)\/>/g)) {
        const a = attrs(m[2]);
        const style = a['w:val'] || 'single';
        if (style === 'nil' || style === 'none') continue;
        out[m[1]] = {
          style,
          widthPx: Math.max(1, Math.round((Number(a['w:sz'] ?? 4) / 8) * (96 / 72))),
          colour: a['w:color'] && a['w:color'] !== 'auto' ? '#' + a['w:color'].toUpperCase() : '#000000',
          spacePt: Number(a['w:space'] ?? 0) || 0,
        };
      }
      return Object.keys(out).length ? out : undefined;
    })(),
    // A style's shading and tab stops: a Title band, a TOC entry's dotted
    // right tab. The same shape the paragraph's own decor uses.
    shading: (() => {
      const shd = first(pPr, 'w:shd');
      const fill = shd ? attrs(shd)['w:fill'] : null;
      return fill && /^[0-9A-Fa-f]{6}$/.test(fill) ? '#' + fill.toUpperCase() : undefined;
    })(),
    tabs: (() => {
      const tabs = first(pPr, 'w:tabs');
      if (!tabs) return undefined;
      const stops = [];
      for (const m of tabs.matchAll(/<w:tab\b([^>]*)\/>/g)) {
        const a = attrs(m[1]);
        if (a['w:val'] === 'clear' || a['w:pos'] === undefined) continue;
        stops.push({ align: a['w:val'] || 'left', posPx: twipsToPx(a['w:pos']), leader: a['w:leader'] && a['w:leader'] !== 'none' ? a['w:leader'] : null });
      }
      return stops.length ? stops.sort((x, y) => x.posPx - y.posPx) : undefined;
    })(),
  };
  return out;
}

/** Markers for a font named by theme slot, resolved after the chains flatten. */
const THEME_MAJOR = '@theme:major';
const THEME_MINOR = '@theme:minor';

/**
 * The two Latin faces a theme names: headings (`majorFont`) and body
 * (`minorFont`). Word's own defaults are "Calibri Light" and "Calibri", and
 * nearly every styles.xml Word writes names its fonts through these slots
 * rather than outright — so without the theme every document is Calibri.
 */
export function readThemeFonts(themeXml) {
  const out = { major: 'Calibri Light', minor: 'Calibri' };
  if (!themeXml) return out;
  const face = (slot) => {
    const block = first(themeXml, 'a:' + slot + 'Font');
    const latin = block ? first(block, 'a:latin') : null;
    const typeface = latin ? attrs(latin).typeface : null;
    return typeface && typeface.trim() ? typeface : null;
  };
  out.major = face('major') ?? out.major;
  out.minor = face('minor') ?? out.minor;
  return out;
}

/**
 * The theme's colour scheme, slot -> hex without the hash. Office's own
 * defaults fill any slot a theme omits, and a system colour resolves to the
 * value the writer last saw (`lastClr`).
 */
export function readThemeColours(themeXml) {
  const out = {
    dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6',
    accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47',
    hlink: '0563C1', folHlink: '954F72',
  };
  const scheme = themeXml ? first(themeXml, 'a:clrScheme') : null;
  if (!scheme) return out;
  for (const slot of Object.keys(out)) {
    const el = first(scheme, 'a:' + slot);
    if (!el) continue;
    const srgb = /<a:srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(el);
    const sys = /<a:sysClr\b[^>]*\blastClr="([0-9A-Fa-f]{6})"/.exec(el);
    if (srgb) out[slot] = srgb[1].toUpperCase();
    else if (sys) out[slot] = sys[1].toUpperCase();
  }
  return out;
}

/**
 * Every paragraph style, chains flattened, in CSS pixels.
 *
 * @param {string|null} stylesXml   word/styles.xml
 * @param {{major:string,minor:string}} [themeFonts]  from `readThemeFonts`
 * @returns {Record<string, {sizePx?:number,bold?:boolean,italic?:boolean,colour?:string,fontName?:string,
 *   align?:string,spaceBeforePx?:number,spaceAfterPx?:number,indentPx?:number,name?:string}>}
 */
export function readParagraphStyles(stylesXml, themeFonts = null) {
  if (!stylesXml) return {};
  const theme = themeFonts ?? readThemeFonts(null);
  const raw = new Map();

  for (const m of String(stylesXml).matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
    const a = attrs(m[1]);
    if (a['w:type'] !== 'paragraph') continue;
    const id = a['w:styleId'];
    if (!id) continue;
    raw.set(id, {
      basedOn: val(m[2], 'w:basedOn'),
      name: val(m[2], 'w:name'),
      props: readProps(m[2]),
      isDefault: a['w:default'] === '1' || a['w:default'] === 'true',
    });
  }

  // docDefaults is the floor of every chain.
  const defaults = readProps(first(stylesXml, 'w:docDefaults') ?? '');

  const resolved = {};
  const resolve = (id, seen = new Set()) => {
    if (resolved[id]) return resolved[id];
    const entry = raw.get(id);
    if (!entry || seen.has(id)) return { ...defaults };
    seen.add(id);
    const base = entry.basedOn ? resolve(entry.basedOn, seen) : { ...defaults };
    const merged = { ...base };
    for (const [k, v] of Object.entries(entry.props)) if (v !== undefined) merged[k] = v;
    merged.name = entry.name ?? id;
    resolved[id] = merged;
    return merged;
  };
  for (const id of raw.keys()) resolve(id);

  // The style a paragraph with no pStyle gets: the declared default style if
  // there is one, else docDefaults bare — an unstyled paragraph still obeys the
  // document's base font size, or an 11pt house document renders at our 10.5.
  const defaultId = [...raw.entries()].find(([, e]) => e.isDefault)?.[0];
  resolved['*default*'] = defaultId ? resolved[defaultId] : { ...defaults };

  // A font named by theme slot becomes the theme's face — after flattening, so
  // a heading based on Normal that names the major slot keeps it.
  for (const s of Object.values(resolved)) {
    if (s.fontName === THEME_MAJOR) s.fontName = theme.major;
    else if (s.fontName === THEME_MINOR) s.fontName = theme.minor;
  }
  return resolved;
}

/**
 * Every CHARACTER style, chains flattened, fonts through the theme — the
 * look a run gets from its `w:rStyle`: Hyperlink is blue and underlined,
 * FootnoteReference is superscript, Strong is bold. A run's own properties
 * beat these; these beat the paragraph's. Only run-level fields are kept.
 */
export function readCharacterStyles(stylesXml, themeFonts = null) {
  if (!stylesXml) return {};
  const theme = themeFonts ?? readThemeFonts(null);
  const raw = new Map();
  for (const m of String(stylesXml).matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
    const a = attrs(m[1]);
    if (a['w:type'] !== 'character') continue;
    const id = a['w:styleId'];
    if (!id) continue;
    const rPr = first(m[2], 'w:rPr');
    const props = readProps('<w:style>' + (rPr || '') + '</w:style>');
    const vert = val(rPr, 'w:vertAlign');
    raw.set(id, {
      basedOn: val(m[2], 'w:basedOn'),
      name: val(m[2], 'w:name'),
      props: {
        bold: props.bold, italic: props.italic, underline: props.underline, caps: props.caps, smallCaps: props.smallCaps,
        sizePx: props.sizePx, colour: props.colour, fontName: props.fontName,
        vertAlign: vert && vert !== 'baseline' ? vert : undefined,
      },
    });
  }
  const resolved = {};
  const resolve = (id, seen = new Set()) => {
    if (resolved[id]) return resolved[id];
    const entry = raw.get(id);
    if (!entry || seen.has(id)) return {};
    seen.add(id);
    const base = entry.basedOn ? resolve(entry.basedOn, seen) : {};
    const merged = { ...base };
    for (const [k, v] of Object.entries(entry.props)) if (v !== undefined) merged[k] = v;
    merged.name = entry.name ?? id;
    resolved[id] = merged;
    return merged;
  };
  for (const id of raw.keys()) resolve(id);
  for (const s of Object.values(resolved)) {
    if (s.fontName === THEME_MAJOR) s.fontName = theme.major;
    else if (s.fontName === THEME_MINOR) s.fontName = theme.minor;
  }
  return resolved;
}

/** How a level formats its counter. Everything else falls back to decimal. */
const FORMATS = new Set(['decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman', 'bullet', 'none']);

/**
 * Numbering definitions: numId -> level -> how to label it.
 *
 * Two-layer on purpose, because the format is: `w:num` maps a numId to an
 * abstract definition, and the abstract definition carries the levels. A file
 * can point ten lists at one abstract definition — that is how "restart
 * numbering" works — so collapsing the layers here would lose nothing we use
 * but is where overrides would land later.
 *
 * @returns {Record<string, Array<{format:string,lvlText:string,start:number,indentPx:number}>>}
 */
export function readNumberingDefs(numberingXml) {
  if (!numberingXml) return {};

  const abstract = new Map();
  for (const m of String(numberingXml).matchAll(/<w:abstractNum\b([^>]*)>([\s\S]*?)<\/w:abstractNum>/g)) {
    const id = attrs(m[1])['w:abstractNumId'];
    const levels = [];
    for (const lvl of m[2].matchAll(/<w:lvl\b([^>]*)>([\s\S]*?)<\/w:lvl>/g)) {
      const ilvl = Number(attrs(lvl[1])['w:ilvl'] ?? levels.length);
      const fmt = val(lvl[2], 'w:numFmt') ?? 'decimal';
      const ind = first(lvl[2], 'w:ind');
      const fonts = first(lvl[2], 'w:rFonts');
      levels[ilvl] = {
        format: FORMATS.has(fmt) ? fmt : 'decimal',
        lvlText: val(lvl[2], 'w:lvlText') ?? '%' + (ilvl + 1) + '.',
        // The marker hangs to the left of the text by this much: Word's
        // bullet at a quarter inch with the text at a half.
        hangingPx: ind && attrs(ind)['w:hanging'] !== undefined ? twipsToPx(attrs(ind)['w:hanging']) : null,

        // The font the bullet character was stored for: U+F0A7 is a square
        // in Wingdings and a club in Symbol.
        font: fonts ? (attrs(fonts)['w:ascii'] ?? attrs(fonts)['w:hAnsi'] ?? null) : null,

        start: Number(val(lvl[2], 'w:start') ?? 1) || 1,
        indentPx: ind && attrs(ind)['w:left'] !== undefined ? twipsToPx(attrs(ind)['w:left']) : (ilvl + 1) * 24,
      };
    }
    abstract.set(id, levels);
  }

  const out = {};
  for (const m of String(numberingXml).matchAll(/<w:num\b([^>]*)>([\s\S]*?)<\/w:num>/g)) {
    const numId = attrs(m[1])['w:numId'];
    const abstractId = val(m[2], 'w:abstractNumId');
    if (numId && abstract.has(abstractId)) out[numId] = abstract.get(abstractId);
  }
  return out;
}

export { halfPointsToPx as _halfPointsToPx };

/**
 * The standard paragraph styles a fresh document ships with — and the
 * catalogue the editor falls back to when a file has no styles part at all.
 * One definition, two consumers: `buildDocx({styles: true})` writes it into
 * new templates, and `Document.ensureParagraphStyles()` writes it into an
 * opened file the moment somebody picks a style it has nowhere to store.
 * Sizes are half-points, spacing twentieths; the accent is the house blue.
 */
export const STANDARD_PARAGRAPH_STYLES = [
  { id: 'Normal', name: 'Normal' },
  { id: 'Title', name: 'Title' },
  { id: 'Heading1', name: 'heading 1' },
  { id: 'Heading2', name: 'heading 2' },
  { id: 'Heading3', name: 'heading 3' },
];

export const STANDARD_STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:spacing w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="44"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="1F5F8B"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:spacing w:before="200" w:after="100"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="1F5F8B"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>' +
  '</w:styles>';
