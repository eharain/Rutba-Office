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
  const out = {
    bold: toggle(rPr, 'w:b'),
    italic: toggle(rPr, 'w:i'),
    sizePx: sz !== null ? halfPointsToPx(sz) : undefined,
    colour: (() => {
      const c = val(rPr, 'w:color');
      return c && c !== 'auto' ? '#' + c.toLowerCase() : undefined;
    })(),
    align: val(pPr, 'w:jc') ?? undefined,
    spaceBeforePx: spacing ? (attrs(spacing)['w:before'] !== undefined ? twipsToPx(attrs(spacing)['w:before']) : undefined) : undefined,
    spaceAfterPx: spacing ? (attrs(spacing)['w:after'] !== undefined ? twipsToPx(attrs(spacing)['w:after']) : undefined) : undefined,
    indentPx: ind ? (attrs(ind)['w:left'] !== undefined ? twipsToPx(attrs(ind)['w:left']) : undefined) : undefined,
  };
  return out;
}

/**
 * Every paragraph style, chains flattened, in CSS pixels.
 *
 * @returns {Record<string, {sizePx?:number,bold?:boolean,italic?:boolean,colour?:string,
 *   align?:string,spaceBeforePx?:number,spaceAfterPx?:number,indentPx?:number,name?:string}>}
 */
export function readParagraphStyles(stylesXml) {
  if (!stylesXml) return {};
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
      levels[ilvl] = {
        format: FORMATS.has(fmt) ? fmt : 'decimal',
        lvlText: val(lvl[2], 'w:lvlText') ?? '%' + (ilvl + 1) + '.',
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
