/**
 * Strict Open XML (ISO/IEC 29500 Strict), read and written through one
 * mapping layer.
 *
 * Every engine here reads and writes Transitional: its namespaces, its
 * relationship types, its spellings of values. A Strict package differs from
 * a Transitional one in a known, finite set of ways, so rather than teach
 * each engine both, the package maps a Strict part's text to Transitional as
 * it is read and back to Strict as it is written:
 *
 *   - the namespaces and relationship types: http://purl.oclc.org/ooxml/...
 *     for http://schemas.openxmlformats.org/.../2006/...;
 *   - DrawingML's percentages written "75%" where Transitional writes 75000;
 *   - WordprocessingML's measures written "12pt" where Transitional writes
 *     twips, its start and end where Transitional has left and right, and
 *     its shapes and text boxes in the wordprocessingDrawing namespace;
 *   - SpreadsheetML's dates, ISO 8601 text in a cell of type d;
 *   - VML, which Strict does not have: Office writes it inside
 *     mc:AlternateContent, a choice that requires VML, and that choice is
 *     taken here as a reader that knows VML takes it.
 *
 * Mapping on read sets the inflated text of a part without marking it
 * changed, so a part nothing edits is written back as the Strict bytes it
 * came as. On write, a part an engine changed (or every part, when a
 * Transitional file is saved as Strict) is mapped back. What this layer
 * does not cover is listed in docs/OOXML-NOTES.md.
 */

const PURL = 'http://purl.oclc.org/ooxml/';
const T = 'http://schemas.openxmlformats.org/';
const T_REL = T + 'officeDocument/2006/relationships';
const T_THUMB = T + 'package/2006/relationships/metadata/thumbnail';
const VML = 'urn:schemas-microsoft-com:vml';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const WNE = 'http://schemas.microsoft.com/office/word/2006/wordml';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const WPG = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';

/** The names Strict spells as words run together where Transitional hyphenates them. */
const RENAMED = { extendedProperties: 'extended-properties', customProperties: 'custom-properties' };
const RENAMED_BACK = Object.fromEntries(Object.entries(RENAMED).map(([s, t]) => [t, s]));

/** A Strict namespace or relationship type as Transitional names it. */
export function transitionalUri(uri) {
  if (!uri.startsWith(PURL)) return uri;
  const [area, ...more] = uri.slice(PURL.length).split('/');
  if (area === 'officeDocument' && more[0] === 'relationships') {
    if (more.length === 1) return T_REL;
    const type = more.slice(1).join('/');
    if (type === 'metadata/thumbnail') return T_THUMB;
    return T_REL + '/' + (RENAMED[type] ?? type);
  }
  const name = more.join('/');
  return T + area + '/2006/' + (RENAMED[name] ?? name);
}

/**
 * A Transitional namespace or relationship type as Strict names it. The
 * relationship to a VML drawing stays Transitional, as Excel writes it:
 * Strict has no VML of its own to name. So does the schema library's
 * namespace, as Word writes it in a Strict document's settings.
 */
export function strictUri(uri) {
  if (uri === T_THUMB) return PURL + 'officeDocument/relationships/metadata/thumbnail';
  const m = /^http:\/\/schemas\.openxmlformats\.org\/(officeDocument|wordprocessingml|spreadsheetml|presentationml|drawingml)\/2006\/(.+)$/.exec(uri);
  if (!m) return uri;
  const [, area, rest] = m;
  if (area === 'officeDocument' && (rest === 'relationships' || rest.startsWith('relationships/'))) {
    const type = rest.slice('relationships/'.length);
    if (!type) return PURL + 'officeDocument/relationships';
    if (type === 'vmlDrawing') return uri;
    return PURL + 'officeDocument/relationships/' + (RENAMED_BACK[type] ?? type);
  }
  return PURL + area + '/' + (RENAMED_BACK[rest] ?? rest);
}

/** The attributes that carry a namespace or a relationship type: declarations, Type, and graphicData's uri. */
const URI_ATTR = /(\s(?:xmlns(?::[\w.-]+)?|Type|uri)\s*=\s*")([^"]*)"/g;
const mapUris = (xml, fn) => xml.replace(URI_ATTR, (all, head, uri) => {
  const to = fn(uri);
  return to === uri ? all : head + to + '"';
});

/** Whether a package is Strict: what its root relationship to the main part is called. */
export function isStrictRels(relsXml) {
  return /Type\s*=\s*"http:\/\/purl\.oclc\.org\/ooxml\/officeDocument\/relationships\/officeDocument"/.test(relsXml);
}

/** Every prefix a part declares, and the default namespace under '', mapped to Transitional. */
function prefixes(xml) {
  const out = new Map();
  for (const m of xml.matchAll(/\sxmlns(?::([\w.-]+))?\s*=\s*"([^"]*)"/g)) {
    const p = m[1] ?? '';
    if (!out.has(p)) out.set(p, transitionalUri(m[2]));
  }
  return out;
}

/* ── DrawingML percentages ──────────────────────────────────────────────
 *
 * Strict writes ST_Percentage and its kin as "75%"; Transitional writes the
 * same value in thousandths, 75000. Read: any plain attribute of a DrawingML
 * or PresentationML element that is a number and a percent sign, the free
 * text ones aside. Write: the attributes these types are known on.
 */

const DRAWING_NS = /^(http:\/\/schemas\.openxmlformats\.org\/(drawingml|presentationml)\/|http:\/\/schemas\.microsoft\.com\/office\/(drawing|powerpoint)\/)/;
const FREE_TEXT = new Set(['name', 'descr', 'title', 'typeface', 'macro', 'textlink', 'fmla', 'uri', 'lang', 'altLang', 'id', 'val_text']);
const COLOUR_MODS = ['tint', 'shade', 'alpha', 'alphaOff', 'alphaMod', 'hueMod', 'sat', 'satOff', 'satMod', 'lum', 'lumOff', 'lumMod',
  'red', 'redOff', 'redMod', 'green', 'greenOff', 'greenMod', 'blue', 'blueOff', 'blueMod'];
/** The percentage attributes Strict spells with a percent sign, by element: the ones written back. */
export const PERCENT_ATTRS = new Map([
  ...COLOUR_MODS.map((e) => [e, ['val']]),
  ['tint', ['val', 'amt']],
  ['scrgbClr', ['r', 'g', 'b']], ['hslClr', ['sat', 'lum']], ['hsl', ['sat', 'lum']],
  ['gs', ['pos']], ['fillToRect', ['l', 't', 'r', 'b']], ['tileRect', ['l', 't', 'r', 'b']], ['srcRect', ['l', 't', 'r', 'b']], ['fillRect', ['l', 't', 'r', 'b']],
  ['tile', ['sx', 'sy']], ['miter', ['lim']],
  ['rPr', ['baseline']], ['defRPr', ['baseline']], ['endParaRPr', ['baseline']],
  ['spcPct', ['val']], ['buSzPct', ['val']], ['normAutofit', ['fontScale', 'lnSpcReduction']],
  ['alphaModFix', ['amt']], ['alphaBiLevel', ['thresh']], ['biLevel', ['thresh']], ['lum', ['val', 'bright', 'contrast']],
  ['outerShdw', ['sx', 'sy']], ['reflection', ['stA', 'stPos', 'endA', 'endPos', 'sx', 'sy']],
  ['tav', ['tm']], ['restoredLeft', ['sz']], ['restoredTop', ['sz']], ['cTn', ['spd', 'accel', 'decel']], ['tmPct', ['val']],
  ['by', ['x', 'y']], ['from', ['x', 'y']], ['to', ['x', 'y']], ['rCtr', ['x', 'y']],
]);

const TAG = /<([\w.-]+:)?([\w.-]+)(\s[^<>]*?)?(\/?)>/g;

/** Rewrite the plain attributes of the drawing elements of a part, `fn(local, attr, value)` answering a new value or undefined. */
function drawingAttrs(xml, fn) {
  const ns = prefixes(xml);
  if (![...ns.values()].some((u) => DRAWING_NS.test(u))) return xml;
  return xml.replace(TAG, (all, pfx, local, rest, close) => {
    if (!rest) return all;
    const uri = ns.get(pfx ? pfx.slice(0, -1) : '');
    if (!uri || !DRAWING_NS.test(uri)) return all;
    let changed = false;
    const next = rest.replace(/(\s)([\w.-]+)(\s*=\s*")([^"]*)"/g, (a, sp, attr, eq, value) => {
      const to = fn(local, attr, value);
      if (to === undefined || to === value) return a;
      changed = true;
      return sp + attr + eq + to + '"';
    });
    return changed ? '<' + (pfx || '') + local + next + close + '>' : all;
  });
}

const PCT = /^(-?\d+(?:\.\d+)?)%$/;
const percentsToThousandths = (xml) => (xml.includes('%"') ? drawingAttrs(xml, (local, attr, v) => {
  const m = PCT.exec(v);
  return m && !FREE_TEXT.has(attr) ? String(Math.round(Number(m[1]) * 1000)) : undefined;
}) : xml);
const thousandthsToPercents = (xml) => drawingAttrs(xml, (local, attr, v) => {
  if (!PERCENT_ATTRS.get(local)?.includes(attr) || !/^-?\d+$/.test(v)) return undefined;
  return String(Number(v) / 1000) + '%';
});

/* ── WordprocessingML ─────────────────────────────────────────────────── */

const WML = T + 'wordprocessingml/2006/main';
/** Points in each universal measure. */
const PER_UNIT = { pt: 1, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, pc: 12, pi: 12 };
/** Elements whose measure is in half-points; every other measure is in twips. */
const HALF_POINTS = new Set(['sz', 'szCs', 'kern', 'position', 'hps', 'hpsRaise', 'hpsBaseText']);
const MEASURE = /^(-?\d+(?:\.\d+)?)(pt|in|cm|mm|pc|pi)$/;
const JC = { start: 'left', end: 'right' };
const JC_BACK = { left: 'start', right: 'end' };
const IND = { 'w:start': 'w:left', 'w:end': 'w:right', 'w:startChars': 'w:leftChars', 'w:endChars': 'w:rightChars' };
const IND_BACK = Object.fromEntries(Object.entries(IND).map(([s, t]) => [t, s]));
/**
 * A font's character set: Strict names it (w:characterSet, as Word writes
 * the names), Transitional numbers it (w:val, the Windows charset in hex).
 */
const CHARSETS_BACK = {
  '00': 'iso-8859-1', 80: 'shift_jis', 81: 'ks_c-5601-1987', 86: 'GB2312', 88: 'Big5', A1: 'windows-1253', A2: 'windows-1254',
  A3: 'windows-1258', B1: 'windows-1255', B2: 'windows-1256', BA: 'windows-1257', CC: 'windows-1251', DE: 'windows-874', EE: 'windows-1250',
};
const CHARSETS = Object.fromEntries(Object.entries(CHARSETS_BACK).map(([n, name]) => [name.toLowerCase(), n]));
/** The containers whose sides Strict calls start and end. */
const SIDED = /<w:(tblBorders|tcBorders|tblCellMar|tcMar)\b[^>]*>[\s\S]*?<\/w:\1>/g;

/** Each w: tag's attributes rewritten by `fn(local, attrs)`, which edits the [name, value] list in place. */
function wTags(xml, fn) {
  return xml.replace(/<w:([\w.-]+)(\s[^<>]*?)?(\/?)>/g, (all, local, rest, close) => {
    if (!rest) return all;
    const list = [...rest.matchAll(/\s([\w.:-]+)\s*=\s*"([^"]*)"/g)].map((m) => [m[1], m[2]]);
    const before = JSON.stringify(list);
    fn(local, list);
    if (JSON.stringify(list) === before) return all;
    return '<w:' + local + list.map(([n, v]) => ` ${n}="${v}"`).join('') + close + '>';
  });
}

function wordToTransitional(xml) {
  let out = wTags(xml, (local, list) => {
    const pct = list.some(([n, v]) => n === 'w:type' && v === 'pct');
    for (const pair of list) {
      const [n, v] = pair;
      if (local === 'ind' && IND[n]) pair[0] = IND[n];
      if (local === 'charset' && n === 'w:characterSet' && CHARSETS[v.toLowerCase()]) pair.splice(0, 2, 'w:val', CHARSETS[v.toLowerCase()]);
      const m = MEASURE.exec(v);
      const p = PCT.exec(v);
      if (m) pair[1] = String(Math.round(Number(m[1]) * PER_UNIT[m[2]] * (HALF_POINTS.has(local) ? 2 : 20)));
      else if (p) pair[1] = String(Math.round(Number(p[1]) * (pct ? 50 : 1)));
      else if ((local === 'jc' || local === 'lvlJc' || local === 'tab') && n === 'w:val' && JC[v]) pair[1] = JC[v];
    }
  });
  out = out.replace(SIDED, (block) => block.replace(/<(\/?)w:start\b/g, '<$1w:left').replace(/<(\/?)w:end\b/g, '<$1w:right'));
  return shapesToTransitional(out);
}

function wordToStrict(xml) {
  let out = shapesToStrict(xml);
  out = out.replace(SIDED, (block) => block.replace(/<(\/?)w:left\b/g, '<$1w:start').replace(/<(\/?)w:right\b/g, '<$1w:end'));
  out = wTags(out, (local, list) => {
    const pct = list.some(([n, v]) => n === 'w:type' && v === 'pct');
    for (const pair of list) {
      const [n, v] = pair;
      if (local === 'charset' && n === 'w:val' && CHARSETS_BACK[v.toUpperCase()]) pair.splice(0, 2, 'w:characterSet', CHARSETS_BACK[v.toUpperCase()]);
      if ((local === 'jc' || local === 'lvlJc' || local === 'tab') && n === 'w:val' && JC_BACK[v]) pair[1] = JC_BACK[v];
      if (local === 'ind' && IND_BACK[n]) pair[0] = IND_BACK[n];
      if (/^-?\d+(\.\d+)?$/.test(v) && ((pct && n === 'w:w') || (local === 'zoom' && n === 'w:percent') || (local === 'w' && n === 'w:val'))) {
        pair[1] = String(pct && n === 'w:w' ? Number(v) / 50 : Number(v)) + '%';
      }
    }
  });
  // Settings' VML shape defaults: Word leaves them out of a Strict file.
  out = out.replace(/<w:(hdrShapeDefaults|shapeDefaults)\b[^>]*?(\/>|>[\s\S]*?<\/w:\1>)/g, '');
  return out;
}

/**
 * Word's Strict shapes: a shape, a group and a text box are written in the
 * wordprocessingDrawing namespace (wp:wsp, wp:wgp, wp:txbx) with the text
 * box's content as wne:txbxContent, where Transitional writes wps:, wpg: and
 * w:txbxContent. Read, the tags inside a shape or a group take the prefixes
 * Transitional gives them, a drawing within a text box keeping its own.
 */
function shapesToTransitional(xml) {
  if (!/<wp:(wsp|wgp)\b/.test(xml)) return xml;
  const stack = [];
  let out = xml.replace(/<(\/?)wp:([\w.-]+)([^<>]*?)(\/?)>/g, (all, end, local, rest, self) => {
    if (end) {
      const top = stack.pop() ?? 'wp';
      return top === 'wp' ? all : `</${top}:${local}>`;
    }
    const top = stack[stack.length - 1] ?? 'wp';
    let as = top;
    if (local === 'wsp') as = 'wps';
    else if (local === 'wgp' || local === 'grpSp') as = 'wpg';
    else if (local === 'anchor' || local === 'inline') as = 'wp';
    else if (top === 'wpg' && !['cNvPr', 'cNvGrpSpPr', 'grpSpPr', 'graphicFrame', 'cNvFrPr', 'xfrm', 'extLst'].includes(local)) as = 'wp';
    if (!self) stack.push(as);
    return as === 'wp' ? all : `<${as}:${local}${rest}${self}>`;
  });
  out = out.replace(/<wne:txbxContent\b/g, '<w:txbxContent').replace(/<\/wne:txbxContent>/g, '</w:txbxContent>');
  return declare(declare(out, 'wps', WPS), 'wpg', WPG);
}

function shapesToStrict(xml) {
  // A shape's choice is taken, and its VML fallback left out, as Word does.
  let out = resolveChoices(xml, (req) => /^(wps|wpg)$/.test(req.trim()));
  if (!/<\/?(wps|wpg):/.test(out)) return out;
  out = out.replace(/<(\/?)(?:wps|wpg):/g, '<$1wp:');
  out = out.replace(/<wp:txbx\b[^>]*>[\s\S]*?<\/wp:txbx>/g, (box) => box
    .replace(/<w:txbxContent\b/g, '<wne:txbxContent').replace(/<\/w:txbxContent>/g, '</wne:txbxContent>'));
  return ignorable(declare(out, 'wne', WNE), 'wne');
}

/* ── SpreadsheetML dates ─────────────────────────────────────────────── */

const DAY = 86400000;
/**
 * A cell of type d holds an ISO 8601 date, time or both. As a serial number
 * in the workbook's date system: days since 30 December 1899 (or 1 January
 * 1904), the fraction the time of day, and the 1900 system's day for
 * 29 February 1900 kept, unless the workbook says it does not keep it.
 */
export function isoToSerial(text, { date1904 = false, compat = true } = {}) {
  const m = /^(?:(-?\d{4,})-(\d{2})-(\d{2}))?(?:T?(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?(?:Z|[+-]\d{2}:?\d{2})?$/.exec(String(text).trim());
  if (!m || (!m[1] && !m[4])) return null;
  let days = 0;
  if (m[1]) {
    const t = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    days = (t - (date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30))) / DAY;
    if (!date1904 && compat && t < Date.UTC(1900, 2, 1)) days -= 1;
  }
  const time = m[4] ? (Number(m[4]) * 3600 + Number(m[5]) * 60 + Number(m[6] || 0)) / 86400 : 0;
  return days + time;
}

function cellDates(xml, system) {
  if (!xml.includes('t="d"')) return xml;
  return xml.replace(/<c\b([^>]*?)\st="d"([^>]*)>(\s*(?:<f\b[^>]*\/>|<f\b[^>]*>[^<]*<\/f>)?\s*<v>)([^<]*)(<\/v>)/g, (all, a, b, open, value, close) => {
    const serial = isoToSerial(value, system);
    return serial === null ? all : `<c${a}${b}>${open}${Number(serial.toFixed(10))}${close}`;
  });
}

/** A cell border's sides: start and end in Strict, left and right in Transitional. */
const borderSides = (xml, [from, to]) => (xml.includes('<border') ? xml.replace(/<border\b[^>]*?(?:\/>|>[\s\S]*?<\/border>)/g, (border) => border
  .replace(new RegExp(`<(/?)${from[0]}\\b`, 'g'), (_, slash) => `<${slash}${to[0]}`).replace(new RegExp(`<(/?)${from[1]}\\b`, 'g'), (_, slash) => `<${slash}${to[1]}`)) : xml);

/* ── markup compatibility and VML ────────────────────────────────────── */

/** The end of the element that opens at `at`, counting the same name nested within. */
function elementEnd(xml, at, name) {
  const re = new RegExp(`<(/?)${name}\\b[^>]*?(/?)>`, 'g');
  re.lastIndex = at;
  let depth = 0;
  for (let m; (m = re.exec(xml));) {
    if (m[1]) depth -= 1;
    else if (!m[2]) depth += 1;
    if (depth === 0) return m.index + m[0].length;
  }
  return -1;
}

/** Each mc:AlternateContent whose first choice's Requires `pick` accepts, replaced by that choice's content. */
function resolveChoices(xml, pick) {
  if (!xml.includes('mc:AlternateContent')) return xml;
  let out = '';
  let from = 0;
  for (let at = xml.indexOf('<mc:AlternateContent', from); at >= 0; at = xml.indexOf('<mc:AlternateContent', from)) {
    const end = elementEnd(xml, at, 'mc:AlternateContent');
    if (end < 0) break;
    const whole = xml.slice(at, end);
    const choice = /<mc:Choice\b([^>]*?)(\/?)>/.exec(whole);
    const req = choice && /Requires\s*=\s*"([^"]*)"/.exec(choice[1])?.[1];
    if (!choice || req == null || !pick(req)) {
      out += xml.slice(from, at + 1);
      from = at + 1;
      continue;
    }
    let inner = '';
    if (!choice[2]) {
      const open = choice.index + choice[0].length;
      const close = elementEnd(whole, choice.index, 'mc:Choice');
      inner = whole.slice(open, close - '</mc:Choice>'.length);
    }
    out += xml.slice(from, at) + inner;
    from = end;
  }
  return out + xml.slice(from);
}

/** VML's choices taken, as a reader that knows VML takes them: where Transitional writes the VML itself. */
const vmlChoices = (xml) => (xml.includes('Requires="v"') ? resolveChoices(xml, (req) => req.trim() === 'v') : xml);

/** VML wrapped, as Office wraps it in a Strict file: a choice that requires it, so a reader without VML passes it by. */
function wrapVml(xml, tags) {
  let out = xml;
  for (const tag of tags) {
    const re = new RegExp(`<${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</${tag}>)`, 'g');
    out = out.replace(re, (el) => `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:v="${VML}" Requires="v">${el}</mc:Choice></mc:AlternateContent>`);
  }
  return out;
}

/* ── the root element ────────────────────────────────────────────────── */

/** The root element's start tag: its span in `xml`. */
function rootTag(xml) {
  const m = /<(?![?!])[^>]+>/.exec(xml);
  return m ? { at: m.index, tag: m[0] } : null;
}

/** A prefix declared on the root, unless the part declares it already. */
function declare(xml, prefix, uri) {
  if (new RegExp(`\\sxmlns:${prefix}\\s*=`).test(xml)) return xml;
  const root = rootTag(xml);
  if (!root) return xml;
  const tag = root.tag.replace(/(\/?)>$/, (_, slash) => ` xmlns:${prefix}="${uri}"${slash}>`);
  return xml.slice(0, root.at) + tag + xml.slice(root.at + root.tag.length);
}

/** A prefix added to the root's mc:Ignorable, when the root has one. */
function ignorable(xml, prefix) {
  const root = rootTag(xml);
  if (!root) return xml;
  const m = /\smc:Ignorable\s*=\s*"([^"]*)"/.exec(root.tag);
  if (!m || m[1].split(/\s+/).includes(prefix)) return xml;
  const tag = root.tag.replace(m[0], () => ` mc:Ignorable="${(m[1] + ' ' + prefix).trim()}"`);
  return xml.slice(0, root.at) + tag + xml.slice(root.at + root.tag.length);
}

/** The root's conformance attribute: taken away (null) or set. */
function conformance(xml, value, word) {
  const root = rootTag(xml);
  if (!root) return xml;
  let tag = root.tag.replace(/\s(?:w:)?conformance\s*=\s*"[^"]*"/, '');
  if (value) tag = tag.replace(/(\/?)>$/, (_, slash) => ` ${word ? 'w:' : ''}conformance="${value}"${slash}>`);
  return tag === root.tag ? xml : xml.slice(0, root.at) + tag + xml.slice(root.at + root.tag.length);
}

/* ── the two directions ──────────────────────────────────────────────── */

/** Parts whose text is mapped: XML and relationships, VML drawings among them. */
export const isMappedPart = (name) => /\.(xml|rels|vml)$/i.test(name);

/**
 * One part of a Strict package as Transitional writes it. `context` carries
 * what a part cannot say for itself: { main, dates }, the main part's name
 * and the workbook's date system.
 */
export function partToTransitional(name, xml, context = {}) {
  let out = mapUris(xml, transitionalUri);
  if (name === context.main) out = conformance(out, null);
  out = vmlChoices(out);
  out = percentsToThousandths(out);
  const ns = prefixes(out);
  if ([...ns.values()].includes(WML)) out = wordToTransitional(out);
  if (context.dates && /^xl\/worksheets\/[^/]+\.xml$/.test(name)) out = cellDates(out, context.dates);
  if (name === 'xl/styles.xml') out = borderSides(out, [['start', 'end'], ['left', 'right']]);
  return out;
}

/** One part, as Transitional writes it, written as Strict. */
export function partToStrict(name, xml, context = {}) {
  let out = xml;
  const ns = prefixes(out);
  if ([...ns.values()].includes(WML)) {
    out = wordToStrict(out);
    out = wrapVml(out, ['w:pict', 'w:object']);
  }
  // A sheet's legacyDrawing, its notes' VML, is left as it is: Excel, reading
  // Strict, passes over a choice that requires VML (its own Strict files say
  // each note again in DrawingML, in the fallback), and reads a bare one.
  if (name === 'xl/styles.xml') out = borderSides(out, [['left', 'right'], ['start', 'end']]);
  out = thousandthsToPercents(out);
  out = mapUris(out, strictUri);
  if (name === context.main) out = conformance(out, 'strict', [...ns.values()].includes(WML));
  return out;
}

/** The workbook's date system, read from its workbookPr. */
export function dateSystem(workbookXml) {
  const pr = /<(?:\w+:)?workbookPr\b([^>]*)>/.exec(workbookXml)?.[1] ?? '';
  const flag = (n) => new RegExp(`\\b${n}\\s*=\\s*"(1|true)"`).test(pr);
  const off = (n) => new RegExp(`\\b${n}\\s*=\\s*"(0|false)"`).test(pr);
  return { date1904: flag('date1904'), compat: !off('dateCompatibility') };
}
