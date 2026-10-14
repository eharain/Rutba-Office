/**
 * Redaction over a whole package, a document, a workbook or a deck, as
 * Apply takes words away: every part of it, read as the words a person
 * sees and the words a reader of the file would find, the marked ones
 * replaced by black bars of the same length.
 *
 * Every XML part is read the same way. Its runs of words (w:t, a:t, t,
 * w:delText) are read together, paragraph by paragraph, so a name split
 * across two runs is still found; a field's code (w:instrText) on its own;
 * every other element's words (a title, an author in the properties, a
 * cell's value) one by one; and the attributes that hold words a person
 * wrote (an author, alt text, a list item, a watermark's words, a link's
 * address) each by itself. A worksheet's number found by a pattern becomes
 * a text cell of bars, as a number of bars cannot be. A picture is replaced
 * by the bytes the window blacked out. A package inside (an embedded
 * workbook) is redacted in its turn. Other bytes are searched as text and
 * a match is overwritten in place with marks of the same length, so
 * nothing round it moves. The thumbnail (a picture of the first page) is
 * dropped.
 */
import { OoxmlPackage, esc } from './package.js';
import { unesc } from './workbook.js';
import { compileMarks, findSpans, barsOver, BAR } from './redact-match.js';

export { compileMarks, findSpans, barsOver, BAR, PATTERNS, luhn } from './redact-match.js';

/** Elements whose words run on from one to the next within a paragraph. */
const RUN_TEXT = new Set(['w:t', 'w:delText', 'a:t', 't']);
/** A field's code, read as one stretch of words. */
const INSTR = new Set(['w:instrText', 'w:delInstrText']);
/** Elements that end a paragraph's words: what is on either side is not one line. */
const BOUNDARY = /^(?:w:p|a:p|si|is|r|w:tc|w:tr|w:tbl|w:txbxContent|w:comment|w:footnote|w:endnote|w:sdtContent|a:tc|a:tr|p:txBody|c:rich|w:fldChar|text|comment)$/;
/** A tab or a break inside a paragraph, which the words are read across. */
const SEPARATOR = { 'w:tab': '\t', 'w:br': '\n', 'w:cr': '\n', 'a:br': '\n', 'w:noBreakHyphen': '-' };

/** Attributes that hold words a person wrote, by their local name. */
const WORD_ATTRS = new Set(['author', 'initials', 'descr', 'title', 'name', 'displayName', 'userId', 'displayText', 'value', 'lastValue', 'instr', 'string', 'alt', 'tooltip', 'providerId']);
/** Elements whose `w:val` holds words: a control's title and tag, a form field's name, help and entries. */
const VAL_ELEMENTS = new Set(['w:alias', 'w:tag', 'w:name', 'w:helpText', 'w:statusText', 'w:default', 'w:listEntry', 'w:format']);

const localName = (name) => name.slice(name.indexOf(':') + 1);
/** Words between tags written back as they were written, quotes and all: only what must be escaped is. */
const escText = (s) => esc(s).replace(/&quot;/g, '"');

/**
 * One XML part's words redacted. Answers the new XML and how many matches
 * were taken away (0 leaves the part as it was). `found`, when given,
 * collects each match's words (for the scan, not the apply).
 */
export function redactXml(xml, finders, { found = null, external = null } = {}) {
  const tokens = String(xml).split(/(<[^>]*>)/);
  let count = 0;
  const stack = [];
  // Open stretches: run words and field codes, each a list of pieces { i, text } (i null for a separator).
  let words = [];
  let codes = [];
  const flush = (pieces) => {
    if (!pieces.length) return;
    const joined = pieces.map((p) => p.text).join('');
    const spans = findSpans(joined, finders);
    if (spans.length) {
      count += spans.length;
      if (found) for (const [a, b] of spans) found.push(joined.slice(a, b));
      let at = 0;
      for (const p of pieces) {
        const from = at;
        const to = at + p.text.length;
        at = to;
        if (p.i == null) continue;
        const local = spans.filter(([a, b]) => a < to && b > from).map(([a, b]) => [Math.max(a, from) - from, Math.min(b, to) - from]);
        if (local.length) tokens[p.i] = escText(barsOver(p.text, local));
      }
    }
  };
  const flushAll = () => {
    flush(words);
    flush(codes);
    words = [];
    codes = [];
  };
  for (let i = 0; i < tokens.length; i++) {
    const tok = tokens[i];
    if (!tok) continue;
    if (tok[0] === '<') {
      if (tok.startsWith('<?') || tok.startsWith('<!')) continue;
      const close = tok.startsWith('</');
      const self = tok.endsWith('/>');
      const name = /^<\/?([^\s/>]+)/.exec(tok)?.[1] || '';
      if (BOUNDARY.test(name)) flushAll();
      if (self && SEPARATOR[name] && words.length) words.push({ i: null, text: SEPARATOR[name] });
      if (!close) {
        const next = redactAttrs(tok, name, finders, found, external);
        if (next.count) {
          tokens[i] = next.tag;
          count += next.count;
        }
      }
      if (close) {
        while (stack.length && stack.pop() !== name) { /* an unbalanced part is read as far as it goes */ }
      } else if (!self) stack.push(name);
      continue;
    }
    const parent = stack[stack.length - 1] || '';
    const text = unesc(tok);
    if (!text.trim() && !RUN_TEXT.has(parent)) continue;
    if (RUN_TEXT.has(parent)) words.push({ i, text });
    else if (INSTR.has(parent)) codes.push({ i, text });
    else {
      // Words of their own: a property, a cell's value, an author's name.
      const spans = findSpans(text, finders);
      if (spans.length) {
        count += spans.length;
        if (found) for (const [a, b] of spans) found.push(text.slice(a, b));
        tokens[i] = escText(barsOver(text, spans));
      }
    }
  }
  flushAll();
  return { xml: count ? tokens.join('') : String(xml), count };
}

/** A start tag's word-holding attributes redacted. */
function redactAttrs(tag, name, finders, found, external) {
  let count = 0;
  const next = tag.replace(/(\s)([\w:.-]+)="([^"]*)"/g, (whole, sp, attr, raw) => {
    const local = localName(attr);
    const words = WORD_ATTRS.has(local) || (local === 'val' && VAL_ELEMENTS.has(name)) || (local === 'v' && name === 's')
      || (attr === 'Target' && external?.(tag));
    if (!words) return whole;
    const value = unesc(raw);
    const spans = findSpans(value, finders);
    if (!spans.length) return whole;
    count += spans.length;
    if (found) for (const [a, b] of spans) found.push(value.slice(a, b));
    return sp + attr + '="' + esc(barsOver(value, spans)) + '"';
  });
  return { tag: next, count };
}

/**
 * A worksheet's cells with a number a mark finds: made a text cell of bars
 * (an inline string, its formula dropped), since bars are not a number.
 * Words in cells are left to the XML pass.
 */
function redactSheetNumbers(xml, finders) {
  let count = 0;
  const out = String(xml).replace(/<c\b([^>]*)>([\s\S]*?)<\/c>/g, (whole, a, inner) => {
    const type = /\bt="([^"]*)"/.exec(a)?.[1] || 'n';
    if (type !== 'n') return whole;
    const v = /<v>([^<]*)<\/v>/.exec(inner)?.[1];
    if (v == null) return whole;
    const spans = findSpans(v, finders);
    if (!spans.length) return whole;
    count += spans.length;
    const attrs = a.replace(/\s+t="[^"]*"/, '');
    return '<c' + attrs + ' t="inlineStr"><is><t>' + esc(barsOver(v, spans)) + '</t></is></c>';
  });
  return { xml: out, count };
}

/** Bytes that are not XML or a package, searched as UTF-8 and UTF-16 words, each match overwritten in place. */
function redactBytes(bytes, finders) {
  let count = 0;
  const out = Buffer.from(bytes);
  for (const enc of ['latin1', 'utf16le']) {
    const text = out.toString(enc);
    const spans = findSpans(text, finders);
    for (const [a, b] of spans) {
      count += 1;
      if (enc === 'utf16le') {
        for (let k = a; k < b; k++) out.writeUInt16LE(BAR.charCodeAt(0), k * 2);
      } else {
        out.fill(0x2A, a, b);
      }
    }
  }
  return { bytes: out, count };
}

/** What a part is, for saying where something was found. */
export function partKind(name) {
  const n = String(name);
  if (/^docProps\/thumbnail/.test(n)) return 'thumbnail';
  if (/^docProps\//.test(n) || /^customXml\/itemProps/.test(n)) return 'properties';
  if (/^customXml\//.test(n)) return 'customXml';
  if (/^word\/document\.xml$/.test(n)) return 'body';
  if (/^word\/header\d*\.xml$/.test(n)) return 'headers';
  if (/^word\/footer\d*\.xml$/.test(n)) return 'footers';
  if (/^word\/(?:foot|end)notes\.xml$/.test(n) || /^ppt\/notesSlides\//.test(n)) return 'notes';
  if (/comment|people|persons|authors/i.test(n)) return 'comments';
  if (/^word\/glossary\//.test(n)) return 'placeholders';
  if (/\/charts\//.test(n) || /\/diagrams\//.test(n)) return 'charts';
  if (/\/embeddings\//.test(n)) return 'embedded';
  if (/\/media\//.test(n)) return 'pictures';
  if (/^ppt\/slides\//.test(n)) return 'slides';
  if (/^ppt\/(?:slideMasters|slideLayouts|notesMasters|handoutMasters)\//.test(n)) return 'masters';
  if (/^xl\/(?:worksheets|sharedStrings)/.test(n)) return 'cells';
  if (/^xl\/workbook\.xml$/.test(n)) return 'workbook';
  if (/_rels\//.test(n)) return 'links';
  return 'other';
}

const isXml = (name, pkg) => /\.(?:xml|rels|vml)$/i.test(name) || /xml/.test(pkg.contentTypeOf(name) || '');
const isPackage = (name) => /\.(?:docx|docm|xlsx|xlsm|pptx|pptm)$/i.test(name);

/** An external relationship's Target holds an address a person wrote; an internal one names a part. */
const externalRel = (tag) => /TargetMode="External"/.test(tag);

/**
 * Scan or apply. With `apply` false, nothing changes and the answer says
 * what would go and where: `{ total, parts: [{ part, kind, count }], words }`.
 * With `apply`, the package is redacted in place (`pictures` maps a media
 * part to the blacked-out bytes that replace it) and the same is answered.
 */
export function redactPackage(pkg, { marks = [], pictures = {}, apply = true, depth = 0 } = {}) {
  const finders = compileMarks(marks);
  const parts = [];
  const words = [];
  let total = 0;
  const note = (part, count) => {
    if (!count) return;
    total += count;
    parts.push({ part, kind: partKind(part), count });
  };
  for (const name of pkg.partNames()) {
    if (name === '[Content_Types].xml') continue;
    if (pictures[name] && apply) {
      pkg.write_(name, Buffer.from(pictures[name]));
      note(name, 1);
      continue;
    }
    if (!finders.length) continue;
    if (isXml(name, pkg)) {
      const found = apply ? null : words;
      const sheet = /^xl\/worksheets\/[^/]+\.xml$/.test(name);
      let xml = pkg.text(name);
      let count = 0;
      if (sheet) {
        const r = redactSheetNumbers(xml, finders);
        if (!apply && r.count) words.push(...findSpansText(xml, finders));
        xml = r.xml;
        count += r.count;
      }
      const r = redactXml(xml, finders, { found, external: name.endsWith('.rels') ? externalRel : null });
      count += r.count;
      if (apply && count) pkg.write_(name, r.xml);
      note(name, count);
    } else if (isPackage(name) && depth < 3) {
      try {
        const inner = OoxmlPackage.read(pkg.read(name));
        const r = redactPackage(inner, { marks, apply, depth: depth + 1 });
        if (r.total) {
          if (apply) pkg.write_(name, Buffer.from(inner.write()));
          words.push(...r.words);
          note(name, r.total);
        }
      } catch { /* not a package after all: read as bytes below */ }
    } else if (!/\/media\//.test(name) && !/^docProps\/thumbnail/.test(name)) {
      const r = redactBytes(pkg.read(name), finders);
      if (apply && r.count) pkg.write_(name, r.bytes);
      note(name, r.count);
    }
  }
  // The thumbnail is a picture of the first page, words and all.
  const thumbs = pkg.partNames().filter((n) => /^docProps\/thumbnail\./i.test(n));
  if (apply && (total || Object.keys(pictures).length) && thumbs.length) {
    for (const thumb of thumbs) pkg.removePart(thumb);
    if (pkg.has('_rels/.rels')) pkg.write_('_rels/.rels', pkg.text('_rels/.rels').replace(/<Relationship\b[^>]*Target="\/?docProps\/thumbnail\.[^"]*"[^>]*\/>/g, ''));
  }
  return { total, parts, words };
}

/** The words a sheet's numbers would lose, for the scan. */
function findSpansText(xml, finders) {
  const out = [];
  for (const m of String(xml).matchAll(/<c\b(?![^>]*\bt="(?:s|str|inlineStr|b|e)")[^>]*>[\s\S]*?<v>([^<]*)<\/v>/g)) {
    for (const [a, b] of findSpans(m[1], finders)) out.push(m[1].slice(a, b));
  }
  return out;
}

/**
 * Which parts of a package still hold any of `strings`, read as words in
 * every XML part (entities undone), as UTF-8 and UTF-16 in every other part,
 * and inside every package within. The proof Apply is checked against: for
 * a redacted file it answers nothing.
 */
export function partsHolding(bytes, strings, depth = 0) {
  const pkg = Buffer.isBuffer(bytes) || bytes instanceof Uint8Array ? OoxmlPackage.read(Buffer.from(bytes)) : bytes;
  const wanted = strings.map((s) => String(s).toLowerCase()).filter(Boolean);
  const hits = [];
  for (const name of pkg.partNames()) {
    const raw = pkg.read(name);
    if (isPackage(name) && depth < 3) {
      try {
        for (const inner of partsHolding(raw, strings, depth + 1)) hits.push(name + '!' + inner);
        continue;
      } catch { /* bytes, then */ }
    }
    const texts = isXml(name, pkg)
      ? [unesc(raw.toString('utf8')).replace(/<[^>]*>/g, ''), unesc(raw.toString('utf8'))]
      : [raw.toString('latin1'), raw.toString('utf16le'), raw.toString('utf8')];
    if (texts.some((t) => wanted.some((w) => t.toLowerCase().includes(w)))) hits.push(name);
  }
  return hits;
}
