/**
 * References — citations and a bibliography, written as Word writes them.
 *
 * These are Document methods kept in a file of their own (installed onto
 * the class by `installReferences`, called once from document.js) so the
 * References tab's engine can grow without every change landing in the
 * four-thousand-line reader.
 *
 * Citations: the sources are the document's custom XML part (see
 * bibliography.js); a citation is a CITATION field inside a run-level
 * content control carrying `w:citation`, its result the in-text citation in
 * the style the part names. The bibliography is a BIBLIOGRAPHY field inside
 * a content control carrying `w:bibliography`, itself inside the building
 * block Word's Bibliography gallery inserts (`docPartGallery`
 * "Bibliographies") with a heading — Bibliography, References or Works
 * Cited. Update Citations and Bibliography rewrites each citation's result
 * and the bibliography's entries from the sources as they now stand.
 */


/** Text content escaped: a field code keeps its quotes as Word writes them. */
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
import { textOf, renderRun, topComplexFields } from './runs.js';
import {
  parseSources, sourcesXml, formatCitation, formatBibliography, parseCitationInstr, citationInstr,
  citationNumbers, segmentsText, DEFAULT_STYLE, styleById,
} from './bibliography.js';
import { indexMethods } from './references-index.js';

const CUSTOMXML_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';
const CUSTOMXML_PROPS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps';
const CUSTOMXML_PROPS_CT = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml';
const BIB_SCHEMA = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography';

/** A run-level content control round a citation (no content control inside it). */
const CITATION_SDT_RE = () => /<w:sdt\b[^>]*>(?:(?!<w:sdt\b)[\s\S])*?<w:citation\s*\/>(?:(?!<w:sdt\b)[\s\S])*?<\/w:sdt>/g;

/** A content control's `w:id`: a signed 32-bit number, as Word mints them. */
const sdtId = () => String((Math.floor(Math.random() * 0xffffffff) | 0) || 1);

/** Segments as runs: Word writes a citation's and a bibliography's result `noProof`, a title italic. */
export function segmentRuns(segments, { bold = false } = {}) {
  return (segments || []).map((s) => {
    const rPr = '<w:rPr>' + (bold ? '<w:b/><w:bCs/>' : '') + (s.italic ? '<w:i/><w:iCs/>' : '') + '<w:noProof/></w:rPr>';
    return renderRun(rPr, s.text);
  }).join('');
}

/** A complex field's begin, code and separate — its result follows, then `FIELD_END`. */
export const fieldBegin = (instr, rPr = '') =>
  '<w:r>' + rPr + '<w:fldChar w:fldCharType="begin"/></w:r>'
  + '<w:r>' + rPr + '<w:instrText xml:space="preserve">' + esc(instr) + '</w:instrText></w:r>'
  + '<w:r>' + rPr + '<w:fldChar w:fldCharType="separate"/></w:r>';
export const FIELD_END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

/** A citation as Word writes one: the CITATION field in a content control marked `w:citation`. */
export function citationControlXml(instr, segments) {
  return '<w:sdt><w:sdtPr><w:id w:val="' + sdtId() + '"/><w:citation/></w:sdtPr><w:sdtContent>'
    + fieldBegin(instr) + segmentRuns(segments) + FIELD_END
    + '</w:sdtContent></w:sdt>';
}

/** Every `<w:sdt>` in a fragment, outermost first, nesting honoured: `{ start, end, prEnd }`. */
function sdtSpans(xml) {
  const out = [];
  const stack = [];
  const re = /<w:sdt\b[^>]*?(\/?)>|<\/w:sdt>/g;
  let m;
  while ((m = re.exec(xml))) {
    if (m[0].startsWith('</')) {
      const open = stack.pop();
      if (open) out.push({ start: open.start, end: m.index + m[0].length, depth: stack.length });
    } else if (m[1] !== '/') stack.push({ start: m.index });
  }
  return out.sort((a, b) => a.start - b.start);
}

/** The `w:sdtPr` of the control starting at `start`. */
const sdtPrOf = (xml, span) => {
  const m = /^<w:sdt\b[^>]*>\s*(<w:sdtPr\b[^>]*>[\s\S]*?<\/w:sdtPr>|<w:sdtPr\b[^>]*\/>)?/.exec(xml.slice(span.start, span.end));
  return m?.[1] || '';
};

const BIB_GALLERY = /<w:docPartGallery\b[^>]*\bw:val="Bibliographies"/;

const methods = {
  // ---- the sources part ----------------------------------------------------

  /** The custom XML part holding `<b:Sources>`, or null — found by what it holds, as Word finds it. */
  _bibliographyPart() {
    for (const name of this.pkg.partNames()) {
      if (!/^customXml\/item\d+\.xml$/i.test(name)) continue;
      const head = this.pkg.text(name).slice(0, 600);
      if (/<(?:\w+:)?Sources\b/.test(head) && head.includes(BIB_SCHEMA)) return name;
    }
    return null;
  },

  /** The document's sources — Manage Sources' Current List — and the style they are cited in. */
  bibliographySources() {
    const part = this._bibliographyPart();
    if (!part) return { style: DEFAULT_STYLE, sources: [] };
    return parseSources(this.pkg.text(part));
  },

  /**
   * Make the sources part real — `customXml/itemN.xml` with its
   * `itemPropsN.xml` (the datastore item naming the bibliography schema) and
   * the relationships Word writes from the main part and from the item —
   * and put it on the undo list, BEFORE the edit's snapshot is taken, as a
   * comment's part is (see `registerCommentUndo`).
   */
  registerBibliographyUndo() {
    let part = this._bibliographyPart();
    if (!part) {
      const n = this.pkg.nextPartNumber('customXml/', 'item');
      part = 'customXml/item' + n + '.xml';
      this.pkg.ensureDefault('xml', 'application/xml');
      this.pkg.addPart(part, sourcesXml({ style: DEFAULT_STYLE, sources: [] }), null);
      const guid = '{' + (globalThis.crypto?.randomUUID?.() || '00000000-0000-4000-8000-' + String(Date.now()).padStart(12, '0').slice(-12)).toUpperCase() + '}';
      this.pkg.addPart('customXml/itemProps' + n + '.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="no"?>'
        + '<ds:datastoreItem ds:itemID="' + guid + '" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">'
        + '<ds:schemaRefs><ds:schemaRef ds:uri="' + BIB_SCHEMA + '"/></ds:schemaRefs></ds:datastoreItem>',
        CUSTOMXML_PROPS_CT);
      this.pkg.addRelationshipTo(part, CUSTOMXML_PROPS_REL, 'itemProps' + n + '.xml');
      this.pkg.addRelationshipTo(this.mainPart, CUSTOMXML_REL, '../customXml/item' + n + '.xml');
      this.dirty = true;
    }
    this._undoParts.add(part);
    return part;
  },

  /** Write the sources and the style — the whole part, the way Word rewrites it. */
  setBibliographySources({ style, sources } = {}) {
    const part = this.registerBibliographyUndo();
    const now = this.bibliographySources();
    this.pkg.write_(part, sourcesXml({ style: style || now.style, sources: sources || now.sources }));
    this.dirty = true;
    return this;
  },

  // ---- citations -----------------------------------------------------------

  /** Every citation in the body, in document order: `{ instr, tags, pages, … }`. */
  citations() {
    const { body } = this._body();
    const out = [];
    for (const m of body.matchAll(CITATION_SDT_RE())) {
      const field = topComplexFields(m[0])[0];
      const parsed = field ? parseCitationInstr(field.instr) : null;
      if (parsed) out.push({ ...parsed, instr: field.instr, text: textOf(m[0].slice(m[0].indexOf('<w:sdtContent'))) });
    }
    return out;
  },

  /** The words a citation shows now, from the sources and the style: segments. */
  formatCitationFor(spec, { style, sources } = this.bibliographySources(), numbers = null) {
    const byTag = new Map(sources.map((s) => [s.tag, s]));
    const entries = (spec.tags || []).map((tag) => byTag.get(tag) || { tag, placeholder: true });
    const nums = numbers || (style === 'ieee' ? citationNumbers(sources, [...this.citations().flatMap((c) => c.tags), ...(spec.tags || [])]) : null);
    return formatCitation(entries, style, { ...spec, numbers: nums });
  },

  /** A new citation's XML, ready to go into a paragraph: `{ instr, text, xml }`. */
  citationFor({ tags, pages = '', lcid = 1033, suppressAuthor, suppressYear, suppressTitle } = {}) {
    if (!tags?.length) throw new Error('a citation names at least one source');
    const instr = citationInstr({ tags, lcid, pages, suppressAuthor, suppressYear, suppressTitle });
    const segments = this.formatCitationFor({ tags, pages, suppressAuthor, suppressYear, suppressTitle });
    return { instr, text: segmentsText(segments), xml: citationControlXml(instr, segments) };
  },

  /**
   * Update Citations and Bibliography (and F9): every citation's result
   * written again from the sources and the style, the RefOrder of each cited
   * source set in the order it is first cited (a numeric style numbers by
   * it), and the bibliography's entries rebuilt. Returns how many citations
   * changed.
   */
  updateCitations() {
    const bib = this.bibliographySources();
    const cited = this.citations().flatMap((c) => c.tags);
    const numbers = citationNumbers(bib.sources, cited);
    let changed = 0;
    const { prefix, body, suffix } = this._body();
    const next = body.replace(CITATION_SDT_RE(), (whole) => {
      const content = /<w:sdtContent\b[^>]*>/.exec(whole);
      if (!content) return whole;
      const base = content.index + content[0].length;
      const inner = whole.slice(base, whole.lastIndexOf('</w:sdtContent>'));
      const field = topComplexFields(inner)[0];
      const spec = field ? parseCitationInstr(field.instr) : null;
      if (!spec) return whole;
      const was = inner.slice(field.resultStart, field.resultEnd);
      // A result Word wrote with a space in front of it keeps the space.
      const lead = /^\s/.test(textOf(was)) ? [{ text: ' ' }] : [];
      const segments = [...lead, ...this.formatCitationFor(spec, bib, bib.style === 'ieee' ? numbers : null)];
      if (textOf(was) === segmentsText(segments)) return whole;
      changed += 1;
      return whole.slice(0, base + field.resultStart) + segmentRuns(segments) + whole.slice(base + field.resultEnd);
    });
    if (next !== body) {
      this.xml = prefix + next + suffix;
      this.dirty = true;
    }
    // RefOrder: the order a source is first cited, as Word keeps it.
    if (this._bibliographyPart()) {
      const order = new Map();
      for (const tag of cited) if (!order.has(tag)) order.set(tag, order.size + 1);
      const sources = bib.sources.map((s) => ({ ...s, refOrder: order.get(s.tag) || null }));
      if (sources.some((s, i) => s.refOrder !== bib.sources[i].refOrder)) this.setBibliographySources({ style: bib.style, sources });
    }
    if (this._bibliographySpan()) this.updateBibliography();
    return changed;
  },

  /** References → Style: the citations and the bibliography in another style, at once. */
  setBibliographyStyle(style) {
    const bib = this.bibliographySources();
    this.setBibliographySources({ style: styleById(style).id, sources: bib.sources });
    this.updateCitations();
    return this;
  },

  // ---- the bibliography ----------------------------------------------------

  /** The "Bibliography" paragraph style, once — Word's own, based on Normal. */
  _ensureBibliographyStyle() {
    this.ensureParagraphStyles();
    const part = 'word/styles.xml';
    const xml = this.pkg.text(part);
    if (/<w:style\b[^>]*\bw:styleId="Bibliography"/.test(xml)) return false;
    const style = '<w:style w:type="paragraph" w:styleId="Bibliography"><w:name w:val="Bibliography"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="37"/><w:unhideWhenUsed/></w:style>';
    this.pkg.write_(part, xml.replace('</w:styles>', style + '</w:styles>'));
    return true;
  },

  /**
   * The bibliography field's paragraphs: the field begins in the first
   * entry, as Word writes it, and ends in a paragraph of its own. Every
   * entry hangs half an inch.
   */
  _bibliographyFieldXml() {
    const bib = this.bibliographySources();
    const cited = this.citations().flatMap((c) => c.tags);
    const entries = formatBibliography(bib.sources, bib.style, { citedTags: cited });
    const pPr = (bib.style === 'ieee')
      ? '<w:pPr><w:pStyle w:val="Bibliography"/><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="720"/><w:rPr><w:noProof/></w:rPr></w:pPr>'
      : '<w:pPr><w:pStyle w:val="Bibliography"/><w:ind w:left="720" w:hanging="720"/><w:rPr><w:noProof/></w:rPr></w:pPr>';
    const begin = fieldBegin(' BIBLIOGRAPHY ');
    if (!entries.length) {
      return '<w:p>' + pPr + begin + segmentRuns([{ text: 'There are no sources in the current document.' }], { bold: true }) + '</w:p>'
        + '<w:p><w:pPr><w:pStyle w:val="Bibliography"/></w:pPr>' + FIELD_END + '</w:p>';
    }
    return entries.map((e, i) => '<w:p>' + pPr + (i === 0 ? begin : '') + segmentRuns(e.segments) + '</w:p>').join('')
      + '<w:p><w:pPr><w:pStyle w:val="Bibliography"/></w:pPr>' + FIELD_END + '</w:p>';
  },

  /** The control holding the BIBLIOGRAPHY field (`w:bibliography`), and the building block round it if there is one. */
  _bibliographySpan() {
    const { body } = this._body();
    const spans = sdtSpans(body);
    const inner = spans.find((s) => /<w:bibliography\s*\/>/.test(sdtPrOf(body, s)));
    if (!inner) return null;
    const outer = spans.find((s) => s.start < inner.start && s.end > inner.end && BIB_GALLERY.test(sdtPrOf(body, s)));
    return { inner, outer: outer || null };
  },

  /** True when the body carries a bibliography. */
  hasBibliography() { return this._bibliographySpan() !== null; },

  /**
   * References → Bibliography, after paragraph `at` (the edit address): the
   * gallery's building block with `heading` — 'Bibliography', 'References'
   * or 'Works Cited' as Heading 1 — or, with no heading, Insert
   * Bibliography's bare field.
   */
  insertBibliography({ at, heading = 'Bibliography' } = {}) {
    const p = this.editParagraph(at);
    if (!p) throw new Error('no paragraph at index ' + at);
    if (p.container != null) throw new Error('a bibliography goes in the body, not in a table');
    this._ensureBibliographyStyle();
    const field = '<w:sdt><w:sdtPr><w:id w:val="' + sdtId() + '"/><w:bibliography/></w:sdtPr><w:sdtContent>' + this._bibliographyFieldXml() + '</w:sdtContent></w:sdt>';
    const xml = heading
      ? '<w:sdt><w:sdtPr><w:id w:val="' + sdtId() + '"/><w:docPartObj><w:docPartGallery w:val="Bibliographies"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>'
        + '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' + renderRun(null, heading) + '</w:p>'
        + field + '</w:sdtContent></w:sdt>'
      : field;
    this._spliceBody(p.end, p.end, xml);
    this.dirty = true;
    return this;
  },

  /** The bibliography's entries rebuilt from the sources as they stand, its heading kept. */
  updateBibliography() {
    const span = this._bibliographySpan();
    if (!span) return false;
    const { body } = this._body();
    const control = body.slice(span.inner.start, span.inner.end);
    const open = /<w:sdtContent\b[^>]*>/.exec(control);
    const close = control.lastIndexOf('</w:sdtContent>');
    if (!open || close < 0) return false;
    this._ensureBibliographyStyle();
    const next = control.slice(0, open.index + open[0].length) + this._bibliographyFieldXml() + control.slice(close);
    if (next === control) return false;
    this._spliceBody(span.inner.start, span.inner.end, next);
    return true;
  },

  /** The bibliography's entries as the body holds them now — for the window and the tests. */
  bibliographyEntries() {
    const span = this._bibliographySpan();
    if (!span) return null;
    const { body } = this._body();
    const xml = body.slice(span.inner.start, span.inner.end);
    const heading = span.outer ? textOf(/<w:p\b[\s\S]*?<\/w:p>/.exec(body.slice(span.outer.start, span.inner.start))?.[0] || '') : null;
    const entries = [...xml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)].map((m) => textOf(m[0])).filter((t) => t.trim());
    return { heading, entries };
  },

  /** What the References tab reads: the sources, the style, the citations and whether a bibliography is in. */
  referencesInfo() {
    const bib = this.bibliographySources();
    return {
      style: bib.style,
      sources: bib.sources,
      citations: this.citations().map((c) => ({ tags: c.tags, pages: c.pages, text: c.text })),
      bibliography: this.hasBibliography(),
      // The index, when there is one: the options Insert Index opens with.
      index: (() => {
        const r = this.indexResult();
        return r ? { columns: r.columns, rightAlign: r.rightAlign, runIn: r.runIn, leader: r.leader } : null;
      })(),
    };
  },
};

/** Put the References methods on the Document class. */
export function installReferences(Document) {
  for (const [name, fn] of Object.entries({ ...methods, ...indexMethods })) {
    if (!Object.prototype.hasOwnProperty.call(Document.prototype, name)) Document.prototype[name] = fn;
  }
  // F9 refreshes citations and the bibliography with every other field.
  const refresh = Document.prototype.refreshRefFields;
  if (refresh && !refresh.withReferences) {
    const wrapped = function refreshRefFields(...args) {
      let changed = refresh.apply(this, args);
      if (this.citations().length || this.hasBibliography()) changed += this.updateCitations();
      return changed;
    };
    wrapped.withReferences = true;
    Document.prototype.refreshRefFields = wrapped;
  }
}
