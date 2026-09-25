/**
 * References → Index: Mark Entry's XE fields read back, and the INDEX field
 * Insert Index writes — its result every marked entry, sorted, under its
 * letter, with the pages the window laid the entries on. Document methods,
 * put on the class by references.js with the rest of the References tab.
 */


/** Text content escaped: a field code keeps its quotes as Word writes them. */
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
import { renderRun, topComplexFields } from './runs.js';
import { parseXeInstr, xeInstr, indexInstr, parseIndexInstr, buildIndex, lineTail } from './wordindex.js';

/** An XE field as Word writes one: hidden text, begin, code, end — an XE has no result. */
export function xeFieldXml(instr) {
  const rPr = '<w:rPr><w:vanish/></w:rPr>';
  return '<w:r>' + rPr + '<w:fldChar w:fldCharType="begin"/></w:r>'
    + '<w:r>' + rPr + '<w:instrText xml:space="preserve">' + esc(instr) + '</w:instrText></w:r>'
    + '<w:r>' + rPr + '<w:fldChar w:fldCharType="end"/></w:r>';
}

const INDEX_STYLES = [
  ['IndexHeading', 'index heading', '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="60"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:b/><w:bCs/></w:rPr>', 'Index1'],
  ['Index1', 'index 1', '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="220" w:hanging="220"/></w:pPr>'],
  ['Index2', 'index 2', '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="440" w:hanging="220"/></w:pPr>'],
  ['Index3', 'index 3', '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="660" w:hanging="220"/></w:pPr>'],
];

/** Segments as runs, bold and italic as each asks. */
const runsOf = (segs) => segs.map((s) => {
  const rPr = s.bold || s.italic ? '<w:rPr>' + (s.bold ? '<w:b/><w:bCs/>' : '') + (s.italic ? '<w:i/><w:iCs/>' : '') + '</w:rPr>' : null;
  return renderRun(rPr, s.text);
}).join('');

export const indexMethods = {
  /** Mark Entry's field: its code and its XML, for the editor to put in a paragraph as one run. */
  xeFor(spec) {
    const instr = xeInstr(spec);
    return { instr, xml: xeFieldXml(instr) };
  },

  /**
   * Every XE field in the edit address space, in document order: its
   * paragraph (`block`) and what it marks. A table cell's paragraph counts,
   * as Word indexes one.
   */
  indexEntries() {
    const out = [];
    this.editParagraphs().forEach((p, block) => {
      if (!p.xml.includes('XE')) return;
      for (const f of topComplexFields(p.xml)) {
        const e = parseXeInstr(f.instr);
        if (e) out.push({ block, instr: f.instr, ...e });
      }
    });
    return out;
  },

  /** A bookmark's first and last paragraph in the edit address space — a page range's two ends. */
  _bookmarkEditSpan(name) {
    const { body } = this._body();
    let id = null;
    let start = -1;
    for (const m of body.matchAll(/<w:bookmarkStart\b([^>]*)\/>/g)) {
      if (new RegExp('\\bw:name="' + String(name).replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"').test(m[1])) {
        id = /\bw:id="([^"]*)"/.exec(m[1])?.[1];
        start = m.index;
        break;
      }
    }
    if (start < 0) return null;
    const endM = id !== null ? new RegExp('<w:bookmarkEnd\\b[^>]*\\bw:id="' + id + '"').exec(body) : null;
    const end = endM ? endM.index : start;
    const paras = this.editParagraphs();
    const at = (offset) => {
      let hit = 0;
      for (let i = 0; i < paras.length; i++) {
        if (paras[i].start <= offset) hit = i;
        else break;
      }
      return hit;
    };
    return { from: at(start), to: at(end) };
  },

  /** The "Index 1–3" and "Index Heading" styles, once — Word's own defaults. */
  _ensureIndexStyles() {
    this.ensureParagraphStyles();
    const part = 'word/styles.xml';
    let xml = this.pkg.text(part);
    let changed = false;
    for (const [id, name, body, next] of INDEX_STYLES) {
      if (new RegExp('<w:style\\b[^>]*\\bw:styleId="' + id + '"').test(xml)) continue;
      xml = xml.replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="' + id + '"><w:name w:val="' + name + '"/><w:basedOn w:val="Normal"/>'
        + (next ? '<w:next w:val="' + next + '"/>' : '') + '<w:uiPriority w:val="99"/><w:unhideWhenUsed/>' + body + '</w:style></w:styles>');
      changed = true;
    }
    if (changed) this.pkg.write_(part, xml);
    return changed;
  },

  /**
   * Where the INDEX field sits: the paragraph holding its begin and the one
   * holding its end (`start`/`end` in the body), and its code. Null when
   * there is none.
   */
  _indexSpan() {
    const { body } = this._body();
    const re = /<w:fldChar\b[^>]*\bw:fldCharType="(begin|end)"[^>]*\/>/g;
    let m;
    while ((m = re.exec(body))) {
      if (m[1] !== 'begin') continue;
      const from = m.index + m[0].length;
      const next = body.indexOf('<w:fldChar', from);
      const after = body.slice(from, next < 0 ? from + 600 : next);
      const instr = [...after.matchAll(/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g)].map((x) => x[1]).join('');
      if (!/^\s*INDEX\b/.test(instr.replace(/&quot;/g, '"'))) continue;
      let depth = 0;
      re.lastIndex = m.index;
      let end = -1;
      let mm;
      while ((mm = re.exec(body))) {
        depth += mm[1] === 'begin' ? 1 : -1;
        if (depth === 0) { end = mm.index; break; }
      }
      if (end < 0) return null;
      const start = Math.max(body.lastIndexOf('<w:p>', m.index), body.lastIndexOf('<w:p ', m.index));
      const close = body.indexOf('</w:p>', end);
      const code = instr.replace(/&quot;/g, '"').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
      return { start, end: close + '</w:p>'.length, instr: code };
    }
    return null;
  },

  hasIndex() { return this._indexSpan() !== null; },

  /** The index as the body holds it now: its options and each line's words — for the window and the tests. */
  indexResult() {
    const span = this._indexSpan();
    if (!span) return null;
    const { body } = this._body();
    const xml = body.slice(span.start, span.end);
    const lines = [...xml.matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g)].map((m) => {
      const style = /<w:pStyle w:val="([^"]+)"/.exec(m[1])?.[1] || '';
      const text = [...m[1].matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\/>/g)].map((t) => (t[0] === '<w:tab/>' ? '\t' : t[1])).join('')
        .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"');
      return { style, text };
    }).filter((l) => l.text);
    const leader = /<w:tab w:val="right" w:leader="([a-z]+)"/.exec(xml)?.[1] || null;
    return { ...parseIndexInstr(span.instr), leader, lines };
  },

  /**
   * The INDEX field, whole: its begin (and code) in a paragraph of its own,
   * the letters and entries, and its end in a last paragraph — the shape
   * the table of contents takes here, one paragraph wider at each end than
   * Word's own, so every paragraph's own runs stay whole. `pages` maps an
   * edit-space paragraph to the page the window laid it on.
   */
  _indexFieldXml({ columns = 2, rightAlign = false, leader = 'dot', runIn = false, lcid = 1033 } = {}, pages = null, cached = null) {
    const instr = indexInstr({ columns, rightAlign, runIn, headings: true, lcid });
    const pageOf = (block) => {
      if (!pages) return null;
      const v = pages[block] ?? pages[String(block)];
      return v === undefined || v === null || v === '' ? null : Number(v);
    };
    const entries = this.indexEntries().map((e) => {
      const range = e.bookmark ? this._bookmarkEditSpan(e.bookmark) : null;
      return { ...e, page: pageOf(range ? range.from : e.block), pageEnd: range ? pageOf(range.to) : null };
    });
    const groups = buildIndex(entries, { runIn });
    // Where a right-aligned page number stands: the column's right edge.
    const sect = this.section?.() || null;
    const content = sect?.contentWidthPx ? Math.round(sect.contentWidthPx * 15) : 9026;
    const cols = Math.max(1, Math.min(4, Number(columns) || 1));
    const colWidth = Math.max(1440, Math.round((content - (cols - 1) * 720) / cols));
    const tabs = rightAlign && !runIn
      ? '<w:tabs><w:tab w:val="right" w:leader="' + (leader && leader !== 'none' ? leader : 'none') + '" w:pos="' + colWidth + '"/></w:tabs>'
      : '';
    const pPr = (style) => '<w:pPr><w:pStyle w:val="' + style + '"/>' + (style === 'IndexHeading' ? '' : tabs) + '<w:rPr><w:noProof/></w:rPr></w:pPr>';
    const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">' + esc(instr) + '</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
    const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const tail = (line) => {
      if (!pages && cached) {
        const was = cached.get(line.level + ':' + line.text);
        if (was !== undefined) return [{ text: was }];
      }
      return lineTail(line, { rightAlign: rightAlign && !runIn });
    };
    // As Word writes it: the field begins in the first paragraph, ahead of
    // its words, and ends in the last, after them — an index holds no field
    // of its own that a paragraph could pair with the wrong end.
    const paras = [];
    for (const g of groups) {
      paras.push([pPr('IndexHeading'), renderRun(null, g.letter)]);
      for (const line of g.lines) {
        let segs = [{ text: line.text }, ...tail(line)];
        if (runIn && line.runIn?.length) {
          segs.push({ text: ': ' });
          line.runIn.forEach((sub, i) => {
            if (i) segs.push({ text: '; ' });
            segs = segs.concat([{ text: sub.text }, ...tail({ ...sub, level: 2 })]);
          });
        }
        paras.push([pPr('Index' + Math.min(3, line.level)), runsOf(segs)]);
      }
    }
    if (!paras.length) paras.push([pPr('Index1'), renderRun('<w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr>', 'No index entries found.')]);
    return paras.map(([ppr, runs], i) => '<w:p>' + ppr + (i === 0 ? begin : '') + runs + (i === paras.length - 1 ? end : '') + '</w:p>').join('');
  },

  /** What each line said after its entry last time — kept by an Update Index that has no pages to hand. */
  _indexCachedTails() {
    const now = this.indexResult();
    if (!now) return null;
    const map = new Map();
    for (const l of now.lines) {
      const level = Number(/Index(\d)/.exec(l.style)?.[1]) || 0;
      if (!level) continue;
      const m = /^(.*?)((?:\t|, )[\s\S]*)$/.exec(l.text);
      if (m) map.set(level + ':' + m[1], m[2]);
    }
    return map;
  },

  /**
   * References → Insert Index, after paragraph `at` (the edit address): the
   * INDEX field built from every XE in the document. An index already in
   * the document is replaced where it stands, as Word asks to.
   */
  insertIndex({ at, columns = 2, rightAlign = false, leader = 'dot', runIn = false, lcid = 1033, pages = null } = {}) {
    this._ensureIndexStyles();
    const xml = this._indexFieldXml({ columns, rightAlign, leader, runIn, lcid }, pages);
    const span = this._indexSpan();
    if (span) {
      this._spliceBody(span.start, span.end, xml);
    } else {
      const p = this.editParagraph(at);
      if (!p) throw new Error('no paragraph at index ' + at);
      const top = this.paragraphs().find((q) => q.start <= p.start && q.end >= p.end) || p;
      this._spliceBody(top.end, top.end, xml);
    }
    this.dirty = true;
    return this;
  },

  /** References → Update Index: the entries and pages as they stand, the options kept. */
  updateIndex({ pages = null } = {}) {
    const now = this.indexResult();
    const span = this._indexSpan();
    if (!now || !span) return false;
    const cached = pages ? null : this._indexCachedTails();
    this._ensureIndexStyles();
    const xml = this._indexFieldXml({ columns: now.columns, rightAlign: now.rightAlign, leader: now.leader || 'dot', runIn: now.runIn, lcid: now.lcid }, pages, cached);
    this._spliceBody(span.start, span.end, xml);
    this.dirty = true;
    return true;
  },
};

export { xeInstr };
