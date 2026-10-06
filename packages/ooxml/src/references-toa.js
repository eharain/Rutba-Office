/**
 * References → Table of Authorities: Mark Citation's TA fields read back,
 * and the TOA fields Insert Table of Authorities writes — one per category,
 * its heading, then each authority cited in it, sorted, with the pages the
 * window laid its citations on. Document methods, put on the class by
 * references.js with the rest of the References tab.
 */


/** Text content escaped: a field code keeps its quotes as Word writes them. */
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
import { renderRun, topComplexFields } from './runs.js';
import { taInstr, parseTaInstr, toaInstr, parseToaInstr, buildAuthorities, authorityTail, categoryName } from './wordtoa.js';
import { xeFieldXml } from './references-index.js';

const TOA_STYLES = [
  ['TOAHeading', 'toa heading', '<w:pPr><w:keepNext/><w:spacing w:before="120"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'],
  ['TableofAuthorities', 'table of authorities', '<w:pPr><w:ind w:left="220" w:hanging="220"/></w:pPr>'],
];

/** Segments as runs, bold and italic as each asks. */
const runsOf = (segs) => segs.map((s) => {
  const rPr = '<w:rPr>' + (s.bold ? '<w:b/><w:bCs/>' : '') + (s.italic ? '<w:i/><w:iCs/>' : '') + '<w:noProof/></w:rPr>';
  return s.text === '\t' ? '<w:r>' + rPr + '<w:tab/></w:r>' : renderRun(rPr, s.text);
}).join('');

const unesc = (s) => s.replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');

export const toaMethods = {
  /** Mark Citation's field: its code and its XML — hidden, as Word writes a TA, with no result. */
  taFor(spec) {
    const instr = taInstr(spec);
    return { instr, xml: xeFieldXml(instr) };
  },

  /**
   * Every TA field in the edit address space, in document order: its
   * paragraph (`block`) and what it cites. A table cell's paragraph counts.
   */
  authorityEntries() {
    const out = [];
    this.editParagraphs().forEach((p, block) => {
      if (!p.xml.includes('TA')) return;
      for (const f of topComplexFields(p.xml)) {
        const e = parseTaInstr(unesc(f.instr));
        if (e) out.push({ block, instr: f.instr, ...e });
      }
    });
    return out;
  },

  /** The authorities marked so far, each once: what Mark Citation lists and Insert Table of Authorities previews. */
  authorities() {
    return buildAuthorities(this.authorityEntries(), { usePassim: false }).map(({ long, short, category }) => ({ long, short, category }));
  },

  /** The "TOA Heading" and "Table of Authorities" styles, once — Word's own. */
  _ensureToaStyles() {
    this.ensureParagraphStyles();
    const part = 'word/styles.xml';
    let xml = this.pkg.text(part);
    let changed = false;
    for (const [id, name, body] of TOA_STYLES) {
      if (new RegExp('<w:style\\b[^>]*\\bw:styleId="' + id + '"').test(xml)) continue;
      xml = xml.replace('</w:styles>', () => '<w:style w:type="paragraph" w:styleId="' + id + '"><w:name w:val="' + name + '"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>' + body + '</w:style></w:styles>');
      changed = true;
    }
    if (changed) this.pkg.write_(part, xml);
    return changed;
  },

  /**
   * Every TOA field in the body, in order: the paragraph holding its begin
   * and the one holding its end (`start`/`end` in the body), its code and
   * the choices it was made with.
   */
  _toaSpans() {
    const { body } = this._body();
    const out = [];
    const re = /<w:fldChar\b[^>]*\bw:fldCharType="(begin|end)"[^>]*\/>/g;
    let m;
    while ((m = re.exec(body))) {
      if (m[1] !== 'begin') continue;
      const from = m.index + m[0].length;
      const next = body.indexOf('<w:fldChar', from);
      const instr = unesc([...body.slice(from, next < 0 ? from + 600 : next).matchAll(/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g)].map((x) => x[1]).join(''));
      const opts = parseToaInstr(instr);
      if (!opts) continue;
      let depth = 0;
      const scan = new RegExp(re.source, 'g');
      scan.lastIndex = m.index;
      let end = -1;
      let mm;
      while ((mm = scan.exec(body))) {
        depth += mm[1] === 'begin' ? 1 : -1;
        if (depth === 0) { end = mm.index; break; }
      }
      if (end < 0) break;
      const start = Math.max(body.lastIndexOf('<w:p>', m.index), body.lastIndexOf('<w:p ', m.index));
      const close = body.indexOf('</w:p>', end) + '</w:p>'.length;
      const leader = /<w:tab w:val="right" w:leader="([a-z]+)"/.exec(body.slice(start, close))?.[1] || 'none';
      out.push({ start, end: close, instr, ...opts, leader });
      re.lastIndex = close;
    }
    return out;
  },

  /** The tables of authorities as the body holds them now: each one's choices and its lines' words — for the window and the tests. */
  tablesOfAuthorities() {
    const { body } = this._body();
    return this._toaSpans().map((t) => {
      const xml = body.slice(t.start, t.end);
      const lines = [...xml.matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g)].map((m) => {
        const style = /<w:pStyle w:val="([^"]+)"/.exec(m[1])?.[1] || '';
        const text = unesc([...m[1].matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\/>/g)].map((x) => (x[0] === '<w:tab/>' ? '\t' : x[1])).join(''));
        return { style, text };
      }).filter((l) => l.text);
      return { category: t.category, heading: t.heading, passim: t.passim, keepFormatting: t.keepFormatting, leader: t.leader, lines };
    });
  },

  /**
   * One TOA field's paragraphs: its heading (when `heading`) and each
   * authority in its category, the field beginning in the first and ending
   * in the last — the index's shape. `pages` maps an edit-space paragraph
   * to the page the window laid it on; `cached` keeps what each line said
   * after its citation when there are no pages to hand.
   */
  _toaFieldXml({ category = 1, heading = true, passim = true, keepFormatting = true, leader = 'dot' } = {}, pages = null, cached = null) {
    const instr = toaInstr({ category, heading, passim, keepFormatting });
    const pageOf = (block) => {
      if (!pages) return null;
      const v = pages[block] ?? pages[String(block)];
      return v === undefined || v === null || v === '' ? null : Number(v);
    };
    const entries = this.authorityEntries().map((e) => {
      const range = e.bookmark ? this._bookmarkEditSpan(e.bookmark) : null;
      return { ...e, page: pageOf(range ? range.from : e.block), pageEnd: range ? pageOf(range.to) : null };
    });
    const list = buildAuthorities(entries, { category, usePassim: passim });
    const tabs = '<w:tabs><w:tab w:val="right" w:leader="' + (leader && leader !== 'none' ? leader : 'none') + '" w:pos="' + this._tocTabPositionTwips() + '"/></w:tabs>';
    const pPr = (style) => '<w:pPr><w:pStyle w:val="' + style + '"/>' + (style === 'TOAHeading' ? '' : tabs) + '<w:rPr><w:noProof/></w:rPr></w:pPr>';
    const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">' + esc(instr) + '</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
    const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const paras = [];
    if (heading) paras.push([pPr('TOAHeading'), renderRun('<w:rPr><w:noProof/></w:rPr>', categoryName(category))]);
    for (const a of list) {
      let tail = authorityTail(a);
      if (!pages && cached?.has(a.long)) tail = [{ text: '\t' }, { text: cached.get(a.long) }];
      paras.push([pPr('TableofAuthorities'), runsOf([{ text: a.long }, ...tail.filter((s) => s.text !== '')])]);
    }
    if (!list.length) paras.push([pPr('TableofAuthorities'), renderRun('<w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr>', 'No table of authorities entries found.')]);
    return paras.map(([ppr, runs], i) => '<w:p>' + ppr + (i === 0 ? begin : '') + runs + (i === paras.length - 1 ? end : '') + '</w:p>').join('');
  },

  /** What each authority's line said after the tab last time, by its long citation — kept by an update that has no pages. */
  _toaCachedTails() {
    const map = new Map();
    for (const t of this.tablesOfAuthorities()) {
      for (const l of t.lines) {
        if (l.style !== 'TableofAuthorities') continue;
        const at = l.text.lastIndexOf('\t');
        if (at > 0) map.set(l.text.slice(0, at), l.text.slice(at + 1));
      }
    }
    return map;
  },

  /**
   * References → Insert Table of Authorities, after paragraph `at` (the
   * edit address). `category` is a number, or 'all' for a TOA field per
   * category that has authorities in it, one after another, as Word inserts
   * them. A table of a category already in the document is replaced where
   * it stands — with All, every one is, the new tables in the first one's
   * place.
   */
  insertTableOfAuthorities({ at, category = 'all', passim = true, keepFormatting = true, leader = 'dot', pages = null } = {}) {
    this._ensureToaStyles();
    const opts = { heading: true, passim, keepFormatting, leader };
    let cats;
    if (category === 'all' || category === 0 || category === '0') {
      const used = new Set(this.authorities().map((a) => a.category));
      cats = [...used].sort((a, b) => a - b);
      if (!cats.length) cats = [1];
    } else cats = [Number(category) || 1];
    const xml = cats.map((c) => this._toaFieldXml({ ...opts, category: c }, pages)).join('');
    const spans = this._toaSpans();
    const replace = category === 'all' || category === 0 || category === '0' ? spans : spans.filter((s) => s.category === cats[0]);
    if (replace.length) {
      // Last first, so the earlier spans stay where they are; the new tables where the first stood.
      for (let k = replace.length - 1; k > 0; k--) this._spliceBody(replace[k].start, replace[k].end, '');
      this._spliceBody(replace[0].start, replace[0].end, xml);
    } else {
      const p = this.editParagraph(at);
      if (!p) throw new Error('no paragraph at index ' + at);
      const top = this.paragraphs().find((q) => q.start <= p.start && q.end >= p.end) || p;
      this._spliceBody(top.end, top.end, xml);
    }
    this.dirty = true;
    return this;
  },

  /** Update Table (Table of Authorities group, and F9): every table rebuilt from the citations and pages as they stand, its choices kept. */
  updateTablesOfAuthorities({ pages = null } = {}) {
    const spans = this._toaSpans();
    if (!spans.length) return 0;
    const cached = pages ? null : this._toaCachedTails();
    this._ensureToaStyles();
    const fresh = this._toaSpans();
    for (let k = fresh.length - 1; k >= 0; k--) {
      const t = fresh[k];
      this._spliceBody(t.start, t.end, this._toaFieldXml(t, pages, cached));
    }
    this.dirty = true;
    return spans.length;
  },

  hasTableOfAuthorities() { return this._toaSpans().length > 0; },
};
