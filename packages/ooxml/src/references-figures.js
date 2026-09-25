/**
 * References → Insert Table of Figures, and the document fields Insert →
 * Quick Parts → Field puts in the body. Document methods, put on the class
 * by references.js with the rest of the References tab.
 *
 * A table of figures is Word's TOC field with `\c "Figure"` (or `\a` when
 * the label and number are left off): one paragraph per caption carrying
 * that label's SEQ field, each a hyperlink to a `_Toc` bookmark on the
 * caption with a PAGEREF for its page — the table of contents' shape,
 * built from captions instead of headings. Update Table rebuilds it from
 * the captions as they stand.
 */


/** Text content escaped: a field code keeps its quotes as Word writes them. */
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
import { renderRun, textOf, mapComplexFieldResults, topComplexFields } from './runs.js';
import { evaluateDocField, parseDocFieldInstr, docFieldInstr } from './docfields.js';

/** A table of figures' bookmarks are numbered apart from a table of contents' (`_Toc1…`), so neither rebuild takes the other's. */
const TOF_BASE = 200000001;

/** A TOC field's code that builds from captions: `\c "Label"` or `\a "Label"`. */
const figuresOf = (instr) => /^\s*TOC\b[\s\S]*?\\([ca])\s+"?([^"\s\\]+)"?/.exec(String(instr));

export const figuresMethods = {
  /** The "Table of Figures" paragraph style, once — Word's own. */
  _ensureFiguresStyle() {
    this.ensureParagraphStyles();
    const part = 'word/styles.xml';
    const xml = this.pkg.text(part);
    if (/<w:style\b[^>]*\bw:styleId="TableofFigures"/.test(xml)) return false;
    this.pkg.write_(part, xml.replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="TableofFigures"><w:name w:val="table of figures"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0"/></w:pPr></w:style></w:styles>'));
    return true;
  },

  /**
   * Every caption with this label, in document order: a top-level
   * paragraph carrying a SEQ field for it — its `paragraphs()` index, its
   * edit-space index (for its page), its words, and its words without the
   * label and number.
   */
  captionsFor(label) {
    const edit = new Map(this.editParagraphs().map((p, i) => [p.start, i]));
    const out = [];
    for (const p of this.paragraphs()) {
      if (!p.xml.includes('SEQ')) continue;
      const seq = topComplexFields(p.xml).find((f) => new RegExp('^\\s*SEQ\\s+' + label + '\\b', 'i').test(f.instr));
      if (!seq) continue;
      const text = p.text.trim();
      // Without the label and number: what follows the SEQ field, less the separator.
      const after = textOf(p.xml.slice(seq.end)).replace(/^\s*[:.\-–—]?\s*/, '').trim();
      out.push({ index: p.index, block: edit.get(p.start) ?? null, text, bare: after });
    }
    return out;
  },

  /** Every table of figures in the body: `{ start, end, label, includeLabel, instr }`, in order. */
  _figureTables() {
    const { body } = this._body();
    const out = [];
    const re = /<w:fldChar\b[^>]*\bw:fldCharType="(begin|end)"[^>]*\/>/g;
    let m;
    while ((m = re.exec(body))) {
      if (m[1] !== 'begin') continue;
      const from = m.index + m[0].length;
      const next = body.indexOf('<w:fldChar', from);
      const instr = [...body.slice(from, next < 0 ? from + 400 : next).matchAll(/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g)]
        .map((x) => x[1].replace(/&quot;/g, '"').replace(/&amp;/g, '&')).join('');
      const fig = figuresOf(instr);
      if (!fig) continue;
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
      const xml = body.slice(start, close);
      out.push({
        start, end: close, label: fig[2], includeLabel: fig[1] === 'c', instr,
        // The choices Update Table keeps: page numbers (`\n` leaves them off),
        // right-aligned (`\p` puts a separator in place of the tab) and the leader.
        pageNumbers: !/\\n\b/.test(instr),
        rightAlign: !/\\p\b/.test(instr),
        leader: /<w:tab w:val="right" w:leader="([a-z]+)"/.exec(xml)?.[1] || 'none',
      });
      re.lastIndex = close;
    }
    return out;
  },

  /** The tables of figures read back — label, entries, cached pages — for the window and the tests. */
  tablesOfFigures() {
    const { body } = this._body();
    return this._figureTables().map((t) => {
      const xml = body.slice(t.start, t.end);
      const entries = [...xml.matchAll(/<w:hyperlink\b[^>]*\bw:anchor="([^"]+)"[^>]*>([\s\S]*?)<\/w:hyperlink>/g)].map((m) => {
        const tab = m[2].indexOf('<w:tab/>');
        const page = /<w:fldChar\b[^>]*"separate"[^>]*\/>[\s\S]*?<w:t\b[^>]*>([\s\S]*?)<\/w:t>/.exec(m[2]);
        return { anchor: m[1], text: textOf(tab < 0 ? m[2] : m[2].slice(0, tab)), page: page ? page[1] || null : null };
      });
      return { label: t.label, includeLabel: t.includeLabel, entries };
    });
  },

  /** The `_Toc` bookmarks a table of figures links to — its own, which its rebuild may take away. */
  _figureAnchors(xml) {
    return new Set([...String(xml).matchAll(/\bw:anchor="(_Toc\d+)"/g)].map((m) => m[1]));
  },

  /**
   * The field's paragraphs: begin in a paragraph of its own and end in a
   * last one, each caption an entry between — the table of contents'
   * shape (see `_tocFieldXml`). `pages` maps an edit-space paragraph to its
   * page; `cached` keeps the pages an entry had when none are given.
   */
  _figuresXml(label, includeLabel, pages, cached = null, { pageNumbers = true, rightAlign = true, leader = 'dot' } = {}) {
    const caps = this.captionsFor(label);
    const tab = this._tocTabPositionTwips();
    const right = pageNumbers && rightAlign;
    const instr = ' TOC \\h \\z ' + (pageNumbers ? '' : '\\n ') + (right || !pageNumbers ? '' : '\\p " " ') + '\\' + (includeLabel ? 'c' : 'a') + ' "' + label + '" ';
    const tabs = right ? '<w:tabs><w:tab w:val="right" w:leader="' + (leader || 'none') + '" w:pos="' + tab + '"/></w:tabs>' : '';
    const pPr = '<w:pPr><w:pStyle w:val="TableofFigures"/>' + tabs + '<w:rPr><w:noProof/></w:rPr></w:pPr>';
    const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">' + esc(instr) + '</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
    const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    if (!caps.length) {
      return '<w:p>' + pPr + begin + renderRun('<w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr>', 'No table of figures entries found.') + end + '</w:p>';
    }
    // Fresh bookmarks on the captions, numbered past every one already there.
    let n = TOF_BASE;
    const taken = new Set([...this.xml.matchAll(/\bw:name="(_Toc\d+)"/g)].map((m) => m[1]));
    const entries = caps.map((c, i) => {
      while (taken.has('_Toc' + n)) n++;
      const name = '_Toc' + n++;
      taken.add(name);
      this._mintBookmark(name, c.index, c.index);
      const pageRaw = pages ? (pages[c.block] ?? pages[String(c.block)]) : cached?.[i];
      const page = pageRaw === undefined || pageRaw === null ? '' : String(pageRaw);
      const text = includeLabel ? c.text : c.bare || c.text;
      const head = '<w:p>' + pPr + '<w:hyperlink w:anchor="' + name + '" w:history="1">' + renderRun('<w:rPr><w:noProof/></w:rPr>', text);
      if (!pageNumbers) return head + '</w:hyperlink></w:p>';
      return head + (right ? '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r>' : renderRun('<w:rPr><w:noProof/><w:webHidden/></w:rPr>', ' '))
        + '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF ' + name + ' \\h </w:instrText></w:r>'
        + '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>' + renderRun('<w:rPr><w:noProof/><w:webHidden/></w:rPr>', page)
        + '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>';
    });
    return '<w:p>' + pPr + begin + '</w:p>' + entries.join('') + '<w:p>' + pPr + end + '</w:p>';
  },

  /**
   * References → Insert Table of Figures, after paragraph `at` (the edit
   * address): built from `label`'s captions, the label and number included
   * unless `includeLabel` is false. A table of figures of the same label
   * already in the document is replaced where it stands, as Word asks to.
   */
  insertTableOfFigures({ at, label = 'Figure', includeLabel = true, pages = null, pageNumbers = true, rightAlign = true, leader = 'dot' } = {}) {
    const opts = { pageNumbers, rightAlign, leader };
    this._ensureFiguresStyle();
    const same = this._figureTables().find((t) => t.label === label);
    if (same) {
      const { body } = this._body();
      for (const name of this._figureAnchors(body.slice(same.start, same.end))) this.removeBookmark(name);
      const again = this._figureTables().find((t) => t.label === label);
      const xml = this._figuresXml(label, includeLabel, pages, null, opts);
      const now = this._figureTables().find((t) => t.label === label) || again;
      this._spliceBody(now.start, now.end, xml);
    } else {
      const p = this.editParagraph(at);
      if (!p) throw new Error('no paragraph at index ' + at);
      const top = this.paragraphs().find((q) => q.start <= p.start && q.end >= p.end) || p;
      // Bookmarks first — they move what follows — then the field at the paragraph's (new) end.
      const xml = this._figuresXml(label, includeLabel, pages, null, opts);
      const after = this.paragraphs().find((q) => q.index === top.index) || top;
      this._spliceBody(after.end, after.end, xml);
    }
    this.dirty = true;
    return this;
  },

  /** Update Table in the Captions group (and F9): every table of figures rebuilt from its captions and pages. */
  updateTablesOfFigures({ pages = null } = {}) {
    const tables = this._figureTables();
    if (!tables.length) return 0;
    const cachedAll = this.tablesOfFigures().map((t) => t.entries.map((e) => e.page));
    // Last first, so the earlier spans stay where they are.
    for (let k = tables.length - 1; k >= 0; k--) {
      const t = tables[k];
      const { body } = this._body();
      for (const name of this._figureAnchors(body.slice(t.start, t.end))) this.removeBookmark(name);
      const now = this._figureTables()[k];
      const xml = this._figuresXml(t.label, t.includeLabel, pages, pages ? null : cachedAll[k], t);
      const fresh = this._figureTables()[k] || now;
      this._spliceBody(fresh.start, fresh.end, xml);
    }
    this.dirty = true;
    return tables.length;
  },

  // ---- document fields -----------------------------------------------------

  /** Author and title from the document's core properties, as AUTHOR and TITLE read them. */
  coreProperties() {
    const part = this.pkg.has('docProps/core.xml') ? 'docProps/core.xml' : null;
    if (!part) return { author: null, title: null };
    const xml = this.pkg.text(part);
    const read = (tag) => {
      const m = new RegExp('<' + tag + '\\b[^>]*>([\\s\\S]*?)</' + tag + '>').exec(xml);
      return m ? m[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&') : null;
    };
    return { author: read('dc:creator'), title: read('dc:title') };
  },

  /** Insert → Field: a field's code and its result now, from the document's own properties and the page, count, time and file name given. */
  docFieldFor(spec, ctx = {}) {
    const instr = docFieldInstr(spec);
    const core = this.coreProperties();
    const text = evaluateDocField(instr, { author: core.author ?? '', title: core.title ?? '', ...ctx }) ?? '';
    return { instr, text };
  },

  /** A document field's whole XML — the complex field Word's dialog writes — with its result. */
  docFieldXml(instr, result, rPr = '') {
    // Every run of the field wears the result's look, as Word writes it; the
    // code keeps its quotes.
    const r = (inner) => '<w:r>' + (rPr || '') + inner + '</w:r>';
    return r('<w:fldChar w:fldCharType="begin"/>') + r('<w:instrText xml:space="preserve">' + esc(instr) + '</w:instrText>')
      + r('<w:fldChar w:fldCharType="separate"/>') + renderRun(rPr || null, String(result ?? '')) + r('<w:fldChar w:fldCharType="end"/>');
  },

  /**
   * F9, and printing: every PAGE, NUMPAGES, DATE, TIME, FILENAME, AUTHOR and
   * TITLE field's result worked out again — a page number from the page the
   * window laid its paragraph on (`pages`, by edit address), the count of
   * them (`pageCount`), today, the file's name. Only the results change.
   * Answers how many fields changed.
   */
  updateDocFields({ pages = null, pageCount = null, now = new Date(), fileName = null, filePath = null } = {}) {
    const core = this.coreProperties();
    const paras = this.editParagraphs();
    let changed = 0;
    // Last first: a result of another length moves every later paragraph.
    for (let i = paras.length - 1; i >= 0; i--) {
      const p = paras[i];
      if (!/<w:instrText\b[^>]*>\s*(?:PAGE|NUMPAGES|DATE|TIME|FILENAME|AUTHOR|TITLE)\b/i.test(p.xml)) continue;
      const page = pages ? pages[i] ?? pages[String(i)] ?? null : null;
      const ctx = { page, pages: pageCount, now, fileName, filePath, author: core.author, title: core.title };
      let n = 0;
      const next = mapComplexFieldResults(p.xml, (instr, text) => {
        const { name } = parseDocFieldInstr(instr);
        if (!['PAGE', 'NUMPAGES', 'DATE', 'TIME', 'FILENAME', 'AUTHOR', 'TITLE'].includes(name)) return undefined;
        const v = evaluateDocField(instr, ctx);
        if (v === undefined || v === text) return undefined;
        n += 1;
        return v;
      });
      if (next !== p.xml) {
        this._spliceBody(p.start, p.end, next);
        changed += n;
      }
    }
    if (changed) this.dirty = true;
    return changed;
  },

  /** Whether the body holds any document field — the window refreshes them before printing. */
  hasDocFields() {
    return /<w:instrText\b[^>]*>\s*(?:PAGE|NUMPAGES|DATE|TIME|FILENAME|AUTHOR|TITLE)\b/i.test(this._body().body);
  },
};
