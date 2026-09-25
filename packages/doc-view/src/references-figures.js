/**
 * References → Insert Table of Figures and Update Table, and Insert → Quick
 * Parts → Field, on DocView — installed onto the class with the rest of the
 * References tab (see references.js). The engine's half is `@rutba/ooxml`'s
 * references-figures.js.
 */
import { locate, sliceRuns, coalesce } from './positions.js';

function engineOf(view, what) {
  const doc = view.doc?.doc;
  if (!doc || typeof doc.insertTableOfFigures !== 'function') throw new Error('this document backend does not support ' + what);
  return doc;
}

export const figuresViewMethods = {
  /** Insert Table of Figures, after the caret's paragraph — or in place of the one for this label. */
  insertTableOfFigures({ label = 'Figure', includeLabel = true, pages = null, pageNumbers = true, rightAlign = true, leader = 'dot' } = {}) {
    const doc = engineOf(this, 'a table of figures');
    return this._edit('table of figures', null, () => {
      const { block } = this.focus;
      if (!this.block(block)) throw new Error('no paragraph at index ' + block);
      doc.insertTableOfFigures({ at: block, label, includeLabel, pages, pageNumbers, rightAlign, leader });
      if (typeof this._stylesWritten === 'function') this._stylesWritten();
      this._invalidate();
      return this;
    });
  },

  /** Update Table (Captions group): every table of figures from its captions and pages. */
  updateTablesOfFigures({ pages = null } = {}) {
    const doc = engineOf(this, 'a table of figures');
    return this._edit('update table of figures', null, () => {
      const n = doc.updateTablesOfFigures({ pages });
      this._invalidate();
      return n;
    });
  },

  /**
   * Insert → Quick Parts → Field: PAGE, NUMPAGES, DATE, TIME, FILENAME,
   * AUTHOR or TITLE at the caret, as the complex field Word's dialog writes,
   * its result worked out now — the caret's page from `page`, the count from
   * `pageCount`, the name from `fileName`.
   */
  insertDocField({ name, picture = '', format = '', path = false, page = null, pageCount = null, fileName = null, filePath = null, now = null } = {}) {
    const doc = engineOf(this, 'fields');
    const made = doc.docFieldFor({ name, picture, format, path }, { page, pages: pageCount, now: now ? new Date(now) : new Date(), fileName, filePath });
    const { instr } = made;
    const result = made.text;
    return this._edit('field', null, () => {
      if (!this.collapsed) this.deleteSelection();
      const { block, offset } = this.focus;
      const b = this._editable(block);
      const { runIndex } = locate(b.runs, offset, 'left');
      const here = b.runs[runIndex];
      const rPr = here && !here.field && !here.noteRef && !here.noteMark && !here.math && !here.del ? here.rPr : null;
      const run = {
        rPr, text: result, field: { instr, kind: String(name).toLowerCase(), name: null },
        complexXml: doc.docFieldXml(instr, result, rPr || ''), fieldCached: result, fieldRPr: rPr,
      };
      const next = [...sliceRuns(b.runs, 0, offset), run, ...sliceRuns(b.runs, offset, Infinity)];
      this.doc.setParagraphRuns(block, coalesce(next));
      this._invalidate();
      this.pendingFormat = null;
      this.collapseTo({ block, offset: offset + result.length });
      return this;
    });
  },

  /**
   * F9 and printing, for what the window's own page layout knows: every
   * table of figures, the index and every document field, in one step.
   * `pages` maps an edit-space paragraph to its page. Answers how many
   * fields changed.
   */
  refreshReferences({ pages = null, pageCount = null, fileName = null, filePath = null, now = null, fieldsOnly = false } = {}) {
    const doc = this.doc?.doc;
    if (!doc || typeof doc.updateDocFields !== 'function') return 0;
    const tables = fieldsOnly ? 0 : doc._figureTables().length;
    const index = fieldsOnly ? false : doc.hasIndex();
    const fields = doc.hasDocFields();
    if (!tables && !index && !fields) return 0;
    return this._edit('update fields', null, () => {
      let n = 0;
      if (tables) n += doc.updateTablesOfFigures({ pages });
      if (index && doc.updateIndex({ pages })) n += 1;
      if (fields) n += doc.updateDocFields({ pages, pageCount, fileName, filePath, now: now ? new Date(now) : new Date() });
      this._invalidate();
      return n;
    });
  },
};
