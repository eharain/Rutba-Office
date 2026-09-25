/**
 * References → Mark Entry, Insert Index and Update Index on DocView — the
 * index's editing verbs, installed onto the class with the rest of the
 * References tab (see references.js). The engine's half is `@rutba/ooxml`'s
 * references-index.js.
 */
import { sliceRuns, coalesce } from './positions.js';

function engineOf(view) {
  const doc = view.doc?.doc;
  if (!doc || typeof doc.indexEntries !== 'function') throw new Error('this document backend does not support an index');
  return doc;
}

/** An offset moved out of a field it falls inside, to the field's end — a mark never cuts a citation in two. */
function outsideFields(runs, offset) {
  let seen = 0;
  for (const r of runs) {
    const end = seen + r.text.length;
    if (r.field && offset > seen && offset < end) return end;
    seen = end;
  }
  return offset;
}

/** An XE field as one run of a paragraph: no words of its own, the field's XML (from the engine) kept whole. */
const xeRun = ({ instr, xml }) => ({
  rPr: null, text: '', field: { instr, kind: 'xe', name: null },
  complexXml: xml, fieldCached: '', fieldRPr: null,
});

export const indexViewMethods = {
  /**
   * Mark Entry (Alt+Shift+X): an XE field right after the selected words —
   * or at the caret — for `main` (and `sub`), with a cross-reference, a
   * bookmark's page range, or a bold or italic page number. `all` marks,
   * instead, the first time `text` (the selected words, by default) appears
   * in each paragraph, matching case, as Mark All does. Answers how many
   * entries were marked.
   */
  markIndexEntry({ main, sub = '', crossRef = '', bookmark = '', bold = false, italic = false, all = false, text = null } = {}) {
    const doc = engineOf(this);
    if (!String(main || '').trim()) throw new Error('an index entry needs its main words');
    const xe = doc.xeFor({ main, sub, crossRef, bookmark, bold, italic });
    return this._edit('mark index entry', null, () => {
      const { from, to } = this.selection;
      if (!all) {
        const b = this._editable(to.block);
        // After the words, not after the space a double-click takes with them.
        let end = to.offset;
        if (from.block === to.block) while (end > from.offset && /\s/.test((b.text || '')[end - 1])) end -= 1;
        const at = outsideFields(b.runs, end);
        const next = [...sliceRuns(b.runs, 0, at), xeRun(xe), ...sliceRuns(b.runs, at, Infinity)];
        this.doc.setParagraphRuns(to.block, coalesce(next));
        this._invalidate();
        return 1;
      }
      const words = text ?? (from.block === to.block ? (this.block(from.block)?.text || '').slice(from.offset, to.offset) : '');
      if (!words) throw new Error('select the words to mark in every paragraph');
      let marked = 0;
      const count = this.blocks.length;
      for (let i = 0; i < count; i++) {
        const b = this.block(i);
        // Never inside a list a field made: an index, a table of contents or of figures.
        if (!b || b.structural || /^(Index|TOC|TableofFigures)/.test(b.style || '')) continue;
        const hit = (b.text || '').indexOf(words);
        if (hit < 0) continue;
        const at = outsideFields(b.runs, hit + words.length);
        const next = [...sliceRuns(b.runs, 0, at), xeRun(xe), ...sliceRuns(b.runs, at, Infinity)];
        this.doc.setParagraphRuns(i, coalesce(next));
        this._invalidate();
        marked += 1;
      }
      return marked;
    });
  },

  /** Insert Index, after the caret's paragraph — or in place of the index already there. */
  insertIndex({ columns = 2, rightAlign = false, leader = 'dot', runIn = false, lcid = 1033, pages = null } = {}) {
    const doc = engineOf(this);
    return this._edit('index', null, () => {
      const { block } = this.focus;
      if (!this.block(block)) throw new Error('no paragraph at index ' + block);
      doc.insertIndex({ at: block, columns, rightAlign, leader, runIn, lcid, pages });
      if (typeof this._stylesWritten === 'function') this._stylesWritten();
      this._invalidate();
      return this;
    });
  },

  /** Update Index: the entries and pages as they stand. */
  updateIndex({ pages = null } = {}) {
    const doc = engineOf(this);
    return this._edit('update index', null, () => {
      const done = doc.updateIndex({ pages });
      this._invalidate();
      return done;
    });
  },
};
