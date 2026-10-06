/**
 * References → Mark Citation, Insert Table of Authorities and Update Table
 * on DocView — installed onto the class with the rest of the References tab
 * (see references.js). The engine's half is `@rutba/ooxml`'s
 * references-toa.js.
 */
import { sliceRuns, coalesce } from './positions.js';

function engineOf(view) {
  const doc = view.doc?.doc;
  if (!doc || typeof doc.authorityEntries !== 'function') throw new Error('this document backend does not support a table of authorities');
  return doc;
}

/** An offset moved out of a field it falls inside, to the field's end — a mark never cuts a field in two. */
function outsideFields(runs, offset) {
  let seen = 0;
  for (const r of runs) {
    const end = seen + r.text.length;
    if (r.field && offset > seen && offset < end) return end;
    seen = end;
  }
  return offset;
}

/** A TA field as one run of a paragraph: no words of its own, the field's XML (from the engine) kept whole. */
const taRun = ({ instr, xml }) => ({
  rPr: null, text: '', field: { instr, kind: 'ta', name: null },
  complexXml: xml, fieldCached: '', fieldRPr: null,
});

/** Where `needle` ends each time it appears in `text`, matching case — `taken` ranges skipped, and added to. */
function endsOf(text, needle, taken) {
  const out = [];
  if (!needle) return out;
  let from = 0;
  for (;;) {
    const at = text.indexOf(needle, from);
    if (at < 0) break;
    const end = at + needle.length;
    if (!taken.some(([a, b]) => at < b && end > a)) { taken.push([at, end]); out.push(end); }
    from = at + 1;
  }
  return out;
}

export const toaViewMethods = {
  /**
   * Mark Citation (Alt+Shift+I): a TA field right after the selected words
   * — the long citation, the short one and the category, or the short one
   * alone where that authority is already marked. `all` then marks every
   * other place the long citation (`text`, the words as the page has them)
   * or the short one appears, matching case, with the short form, as Mark
   * All does. Answers how many citations were marked.
   */
  markCitation({ long, short = '', category = 1, all = false, text = null } = {}) {
    const doc = engineOf(this);
    const longText = String(long || '').trim();
    if (!longText) throw new Error('a citation needs its long form');
    const shortText = String(short || '').trim() || longText;
    const known = doc.authorities().find((a) => a.long.toLowerCase() === longText.toLowerCase());
    const first = known ? doc.taFor({ short: known.short }) : doc.taFor({ long: longText, short: shortText, category });
    const again = doc.taFor({ short: known ? known.short : shortText });
    return this._edit('mark citation', null, () => {
      const { from, to } = this.selection;
      const b = this._editable(to.block);
      // After the words, not after the space a double-click takes with them.
      let end = to.offset;
      if (from.block === to.block) while (end > from.offset && /\s/.test((b.text || '')[end - 1])) end -= 1;
      const at = outsideFields(b.runs, end);
      const here = { block: to.block, from: from.block === to.block ? from.offset : 0, to: end };
      this.doc.setParagraphRuns(to.block, coalesce([...sliceRuns(b.runs, 0, at), taRun(first), ...sliceRuns(b.runs, at, Infinity)]));
      this._invalidate();
      let marked = 1;
      if (!all) return marked;
      const words = String(text ?? (from.block === to.block ? (this.block(from.block)?.text || '').slice(from.offset, to.offset) : '')).trim();
      const count = this.blocks.length;
      for (let i = 0; i < count; i++) {
        const p = this.block(i);
        // Never inside a list a field made: a table of authorities, an index, a table of contents.
        if (!p || p.structural || /^(TableofAuthorities|TOAHeading|Index|TOC|TableofFigures)/.test(p.style || '')) continue;
        const taken = i === here.block ? [[here.from, here.to]] : [];
        // The long form first, so the short one inside it is not marked twice.
        const ends = [...endsOf(p.text || '', words, taken), ...endsOf(p.text || '', shortText, taken)].sort((x, y) => y - x);
        if (!ends.length) continue;
        let runs = this._editable(i).runs;
        for (const e of ends) {
          const pos = outsideFields(runs, e);
          runs = [...sliceRuns(runs, 0, pos), taRun(again), ...sliceRuns(runs, pos, Infinity)];
          marked += 1;
        }
        this.doc.setParagraphRuns(i, coalesce(runs));
        this._invalidate();
      }
      return marked;
    });
  },

  /** Insert Table of Authorities, after the caret's paragraph — or in place of the table (or tables) already there. */
  insertTableOfAuthorities({ category = 'all', passim = true, keepFormatting = true, leader = 'dot', pages = null } = {}) {
    const doc = engineOf(this);
    return this._edit('table of authorities', null, () => {
      const { block } = this.focus;
      if (!this.block(block)) throw new Error('no paragraph at index ' + block);
      doc.insertTableOfAuthorities({ at: block, category, passim, keepFormatting, leader, pages });
      if (typeof this._stylesWritten === 'function') this._stylesWritten();
      this._invalidate();
      return this;
    });
  },

  /** Update Table: every table of authorities from the citations and pages as they stand. */
  updateTablesOfAuthorities({ pages = null } = {}) {
    const doc = engineOf(this);
    return this._edit('update table of authorities', null, () => {
      const n = doc.updateTablesOfAuthorities({ pages });
      this._invalidate();
      return n;
    });
  },
};
