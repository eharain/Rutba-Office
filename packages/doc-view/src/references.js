/**
 * The References tab's editing verbs on DocView — citations and the
 * bibliography — kept beside view.js rather than in it, and installed onto
 * the class once (`installReferenceViews`). Each is one undoable edit; the
 * engine's half is `@rutba/ooxml`'s references.js.
 */
import { locate, sliceRuns, coalesce } from './positions.js';
import { indexViewMethods } from './references-index.js';
import { figuresViewMethods } from './references-figures.js';

/** The Document behind an OOXML backend, or a sentence when this backend has none. */
function engineOf(view, what) {
  const doc = view.doc?.doc;
  if (!doc || typeof doc.bibliographySources !== 'function') throw new Error('this document backend does not support ' + what);
  return doc;
}

const methods = {
  /** Sources, style, citations and whether a bibliography is in — read once per version of the XML. */
  referencesInfo() {
    const doc = this.doc?.doc;
    if (!doc || typeof doc.referencesInfo !== 'function') return null;
    const part = doc._bibliographyPart();
    const key = doc.xml + '\u0000' + (part ? doc.pkg.text(part) : '');
    if (this._refsKey !== key) {
      this._refsKey = key;
      this._refsInfo = doc.referencesInfo();
    }
    return this._refsInfo;
  },

  /**
   * An edit that may write the sources part: the part is made real and put
   * on the undo list BEFORE the snapshot, so Ctrl+Z takes the sources back
   * with the words.
   */
  _refsEdit(label, fn) {
    const doc = engineOf(this, 'citations');
    doc.registerBibliographyUndo();
    return this._edit(label, null, () => fn(doc));
  },

  /**
   * References → Insert Citation, at the caret: the CITATION field in its
   * content control, one run of the paragraph (a selection is replaced, as
   * typing would). A space goes in first when the caret sits right after a
   * word, as Word puts one there.
   */
  insertCitation({ tags, pages = '', lcid = 1033 } = {}) {
    return this._refsEdit('citation', (doc) => {
      if (!this.collapsed) this.deleteSelection();
      const { block, offset } = this.focus;
      const b = this._editable(block);
      const cite = doc.citationFor({ tags, pages, lcid });
      const before = (b.text || '').slice(0, offset);
      const space = before && !/[\s(\[]$/.test(before) ? ' ' : '';
      const { runIndex } = locate(b.runs, offset, 'left');
      const here = b.runs[runIndex];
      const rPr = here && !here.field && !here.noteRef && !here.noteMark && !here.math && !here.del ? here.rPr : null;
      const run = {
        rPr: null, text: cite.text,
        field: { instr: cite.instr, kind: 'citation', name: tags[0] ?? null },
        complexXml: cite.xml, fieldCached: cite.text, fieldRPr: null,
      };
      const next = [
        ...sliceRuns(b.runs, 0, offset),
        ...(space ? [{ rPr, text: space }] : []),
        run,
        ...sliceRuns(b.runs, offset, Infinity),
      ];
      this.doc.setParagraphRuns(block, coalesce(next));
      this._invalidate();
      this.pendingFormat = null;
      this.collapseTo({ block, offset: offset + space.length + cite.text.length });
      return this;
    });
  },

  /**
   * Manage Sources closes, or Create Source is OK'd: the document's list as
   * the dialog left it, and every citation and the bibliography brought
   * into line with it — a source's author corrected is corrected in the
   * text at once, as in Word.
   */
  setSources({ sources }) {
    return this._refsEdit('sources', (doc) => {
      doc.setBibliographySources({ sources });
      doc.updateCitations();
      this._invalidate();
      return this;
    });
  },

  /** References → Style. */
  setBibliographyStyle(style) {
    return this._refsEdit('citation style', (doc) => {
      doc.setBibliographyStyle(style);
      this._invalidate();
      return this;
    });
  },

  /**
   * References → Bibliography: the building block after the caret's
   * paragraph — Bibliography, References or Works Cited — or, `heading`
   * null, Insert Bibliography's bare field.
   */
  insertBibliography({ heading = 'Bibliography' } = {}) {
    return this._refsEdit('bibliography', (doc) => {
      const { block } = this.focus;
      if (!this.block(block)) throw new Error('no paragraph at index ' + block);
      doc.insertBibliography({ at: block, heading });
      if (typeof this._stylesWritten === 'function') this._stylesWritten();
      this._invalidate();
      return this;
    });
  },

  /** Update Citations and Bibliography. Answers how many citations changed. */
  updateCitations() {
    return this._refsEdit('update citations', (doc) => {
      const n = doc.updateCitations();
      this._invalidate();
      return n;
    });
  },
};

/** Put the References verbs on DocView, and the references into every frame it renders. */
export function installReferenceViews(DocView) {
  for (const [name, fn] of Object.entries({ ...methods, ...indexViewMethods, ...figuresViewMethods })) {
    if (!Object.prototype.hasOwnProperty.call(DocView.prototype, name)) DocView.prototype[name] = fn;
  }
  const render = DocView.prototype.render;
  if (render && !render.withReferences) {
    const wrapped = function renderWithReferences(...args) {
      const frame = render.apply(this, args);
      if (frame && typeof frame === 'object') frame.references = this.referencesInfo();
      return frame;
    };
    wrapped.withReferences = true;
    DocView.prototype.render = wrapped;
  }
  // F9 writes citations too: the sources part goes on the undo list first.
  const update = DocView.prototype.updateFields;
  if (update && !update.withReferences) {
    const wrapped = function updateFields(...args) {
      const doc = this.doc?.doc;
      if (doc && typeof doc._bibliographyPart === 'function' && doc._bibliographyPart()) doc.registerBibliographyUndo();
      return update.apply(this, args);
    };
    wrapped.withReferences = true;
    DocView.prototype.updateFields = wrapped;
  }
}
