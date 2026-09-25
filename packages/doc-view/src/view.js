/**
 * DocView — a real `.docx` made editable.
 *
 * The document counterpart of `SheetView`, and built on the same two rules:
 * every semantic lives here where it can be tested in Node, and every write goes
 * through the preserving path so a customer's letter comes back with only the
 * paragraphs we touched rewritten.
 *
 * The one rule that shapes the whole thing: a paragraph carrying a field code,
 * bookmark, comment anchor or content control is NOT freely editable, because
 * rebuilding it from its runs would flatten that structure. Those paragraphs
 * render, they can be read and copied, and their content controls can be filled —
 * but typing into them is refused rather than attempted. In a bank-confirmation
 * letter that is the difference between updating a balance and destroying the
 * reference number that makes the letter valid.
 */
import { assertBackend, supportsContentControls } from './backend.js';
import { History } from '@rutba/editing';
import { paginate } from './paginate.js';
import { measureText, lineHeight as lineHeightOf } from '@rutba/drawing';
import { computeListLabels } from './lists.js';
import { bandForPage, resolveFields } from './bands.js';
import {
  MERGE_KINDS, readInstr, evaluateField, matchFields, mergeOrder, placeholderFor, mergeFieldInstr,
  addressBlockInstr, greetingLineInstr, ifInstr, conditionInstr, findDuplicates, rangeOf as rangeOfOrder,
} from '@rutba/ooxml/mailmerge';
import {
  runsText, locate, clampPosition, comparePositions, orderedRange,
  sliceRuns, removeRange, trackedRemoveRange, coalesce, samePosition,
} from './positions.js';

/**
 * One indent step, in twips (1/20 point) — half an inch, what Tab-to-indent and
 * the ribbon's indent buttons move by. A plain unit constant, not format
 * knowledge: the backend is the only place that knows a twip becomes `<w:ind>`.
 */
const INDENT_STEP = 720;
// The ruler's absolute paragraph properties, written as given.
const RULER_KEYS = ['leftTwips', 'firstLineTwips', 'hangingTwips', 'rightTwips', 'tabs'];

/**
 * Contextual spacing, resolved against the neighbours.
 *
 * Word drops the gap between two consecutive paragraphs of the same style
 * when either of them says "don't add space between paragraphs of the same
 * style" — directly, or through the style, which is how List Paragraph has
 * it. The rendered blocks get an explicit zero on that edge; the painter and
 * the paginator both read an explicit value before the style's.
 */
function contextualSpacing(sources, styles, blocks) {
  const says = (p) => Boolean(p.spacing?.contextual || styles?.[p.style || 'Normal']?.contextualSpacing);
  for (let i = 0; i + 1 < blocks.length; i++) {
    const a = sources[i];
    const b = sources[i + 1];
    if (!a || !b || (a.style || 'Normal') !== (b.style || 'Normal')) continue;
    if (says(a) || says(b)) {
      blocks[i].spaceAfterPx = 0;
      blocks[i + 1].spaceBeforePx = 0;
    }
  }
  return blocks;
}

export class DocView {

  /**
   * @param {object} backend  see backend.js — the format sits behind this port,
   *   which is what lets Mail compose an email with the same editor.
   */
  constructor(backend) {
    this.doc = assertBackend(backend);
    this.anchor = { block: 0, offset: 0 };
    this.focus = { block: 0, offset: 0 };
    this.pendingFormat = null; // a toggle armed with a collapsed caret
    this.touched = false;
    this._blocks = null;
    this._flow = null;
    this._pages = null;
    // Survives edits on purpose: a keystroke changes one paragraph, and the
    // rest of the document wraps to the same lines it did before.
    this._lineCache = new Map();
    this.history = new History();
    this._editDepth = 0;
  }

  /**
   * Run a mutation as ONE undo step.
   *
   * Public methods call each other — insertText delegates to deleteSelection
   * when there is a selection — so the snapshot is taken at the OUTERMOST call
   * only. Without the depth counter, replacing a selection by typing would cost
   * two undos and the first one would put back a state the user never saw.
   *
   * `group` is what makes a run of keystrokes one undo: consecutive edits with
   * the same group merge. Moving the caret breaks the group, which is why
   * setSelection calls history.break().
   */
  _edit(label, group, fn) {
    const outermost = this._editDepth === 0;
    if (outermost && typeof this.doc.snapshot === 'function') {
      this.history.record({
        state: this.doc.snapshot(),
        label,
        group,
        meta: { anchor: { ...this.anchor }, focus: { ...this.focus } },
      });
    }
    this._editDepth += 1;
    try {
      return fn();
    } finally {
      this._editDepth -= 1;
    }
  }

  /** Can this backend be undone at all? A backend without snapshots says no. */
  get canUndo() { return this.canEdit && this.history.canUndo; }
  get canRedo() { return this.canEdit && this.history.canRedo; }

  /**
   * Step back one edit, putting the caret back where it was.
   *
   * Restoring the selection matters more than it sounds: undo that leaves the
   * caret somewhere else makes the user hunt for what changed, and the next
   * keystroke lands in the wrong place.
   */
  undo() {
    if (typeof this.doc.snapshot !== 'function') return false;
    const entry = this.history.undo(
      this.doc.snapshot(),
      { anchor: { ...this.anchor }, focus: { ...this.focus } },
    );
    if (!entry) return false;
    this._apply(entry);
    return true;
  }

  redo() {
    if (typeof this.doc.snapshot !== 'function') return false;
    const entry = this.history.redo(
      this.doc.snapshot(),
      { anchor: { ...this.anchor }, focus: { ...this.focus } },
    );
    if (!entry) return false;
    this._apply(entry);
    return true;
  }

  _apply(entry) {
    this.doc.restore(entry.state);
    this._blocks = null;
    this._flow = null;
    this._pages = null;
    this.touched = true;
    this.pendingFormat = null;
    if (entry.meta) {
      // through clampPosition, so a snapshot from a longer document cannot
      // leave the caret past the end of a shorter one
      this.anchor = clampPosition(this.blocks, entry.meta.anchor);
      this.focus = clampPosition(this.blocks, entry.meta.focus);
    }
  }

  /*
   * There is deliberately no `DocView.open(bytes)`. A format-free editor cannot
   * know what bytes are, and making it guess would put an import of every
   * backend at the top of this file. Each backend exports its own opener:
   *   openDocx(buffer)  from '@rutba/doc-view/backends/ooxml'
   *   openHtml(body)    from '@rutba/doc-view/backends/html'
   */

  /** Paragraphs, decomposed. Recomputed after every edit — the XML is the truth. */
  get blocks() {
    if (!this._blocks) {
      const count = this.doc.paragraphCount();
      this._blocks = [];
      for (let i = 0; i < count; i++) this._blocks.push(this.doc.paragraph(i));
    }
    return this._blocks;
  }

  _invalidate() {
    this._blocks = null;
    this._flow = null;
    this._pages = null;
    this.touched = true;
  }

  /**
   * The body in DOCUMENT ORDER — what a renderer walks.
   *
   * `blocks` above is the editor's address space: every editable paragraph,
   * indexed flat in document order — since tables became editable that
   * includes their cell paragraphs, each carrying the `container` key of its
   * cell. This is the complementary view for the RENDERER: a paragraph entry
   * points back into `blocks` by index, a table entry carries the parsed
   * table with each cell paragraph's `blockIndex` attached, so a caret can
   * live in a drawn cell.
   *
   * A backend with no tables (an email body) omits `blocks()` and gets the plain
   * paragraph sequence, which is exactly right for it.
   */
  get flow() {
    if (!this._flow) {
      this._flow = typeof this.doc.blocks === 'function'
        ? this.doc.blocks()
        : this.blocks.map((b) => ({ kind: 'paragraph', paragraphIndex: b.index }));
    }
    return this._flow;
  }

  /** Page geometry, or null for a format with no pages. */
  get section() {
    return typeof this.doc.section === 'function' ? this.doc.section() : null;
  }

  /**
   * The flow laid onto sheets, with each page's header and footer resolved.
   *
   * Cached with the block list, because pagination costs real work — it wraps
   * every line in the document — and a keystroke must not re-lay a fifty-page
   * report before the character appears.
   */
  get pages() {
    if (this._pages !== null && this._pages !== undefined) return this._pages;
    const section = this.section;
    if (!section) return (this._pages = null);

    // Styles and numbering definitions never change under an edit — no editing
    // operation writes styles.xml or numbering.xml — so they are read once per
    // view rather than per keystroke. The LABELS are recomputed every layout,
    // because deleting item two renumbers item three.
    this._loadDefinitions();
    const listLabels = computeListLabels(this.flow, this.blocks, this._numberingDefs);
    const bands = typeof this.doc.headerFooters === 'function'
      ? this.doc.headerFooters()
      : { headers: {}, footers: {} };
    // The notes go to the paginator numbered, so each page can carry the
    // footnotes its references call for; the watermark rides every page.
    // Each section's own paper, when the sections differ — see paginate.
    const sections = typeof this.doc.sections === 'function' && this.doc.sectionCount?.() > 1 ? this.doc.sections() : null;
    const laid = paginate({
      flow: this.flow, blocks: this.blocks, section, sections,
      cache: this._lineCache, styles: this._docStyles, listLabels,
      notes: this._notes(), watermark: bands.watermark ?? null,
      math: (run, sizePx) => this.mathPrint(run, sizePx),
    });
    if (!laid) return (this._pages = null);

    const opts = { titlePage: section.titlePage, evenAndOdd: section.evenAndOdd };

    // An envelope in front of the letter has no header, footer or watermark,
    // and is not counted: the letter's first page is still page 1.
    const before = laid.pages.filter((p) => p.section?.envelope).length;
    for (const page of laid.pages) {
      if (page.section?.envelope) {
        page.header = null;
        page.footer = null;
        page.watermark = null;
        continue;
      }
      const number = page.number - before;
      const header = bandForPage(bands.headers, number, opts);
      const footer = bandForPage(bands.footers, number, opts);
      // Fields resolve PER PAGE — the whole point of a PAGE field is that it
      // says something different on each sheet.
      page.header = header ? resolveFields(header.paragraphs, { page: number, of: laid.count - before }) : null;
      page.footer = footer ? resolveFields(footer.paragraphs, { page: number, of: laid.count - before }) : null;
    }
    return (this._pages = laid);
  }

  /**
   * Equations as the print path lays them out, measured by something that
   * can lay MathML out — the desktop's own Chromium, which draws each one
   * once and hands back its size in ems and its picture. Keyed by the
   * equation's XML. Setting them lays the pages out again.
   *
   * @param {Map<string, {widthEm:number, heightEm:number, baselineEm:number, raster?:object}>|null} measures
   */
  setMathMeasures(measures) {
    this.mathMeasures = measures || null;
    this._pages = null;
    return this;
  }

  /**
   * One equation run's box on paper at a font size, in px: measured when the
   * desktop measured it, else the linear form set as a line of text — the
   * honest fallback where no MathML layout is to hand (Node, a test).
   */
  mathPrint(run, sizePx) {
    const shown = typeof this.doc.mathView === 'function' ? this.doc.mathView(run.math) : null;
    const jc = shown?.jc || null;
    const m = this.mathMeasures?.get(run.math?.xml);
    if (m) {
      return {
        widthPx: m.widthEm * sizePx, heightPx: m.heightEm * sizePx, baselinePx: m.baselineEm * sizePx,
        jc, raster: m.raster || null, text: null,
      };
    }
    const text = shown?.ascii ?? '';
    const heightPx = lineHeightOf(sizePx);
    return { widthPx: measureText(text, { size: sizePx }), heightPx, baselinePx: heightPx * 0.76, jc, raster: null, text };
  }

  block(index) { return this.blocks[index] ?? null; }
  get selection() { return orderedRange(this.anchor, this.focus); }
  get collapsed() { return samePosition(this.anchor, this.focus); }
  get isDirty() { return this.touched; }

  // ---- selection ---------------------------------------------------------

  setSelection(anchor, focus = anchor) {
    this.anchor = clampPosition(this.blocks, anchor);
    this.focus = clampPosition(this.blocks, focus);
    this.pendingFormat = null;
    // Moving the caret ends the run of keystrokes: "type, click away, type" is
    // two undos, not one. Guarded on edit depth because insertText collapses the
    // caret when it finishes — breaking there would defeat coalescing entirely.
    if (this._editDepth === 0) this.history.break();
    return this;
  }

  collapseTo(position) { return this.setSelection(position, position); }

  /** Arrow-key movement in character space, crossing paragraph boundaries. */
  moveCaret(direction, { extend = false } = {}) {
    const from = this.focus;
    let next = { ...from };
    const blocks = this.blocks;

    if (direction === 'left') {
      if (from.offset > 0) next = { block: from.block, offset: from.offset - 1 };
      else if (from.block > 0) {
        next = { block: from.block - 1, offset: blocks[from.block - 1].text.length };
      }
    } else if (direction === 'right') {
      const len = blocks[from.block].text.length;
      if (from.offset < len) next = { block: from.block, offset: from.offset + 1 };
      else if (from.block < blocks.length - 1) next = { block: from.block + 1, offset: 0 };
    } else if (direction === 'up') {
      next = from.block > 0
        ? { block: from.block - 1, offset: Math.min(from.offset, blocks[from.block - 1].text.length) }
        : { block: 0, offset: 0 };
    } else if (direction === 'down') {
      next = from.block < blocks.length - 1
        ? { block: from.block + 1, offset: Math.min(from.offset, blocks[from.block + 1].text.length) }
        : { block: blocks.length - 1, offset: blocks[blocks.length - 1].text.length };
    } else if (direction === 'home') {
      next = { block: from.block, offset: 0 };
    } else if (direction === 'end') {
      next = { block: from.block, offset: blocks[from.block].text.length };
    } else {
      throw new Error('unknown direction: ' + direction);
    }

    if (extend) this.focus = clampPosition(blocks, next);
    else this.collapseTo(next);
    return this;
  }

  selectAll() {
    const last = this.blocks.length - 1;
    return this.setSelection({ block: 0, offset: 0 }, { block: last, offset: this.blocks[last].text.length });
  }

  // ---- guards ------------------------------------------------------------

  _editable(index) {
    const b = this.block(index);
    if (!b) throw new Error('no paragraph at index ' + index);
    if (b.structural) {
      throw new Error(
        'This paragraph carries ' + b.structuralTags.join(', ') +
        ' and cannot be typed into. Edit its content control instead.',
      );
    }
    return b;
  }

  /** Can the current selection be typed into at all? */
  get canEdit() {
    const { from, to } = this.selection;
    for (let i = from.block; i <= to.block; i++) {
      if (this.blocks[i]?.structural) return false;
    }
    return true;
  }

  /** The cell a block lives in — null for prose. See the backend's `container`. */
  _containerOf(index) { return this.blocks[index]?.container ?? null; }

  /**
   * Refuse a range edit that crosses a cell wall.
   *
   * Deleting from prose into a table (or across two cells) has no honest
   * meaning in this model — Word clears cell contents while keeping the grid,
   * which is a different operation, not a bigger delete. Until somebody needs
   * that, crossing the boundary refuses with a sentence, the same stance the
   * structural guard takes. Formatting is exempt: it rewrites each paragraph
   * in place and never touches the walls.
   */
  _sameContainer(from, to) {
    const home = this._containerOf(from.block);
    for (let i = from.block + 1; i <= to.block; i++) {
      if (this._containerOf(i) !== home) {
        throw new Error('This selection crosses a table boundary — delete inside one cell at a time.');
      }
    }
  }

  // ---- editing -----------------------------------------------------------

  /**
   * Review → Track Changes: is this document recording right now? Read
   * straight from the file's own setting (`Document#trackRevisions`) rather
   * than cached, so a file opened with it already on — or another window's
   * edit landing between two of this one's — is never stale.
   */
  get recording() {
    return typeof this.doc.trackRevisions === 'function' && this.doc.trackRevisions();
  }

  /**
   * Turn recording on or off. Not an undoable edit — Word's own toggle is a
   * document SETTING, not a content change, so Ctrl+Z must not touch it —
   * but it does dirty the document, since `<w:trackRevisions/>` is written
   * to settings.xml on save. `author` sticks for every change this window
   * records until it is given a different one.
   */
  setTrackChanges(on, author) {
    if (typeof this.doc.setTrackRevisions !== 'function') {
      throw new Error('this document backend does not support tracked changes');
    }
    this.doc.setTrackRevisions(Boolean(on));
    if (author) this._trackAuthor = author;
    this.touched = true;
    this._invalidate();
    return this;
  }

  /**
   * The `ins`/`del` tag for the NEXT run recording writes: reused, id and
   * all, when `existing` is already a pending change by the same author —
   * so typing in the middle of your own still-open insertion extends the
   * one `w:ins`, the way Word's own does, rather than opening a fresh one
   * every keystroke. A different author, or nothing pending there, mints a
   * new id off the document's own counter.
   */
  _trackMeta(existing) {
    const author = this._trackAuthor || 'Rutba Office user';
    if (existing && existing.author === author) return existing;
    return { id: String(this.doc.nextTrackChangeId()), author, date: new Date().toISOString() };
  }

  /**
   * Insert text at the caret, replacing any selection.
   *
   * The inserted text inherits the formatting of the run to the LEFT of the
   * caret, which is what a person expects: type after a bold word and you get
   * bold. `pendingFormat` overrides that when the user hit Ctrl+B on a collapsed
   * caret and is about to type.
   */
  insertText(text) {
    if (text === '') return this;
    return this._edit('typing', 'type:' + this.focus.block, () => this._insertText(text));
  }

  _insertText(text) {
    if (!this.collapsed) this.deleteSelection();
    const { block, offset } = this.focus;
    const b = this._editable(block);

    const runs = b.runs.length ? b.runs : [{ rPr: null, text: '' }];
    const { runIndex, runOffset } = locate(runs, offset, 'left');
    const target = runs[runIndex];

    const recording = this.recording;

    // A note reference is a run of its own and takes no words: typing beside
    // it goes into a new run wearing the formatting of the text next to it,
    // never into the reference — whose rebuild writes an element, not text.
    // An equation is the same: one character, its own element, no words of
    // the paragraph's inside it.
    if (target.noteRef || target.noteMark || target.math) {
      const neighbour = runOffset === 0 ? runs[runIndex - 1] : runs[runIndex + 1];
      let rPr = neighbour && !neighbour.noteRef && !neighbour.noteMark && !neighbour.math ? neighbour.rPr : null;
      if (this.pendingFormat) rPr = this._applyPending(rPr);
      const at = runOffset === 0 ? runIndex : runIndex + 1;
      const newRun = recording ? { rPr, text, ins: this._trackMeta(null) } : { rPr, text };
      const next = [...runs.slice(0, at), newRun, ...runs.slice(at)];
      this.doc.setParagraphRuns(block, coalesce(next));
      this._invalidate();
      this.pendingFormat = null;
      this.collapseTo({ block, offset: offset + text.length });
      return this;
    }

    // A field run's text is a cached RESULT, not words to splice into: typing
    // at its start goes before it, and typing anywhere else in it — at its
    // end, or (a caret that should never rest there) inside it — goes after
    // the whole run, exactly as Backspace and Delete take the whole run at
    // once. See positions.js's `sliceRuns` for the same rule on deletion.
    if (target.field) {
      const before = runOffset === 0;
      const runStart = offset - runOffset;
      const insertAt = before ? runStart : runStart + target.text.length;
      const neighbour = before ? runs[runIndex - 1] : runs[runIndex + 1];
      let rPr = neighbour && !neighbour.field && !neighbour.noteRef && !neighbour.noteMark && !neighbour.math ? neighbour.rPr : null;
      if (this.pendingFormat) rPr = this._applyPending(rPr);
      const at = before ? runIndex : runIndex + 1;
      const newRun = recording ? { rPr, text, ins: this._trackMeta(null) } : { rPr, text };
      const next = [...runs.slice(0, at), newRun, ...runs.slice(at)];
      this.doc.setParagraphRuns(block, coalesce(next));
      this._invalidate();
      this.pendingFormat = null;
      this.collapseTo({ block, offset: insertAt + text.length });
      return this;
    }

    let rPr = target.rPr;
    if (this.pendingFormat) rPr = this._applyPending(rPr);

    let next;
    if (recording) {
      // Always split three ways while recording, even where the plain path
      // would just splice into the run in place: the new words need a
      // boundary of their own to carry `ins` on, and `coalesce` (below)
      // folds them straight back into whichever neighbour shares it — the
      // run we typed inside of, when it is already OUR pending insertion.
      const ins = this._trackMeta(target.ins && target.rPr === rPr ? target.ins : null);
      next = [
        ...runs.slice(0, runIndex),
        { ...target, text: target.text.slice(0, runOffset) },
        { rPr, text, ins },
        { ...target, text: target.text.slice(runOffset) },
        ...runs.slice(runIndex + 1),
      ];
    } else if (rPr === target.rPr) {
      next = runs.map((r, i) => (i === runIndex
        ? { ...r, text: r.text.slice(0, runOffset) + text + r.text.slice(runOffset) }
        : r));
    } else {
      // formatting differs from the run we landed in: split it around the insert
      next = [
        ...runs.slice(0, runIndex),
        { ...target, text: target.text.slice(0, runOffset) },
        { rPr, text },
        { ...target, text: target.text.slice(runOffset) },
        ...runs.slice(runIndex + 1),
      ];
    }

    this.doc.setParagraphRuns(block, coalesce(next));
    this._invalidate();
    this.pendingFormat = null;
    this.collapseTo({ block, offset: offset + text.length });
    return this;
  }

  /** Remove the selected range, joining paragraphs if it spans them. */
  deleteSelection() {
    if (this.collapsed) return this;
    return this._edit('delete', 'del:' + this.selection.from.block, () => this._deleteSelection());
  }

  _deleteSelection() {
    const { from, to } = this.selection;
    this._sameContainer(from, to);
    for (let i = from.block; i <= to.block; i++) this._editable(i);

    const recording = this.recording;
    // One id for the whole delete, however many runs (or paragraphs) it
    // crosses — Word does not mint a fresh one per run either.
    const meta = recording ? this._trackMeta(null) : null;

    if (from.block === to.block) {
      const b = this.block(from.block);
      this.doc.setParagraphRuns(from.block, coalesce(trackedRemoveRange(b.runs, from.offset, to.offset, recording, meta)));
      this._invalidate();
      this.collapseTo(from);
      return this;
    }

    // Keep the head of the first paragraph and the tail of the last, mark
    // (or drop) what is between, then merge the two survivors into one. The
    // paragraph mark itself is not tracked — joining two paragraphs across a
    // tracked delete happens for real, a stated simplification (see the
    // module doc on paragraph marks).
    const first = this.block(from.block);
    const last = this.block(to.block);
    const head = trackedRemoveRange(first.runs, from.offset, first.text.length, recording, meta);
    const tail = trackedRemoveRange(last.runs, 0, to.offset, recording, meta);

    this.doc.setParagraphRuns(from.block, coalesce([...head, ...tail]));
    this._invalidate();
    for (let i = to.block; i > from.block; i--) {
      this.doc.removeParagraph(i);
      this._invalidate();
    }
    this.collapseTo(from);
    return this;
  }

  /** Backspace. At offset 0 it merges with the paragraph above. */
  deleteBackward() {
    if (!this.collapsed) return this.deleteSelection();
    return this._edit('delete', 'del:' + this.focus.block, () => this._deleteBackward());
  }

  _deleteBackward() {
    const { block, offset } = this.focus;
    if (offset > 0) {
      this.setSelection({ block, offset: offset - 1 }, { block, offset });
      return this.deleteSelection();
    }
    if (block === 0) return this;
    // Backspace at the top of a cell stops at the wall, exactly as Word's
    // does: the paragraph above belongs to another cell (or to the prose the
    // table sits in), and merging across would tear the grid.
    if (this._containerOf(block - 1) !== this._containerOf(block)) return this;
    const previous = this._editable(block - 1);
    this._editable(block);
    const joinAt = previous.text.length;
    this.doc.mergeWithNext(block - 1);
    this._invalidate();
    this.collapseTo({ block: block - 1, offset: joinAt });
    return this;
  }

  /** Delete. At the end of a paragraph it pulls the next one up. */
  deleteForward() {
    if (!this.collapsed) return this.deleteSelection();
    return this._edit('delete', 'del:' + this.focus.block, () => this._deleteForward());
  }

  _deleteForward() {
    const { block, offset } = this.focus;
    const b = this.block(block);
    if (offset < b.text.length) {
      this.setSelection({ block, offset }, { block, offset: offset + 1 });
      return this.deleteSelection();
    }
    if (block >= this.blocks.length - 1) return this;
    // Delete at the end of a cell stops at the wall too — the mirror of
    // backspace at its top.
    if (this._containerOf(block + 1) !== this._containerOf(block)) return this;
    this._editable(block);
    this._editable(block + 1);
    this.doc.mergeWithNext(block);
    this._invalidate();
    this.collapseTo({ block, offset });
    return this;
  }

  /** Enter. Splits at the caret; the new paragraph inherits the style. */
  splitParagraph() {
    // A new paragraph is always its own undo step: nobody expects Enter to be
    // swallowed into the sentence they typed before it.
    return this._edit('new paragraph', null, () => this._splitParagraph());
  }

  _splitParagraph() {
    if (!this.collapsed) this.deleteSelection();
    const { block, offset } = this.focus;
    const b = this._editable(block);
    const runs = b.runs.length ? b.runs : [{ rPr: null, text: '' }];
    const { runIndex, runOffset } = locate(runs, offset, 'left');
    this.doc.splitParagraph(block, runIndex, runOffset);
    this._invalidate();
    this.collapseTo({ block: block + 1, offset: 0 });
    return this;
  }

  /**
   * Toggle bold / italic / underline / strikethrough over the selection.
   *
   * With a collapsed caret it arms `pendingFormat` instead — Ctrl+B then typing
   * should produce bold text, not silently do nothing.
   *
   * The toggle is "make it all on unless it is already all on", which is what
   * every word processor does: select a mix of bold and plain and one press
   * makes the lot bold rather than inverting each run.
   */
  toggleFormat(tag) {
    if (!['b', 'i', 'u', 's'].includes(tag)) throw new Error('unknown format: ' + tag);
    return this._edit('format', null, () => this._toggleFormat(tag));
  }

  _toggleFormat(tag) {
    const key = { b: 'bold', i: 'italic', u: 'underline', s: 'strike' }[tag];

    if (this.collapsed) {
      const active = this.formatAtCaret();
      this.pendingFormat = { ...(this.pendingFormat ?? {}), [tag]: !active[key] };
      return this;
    }

    const { from, to } = this.selection;
    for (let i = from.block; i <= to.block; i++) this._editable(i);

    const spans = [];
    for (let i = from.block; i <= to.block; i++) {
      const b = this.block(i);
      const start = i === from.block ? from.offset : 0;
      const end = i === to.block ? to.offset : b.text.length;
      spans.push({ index: i, start, end, covered: sliceRuns(b.runs, start, end) });
    }
    const allOn = spans.every((s) => s.covered.length > 0 && s.covered.every((r) => r[key]));
    const turnOn = !allOn;

    // last to first, so earlier paragraph offsets stay valid as we rewrite
    for (const span of [...spans].reverse()) {
      const b = this.block(span.index);
      const next = [
        ...sliceRuns(b.runs, 0, span.start),
        ...sliceRuns(b.runs, span.start, span.end).map((r) => ({ ...r, rPr: this.doc.toggleRunFormat(r.rPr, tag, turnOn) })),
        ...sliceRuns(b.runs, span.end, Infinity),
      ];
      this.doc.setParagraphRuns(span.index, coalesce(next));
      this._invalidate();
    }
    return this;
  }

  /**
   * Set or clear a VALUE character format — font family, size (points) or
   * colour — over the selection. The sibling of `toggleFormat`: a toggle is
   * on/off, this carries a value. `delta` names the properties to change
   * (`fontName`, `fontSize`, `fontColour`); a value sets, `null` clears, an
   * absent key is left alone.
   *
   * With a collapsed caret it arms `pendingFormat` instead, so the choice lands
   * on the next character typed — exactly as `toggleFormat` arms a toggle.
   */
  setRunFormat(delta) {
    if (typeof this.doc.setRunProp !== 'function') {
      throw new Error('this document backend does not support character formatting');
    }
    const props = {};
    for (const key of ['fontName', 'fontSize', 'fontColour', 'highlight', 'vertAlign', 'outline', 'shadow', 'glow']) {
      if (delta && key in delta) props[key] = delta[key];
    }
    if (Object.keys(props).length === 0) return this;
    return this._edit('format', null, () => this._setRunFormat(props));
  }

  /**
   * Clear ALL direct character formatting over the selection — the ribbon's
   * eraser. Stronger than clearing the properties this editor models: the
   * covered runs' `rPr` goes entirely, so an effect we merely preserve (a
   * shadow, a spacing tweak) goes with it. That is what "clear formatting"
   * has always meant, and an explicit request is exactly the moment the
   * only-write-what-was-set rule does not apply.
   *
   * With a collapsed caret it arms the next character to be written plain,
   * the way the toggles arm themselves.
   */
  clearFormat() {
    return this._edit('format', null, () => this._clearFormat());
  }

  _clearFormat() {
    if (this.collapsed) {
      this.pendingFormat = { ...(this.pendingFormat ?? {}), clear: true };
      return this;
    }

    const { from, to } = this.selection;
    for (let i = from.block; i <= to.block; i++) this._editable(i);

    const spans = [];
    for (let i = from.block; i <= to.block; i++) {
      const b = this.block(i);
      spans.push({
        index: i,
        start: i === from.block ? from.offset : 0,
        end: i === to.block ? to.offset : b.text.length,
      });
    }
    // last to first, so earlier paragraph offsets stay valid as we rewrite
    for (const span of [...spans].reverse()) {
      const b = this.block(span.index);
      const next = [
        ...sliceRuns(b.runs, 0, span.start),
        ...sliceRuns(b.runs, span.start, span.end).map((r) => ({ ...r, rPr: null })),
        ...sliceRuns(b.runs, span.end, Infinity),
      ];
      this.doc.setParagraphRuns(span.index, coalesce(next));
      this._invalidate();
    }
    return this;
  }

  _setRunFormat(props) {
    const keys = Object.keys(props);

    if (this.collapsed) {
      this.pendingFormat = { ...(this.pendingFormat ?? {}) };
      for (const key of keys) this.pendingFormat[key] = props[key];
      return this;
    }

    const { from, to } = this.selection;
    for (let i = from.block; i <= to.block; i++) this._editable(i);

    const spans = [];
    for (let i = from.block; i <= to.block; i++) {
      const b = this.block(i);
      spans.push({
        index: i,
        start: i === from.block ? from.offset : 0,
        end: i === to.block ? to.offset : b.text.length,
      });
    }

    // last to first, so earlier paragraph offsets stay valid as we rewrite
    for (const span of [...spans].reverse()) {
      const b = this.block(span.index);
      const next = [
        ...sliceRuns(b.runs, 0, span.start),
        ...sliceRuns(b.runs, span.start, span.end).map((r) => {
          let rPr = r.rPr;
          for (const key of keys) rPr = this.doc.setRunProp(rPr, key, props[key]);
          return { ...r, rPr };
        }),
        ...sliceRuns(b.runs, span.end, Infinity),
      ];
      this.doc.setParagraphRuns(span.index, coalesce(next));
      this._invalidate();
    }
    return this;
  }

  /**
   * Make the selection a link, or unlink it.
   *
   * `url` non-null: the covered runs get one link token (minted by the
   * backend — a relationship plus wrapper attributes), and the classic look
   * (the docx hyperlink blue, underlined) lands as DIRECT formatting so the
   * link reads as one in Word too, not only under our painter. With nothing
   * selected the address itself is inserted and linked, which is what every
   * editor does with a bare Ctrl+K.
   *
   * `url` null: the link comes off; the formatting deliberately stays — the
   * eraser exists, and silently unstyling text a person may have coloured
   * themselves is the kind of guess this editor does not make.
   */
  setLink(url) {
    if (typeof this.doc.makeLink !== 'function') {
      throw new Error('this document backend does not support links');
    }
    return this._edit('link', null, () => this._setLink(url));
  }

  _setLink(url) {
    if (this.collapsed) {
      if (url == null) return this; // nothing selected, nothing to unlink
      const { block, offset } = this.focus;
      this._insertText(String(url));
      this.setSelection({ block, offset }, { block, offset: offset + String(url).length });
    }
    const token = url == null ? null : this.doc.makeLink(String(url));

    const { from, to } = this.selection;
    for (let i = from.block; i <= to.block; i++) this._editable(i);
    const spans = [];
    for (let i = from.block; i <= to.block; i++) {
      const b = this.block(i);
      spans.push({
        index: i,
        start: i === from.block ? from.offset : 0,
        end: i === to.block ? to.offset : b.text.length,
      });
    }
    // last to first, so earlier paragraph offsets stay valid as we rewrite
    for (const span of [...spans].reverse()) {
      const b = this.block(span.index);
      const next = [
        ...sliceRuns(b.runs, 0, span.start),
        ...sliceRuns(b.runs, span.start, span.end).map((r) => {
          let rPr = r.rPr;
          if (token !== null && typeof this.doc.setRunProp === 'function') {
            rPr = this.doc.setRunProp(rPr, 'fontColour', '#0563C1');
            rPr = this.doc.toggleRunFormat(rPr, 'u', true);
          }
          const out = { ...r, rPr };
          if (token === null) delete out.link;
          else out.link = token;
          return out;
        }),
        ...sliceRuns(b.runs, span.end, Infinity),
      ];
      this.doc.setParagraphRuns(span.index, coalesce(next));
      this._invalidate();
    }
    return this;
  }

  /**
   * Paragraph-level formatting — alignment and indentation.
   *
   * The paragraph sibling of `setRunFormat`. Where that carries a value onto the
   * RUNS a selection covers, this carries onto the PARAGRAPHS it spans: from the
   * anchor block to the focus block inclusive, so a selection dragged across
   * three paragraphs aligns all three, and a collapsed caret aligns the one it
   * sits in. There is no collapsed-caret "pending" case — a paragraph always
   * exists, so the change lands immediately.
   *
   * `delta` names what to change:
   *   - `align`: 'left' | 'center' | 'right' | 'justify', or null to clear back
   *     to the document default. Justify becomes the docx 'both' at the backend.
   *   - `indentDelta`: +1 or -1 — one `INDENT_STEP` in or out. Indentation is
   *     clamped at zero; you cannot out-dent past the margin.
   *   - `styleId`: a named paragraph style from `paragraphStyles` (the
   *     document's own catalogue), or null to clear back to the default look.
   *     Applied FIRST, before any direct property in the same delta, because
   *     that is what a style is: the base the paragraph's own alignment,
   *     indent and list membership sit on top of — none of which it disturbs.
   *
   * One undo step, and it marks the document dirty like any other edit. A
   * backend with no paragraph properties (an email body) simply does not offer
   * the port, and this refuses rather than pretending.
   */
  setParagraphFormat(delta) {
    if (!delta) return this;
    const hasAlignOrIndent = ('align' in delta) || Boolean(delta.indentDelta) || RULER_KEYS.some((k) => k in delta);
    const hasStyle = ('styleId' in delta);
    const hasList = ('list' in delta);
    const hasSpacing = ('lineSpacing' in delta) || ('spaceBefore' in delta) || ('spaceAfter' in delta);
    const hasLook = ('shading' in delta) || ('borders' in delta);
    if ((hasAlignOrIndent || hasStyle || hasSpacing || hasLook) && typeof this.doc.setParagraphProp !== 'function') {
      throw new Error('this document backend does not support paragraph formatting');
    }
    if (hasList && typeof this.doc.setParagraphList !== 'function') {
      throw new Error('this document backend does not support list formatting');
    }
    if (!hasAlignOrIndent && !hasList && !hasStyle && !hasSpacing && !hasLook) return this;
    return this._edit('paragraph', null, () => this._setParagraphFormat(delta));
  }

  /**
   * The selected paragraphs in the order of their words — A to Z, or Z to
   * A. With nothing selected, the whole body, as Word offers. One undo
   * step; the caret lands on the first of them.
   */
  sortParagraphs({ descending = false } = {}) {
    if (typeof this.doc.sortParagraphs !== 'function') throw new Error('this document backend cannot sort paragraphs');
    const { from, to } = this.selection;
    const collapsed = from.block === to.block && from.offset === to.offset;
    const first = collapsed ? 0 : from.block;
    const last = collapsed ? this.blocks.length - 1 : to.block;
    if (last <= first) return this;
    return this._edit('sort paragraphs', null, () => {
      this.doc.sortParagraphs(first, last, { descending });
      this._invalidate();
      this.collapseTo({ block: first, offset: 0 });
      return this;
    });
  }

  _setParagraphFormat(delta) {
    const { from, to } = this.selection;
    for (let i = from.block; i <= to.block; i++) this._editable(i);

    for (let i = from.block; i <= to.block; i++) {
      if ('styleId' in delta) {
        this.doc.setParagraphProp(i, 'style', delta.styleId ?? null);
      }
      if ('align' in delta) {
        this.doc.setParagraphProp(i, 'align', delta.align ?? null);
      }
      if (delta.indentDelta) {
        const pp = this.doc.getParagraphProps(i);
        // In a list, in and out are the list's levels — 1. to a. to i. — as
        // Word's Increase indent has them; elsewhere, the paragraph's own indent.
        if (pp?.listType && typeof this.doc.setParagraphListLevel === 'function') {
          this.doc.setParagraphListLevel(i, (pp.listLevel || 0) + delta.indentDelta);
        } else {
          const current = pp?.indentTwips ?? 0;
          this.doc.setParagraphProp(i, 'indentTwips', Math.max(0, current + delta.indentDelta * INDENT_STEP));
        }
      }
      // The ruler: a marker was dragged to a place, so the value is absolute.
      for (const key of RULER_KEYS) {
        if (key in delta) this.doc.setParagraphProp(i, key, delta[key] ?? null);
      }
      if ('list' in delta) {
        this.doc.setParagraphList(i, delta.list ?? null);
      }
      // Spacing: the view speaks points and a line multiplier; the backend
      // knows what a twentieth of a point is.
      if ('lineSpacing' in delta) {
        this.doc.setParagraphProp(i, 'lineSpacing', delta.lineSpacing ?? null);
      }
      if ('spaceBefore' in delta) {
        this.doc.setParagraphProp(i, 'spaceBeforePts', delta.spaceBefore ?? null);
      }
      if ('spaceAfter' in delta) {
        this.doc.setParagraphProp(i, 'spaceAfterPts', delta.spaceAfter ?? null);
      }
      // A colour behind the paragraph and lines round it, as Word keeps them.
      if ('shading' in delta) this.doc.setParagraphProp(i, 'shading', delta.shading ?? null);
      if ('borders' in delta) this.doc.setParagraphProp(i, 'borders', delta.borders ?? null);
    }
    // A list toggle can add a definition to numbering.xml, and applying a
    // named style can add the standard styles part to a file that had none —
    // either way the copies cached for pagination (read once on the
    // assumption those parts never change under an edit) are now stale.
    if ('list' in delta || 'styleId' in delta) {
      this._docStyles = undefined;
      this._styleCatalogue = undefined;
    }
    this._invalidate();
    return this;
  }

  /** Fold every armed format — toggles and values alike — onto an rPr. */
  _applyPending(rPr) {
    // An armed clear wipes the inherited formatting FIRST, so anything armed
    // after it (clear, then bold) still lands on the blank slate.
    if (this.pendingFormat.clear) rPr = null;
    for (const [key, value] of Object.entries(this.pendingFormat)) {
      if (key === 'clear') continue;
      if (key === 'b' || key === 'i' || key === 'u' || key === 's') {
        rPr = this.doc.toggleRunFormat(rPr, key, value);
      } else if (typeof this.doc.setRunProp === 'function') {
        rPr = this.doc.setRunProp(rPr, key, value);
      }
    }
    return rPr;
  }

  /** What formatting a keystroke would carry right now — for the toolbar state. */
  formatAtCaret() {
    const { block, offset } = this.focus;
    const b = this.block(block);
    const base = {
      bold: false, italic: false, underline: false, strike: false,
      fontName: null, fontSize: null, fontColour: null, highlight: null,
      outline: false, shadow: false, glow: null,
      link: null,
    };
    if (b && b.runs.length) {
      const { runIndex } = locate(b.runs, Math.max(0, offset - (offset > 0 ? 1 : 0)), 'left');
      const run = b.runs[runIndex];
      if (run) {
        Object.assign(base, {
          bold: run.bold, italic: run.italic, underline: run.underline,
          strike: Boolean(run.strike),
        });
        if (typeof this.doc.readRunProps === 'function') Object.assign(base, this.doc.readRunProps(run.rPr));
        if (run.link != null && typeof this.doc.linkTarget === 'function') {
          base.link = this.doc.linkTarget(run.link);
        }
      }
    }
    // Paragraph-level state for the toolbar, named apart from the run fields on
    // purpose: alignment and indent belong to the PARAGRAPH the caret sits in,
    // not to the character under it. A backend without paragraph properties
    // (Mail's HTML body) leaves the sensible defaults in place.
    base.paragraphAlign = 'left';
    base.indentLevel = 0;
    // Whether the caret paragraph is a list, and which kind — 'bullet'|'number'|
    // null. Named apart from the run fields for the same reason alignment is: it
    // belongs to the PARAGRAPH, not the character under the caret.
    base.listType = null;
    base.listLevel = null;
    // The caret paragraph's NAMED style id, or null for the default — what the
    // ribbon's Style dropdown shows as selected.
    base.paragraphStyle = null;
    // Spacing is paragraph state too — what the ribbon's spacing menu ticks.
    base.lineSpacing = null;
    base.spaceBefore = null;
    base.spaceAfter = null;
    // Whether the caret sits in a drop cap's letter or its body — { kind,
    // lines } or null — for the ribbon's Drop Cap menu to tick.
    base.dropCap = null;
    if (b && typeof this.doc.getParagraphProps === 'function') {
      const pp = this.doc.getParagraphProps(block);
      if (pp) {
        base.paragraphAlign = pp.align ?? 'left';
        base.indentLevel = pp.listType ? (pp.listLevel || 0) : pp.indentTwips ? Math.round(pp.indentTwips / INDENT_STEP) : 0;
        base.listType = pp.listType ?? null;
        base.listLevel = pp.listType ? (pp.listLevel || 0) : null;
        base.paragraphStyle = pp.style ?? null;
        base.lineSpacing = pp.lineSpacing ?? null;
        base.spaceBefore = pp.spaceBeforePts ?? null;
        base.spaceAfter = pp.spaceAfterPts ?? null;
        // A drop cap's own paragraph carries it; the caret usually lands in
        // the BODY straight after setting one, so the ribbon checks the
        // pair's spec there too rather than only on the letter's paragraph.
        base.dropCap = pp.dropCap ?? (block > 0 ? this.doc.getParagraphProps(block - 1)?.dropCap ?? null : null);
      }
    }
    if (this.pendingFormat) {
      if (this.pendingFormat.clear) {
        // The next character types plain: the toolbar must not keep showing
        // the formatting the eraser just disarmed.
        Object.assign(base, {
          bold: false, italic: false, underline: false, strike: false,
          fontName: null, fontSize: null, fontColour: null, highlight: null,
          outline: false, shadow: false, glow: null,
        });
      }
      for (const [key, value] of Object.entries(this.pendingFormat)) {
        if (key === 'clear') continue;
        if (key === 'b' || key === 'i' || key === 'u' || key === 's') {
          base[{ b: 'bold', i: 'italic', u: 'underline', s: 'strike' }[key]] = value;
        } else base[key] = value;
      }
    }
    return base;
  }

  /**
   * The document's named-style catalogue — `[{ id, name }]`, paragraph styles
   * only — for the ribbon's Style dropdown. Computed once per load and cached,
   * like the resolved styles pagination reads: no editing operation writes
   * styles.xml, so re-asking every frame would buy nothing. A backend without a
   * catalogue (an email body) yields an empty list, and the dropdown is simply
   * not rendered — the no-dead-controls rule.
   */
  /** Styles and numbering, read once per view — see `_loadDefinitions`. */
  get docStyles() {
    this._loadDefinitions();
    return this._docStyles;
  }

  _loadDefinitions() {
    if (this._docStyles !== undefined) return;
    this._docStyles = typeof this.doc.paragraphStyles === 'function' ? this.doc.paragraphStyles() : null;
    this._charStyles = typeof this.doc.characterStyles === 'function' ? this.doc.characterStyles() : null;
    this._numberingDefs = typeof this.doc.numberingDefs === 'function' ? this.doc.numberingDefs() : null;
  }

  get paragraphStyles() {
    if (this._styleCatalogue === undefined) {
      this._styleCatalogue = typeof this.doc.paragraphStyleCatalogue === 'function'
        ? this.doc.paragraphStyleCatalogue()
        : [];
    }
    return this._styleCatalogue;
  }

  /** Plain text of the selection, for copy. */
  copyText() {
    const { from, to } = this.selection;
    if (from.block === to.block) return this.block(from.block).text.slice(from.offset, to.offset);
    const parts = [this.block(from.block).text.slice(from.offset)];
    for (let i = from.block + 1; i < to.block; i++) parts.push(this.block(i).text);
    parts.push(this.block(to.block).text.slice(0, to.offset));
    return parts.join('\n');
  }

  /** Paste. Newlines become paragraph splits, so structure survives the round trip. */
  pasteText(text) {
    // A paste is one undo however many paragraphs it spans — which is what the
    // depth counter in _edit buys: the nested insert/split calls do not record.
    return this._edit('paste', null, () => this._pasteText(text));
  }

  _pasteText(text) {
    const lines = String(text).replace(/\r\n?/g, '\n').split('\n');
    this.insertText(lines[0]);
    for (let i = 1; i < lines.length; i++) {
      this.splitParagraph();
      this.insertText(lines[i]);
    }
    return this;
  }

  /**
   * A page break at the caret — Ctrl+Enter, the way Word means it: the text
   * after the caret starts a fresh page. Stored as `w:pageBreakBefore` on the
   * paragraph the split creates, a property the paginator already honours,
   * rather than a run-level `w:br` — a break that IS a paragraph property
   * cannot end up half-deleted inside a run.
   *
   * Refused in a table cell: a page cannot break inside one (the paginator
   * breaks tables between rows), and setting a property nothing will honour is
   * the dishonesty the no-dead-controls rule exists to prevent.
   */
  insertPageBreak() {
    if (typeof this.doc.setParagraphProp !== 'function' || !this.section) {
      throw new Error('this document backend has no pages to break');
    }
    if (this._containerOf(this.focus.block) !== null) {
      throw new Error('A page break cannot start inside a table cell — move the caret out of the table.');
    }
    return this._edit('page break', null, () => {
      this._splitParagraph();
      this.doc.setParagraphProp(this.focus.block, 'pageBreakBefore', true);
      this._invalidate();
      return this;
    });
  }

  /**
   * A drop cap — Insert → Drop Cap. Word's own trick: the first letter
   * becomes its own paragraph, framed (`framePr`) to stand `lines` lines
   * deep, its run sized to match; the paragraph after it is the body,
   * unchanged. `{ lines = 3, kind = 'drop' | 'margin' }` sets one, on the
   * caret's paragraph — or, when the caret already sits in a drop cap's
   * letter or its body, on that existing pair, so picking a different
   * style restyles it rather than nesting a second split inside the first.
   * `null` clears it, merging the letter back into its paragraph.
   */
  setDropCap(spec) {
    if (typeof this.doc.setParagraphProp !== 'function') {
      throw new Error('this document backend does not support drop caps');
    }
    return this._edit('drop cap', null, () => (spec ? this._setDropCap(spec) : this._clearDropCap()));
  }

  /** The drop-cap pair a block belongs to — its own, or the one it is the body of. */
  _dropCapPairAt(block) {
    const b = this.block(block);
    if (b?.dropCap) return { dropIndex: block, bodyIndex: block + 1 };
    const prev = block > 0 ? this.block(block - 1) : null;
    if (prev?.dropCap) return { dropIndex: block - 1, bodyIndex: block };
    return null;
  }

  _setDropCap({ lines = 3, kind = 'drop' } = {}) {
    const { block } = this.focus;
    const pair = this._dropCapPairAt(block);
    let dropIndex;
    if (pair) {
      dropIndex = pair.dropIndex;
    } else {
      const b = this.block(block);
      // The same refusal for every reason the caret cannot start one here —
      // an empty paragraph, one that opens with a picture (no text at all,
      // the same shape), a table cell, or a structural paragraph. One
      // sentence rather than four keeps the control's excuse honest without
      // making the person read a taxonomy of docx paragraph kinds.
      const text = b && !b.structural && b.container == null ? (b.text ?? '') : '';
      const firstLetter = /\S/.exec(text);
      if (!firstLetter) throw new Error('Put the caret in a paragraph that starts with a letter');
      const runs = b.runs.length ? b.runs : [{ rPr: null, text: '' }];
      const { runIndex, runOffset } = locate(runs, firstLetter.index + 1, 'left');
      this.doc.splitParagraph(block, runIndex, runOffset);
      this._invalidate();
      dropIndex = block;
    }

    const dropBlock = this._editable(dropIndex);
    // The letter's size: its own run's, the paragraph style's resolved size,
    // or 11pt — the same resolution `formatAtCaret` gives the toolbar, so a
    // drop cap looks proportioned to text that had no explicit size at all.
    const run0 = dropBlock.runs[0] ?? { rPr: null };
    let base = typeof this.doc.readRunProps === 'function' ? this.doc.readRunProps(run0.rPr).fontSize : null;
    if (base == null) {
      const resolved = this.docStyles?.[dropBlock.style || 'Normal'] ?? this.docStyles?.['*default*'];
      if (resolved?.sizePx) base = resolved.sizePx * 72 / 96;
    }
    if (base == null) base = 11;
    const fontSize = Math.round(lines * base * 1.15 * 2) / 2;

    this.doc.setParagraphProp(dropIndex, 'dropCap', { kind, lines });
    this.doc.setParagraphRuns(dropIndex, dropBlock.runs.map((r) => ({ ...r, rPr: this.doc.setRunProp(r.rPr, 'fontSize', fontSize) })));
    this._invalidate();
    this.collapseTo({ block: dropIndex + 1, offset: 0 });
    return this;
  }

  _clearDropCap() {
    const pair = this._dropCapPairAt(this.focus.block);
    if (!pair) return this;
    const { dropIndex } = pair;
    const dropBlock = this._editable(dropIndex);
    this.doc.setParagraphProp(dropIndex, 'dropCap', null);
    // The size the drop set is an override this letter never asked for on
    // its own account; clearing it puts the run back to riding the style,
    // the same honest simplification the drop's OWN size resolution makes.
    // An rPr left empty by the clearing is no rPr at all, so the letter's
    // run reads the same as the body's and the two join back into one run
    // rather than leaving a one-letter run at the paragraph's head.
    const bare = (rPr) => (rPr && /^<w:rPr\b[^>]*>\s*<\/w:rPr>$|^<w:rPr\b[^>]*\/>$/.test(rPr) ? null : rPr);
    this.doc.setParagraphRuns(dropIndex, dropBlock.runs.map((r) => ({ ...r, rPr: bare(this.doc.setRunProp(r.rPr, 'fontSize', null)) })));
    this._invalidate();
    this.doc.mergeWithNext(dropIndex);
    this._invalidate();
    this.doc.setParagraphRuns(dropIndex, coalesce(this.block(dropIndex).runs));
    this._invalidate();
    this.collapseTo({ block: dropIndex, offset: 0 });
    return this;
  }

  /**
   * Replace every occurrence of `find` in one pass — one action, one undo.
   *
   * Single-replace needs no engine support (the client selects the match and
   * types over it, which is what those actions already do); replace-ALL does,
   * because a hundred round trips would be a hundred undo steps and a hundred
   * repaints. Matches live within one paragraph — a needle spanning a
   * paragraph break matches nothing, the same stance Word's plain search
   * takes. Structural paragraphs are skipped rather than refused: a letter
   * with one field-code line should still have its other nineteen "2025"s
   * replaced. The replacement inherits the formatting of the text it replaces.
   *
   * @returns {number} how many occurrences were replaced
   */
  replaceAll(find, replaceWith, { matchCase = false } = {}) {
    const needle = String(find ?? '');
    if (needle === '') return 0;
    const replacement = String(replaceWith ?? '');
    const target = matchCase ? needle : needle.toLowerCase();

    // Find everything FIRST: a replace-all that matches nothing must not
    // record an undo step for a document it never changed.
    const found = new Map();
    for (let i = 0; i < this.blocks.length; i++) {
      const b = this.block(i);
      if (b.structural) continue;
      const haystack = matchCase ? b.text : b.text.toLowerCase();
      const matches = [];
      for (let at = haystack.indexOf(target); at !== -1; at = haystack.indexOf(target, at + target.length)) {
        matches.push(at);
      }
      if (matches.length) found.set(i, matches);
    }
    if (!found.size) return 0;

    let count = 0;
    this._edit('replace all', null, () => {
      // Last block to first, so nothing this loop rewrites moves an index it
      // has not visited yet.
      for (const i of [...found.keys()].sort((a, z) => z - a)) {
        const b = this.block(i);
        const matches = found.get(i);
        let runs = b.runs;
        for (const at of matches.reverse()) {
          const covered = sliceRuns(runs, at, at + needle.length);
          runs = [
            ...sliceRuns(runs, 0, at),
            { rPr: covered[0]?.rPr ?? null, text: replacement },
            ...sliceRuns(runs, at + needle.length, Infinity),
          ];
          count += 1;
        }
        this.doc.setParagraphRuns(i, coalesce(runs));
        this._invalidate();
      }
      // Replaced text may have been under the caret; put it somewhere real.
      this.anchor = clampPosition(this.blocks, this.anchor);
      this.focus = clampPosition(this.blocks, this.focus);
      return this;
    });
    return count;
  }

  /**
   * Insert a fresh table after the caret's paragraph — the ribbon's Insert
   * table. The caret lands in the first cell, which in document order is
   * simply the next block: the table follows the anchor paragraph, and its
   * first cell's paragraph is the first editable thing inside it.
   *
   * Refused inside a table (growing a nested table is a decision for the day
   * somebody needs it) and on a backend without tables (an email body).
   */
  insertTable({ rows = 2, cols = 2 } = {}) {
    if (typeof this.doc.insertTable !== 'function') {
      throw new Error('this document backend does not support tables');
    }
    return this._edit('insert table', null, () => this._insertTable(rows, cols));
  }

  _insertTable(rows, cols) {
    const { block } = this.focus;
    // Existence, not editability: the table goes AFTER this paragraph, so a
    // field-code line is a perfectly good anchor — nothing of it is rebuilt.
    if (!this.block(block)) throw new Error('no paragraph at index ' + block);
    this.doc.insertTable(block, rows, cols);
    this._invalidate();
    this.collapseTo({ block: block + 1, offset: 0 });
    return this;
  }

  /**
   * Insert a picture after the caret's paragraph — the ribbon's Insert
   * picture. The image lands as its OWN paragraph (the painter has always
   * drawn images as a block under the text, so this is the honest shape),
   * and the caret moves onto it so an immediate Backspace-equivalent gesture
   * stays possible via undo.
   *
   * Refused inside a table cell: images in cell paragraphs are not painted
   * yet (a known D5 remainder), and an insert nothing shows is a dead
   * control wearing a working one's clothes.
   */
  /**
   * How a picture sits: in the line, or floating with the text wrapping round
   * it (square, tight, topAndBottom), or behind or in front of the text —
   * Word's Wrap Text menu. `hAlign` puts a floating picture at the left,
   * centre or right of the margins. The picture is named by its paragraph
   * and its index among that paragraph's pictures.
   */
  setImageLayout({ block, image = 0, wrap = 'inline', hAlign = 'left' } = {}) {
    if (typeof this.doc.setImageLayout !== 'function') {
      throw new Error('this document backend does not support floating pictures');
    }
    if (!this.block(block)) throw new Error('no paragraph at index ' + block);
    return this._edit('picture layout', null, () => {
      this.doc.setImageLayout(block, image, { wrap, hAlign });
      this._invalidate();
      return this;
    });
  }

  /** A picture's size in pixels — a corner handle dragged. */
  setImageSize({ block, image = 0, widthPx, heightPx } = {}) {
    if (typeof this.doc.setImageSize !== 'function') {
      throw new Error('this document backend does not support resizing pictures');
    }
    if (!this.block(block)) throw new Error('no paragraph at index ' + block);
    return this._edit('picture size', null, () => {
      this.doc.setImageSize(block, image, { widthPx, heightPx });
      this._invalidate();
      return this;
    });
  }

  insertImage({ name, contentType, data, widthPx, heightPx } = {}) {
    if (typeof this.doc.insertImage !== 'function') {
      throw new Error('this document backend does not support pictures');
    }
    if (this._containerOf(this.focus.block) !== null) {
      throw new Error('A picture cannot go inside a table cell yet — move the caret out of the table.');
    }
    return this._edit('insert picture', null, () => {
      const { block } = this.focus;
      if (!this.block(block)) throw new Error('no paragraph at index ' + block);
      this.doc.insertImage(block, { name, contentType, data, widthPx, heightPx });
      this._invalidate();
      this.collapseTo({ block: block + 1, offset: 0 });
      return this;
    });
  }

  /** The picked picture out of its paragraph — the paragraph too when it held nothing else. One undo step. */
  removeImage({ block, image = 0 } = {}) {
    if (typeof this.doc.removeImage !== 'function') {
      throw new Error('this document backend does not support pictures');
    }
    return this._edit('remove picture', null, () => {
      this.doc.removeImage(block, image);
      this._invalidate();
      const last = Math.max(0, this.blocks.length - 1);
      this.collapseTo({ block: Math.min(block, last), offset: 0 });
      return this;
    });
  }

  /**
   * Insert a chart AFTER the caret's table, drawn from the table's own
   * figures: the header row names the series, the first column names the
   * categories, and every other cell is a number (currency symbols and
   * thousands separators forgiven). The data lives in the chart part itself,
   * so Word shows the same chart with no workbook behind it.
   */
  insertChart({ kind = 'column', title = null } = {}) {
    if (typeof this.doc.insertChartAfterTable !== 'function') {
      throw new Error('this document backend does not support charts');
    }
    const { tableStart } = this._caretTable(); // throws its own sentence outside a table

    // The table's VISIBLE cells, gridded by their container coordinates.
    const grid = new Map(); // 'r:c' -> text
    let maxR = 0;
    let maxC = 0;
    for (const b of this.blocks) {
      if (b.hiddenCell) continue;
      const key = b.container ?? '';
      const m = new RegExp('^t' + tableStart + ':r(\\d+):c(\\d+)$').exec(key);
      if (!m) continue;
      const r = Number(m[1]);
      const c = Number(m[2]);
      const at = m[1] + ':' + m[2];
      grid.set(at, ((grid.get(at) ?? '') + ' ' + b.text).trim());
      maxR = Math.max(maxR, r);
      maxC = Math.max(maxC, c);
    }
    if (maxR < 1 || maxC < 1) {
      throw new Error('A chart needs a table with a header row, a category column and at least one number.');
    }
    const cell = (r, c) => grid.get(r + ':' + c) ?? '';
    const toNumber = (t) => {
      const cleaned = String(t).replace(/[^\d.eE+-]/g, '');
      const v = cleaned === '' ? NaN : Number(cleaned);
      return Number.isFinite(v) ? v : null;
    };

    const categories = [];
    for (let r = 1; r <= maxR; r += 1) categories.push(cell(r, 0));
    const series = [];
    for (let c = 1; c <= maxC; c += 1) {
      const values = [];
      for (let r = 1; r <= maxR; r += 1) values.push(toNumber(cell(r, c)));
      if (values.some((v) => v !== null)) {
        series.push({ name: cell(0, c) || 'Series ' + c, values });
      }
    }
    if (!series.length) {
      throw new Error('No numbers found — a chart wants the table\'s body cells to be numeric.');
    }
    // A pie or doughnut shows ONE series — Word plots only the first too, so
    // writing the rest would be data the chart silently ignores.
    const plotted = (kind === 'pie' || kind === 'doughnut') ? series.slice(0, 1) : series;

    return this._edit('insert chart', null, () => {
      this.doc.insertChartAfterTable(tableStart, { kind, title, categories, series: plotted });
      this._invalidate();
      // The caret lands on the chart's own paragraph — the first block after
      // the table it was drawn from.
      let last = -1;
      for (let i = 0; i < this.blocks.length; i += 1) {
        if ((this.blocks[i].container ?? '').startsWith('t' + tableStart + ':')) last = i;
      }
      this.anchor = clampPosition(this.blocks, this.anchor);
      this.focus = clampPosition(this.blocks, this.focus);
      if (this.block(last + 1)) this.collapseTo({ block: last + 1, offset: 0 });
      return this;
    });
  }

  /**
   * Insert a preset shape as its own paragraph after the caret's — the same
   * honest placement pictures take, painted through the drawing commons so
   * what the ribbon offered is what the page shows.
   */
  insertShape({ preset, widthPx, heightPx, fill, outline } = {}) {
    if (typeof this.doc.insertShapeParagraph !== 'function') {
      throw new Error('this document backend does not support shapes');
    }
    return this._edit('insert shape', null, () => {
      const { block } = this.focus;
      if (!this.block(block)) throw new Error('no paragraph at index ' + block);
      this.doc.insertShapeParagraph(block, { preset, widthPx, heightPx, fill, outline });
      this._invalidate();
      this.collapseTo({ block: block + 1, offset: 0 });
      return this;
    });
  }

  /**
   * Page setup — orientation, paper size, margins. One undo step; the layout
   * cache keys on the usable width, so every paragraph re-wraps to the new
   * page on the next render without any special invalidation.
   */
  /** A colour behind every page ('#RRGGBB'), or none. One undo step. */
  setPageColour(colour) {
    if (typeof this.doc.setPageColour !== 'function') throw new Error('this document backend has no page to colour');
    return this._edit('page colour', null, () => {
      this.doc.setPageColour(colour ?? null);
      this._invalidate();
      return this;
    });
  }

  /** A frame around every page — sides of { style, widthPx, colour, spacePt } — or none. One undo step. */
  setPageBorders(borders) {
    if (typeof this.doc.setPageBorders !== 'function') throw new Error('this document backend has no page to border');
    return this._edit('page borders', null, () => {
      this.doc.setPageBorders(borders ?? null);
      this._invalidate();
      return this;
    });
  }

  /** A number beside every line, down the left margin — or none. One undo step. */
  setLineNumbers(spec) {
    if (typeof this.doc.setLineNumbers !== 'function') throw new Error('this document backend has no lines to number');
    return this._edit('line numbers', null, () => {
      this.doc.setLineNumbers(spec ?? null);
      this._invalidate();
      return this;
    });
  }

  setPageSetup(spec) {
    if (typeof this.doc.setPageSetup !== 'function' || !this.section) {
      throw new Error('this document backend has no page to set up');
    }
    return this._edit('page setup', null, () => {
      this.doc.setPageSetup(spec ?? {});
      this._invalidate();
      return this;
    });
  }

  /**
   * Tab, the way a word processor means it.
   *
   * In a table it walks cells — the next cell's first paragraph, or the
   * previous cell's first paragraph with Shift — and in the LAST cell it grows
   * the table by a row, which is how every table in Word has ever been
   * extended. Outside a table it falls back to indent: one step in, or out
   * with Shift, on the caret's paragraph. Returns this either way; the no-op
   * cases (Shift+Tab in the first cell, Tab on a backend with no paragraph
   * properties) simply leave the caret where it was.
   */
  tabCell({ back = false } = {}) {
    const { block } = this.focus;
    const home = this._containerOf(block);

    if (home === null) {
      // Prose: Tab indents, Shift+Tab outdents — the paragraph, not the caret.
      if (typeof this.doc.setParagraphProp !== 'function') return this;
      return this.setParagraphFormat({ indentDelta: back ? -1 : 1 });
    }

    if (back) {
      // The previous VISIBLE cell: skip the rest of this cell, then any
      // vMerge continuations (they exist in the file, not on the page), then
      // walk to the FIRST block of whatever cell precedes.
      let i = block - 1;
      while (i >= 0 && (this._containerOf(i) === home || this.blocks[i]?.hiddenCell)) i -= 1;
      const target = i >= 0 ? this._containerOf(i) : null;
      if (target === null) return this; // the first cell — nowhere left to go
      while (i > 0 && this._containerOf(i - 1) === target) i -= 1;
      return this.collapseTo({ block: i, offset: 0 });
    }

    for (let i = block + 1; i < this.blocks.length; i++) {
      const c = this._containerOf(i);
      if (c === home) continue;
      if (this.blocks[i]?.hiddenCell) continue; // an invisible continuation
      if (c === null) break;          // walked out of the table — last cell
      return this.collapseTo({ block: i, offset: 0 });
    }

    // The last cell: Tab grows the table, as it always has.
    if (typeof this.doc.appendTableRow !== 'function') return this;
    // The innermost table's id is the last t-segment of the container key;
    // appending inside it never moves the table's own start.
    const segments = home.split(':');
    let tableAt = 0;
    for (let s = segments.length - 1; s >= 0; s--) {
      if (segments[s][0] === 't') { tableAt = s; break; }
    }
    const prefix = segments.slice(0, tableAt + 1).join(':');
    return this._edit('add table row', null, () => {
      this.doc.appendTableRow(Number(segments[tableAt].slice(1)));
      this._invalidate();
      for (let i = this.focus.block + 1; i < this.blocks.length; i++) {
        const c = this._containerOf(i);
        if (c === home) continue;
        if (this.blocks[i]?.hiddenCell) continue; // never land in a merged row's ghost
        if (c !== null && c.startsWith(prefix + ':')) {
          this.collapseTo({ block: i, offset: 0 });
          break;
        }
        break;
      }
      return this;
    });
  }

  /**
   * Where the caret's INNERMOST table is, and which of its cells the caret
   * sits in — parsed off the container key (`t<start>:r<i>:c<j>`, possibly
   * nested). The address every table restructuring below acts on.
   */
  _caretTable() {
    const key = this._containerOf(this.focus.block);
    if (!key) throw new Error('The caret is not in a table.');
    const segments = key.split(':');
    let ti = 0;
    for (let i = segments.length - 1; i >= 0; i--) {
      if (segments[i][0] === 't') { ti = i; break; }
    }
    return {
      tableStart: Number(segments[ti].slice(1)),
      rowIndex: Number(segments[ti + 1]?.slice(1) ?? 0),
      cellIndex: Number(segments[ti + 2]?.slice(1) ?? 0),
    };
  }

  /**
   * Restructure the table under the caret. `op` is one of `insertRowAbove`,
   * `insertRowBelow`, `deleteRow`, `insertColumnLeft`, `insertColumnRight`,
   * `deleteColumn`, `deleteTable`. Each is one undo step; deleting the last
   * row or column deletes the table, as Word does. A merged table refuses —
   * the engine says why.
   */
  tableOp(op, arg = {}) {
    const PORT = {
      insertRowAbove: 'insertTableRow',
      insertRowBelow: 'insertTableRow',
      deleteRow: 'deleteTableRow',
      insertColumnLeft: 'insertTableColumn',
      insertColumnRight: 'insertTableColumn',
      deleteColumn: 'deleteTableColumn',
      deleteTable: 'deleteTable',
      mergeRight: 'mergeTableCellRight',
      mergeCells: 'mergeTableCells',
      splitCell: 'splitTableCell',
      columnWidth: 'setTableColumnWidth',
    };
    const method = PORT[op];
    if (!method) throw new Error('unknown table operation: ' + op);
    if (typeof this.doc[method] !== 'function') {
      throw new Error('this document backend does not support tables');
    }
    const { tableStart, rowIndex, cellIndex } = this._caretTable();
    return this._edit('table', null, () => {
      if (op === 'insertRowAbove' || op === 'insertRowBelow') {
        this.doc.insertTableRow(tableStart, rowIndex, op === 'insertRowAbove' ? 'above' : 'below');
      } else if (op === 'deleteRow') {
        this.doc.deleteTableRow(tableStart, rowIndex);
      } else if (op === 'insertColumnLeft' || op === 'insertColumnRight') {
        this.doc.insertTableColumn(tableStart, cellIndex, op === 'insertColumnLeft' ? 'left' : 'right');
      } else if (op === 'deleteColumn') {
        this.doc.deleteTableColumn(tableStart, cellIndex);
      } else if (op === 'mergeRight') {
        this.doc.mergeTableCellRight(tableStart, rowIndex, cellIndex);
      } else if (op === 'mergeCells') {
        // The rectangle is the SELECTION's: anchor cell to focus cell, both
        // in this same innermost table.
        const parse = (key) => {
          const segments = String(key ?? '').split(':');
          let ti = -1;
          for (let i = segments.length - 1; i >= 0; i--) if (segments[i][0] === 't') { ti = i; break; }
          return ti < 0 ? null : {
            t: Number(segments[ti].slice(1)),
            r: Number(segments[ti + 1]?.slice(1) ?? 0),
            c: Number(segments[ti + 2]?.slice(1) ?? 0),
          };
        };
        const a = parse(this._containerOf(this.anchor.block));
        const f = parse(this._containerOf(this.focus.block));
        if (!a || !f || a.t !== f.t || a.t !== tableStart) {
          throw new Error('Select from one cell to another in the same table, then merge.');
        }
        this.doc.mergeTableCells(tableStart, a.r, a.c, f.r, f.c);
        this._invalidate();
        // The survivors' caret home is the merge's top-left cell.
        const prefix = 't' + tableStart + ':r' + Math.min(a.r, f.r) + ':c' + Math.min(a.c, f.c);
        // exact, or a nested table inside it — 'c1' must not match 'c11'
        const landing = this.blocks.findIndex((b) =>
          b.container === prefix || (b.container ?? '').startsWith(prefix + ':'));
        if (landing >= 0) this.collapseTo({ block: landing, offset: 0 });
        return this;
      } else if (op === 'splitCell') {
        this.doc.splitTableCell(tableStart, rowIndex, cellIndex);
      } else if (op === 'columnWidth') {
        // The UI speaks centimetres; the file speaks twips. 1cm = 567.
        const cm = Number(arg.cm);
        if (!Number.isFinite(cm)) throw new Error('a column width needs a number of centimetres');
        this.doc.setTableColumnWidth(tableStart, cellIndex, Math.round(cm * 567));
      } else {
        this.doc.deleteTable(tableStart);
      }
      this._invalidate();
      // Rows moved, cells vanished, possibly the whole table: put the caret
      // back on a block that exists.
      this.anchor = clampPosition(this.blocks, this.anchor);
      this.focus = clampPosition(this.blocks, this.focus);
      return this;
    });
  }

  /**
   * Several columns' widths at once, in twips by column index — what a
   * column border dragged on the page asks for, since the column on each
   * side of it changes. The table is named by its start offset (the `t…` of
   * a block's container), so the caret need not be in it. One undo step.
   */
  setTableColumnWidths({ table, widths } = {}) {
    if (typeof this.doc.setTableColumnWidths !== 'function') throw new Error('this document backend does not support tables');
    return this._edit('table', null, () => {
      this.doc.setTableColumnWidths(Number(table), widths);
      this._invalidate();
      return this;
    });
  }

  /** One row's height in twips (a floor), or null for its content's; one undo step. */
  setTableRowHeight({ table, row, twips } = {}) {
    if (typeof this.doc.setTableRowHeight !== 'function') throw new Error('this document backend does not support tables');
    return this._edit('table', null, () => {
      this.doc.setTableRowHeight(Number(table), Number(row), twips ?? null);
      this._invalidate();
      return this;
    });
  }

  /**
   * Replace the default header's or footer's text, line per paragraph. The
   * bands are not caret-addressable (they live in their own parts, painted
   * per page), so this is panel-shaped rather than keystroke-shaped: one
   * action, one undo — and the undo genuinely reverts the band, because the
   * part is registered for the snapshot BEFORE the edit runs.
   */
  setBand(which, lines) {
    if (typeof this.doc.setBandText !== 'function' || !this.section) {
      throw new Error('this document backend has no headers or footers');
    }
    if (typeof this.doc.registerBandUndo === 'function') this.doc.registerBandUndo(which);
    return this._edit(which, null, () => {
      this.doc.setBandText(which, lines);
      this._invalidate();
      return this;
    });
  }

  /** Faint words across every page — DRAFT, in silver, rising — or none. One undo step. */
  setWatermark(text, options = {}) {
    if (typeof this.doc.setWatermark !== 'function' || !this.section) {
      throw new Error('this document backend has no header to carry a watermark');
    }
    if (typeof this.doc.registerBandUndo === 'function') this.doc.registerBandUndo('header');
    return this._edit('watermark', null, () => {
      this.doc.setWatermark(text ?? null, options);
      this._invalidate();
      return this;
    });
  }

  /**
   * Add a comment at the caret's paragraph — a point comment, so the
   * paragraph it discusses stays editable. Allowed on structural paragraphs
   * too: nothing rebuilds them, and a locked field line is exactly what a
   * colleague wants to remark on. The comments part is registered for undo
   * BEFORE the snapshot, so undoing the very first comment leaves an empty
   * part rather than an orphaned voice.
   */
  addComment(text, { author = '' } = {}) {
    if (typeof this.doc.addComment !== 'function') {
      throw new Error('this document backend does not support comments');
    }
    if (typeof this.doc.registerCommentUndo === 'function') this.doc.registerCommentUndo();
    return this._edit('comment', null, () => {
      this.doc.addComment(this.focus.block, { author, text });
      this._invalidate();
      return this;
    });
  }

  /**
   * A footnote or endnote at the caret. The note's words go into the notes
   * part; the REFERENCE — one character, U+FFFC, the way Word counts it —
   * goes into the paragraph as a run of its own wearing the reference style.
   * The paragraph stays editable: the rebuild writes the element back from
   * the run, a caret steps over the character, Backspace removes it and the
   * note is then simply unreferenced.
   */
  insertNote(kind, text) {
    if (typeof this.doc.addNote !== 'function') throw new Error('this document backend does not support footnotes');
    if (kind !== 'footnote' && kind !== 'endnote') throw new Error('a note is a footnote or an endnote');
    if (typeof this.doc.registerNoteUndo === 'function') this.doc.registerNoteUndo(kind);
    return this._edit('note', null, () => {
      if (!this.collapsed) this.collapseTo(this.selection.to);
      const { block, offset } = this.focus;
      const b = this._editable(block);
      const id = this.doc.addNote(kind, text);
      const styleId = kind === 'footnote' ? 'FootnoteReference' : 'EndnoteReference';
      const ref = {
        rPr: '<w:rPr><w:rStyle w:val="' + styleId + '"/><w:vertAlign w:val="superscript"/></w:rPr>',
        text: '￼', noteRef: { kind, id },
        bold: false, italic: false, underline: false, strike: false,
      };
      const next = [...sliceRuns(b.runs, 0, offset), ref, ...sliceRuns(b.runs, offset, Infinity)];
      this.doc.setParagraphRuns(block, coalesce(next));
      this._invalidate();
      this.pendingFormat = null;
      this.collapseTo({ block, offset: offset + 1 });
      return this;
    });
  }

  /**
   * Insert → Equation: an equation at the caret, replacing any selection —
   * `xml` is the OMML (`m:oMathPara` for a display equation, a bare
   * `m:oMath` for one in the words). One character of the paragraph, like a
   * note reference; the caret lands after it. While Track Changes records,
   * it goes in as an insertion like any typed word.
   */
  insertEquation({ xml }) {
    if (typeof this.doc.ensureMathNamespace !== 'function') throw new Error('this document backend does not support equations');
    if (!/^<m:oMath(Para)?\b/.test(String(xml || ''))) throw new Error('an equation is an m:oMath or m:oMathPara element');
    return this._edit('equation', null, () => {
      if (!this.collapsed) this.deleteSelection();
      const { block, offset } = this.focus;
      const b = this._editable(block);
      this.doc.ensureMathNamespace();
      const run = this._mathRun(xml);
      const next = [...sliceRuns(b.runs, 0, offset), run, ...sliceRuns(b.runs, offset, Infinity)];
      this.doc.setParagraphRuns(block, coalesce(next));
      this._invalidate();
      this.pendingFormat = null;
      this.collapseTo({ block, offset: offset + 1 });
      return this;
    });
  }

  /**
   * The equation at `block`/`offset` replaced by another — what OK in the
   * equation editor does to one opened for editing. The new one is left
   * selected, as Word leaves an equation it has just built.
   */
  replaceEquation({ block, offset, xml }) {
    if (!/^<m:oMath(Para)?\b/.test(String(xml || ''))) throw new Error('an equation is an m:oMath or m:oMathPara element');
    return this._edit('equation', null, () => {
      const b = this._editable(block);
      let at = 0;
      const index = b.runs.findIndex((r) => {
        const hit = at === offset && Boolean(r.math);
        at += r.text.length;
        return hit;
      });
      if (index < 0) throw new Error('there is no equation there to change');
      const runs = b.runs.map((r, i) => (i === index ? { ...r, math: { xml, display: xml.startsWith('<m:oMathPara') } } : r));
      this.doc.setParagraphRuns(block, runs);
      this._invalidate();
      this.setSelection({ block, offset }, { block, offset: offset + 1 });
      return this;
    });
  }

  _mathRun(xml) {
    const run = {
      rPr: null, text: '￼', bold: false, italic: false, underline: false, strike: false,
      math: { xml, display: xml.startsWith('<m:oMathPara') },
    };
    return this.recording ? { ...run, ins: this._trackMeta(null) } : run;
  }

  /**
   * A paste from within the suite that carries equations: `lines`, each a
   * list of pieces — `{ text }` or `{ math: { xml } }` — one paragraph per
   * line, as `pasteText` splits on newlines. The equations come back as the
   * equations they were, not as their pictures or their linear form. One
   * undo, however much it held.
   */
  pasteRuns(lines) {
    const list = Array.isArray(lines) ? lines : [];
    return this._edit('paste', null, () => {
      if (!this.collapsed) this.deleteSelection();
      list.forEach((pieces, i) => {
        if (i > 0) this.splitParagraph();
        for (const piece of pieces || []) {
          if (piece?.math?.xml && /^<m:oMath(Para)?\b/.test(piece.math.xml) && typeof this.doc.ensureMathNamespace === 'function') {
            const { block, offset } = this.focus;
            const b = this._editable(block);
            this.doc.ensureMathNamespace();
            const next = [...sliceRuns(b.runs, 0, offset), this._mathRun(piece.math.xml), ...sliceRuns(b.runs, offset, Infinity)];
            this.doc.setParagraphRuns(block, coalesce(next));
            this._invalidate();
            this.collapseTo({ block, offset: offset + 1 });
          } else if (piece?.text) {
            this.insertText(String(piece.text).replace(/￼/g, ''));
          }
        }
      });
      return this;
    });
  }

  /** Replace a note's words — the number and the reference stay where they are. */
  setNoteText(kind, id, text) {
    if (typeof this.doc.setNoteText !== 'function') throw new Error('this document backend does not support footnotes');
    if (typeof this.doc.registerNoteUndo === 'function') this.doc.registerNoteUndo(kind);
    return this._edit('note', null, () => {
      this.doc.setNoteText(kind, id, text);
      this._invalidate();
      return this;
    });
  }

  /** Bookmark spans, read-only — the dialog's list and what Go To searches. Optional in the port. */
  bookmarks() {
    return typeof this.doc.bookmarks === 'function' ? this.doc.bookmarks() : [];
  }

  /**
   * Name the selected paragraphs — Insert > Bookmark. Ordered anchor -> focus
   * regardless of which way the selection was dragged; a collapsed caret
   * bookmarks its own paragraph. Optional in the port: an email body has no
   * bookmarks.
   */
  addBookmark(name) {
    if (typeof this.doc.addBookmark !== 'function') {
      throw new Error('this document backend does not support bookmarks');
    }
    const { from, to } = this.selection;
    return this._edit('bookmark', null, () => {
      this.doc.addBookmark(name, from.block, to.block);
      this._invalidate();
      return this;
    });
  }

  removeBookmark(name) {
    if (typeof this.doc.removeBookmark !== 'function') {
      throw new Error('this document backend does not support bookmarks');
    }
    return this._edit('bookmark', null, () => {
      this.doc.removeBookmark(name);
      this._invalidate();
      return this;
    });
  }

  /**
   * Jump to a named bookmark, selecting its whole span — Word's Go To. Moving
   * the caret is not an edit: no snapshot, no history entry. False, selection
   * untouched, when the name is not there.
   */
  gotoBookmark(name) {
    const mark = this.bookmarks().find((b) => b.name === name);
    if (!mark) return false;
    this.setSelection(
      { block: mark.from, offset: 0 },
      { block: mark.to, offset: this.blocks[mark.to]?.text.length ?? 0 },
    );
    return true;
  }

  /**
   * Insert → Cross-reference: a REF field at the caret naming a bookmark, its
   * result the bookmark's own words right now — `\h` makes it a hyperlink to
   * the bookmark, so Ctrl+click on the field goes there (`gotoBookmark`
   * again). A selection is replaced, as typing would; the field then behaves
   * like the atomic run it is (see positions.js and `_insertText` above).
   */
  insertCrossReference(name) {
    if (typeof this.doc.addBookmark !== 'function') {
      throw new Error('this document backend does not support bookmarks');
    }
    const mark = this.bookmarks().find((b) => b.name === name);
    if (!mark) throw new Error('no bookmark named ' + name);
    return this._edit('cross-reference', null, () => {
      if (!this.collapsed) this.deleteSelection();
      const { block, offset } = this.focus;
      const b = this._editable(block);
      const words = this.blocks.slice(mark.from, mark.to + 1).map((p) => p.text).join(' ');
      const field = { rPr: null, text: words, field: { instr: ' REF ' + name + ' \\h ', kind: 'ref', name } };
      const next = [...sliceRuns(b.runs, 0, offset), field, ...sliceRuns(b.runs, offset, Infinity)];
      this.doc.setParagraphRuns(block, coalesce(next));
      this._invalidate();
      this.pendingFormat = null;
      this.collapseTo({ block, offset: offset + words.length });
      return this;
    });
  }

  /**
   * Insert → Captions → Insert Caption: a new paragraph after the one the
   * caret sits in — a picture's own paragraph when the caret is in it,
   * since an inline picture is its paragraph's whole content here. `label`
   * is 'Figure' | 'Table' | 'Equation'; the running number is the engine's
   * own count, not the dialog's (see `Document#addCaption`), so a caption
   * inserted ahead of others with the same label still comes out right.
   */
  insertCaption({ label, text = '' }) {
    if (typeof this.doc.addCaption !== 'function') {
      throw new Error('this document backend does not support captions');
    }
    return this._edit('caption', null, () => {
      const { block } = this.focus;
      if (!this.block(block)) throw new Error('no paragraph at index ' + block);
      this.doc.addCaption({ label, text, at: block });
      this._invalidate();
      this.collapseTo({ block: block + 1, offset: 0 });
      return this;
    });
  }

  /**
   * References → Update Fields (F9): every REF field's words refreshed from
   * its bookmark, right now — as a person would run it before printing.
   * Returns how many fields changed, for the toast.
   */
  updateFields() {
    if (typeof this.doc.refreshRefFields !== 'function') {
      throw new Error('this document backend does not support fields');
    }
    return this._edit('update fields', null, () => {
      const n = this.doc.refreshRefFields();
      this._invalidate();
      return n;
    });
  }

  /**
   * References → Table of Contents, at the caret's paragraph: a real field
   * (see `Document#insertTableOfContents`), not the plain text this ribbon
   * button used to write. `pages`, when the caller has it, is the on-screen
   * page (1-based) of each heading this finds, in document order — pages.js
   * knows which sheet a block landed on; without it every entry's page is
   * blank until Update Table supplies one.
   */
  insertTableOfContents({ levels = 3, pages } = {}) {
    if (typeof this.doc.insertTableOfContents !== 'function') {
      throw new Error('this document backend does not support a table of contents');
    }
    return this._edit('table of contents', null, () => {
      const { block } = this.focus;
      if (!this.block(block)) throw new Error('no paragraph at index ' + block);
      this.doc.insertTableOfContents({ at: block, levels, pages });
      this._invalidate();
      return this;
    });
  }

  /**
   * References → Update Table: the entries rebuilt from the current
   * headings and their pages, the field kept where it was. Also what F9
   * does to a table of contents alongside REF/SEQ — see `refreshRefFields`.
   */
  updateTableOfContents({ pages } = {}) {
    if (typeof this.doc.updateTableOfContents !== 'function') {
      throw new Error('this document backend does not support a table of contents');
    }
    return this._edit('update table of contents', null, () => {
      this.doc.updateTableOfContents({ pages });
      this._invalidate();
      return this;
    });
  }

  /**
   * Review → Accept: this change (the paragraph the caret sits in), or
   * `{ all: true }` for every change in the document. An insertion keeps
   * its words, unwrapped; a deletion's words are finally gone.
   */
  acceptChanges({ all = false } = {}) {
    if (typeof this.doc.acceptParagraphChanges !== 'function') {
      throw new Error('this document backend does not support tracked changes');
    }
    return this._edit('accept changes', null, () => {
      const changed = all ? this.doc.acceptAllChanges() : this.doc.acceptParagraphChanges(this.focus.block);
      this._invalidate();
      return changed;
    });
  }

  /**
   * Review → Reject: the reverse — an insertion's words are gone; a
   * deletion's words come back.
   */
  rejectChanges({ all = false } = {}) {
    if (typeof this.doc.rejectParagraphChanges !== 'function') {
      throw new Error('this document backend does not support tracked changes');
    }
    return this._edit('reject changes', null, () => {
      const changed = all ? this.doc.rejectAllChanges() : this.doc.rejectParagraphChanges(this.focus.block);
      this._invalidate();
      return changed;
    });
  }

  /** References → Remove Table of Contents. True when one was there. */
  removeTableOfContents() {
    if (typeof this.doc.removeTableOfContents !== 'function') {
      throw new Error('this document backend does not support a table of contents');
    }
    return this._edit('remove table of contents', null, () => {
      const removed = this.doc.removeTableOfContents();
      this._invalidate();
      return removed;
    });
  }

  /**
   * Fill a content control — the business-data binding path, allowed inside
   * structure. Optional in the port: an email body has no named anchors.
   */
  setContentControl(tag, value) {
    if (!supportsContentControls(this.doc)) {
      throw new Error('this document format has no content controls');
    }
    return this._edit('set ' + tag, null, () => this._setContentControl(tag, value));
  }

  _setContentControl(tag, value) {
    this.doc.setContentControlText(tag, value);
    this._invalidate();
    return this;
  }
  contentControls() {
    return supportsContentControls(this.doc) ? this.doc.contentControls() : [];
  }
  get supportsContentControls() { return supportsContentControls(this.doc); }

  // ---- mailings: mail merge ----------------------------------------------
  //
  // A mail merge is a main document — this one — a recipient list, and the
  // fields that say where each recipient's words go. The document's kind
  // and the list's whereabouts are in the file (settings.xml `w:mailMerge`,
  // as Word keeps them); the list itself, which records are ticked, the
  // order it is sorted in and the record being previewed live here, on the
  // view, for as long as the document is open — the rules each field
  // follows are `@rutba/ooxml/mailmerge`'s.

  /** The merge in hand: `{ type, source, excluded, sort, overrides, preview, record }`. */
  get merge() {
    if (!this._merge) {
      const saved = typeof this.doc.mailMerge === 'function' ? this.doc.mailMerge() : null;
      this._merge = {
        type: saved?.type ?? null,
        saved,
        source: null,
        excluded: [],
        sort: null,
        overrides: {},
        preview: false,
        record: saved?.activeRecord || 1,
        email: { toField: saved?.addressField ?? null, subject: saved?.subject ?? '', sent: Boolean(saved?.addressField) },
      };
    }
    return this._merge;
  }

  /** The records a merge walks, as indices into the source, in merge order. */
  mergeOrder() {
    const m = this.merge;
    return m.source ? mergeOrder(m.source, { excluded: m.excluded, sort: m.sort }) : [];
  }

  /** Match Fields: which column stands for each address field. */
  mergeMapping() {
    const m = this.merge;
    return m.source ? matchFields(m.source.fields, m.overrides) : {};
  }

  /** What the ribbon, the dialogs and the status bar need to know, and nothing as big as the list itself. */
  mergeSummary() {
    const m = this.merge;
    const s = m.source;
    const order = this.mergeOrder();
    return {
      type: m.type,
      source: s ? { kind: s.kind, name: s.name, path: s.path ?? null, sheet: s.sheet ?? null, sheets: s.sheets ?? null, fields: s.fields, count: s.records.length } : null,
      // A saved source the window has not reattached yet (the file moved, or
      // the address book is still to be read) — so the window can say so.
      pending: !s && m.saved?.path ? { path: m.saved.path, sheet: m.saved.sheet } : !s && m.saved?.contacts ? { contacts: true } : null,
      included: order.length,
      excluded: m.excluded.length,
      sort: m.sort,
      record: Math.max(1, Math.min(order.length || 1, m.record || 1)),
      preview: Boolean(m.preview && s),
      mapping: this.mergeMapping(),
      overrides: m.overrides,
      email: m.email,
      duplicates: s ? findDuplicates(s).length : 0,
    };
  }

  /** Write the merge's settings into the file, the way Word keeps them. */
  _writeMerge() {
    if (typeof this.doc.setMailMerge !== 'function') return;
    const m = this.merge;
    if (!m.type) { this.doc.setMailMerge(null); return; }
    const src = m.source || (m.saved?.path ? { kind: /\.xlsx?$/i.test(m.saved.path) ? 'xlsx' : 'csv', path: m.saved.path, sheet: m.saved.sheet } : m.saved?.contacts ? { kind: 'contacts' } : null);
    this.doc.setMailMerge({
      type: m.type,
      source: src && (src.path || src.kind === 'contacts') ? { kind: src.kind, path: src.path ?? null, sheet: src.sheet ?? null } : null,
      destination: m.type === 'email' ? 'email' : null,
      // Kept once a merge has been sent as e-mail, whatever the document's
      // kind — Word keeps the To column and the subject line the same way.
      addressField: m.type === 'email' || m.email.sent ? m.email.toField : null,
      subject: m.type === 'email' || m.email.sent ? m.email.subject : null,
      viewMergedData: Boolean(m.preview && m.source),
      activeRecord: m.source ? this.mergeSummary().record : null,
    });
    this.touched = true;
  }

  /**
   * Start Mail Merge: what kind of main document this is — 'formLetters',
   * 'email', 'envelopes', 'mailingLabels', 'catalog' — or null, Normal Word
   * Document, which takes the merge off the file and the list with it.
   */
  startMailMerge(type) {
    const m = this.merge;
    if (!type) {
      Object.assign(m, { type: null, source: null, saved: null, excluded: [], sort: null, preview: false, record: 1 });
    } else m.type = type;
    this._writeMerge();
    this._invalidate();
    return this;
  }

  /**
   * Select Recipients: the list the merge reads — `{ kind: 'csv' | 'tsv' |
   * 'xlsx' | 'contacts', name, path, sheet, sheets, fields, records }`.
   * A document not yet started as a merge becomes a letter, as Word's own
   * Select Recipients makes it.
   */
  attachMergeSource(source, { restore = false } = {}) {
    if (!source || !Array.isArray(source.fields) || !Array.isArray(source.records)) throw new Error('a recipient list needs fields and records');
    const m = this.merge;
    if (restore) {
      // A merge document opened again: the list read back from where the
      // file says, the record and the preview as the file left them, and
      // nothing written — reopening is not an edit.
      m.source = source;
      m.record = Math.max(1, Math.min(source.records.length || 1, m.saved?.activeRecord || 1));
      m.preview = Boolean(m.saved?.viewMergedData);
      if (!m.type) m.type = m.saved?.type || 'formLetters';
      if (!m.email.toField) m.email.toField = source.fields.find((f) => /e-?mail/i.test(f)) || null;
      this._invalidate();
      this.touched = false;
      return this;
    }
    m.source = source;
    m.excluded = [];
    m.sort = null;
    m.record = 1;
    if (!m.type) m.type = 'formLetters';
    if (!m.email.toField) m.email.toField = source.fields.find((f) => /e-?mail/i.test(f)) || null;
    this._writeMerge();
    this._invalidate();
    return this;
  }

  /** Edit Recipient List: which records are left out, and the column the list is sorted on. */
  setMergeRecipients({ excluded, sort } = {}) {
    const m = this.merge;
    if (!m.source) throw new Error('select recipients first');
    if (excluded !== undefined) m.excluded = [...new Set((excluded || []).map(Number).filter((i) => i >= 0 && i < m.source.records.length))];
    if (sort !== undefined) m.sort = sort && sort.field ? { field: sort.field, descending: Boolean(sort.descending) } : null;
    m.record = Math.max(1, Math.min(this.mergeOrder().length || 1, m.record));
    this._writeMerge();
    this._invalidate();
    return this;
  }

  /** Match Fields: `{ _FIRST0_: 'Forename', … }` — a column, or '' for "(not matched)". */
  setMergeMapping(overrides = {}) {
    this.merge.overrides = { ...overrides };
    this._invalidate();
    return this;
  }

  /** The To column and the Subject line an e-mail merge sends with — kept in the file, as Word keeps them. */
  setMergeEmail({ toField, subject } = {}) {
    const m = this.merge;
    if (toField !== undefined) m.email.toField = toField || null;
    if (subject !== undefined) m.email.subject = String(subject ?? '');
    m.email.sent = true;
    this._writeMerge();
    return this;
  }

  /**
   * Preview Results: on or off, and the record shown (1-based in merge
   * order — the number in the ribbon's record box). While it is on the page
   * shows each field's value for that record in place of «its name».
   */
  setMergePreview({ on, record } = {}) {
    const m = this.merge;
    if (on !== undefined) m.preview = Boolean(on) && Boolean(m.source);
    if (record !== undefined) m.record = Math.max(1, Math.min(this.mergeOrder().length || 1, Math.round(Number(record) || 1)));
    this._writeMerge();
    this._invalidate();
    return this;
  }

  /**
   * Find Recipient: the next record in merge order, after the one shown and
   * round to the start, holding `text` in `field` (or in any field). Shows
   * it and answers its number, or 0 when no record has it.
   */
  findMergeRecipient(text, { field = null } = {}) {
    const m = this.merge;
    const order = this.mergeOrder();
    const want = String(text || '').trim().toLowerCase();
    if (!m.source || !want || !order.length) return 0;
    const col = field ? m.source.fields.indexOf(field) : -1;
    for (let k = 1; k <= order.length; k++) {
      const at = ((m.record - 1 + k) % order.length);
      const rec = m.source.records[order[at]];
      const cells = col >= 0 ? [rec[col]] : rec;
      if (cells.some((v) => String(v ?? '').toLowerCase().includes(want))) {
        this.setMergePreview({ on: true, record: at + 1 });
        return at + 1;
      }
    }
    return 0;
  }

  /**
   * Check for Errors: the MERGEFIELDs that name a column the list does not
   * have — Word's own complaint, "Invalid Merge Field" — each once.
   */
  mergeErrors() {
    const s = this.merge.source;
    if (!s) return [];
    const want = (n) => String(n || '').trim().toLowerCase().replace(/[\s_]+/g, ' ');
    const have = new Set(s.fields.map(want));
    const bad = new Set();
    for (const b of this.blocks) {
      for (const r of b.runs) {
        if (r.field?.kind === 'mergefield' && !have.has(want(r.field.name))) bad.add(r.field.name);
      }
    }
    return [...bad];
  }

  /** A merge field at the caret — one run, the caret's formatting, a selection replaced as typing would. */
  _insertMergeRun(instr, label) {
    return this._edit(label, null, () => {
      if (!this.collapsed) this.deleteSelection();
      const { block, offset } = this.focus;
      const b = this._editable(block);
      const { runIndex } = locate(b.runs, offset, 'left');
      const here = b.runs[runIndex];
      const rPr = here && !here.field && !here.noteRef && !here.noteMark && !here.math && !here.del ? here.rPr : null;
      const { kind, name } = readInstr(instr);
      const m = this.merge;
      const order = this.mergeOrder();
      const ctx = m.source && order.length ? { source: m.source, record: m.source.records[order[Math.min(order.length, m.record) - 1]], mapping: this.mergeMapping() } : null;
      const text = placeholderFor(instr, ctx);
      const run = { rPr, text, field: { instr, kind, name } };
      const next = [...sliceRuns(b.runs, 0, offset), run, ...sliceRuns(b.runs, offset, Infinity)];
      this.doc.setParagraphRuns(block, coalesce(next));
      this._invalidate();
      this.pendingFormat = null;
      this.collapseTo({ block, offset: offset + text.length });
      return this;
    });
  }

  /** Insert Merge Field: «Column», Word's `MERGEFIELD Column \* MERGEFORMAT`. */
  insertMergeField(name) {
    if (!name) throw new Error('which field?');
    return this._insertMergeRun(mergeFieldInstr(name), 'merge field');
  }

  /** Address Block: an ADDRESSBLOCK field with the dialog's choices as its switches. */
  insertAddressBlock(opts = {}) {
    return this._insertMergeRun(addressBlockInstr(opts), 'address block');
  }

  /** Greeting Line: a GREETINGLINE field — salutation, name format, punctuation, and the fallback. */
  insertGreetingLine(opts = {}) {
    return this._insertMergeRun(greetingLineInstr(opts), 'greeting line');
  }

  /**
   * Rules: 'if' (If…Then…Else), 'next' (Next Record), 'nextif', 'skipif'
   * (Skip Record If), 'mergerec' (Merge Record #), 'mergeseq' (Merge
   * Sequence #). The comparisons take `{ field, comparison, value }`, an IF
   * `then` and `otherwise` as well.
   */
  insertMergeRule(kind, opts = {}) {
    const instr = kind === 'if' ? ifInstr(opts)
      : kind === 'nextif' || kind === 'skipif' ? conditionInstr(kind, opts)
      : kind === 'next' ? ' NEXT '
      : kind === 'mergerec' ? ' MERGEREC '
      : kind === 'mergeseq' ? ' MERGESEQ '
      : null;
    if (!instr) throw new Error('no rule called ' + kind);
    return this._insertMergeRun(instr, 'rule');
  }

  /**
   * Each merge field's value for the record in preview, by run — walked in
   * document order because a «Next Record» moves every field after it on to
   * the next record, which is how a sheet of labels previews twenty-one.
   */
  _previewValues() {
    const m = this.merge;
    if (!m.preview || !m.source) return null;
    const order = this.mergeOrder();
    if (!order.length) return null;
    const mapping = this.mergeMapping();
    let pointer = Math.max(0, Math.min(order.length, m.record) - 1);
    const out = new Map();
    for (const b of this.blocks) {
      for (const r of b.runs) {
        if (!r.field || !MERGE_KINDS.has(r.field.kind)) continue;
        const index = pointer < order.length ? order[pointer] : null;
        const record = index === null ? null : m.source.records[index];
        const got = record ? evaluateField(r.field.instr, { source: m.source, record, recordNumber: index + 1, sequence: Math.min(order.length, m.record), mapping }) : { text: '' };
        if (got.next || (!record && r.field.kind === 'next')) pointer += 1;
        out.set(r, got.text);
      }
    }
    return out;
  }

  /** Finish & Merge → Edit Individual Documents (and Print Documents): the merged document's bytes. */
  mergeToDocument({ range = 'all' } = {}) {
    const m = this.merge;
    if (!m.source) throw new Error('Select recipients first — Mailings → Select Recipients.');
    const order = rangeOfOrder(this.mergeOrder(), range, m.record);
    if (!order.length) throw new Error('No recipients are ticked in the list.');
    return this.doc.mergeToDocument(m.source, order, { type: m.type || 'formLetters', mapping: this.mergeMapping() });
  }

  /** Finish & Merge → Send E-mail Messages: a message per record. */
  mergeMessages({ range = 'all', toField, subject, format = 'html' } = {}) {
    const m = this.merge;
    if (!m.source) throw new Error('Select recipients first — Mailings → Select Recipients.');
    const order = rangeOfOrder(this.mergeOrder(), range, m.record);
    return this.doc.mergeMessages(m.source, order, { toField: toField ?? m.email.toField, subject: subject ?? m.email.subject, format, mapping: this.mergeMapping() });
  }

  // ---- mailings: envelopes and labels --------------------------------------

  /**
   * The styles part was written — Envelope Address and Envelope Return added —
   * so the copies read once for the page and the paginator are stale, and the
   * window is told to take the whole model, styles and all, next time.
   */
  _stylesWritten() {
    this._docStyles = undefined;
    this._styleCatalogue = undefined;
    this.stylesChanged = true;
  }

  /** The envelope in front of the letter, as the page and the dialog need it — or null. */
  envelopeSummary() {
    if (typeof this.doc.envelope !== 'function' || (this.doc.sectionCount?.() ?? 1) < 2) return null;
    const e = this.doc.envelope();
    if (!e) return null;
    const s = this.doc.sections()[0];
    return { ...e, margins: s.margins, orientation: s.orientation };
  }

  /** Envelopes → Add to Document / Change Document. One undo. */
  addEnvelope(spec = {}) {
    if (typeof this.doc.addEnvelope !== 'function') throw new Error('this document backend does not support envelopes');
    return this._edit('envelope', null, () => {
      this.doc.addEnvelope(spec);
      this._stylesWritten();
      this._invalidate();
      this.collapseTo({ block: 0, offset: 0 });
      return this;
    });
  }

  /**
   * The whole document as a sheet of labels (Labels → New Document, or Start
   * Mail Merge → Labels with `mode: 'merge'`). One undo.
   */
  setLabelSheet(spec = {}) {
    if (typeof this.doc.setLabelSheet !== 'function') throw new Error('this document backend does not support labels');
    return this._edit('labels', null, () => {
      this.doc.setLabelSheet(spec);
      if (spec.mode === 'merge') this.merge.type = 'mailingLabels';
      this._invalidate();
      this.collapseTo({ block: 0, offset: 0 });
      if (spec.mode === 'merge') this._writeMerge();
      return this;
    });
  }

  /**
   * Start Mail Merge → Envelopes: the document becomes one envelope — its
   * paper, the return address, and an empty delivery address frame for
   * the Address Block. One undo.
   */
  setEnvelopeDocument(spec = {}) {
    if (typeof this.doc.setEnvelopeDocument !== 'function') throw new Error('this document backend does not support envelopes');
    return this._edit('envelope', null, () => {
      this.doc.setEnvelopeDocument(spec);
      this._stylesWritten();
      this.merge.type = 'envelopes';
      this._invalidate();
      // The caret in the delivery address frame, where the Address Block goes.
      const at = this.blocks.findIndex((b) => b.frame);
      this.collapseTo({ block: Math.max(0, at), offset: 0 });
      this._writeMerge();
      return this;
    });
  }

  /** Mailings → Update Labels: the first label's fields copied to every other. Answers how many. */
  updateLabels() {
    if (typeof this.doc.updateLabels !== 'function') throw new Error('this document backend does not support labels');
    return this._edit('update labels', null, () => {
      const n = this.doc.updateLabels();
      this._invalidate();
      return n;
    });
  }

  // ---- rendering ---------------------------------------------------------

  /**
   * One run, serialised for a painter. The toggles are always present; the
   * value properties (family, size, colour, highlight) appear only when set,
   * so a plain document's frame carries no null ballast. The painter reads
   * these rather than `rPr` — the frame stays format-free.
   */
  _renderRun(r, { toc = false } = {}) {
    const out = { text: r.text, bold: r.bold, italic: r.italic, underline: r.underline };
    if (r.strike) out.strike = true;
    // A footnote/endnote reference carries its NUMBER — assigned by `_notes`
    // from where the reference falls in the body — and the mark at the head
    // of a note carries the kind, numbered by `_notes` too.
    if (r.noteRef) out.noteRef = { ...r.noteRef, n: this._noteNumbers?.get(r) ?? null };
    if (r.noteMark) out.noteMark = r.noteMark;
    // A field's own code and kind — REF Summary, PAGE, DATE — so the page can
    // shade it and Ctrl+click can follow a REF to its bookmark.
    if (r.field) out.field = r.field;
    // Preview Results: the record's value where the field's «name» was.
    if (r.field && this._mergeValues?.has(r)) {
      out.text = this._mergeValues.get(r);
      out.merged = true;
    }
    // An equation: the MathML the page draws, the linear form the editor
    // opens with, and the OMML itself — what a copy within the suite
    // carries, so a paste puts back the same equation, not its picture.
    if (r.math) {
      const shown = typeof this.doc.mathView === 'function' ? this.doc.mathView(r.math) : null;
      out.math = { display: Boolean(r.math.display), xml: r.math.xml, ...(shown || {}) };
    }
    // A tracked change — who, when, and (a deletion only) the words it took.
    // The painter reads these for the underline/strike-through and the
    // change bar; Accept/Reject and the Reviewing Pane read them too.
    if (r.ins) out.ins = r.ins;
    if (r.del) out.del = r.del;
    // The run's CHARACTER style — Hyperlink, FootnoteReference, Strong —
    // fills in what the run does not set itself. Word paints a hyperlink
    // blue and underlined only because its style says so; a TOC entry that
    // links to a heading without the style is plain, and stays plain here.
    const rStyle = r.rPr ? /<w:rStyle\b[^>]*\bw:val="([^"]*)"/.exec(r.rPr) : null;
    // Inside a table of contents Word writes the Hyperlink style on every
    // entry and then draws the entries in the TOC style, black — the link is
    // there for Ctrl+click, not for colour. So the one style is skipped there.
    const cs = rStyle && !(toc && rStyle[1] === 'Hyperlink') ? this._charStyles?.[rStyle[1]] : null;
    if (cs) {
      if (cs.bold && !r.bold) out.bold = true;
      if (cs.italic && !r.italic) out.italic = true;
      if (cs.underline && !r.underline) out.underline = true;
      if (out.fontName == null && cs.fontName) out.fontName = cs.fontName;
      if (out.fontSize == null && cs.sizePx) out.fontSize = Math.round((cs.sizePx * 72) / 96 * 2) / 2;
      if (out.fontColour == null && cs.colour) out.fontColour = cs.colour.slice(1).toUpperCase();
      if (out.vertAlign == null && cs.vertAlign) out.vertAlign = cs.vertAlign;
      if (cs.caps && !out.caps) out.caps = true;
      if (cs.smallCaps && !out.smallCaps) out.smallCaps = true;
    }
    if (typeof this.doc.readRunProps === 'function' && r.rPr) {
      const props = this.doc.readRunProps(r.rPr);
      if (props.fontName != null) out.fontName = props.fontName;
      if (props.fontSize != null) out.fontSize = props.fontSize;
      if (props.fontColour != null) out.fontColour = props.fontColour;
      if (props.highlight != null) out.highlight = props.highlight;
      if (props.vertAlign != null) out.vertAlign = props.vertAlign;
      if (props.caps) out.caps = true;
      if (props.smallCaps) out.smallCaps = true;
      // The three text effects: outline and shadow are on/off, glow carries
      // its own colour and radius. Left off the frame when unset, like the
      // other value properties above.
      if (props.outline) out.outline = true;
      if (props.shadow) out.shadow = true;
      if (props.glow) out.glow = props.glow;
    }
    // The painter gets the TARGET, never the token: the frame stays
    // format-free, and a dangling id degrades to plain text.
    if (r.link != null && typeof this.doc.linkTarget === 'function') {
      const href = this.doc.linkTarget(r.link);
      if (href) out.link = href;
    }
    return out;
  }

  /**
   * A DISPLAYED paragraph — inside a text box, a footnote — shaped like a
   * block so the painter draws it with the block code, minus the index and
   * the address a caret would need.
   */
  _liteBlock(p) {
    const PX_PER_STEP = INDENT_STEP * (96 / 1440);
    return {
      style: p.style, align: p.align ?? null,
      indentLevel: p.indentPx ? Math.round(p.indentPx / PX_PER_STEP) : 0,
      indentPx: p.indentPx ?? null,
      ...(p.decor || {}),
      lineSpacing: p.spacing?.lineFactor ?? null,
      lineHeightPx: p.spacing?.lineExactPx ?? null,
      spaceBeforePx: p.spacing?.beforePx ?? null,
      spaceAfterPx: p.spacing?.afterPx ?? null,
      text: p.text,
      runs: p.runs.map((r) => this._renderRun(r)),
      images: (p.images || []).filter((img) => img.href),
    };
  }

  /**
   * Footnotes and endnotes as the page shows them: numbered by where each
   * reference falls in the body — the order Word numbers them, whatever the
   * ids say — each with its paragraphs shaped like blocks. Fills
   * `_noteNumbers`, which `_renderRun` reads to put the number on the
   * reference in the body, so it runs before the blocks are mapped.
   */
  _notes() {
    this._noteNumbers = new Map();
    const out = { footnotes: [], endnotes: [] };
    const source = typeof this.doc.notes === 'function' ? this.doc.notes() : null;
    if (!source) return out;
    const byId = {
      footnote: new Map((source.footnotes || []).map((n) => [n.id, n])),
      endnote: new Map((source.endnotes || []).map((n) => [n.id, n])),
    };
    const counters = { footnote: 0, endnote: 0 };
    for (const b of this.blocks) {
      for (const r of b.runs) {
        if (!r.noteRef) continue;
        const kind = r.noteRef.kind;
        const n = ++counters[kind];
        this._noteNumbers.set(r, n);
        const note = byId[kind].get(r.noteRef.id);
        if (!note) continue;
        const paragraphs = note.paragraphs.map((p) => this._liteBlock(p));
        for (const p of paragraphs) for (const run of p.runs) if (run.noteMark) run.noteMark = { kind: run.noteMark, n };
        out[kind + 's'].push({ n, id: note.id, paragraphs });
      }
    }
    return out;
  }

  /**
   * The frame a painter draws from.
   * @param {{ pages?: boolean }} [opts] `pages: false` skips pagination — the
   *   editor draws from blocks and never reads the sheets.
   */
  render({ pages: withPages = true } = {}) {
    const { from, to } = this.selection;
    // Paragraph properties reach the painter here, once per block, read from
    // what the block already carries — the backend parsed `align`, `indentPx`
    // and `spacing` when it built the block. The first version asked the
    // backend again, per block, per frame; that call re-reads the paragraph's
    // XML and cost seven milliseconds each, which on a three-thousand-paragraph
    // specification was twenty seconds per keystroke.
    const PX_PER_STEP = INDENT_STEP * (96 / 1440);
    // The styles first — the runs below fold their character style in — and
    // the notes next: the blocks read the numbers off the references.
    this._loadDefinitions();
    const notes = this._notes();
    // Preview Results reads every merge field's value first, in document
    // order — a «Next Record» moves the fields after it to the next record.
    this._mergeValues = this._previewValues();
    return {
      ...notes,
      blocks: contextualSpacing(this.blocks, this._docStyles, this.blocks.map((b) => {

        return {
        index: b.index,
        style: b.style,
        align: b.align ?? null,
        indentLevel: b.indentPx ? Math.round(b.indentPx / PX_PER_STEP) : 0,
        // The indent to the pixel, null when the paragraph sets none — an
        // explicit zero beats the style's indent, and a step count cannot say so.
        indentPx: b.indentPx ?? null,
        // What the OOXML reader found beyond alignment and the left indent:
        // tab stops, shading, borders, the first-line/hanging/right indents.
        // Spread flat so the painter reads `block.tabs`, not `block.decor.tabs`.
        ...(b.decor || {}),
        lineSpacing: b.spacing?.lineFactor ?? null,
        // Direct spacing beats the style's: an exact line, a gap before or after.
        lineHeightPx: b.spacing?.lineExactPx ?? null,
        spaceBeforePx: b.spacing?.beforePx ?? null,
        spaceAfterPx: b.spacing?.afterPx ?? null,
        structural: b.structural,
        structuralTags: b.structuralTags,
        // Where the pages fall on screen is the shell's to decide, from what
        // the author said: an explicit break, a heading kept with what
        // follows it, a paragraph kept in one piece.
        ...(b.pageBreakBefore ? { pageBreakBefore: true } : {}),
        ...(b.keepNext ? { keepNext: true } : {}),
        ...(b.keepLines ? { keepLines: true } : {}),
        ...(b.dropCap ? { dropCap: b.dropCap } : {}),
        ...(b.inSdt ? { inSdt: true } : {}),
        // Text boxes anchored here, their paragraphs shaped like blocks so the
        // painter draws them with the same code — read-only, no index.
        ...(b.textBoxes ? {
          textBoxes: b.textBoxes.map((box) => ({ ...box, paragraphs: box.paragraphs.map((p) => this._liteBlock(p)) })),
        } : {}),
        // The cell this block lives in, or null for prose — what lets the
        // shell tell a caret at a cell boundary why Tab and Backspace behave.
        container: b.container ?? null,
        // A top-level table's grid and width ride every paragraph in it, a
        // merged cell its span, a sized row its height — the page draws the
        // columns from these and the ruler moves them.
        ...(b.gridPx ? { gridPx: b.gridPx, tableWidth: b.tableWidth ?? null } : {}),
        ...(b.cellSpan ? { cellSpan: b.cellSpan } : {}),
        ...(b.rowHeightPx ? { rowHeightPx: b.rowHeightPx, rowRule: b.rowRule ?? null } : {}),
        ...(b.tableLook ? { tableLook: b.tableLook } : {}),
        ...(b.cellVAlign ? { cellVAlign: b.cellVAlign } : {}),
        // A paragraph in a frame placed on the page — drawn there.
        ...(b.frame ? { frame: b.frame } : {}),
        ...(() => {
          const runs = b.runs.map((r) => this._renderRun(r, { toc: /^TOC\d/i.test(b.style || '') }));
          // A previewed paragraph's words are the record's, so its text is too.
          const previewed = this._mergeValues && b.runs.some((r) => this._mergeValues.has(r));
          return { text: previewed ? runs.map((r) => r.text).join('') : b.text, runs };
        })(),
        ...(b.tracked ? { tracked: b.tracked } : {}),
        // Every paragraph's pictures — and its charts and shapes, which ride
        // the same pipeline as SVG — paint from here. This used to be cell
        // paragraphs only, on the theory that the paged body gets images
        // through fragments; the editor draws from these blocks, not from
        // fragments, so an inserted picture landed in the file and never on
        // the page.
        ...(b.images?.some((i) => i.href) ? { images: b.images.filter((i) => i.href) } : {}),
        };
      })),
      selection: {

        anchor: { ...this.anchor },
        focus: { ...this.focus },
        from: { ...from },
        to: { ...to },
        collapsed: this.collapsed,
      },
      // Document order, so a table renders where the author put it. Paragraph
      // entries point back into `blocks` — the caret's address space is untouched.
      flow: this.flow,
      // A page has edges; an email body does not. null means "continuous flow".
      section: this.section,
      // The flow laid onto sheets — unless the caller said it does not want it.
      // The editor draws from `blocks` and never reads this, and paginating a
      // three-thousand-paragraph document on every keystroke made typing into a
      // long specification take seconds per character. The PDF writer and the
      // paged tests ask for it; the document service does not.
      pages: withPages ? this.pages : null,
      format: this.formatAtCaret(),
      // The catalogue the Style dropdown offers. Once per load, not per caret —
      // see the getter.
      paragraphStyles: this.paragraphStyles,
      // The same styles RESOLVED — font, size, colour, spacing per style id,
      // chains flattened, `*default*` for an unstyled paragraph. The page
      // paints from these; without them every document is the app's font.
      styles: this.docStyles,
      // The default header and footer, for the ribbon's editing panel:
      // existence, lines, and whether text editing would flatten a field.
      bands: typeof this.doc.bandInfo === 'function' ? this.doc.bandInfo() : null,
      // Comments, read-only, for the margin markers and the Review listing.
      comments: typeof this.doc.comments === 'function' ? this.doc.comments() : [],
      // Bookmark spans, read-only, for the Bookmark dialog's list.
      bookmarks: this.bookmarks(),
      // Every field, read-only — the Caption dialog previews "Figure 3"
      // from this before OK is even pressed, counting the SEQ fields this
      // label already has.
      fields: typeof this.doc.fields === 'function' ? this.doc.fields() : [],
      // The table of contents, read back for the ribbon (Update Table and
      // Remove Table of Contents only make sense once one exists) and the
      // window — its entries live inside an opaque content control, out of
      // `fields`' reach.
      tableOfContents: typeof this.doc.tableOfContents === 'function' ? this.doc.tableOfContents() : null,
      // Mailings: the merge in hand — its kind, the list's name and fields,
      // how many are ticked, the record in preview. The records themselves
      // are asked for when Edit Recipient List opens, not sent every frame.
      mailMerge: this.mergeSummary(),
      // How many sections — a merged letter has one per record.
      sectionCount: typeof this.doc.sectionCount === 'function' ? this.doc.sectionCount() : 1,
      // An envelope in front of the letter: its page, and the last paragraph
      // of its section — the page draws that sheet at the envelope's size.
      envelope: this.envelopeSummary(),
      // Review → Track Changes: is this document recording, right now — read
      // from the file's own setting so a reopened file with it already on
      // shows the ribbon pressed without the window having to ask first.
      trackRevisions: typeof this.doc.trackRevisions === 'function' ? this.doc.trackRevisions() : false,
      // Every list label by block index — how a TABLE CELL's list items get
      // their bullets and numbers, since cells have no fragments to carry
      // one. Computed on the same counters pagination used, so the body and
      // the cells never disagree about which item is third.
      listLabels: (() => {
        if (!this._numberingDefs) return null;
        const labels = computeListLabels(this.flow, this.blocks, this._numberingDefs);
        return labels.size ? Object.fromEntries(labels) : null;
      })(),
      canEdit: this.canEdit,
      // What the toolbar needs: whether the buttons are live and what they say.
      history: this.history.describe(),
      dirty: this.isDirty,
      wordCount: this.blocks.reduce((n, b) => n + (b.text.trim() ? b.text.trim().split(/\s+/).length : 0), 0),
      characterCount: this.blocks.reduce((n, b) => n + b.text.length, 0),
    };
  }

  save() {
    const out = this.doc.save();
    this.touched = false;
    return out;
  }

  /**
   * The same bytes, without claiming the work is saved — for a background
   * draft flush that must not darken anybody's unsaved indicator.
   */
  serialize() {
    const wasTouched = this.touched;
    const out = this.save();
    this.touched = wasTouched;
    return out;
  }

  modifiedParts() { return this.doc.modifiedParts?.() ?? []; }
}

export { runsText, locate, comparePositions, orderedRange, sliceRuns, removeRange, coalesce };
