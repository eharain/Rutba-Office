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
import { computeListLabels } from './lists.js';
import { bandForPage, resolveFields } from './bands.js';
import {
  runsText, locate, clampPosition, comparePositions, orderedRange,
  sliceRuns, removeRange, coalesce, samePosition,
} from './positions.js';

/**
 * One indent step, in twips (1/20 point) — half an inch, what Tab-to-indent and
 * the ribbon's indent buttons move by. A plain unit constant, not format
 * knowledge: the backend is the only place that knows a twip becomes `<w:ind>`.
 */
const INDENT_STEP = 720;

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
    if (this._docStyles === undefined) {
      this._docStyles = typeof this.doc.paragraphStyles === 'function' ? this.doc.paragraphStyles() : null;
      this._numberingDefs = typeof this.doc.numberingDefs === 'function' ? this.doc.numberingDefs() : null;
    }
    const listLabels = computeListLabels(this.flow, this.blocks, this._numberingDefs);
    const laid = paginate({
      flow: this.flow, blocks: this.blocks, section,
      cache: this._lineCache, styles: this._docStyles, listLabels,
    });
    if (!laid) return (this._pages = null);

    const bands = typeof this.doc.headerFooters === 'function'
      ? this.doc.headerFooters()
      : { headers: {}, footers: {} };
    const opts = { titlePage: section.titlePage, evenAndOdd: section.evenAndOdd };

    for (const page of laid.pages) {
      const header = bandForPage(bands.headers, page.number, opts);
      const footer = bandForPage(bands.footers, page.number, opts);
      // Fields resolve PER PAGE — the whole point of a PAGE field is that it
      // says something different on each sheet.
      page.header = header ? resolveFields(header.paragraphs, { page: page.number, of: laid.count }) : null;
      page.footer = footer ? resolveFields(footer.paragraphs, { page: page.number, of: laid.count }) : null;
    }
    return (this._pages = laid);
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

    let rPr = target.rPr;
    if (this.pendingFormat) rPr = this._applyPending(rPr);

    let next;
    if (rPr === target.rPr) {
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

    if (from.block === to.block) {
      const b = this.block(from.block);
      this.doc.setParagraphRuns(from.block, coalesce(removeRange(b.runs, from.offset, to.offset)));
      this._invalidate();
      this.collapseTo(from);
      return this;
    }

    // Keep the head of the first paragraph and the tail of the last, drop what
    // is between, then merge the two survivors into one.
    const first = this.block(from.block);
    const last = this.block(to.block);
    const head = sliceRuns(first.runs, 0, from.offset);
    const tail = sliceRuns(last.runs, to.offset, Infinity);

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
    for (const key of ['fontName', 'fontSize', 'fontColour', 'highlight']) {
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
    const hasAlignOrIndent = ('align' in delta) || Boolean(delta.indentDelta);
    const hasStyle = ('styleId' in delta);
    const hasList = ('list' in delta);
    const hasSpacing = ('lineSpacing' in delta) || ('spaceBefore' in delta) || ('spaceAfter' in delta);
    if ((hasAlignOrIndent || hasStyle || hasSpacing) && typeof this.doc.setParagraphProp !== 'function') {
      throw new Error('this document backend does not support paragraph formatting');
    }
    if (hasList && typeof this.doc.setParagraphList !== 'function') {
      throw new Error('this document backend does not support list formatting');
    }
    if (!hasAlignOrIndent && !hasList && !hasStyle && !hasSpacing) return this;
    return this._edit('paragraph', null, () => this._setParagraphFormat(delta));
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
        const current = this.doc.getParagraphProps(i)?.indentTwips ?? 0;
        this.doc.setParagraphProp(i, 'indentTwips', Math.max(0, current + delta.indentDelta * INDENT_STEP));
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
    // The caret paragraph's NAMED style id, or null for the default — what the
    // ribbon's Style dropdown shows as selected.
    base.paragraphStyle = null;
    // Spacing is paragraph state too — what the ribbon's spacing menu ticks.
    base.lineSpacing = null;
    base.spaceBefore = null;
    base.spaceAfter = null;
    if (b && typeof this.doc.getParagraphProps === 'function') {
      const pp = this.doc.getParagraphProps(block);
      if (pp) {
        base.paragraphAlign = pp.align ?? 'left';
        base.indentLevel = pp.indentTwips ? Math.round(pp.indentTwips / INDENT_STEP) : 0;
        base.listType = pp.listType ?? null;
        base.paragraphStyle = pp.style ?? null;
        base.lineSpacing = pp.lineSpacing ?? null;
        base.spaceBefore = pp.spaceBeforePts ?? null;
        base.spaceAfter = pp.spaceAfterPts ?? null;
      }
    }
    if (this.pendingFormat) {
      if (this.pendingFormat.clear) {
        // The next character types plain: the toolbar must not keep showing
        // the formatting the eraser just disarmed.
        Object.assign(base, {
          bold: false, italic: false, underline: false, strike: false,
          fontName: null, fontSize: null, fontColour: null, highlight: null,
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

  // ---- rendering ---------------------------------------------------------

  /**
   * One run, serialised for a painter. The toggles are always present; the
   * value properties (family, size, colour, highlight) appear only when set,
   * so a plain document's frame carries no null ballast. The painter reads
   * these rather than `rPr` — the frame stays format-free.
   */
  _renderRun(r) {
    const out = { text: r.text, bold: r.bold, italic: r.italic, underline: r.underline };
    if (r.strike) out.strike = true;
    if (typeof this.doc.readRunProps === 'function' && r.rPr) {
      const props = this.doc.readRunProps(r.rPr);
      if (props.fontName != null) out.fontName = props.fontName;
      if (props.fontSize != null) out.fontSize = props.fontSize;
      if (props.fontColour != null) out.fontColour = props.fontColour;
      if (props.highlight != null) out.highlight = props.highlight;
    }
    // The painter gets the TARGET, never the token: the frame stays
    // format-free, and a dangling id degrades to plain text.
    if (r.link != null && typeof this.doc.linkTarget === 'function') {
      const href = this.doc.linkTarget(r.link);
      if (href) out.link = href;
    }
    return out;
  }

  render() {
    const { from, to } = this.selection;
    // Paragraph properties reach the painter here, once per block. The frame
    // carried the style name and nothing else, so a paragraph the engine had
    // centred or indented drew exactly as it had before — the button worked,
    // the file was right, and the page showed nothing.
    const props = typeof this.doc.getParagraphProps === 'function'
      ? (i) => { try { return this.doc.getParagraphProps(i); } catch { return null; } }
      : () => null;
    return {
      blocks: this.blocks.map((b) => {
        const pp = props(b.index);
        return {
        index: b.index,
        style: b.style,
        align: pp?.align ?? null,
        indentLevel: pp?.indentTwips ? Math.round(pp.indentTwips / INDENT_STEP) : 0,
        lineSpacing: pp?.lineSpacing ?? null,
        structural: b.structural,
        structuralTags: b.structuralTags,
        // The cell this block lives in, or null for prose — what lets the
        // shell tell a caret at a cell boundary why Tab and Backspace behave.
        container: b.container ?? null,
        text: b.text,
        runs: b.runs.map((r) => this._renderRun(r)),
        ...(b.tracked ? { tracked: b.tracked } : {}),
        // Every paragraph's pictures — and its charts and shapes, which ride
        // the same pipeline as SVG — paint from here. This used to be cell
        // paragraphs only, on the theory that the paged body gets images
        // through fragments; the editor draws from these blocks, not from
        // fragments, so an inserted picture landed in the file and never on
        // the page.
        ...(b.images?.some((i) => i.href) ? { images: b.images.filter((i) => i.href) } : {}),
        };
      }),
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
      // The flow laid onto sheets. null means continuous flow — render `flow`.
      pages: this.pages,
      format: this.formatAtCaret(),
      // The catalogue the Style dropdown offers. Once per load, not per caret —
      // see the getter.
      paragraphStyles: this.paragraphStyles,
      // The default header and footer, for the ribbon's editing panel:
      // existence, lines, and whether text editing would flatten a field.
      bands: typeof this.doc.bandInfo === 'function' ? this.doc.bandInfo() : null,
      // Comments, read-only, for the margin markers and the Review listing.
      comments: typeof this.doc.comments === 'function' ? this.doc.comments() : [],
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
