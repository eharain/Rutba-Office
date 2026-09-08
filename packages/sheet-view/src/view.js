/**
 * SheetView — a real workbook made renderable.
 *
 * Everything the browser shell needs and nothing it does not: the visible cells
 * with their display text, alignment and colour; the selection; the edit state;
 * and the geometry to place them. No DOM here, which is what lets the hard parts
 * be tested in Node and keeps the shell thin enough to be obviously correct.
 *
 * Two states per cell, and conflating them is the classic spreadsheet bug:
 *   - the DISPLAY value: calculated, formatted, what the grid shows
 *   - the EDIT value: what the user typed, what the formula bar shows and what
 *     goes back into the cell when they start editing
 * `=B2*1.17` must never appear in the grid, and `1,661.40` must never be written
 * back into the file.
 *
 * The view owns a calculation model and a workbook. Edits go to the model first
 * (so dependents update immediately), and to the file only on save — which keeps
 * typing fast and keeps the preserving write path on one well-tested road.
 */
import { Workbook, OoxmlPackage } from '@rutba/ooxml';
import { toSpreadsheet, writeCachedValue, parseDefinedNameRange } from '@rutba/ooxml/recalc';
import {
  readPivots, computePivot, updatePivotLocation, dataFieldLabel, areaRef,
  createPivot, planPivot,
} from '@rutba/ooxml/pivot';
import {
  isError, shiftFormula, calculate, parse, compareValues, serialToDate, dateToSerial,
} from '@rutba/formula';
import { formatValue, BUILTIN_FORMATS, isDateFormat } from './numfmt.js';
import { applyFormat, formatOf, ensureDxf } from './styles-write.js';
import {
  readStyles, readMergedCells, readDataValidations, readConditionalFormatting, unescXml,
  applyTint,
} from './styles.js';
import {
  SheetGeometry, MAX_ROWS, MAX_COLS,
  pixelsToCharWidth, pixelsToPoints, MIN_COL_WIDTH_PX, MIN_ROW_HEIGHT_PX,
} from './geometry.js';
import { Selection, ref, colName } from './selection.js';
import {
  readSheetDrawings, buildChart, buildShape, buildPicture, renderSvg, scene,
  SUPPORTED_GEOMETRY,
} from '@rutba/drawing';
import { drawingAnchorXml, chartPartXml } from '@rutba/ooxml/build';
import { History } from '@rutba/editing';

/**
 * The pseudo-name the sheet-level autofilter answers to in the filter UI.
 * A ListObject's name must start with a letter or underscore, so no real
 * table can collide with it.
 */
export const SHEET_FILTER = '#sheet';

/**
 * The chart kinds the shared writer takes — Studio's vocabulary, the
 * convergence the roadmap asked for. Lock-step with `chartPartXml`'s kind
 * branch; the ribbon offers exactly this list.
 */
export const CHART_KINDS = ['column', 'bar', 'line', 'area', 'pie', 'doughnut'];

export class SheetView {
  constructor(buf, { now, viewportWidth = 900, viewportHeight = 500, mode = 'light' } = {}) {
    this.mode = mode;
    // The clock is kept so a structural edit can rebuild the calc model with
    // the same time source the constructor used.
    this._now = now;
    this.workbook = Workbook.open(buf);
    this.pkg = this.workbook.pkg;
    const { sheet } = toSpreadsheet(this.workbook, { now });
    this.calc = sheet;
    this.calc.recalculate();

    this.styles = readStyles(this.pkg, { BUILTIN_FORMATS });
    /** @type {Map<string, SheetGeometry>} */
    this.geometry = new Map();
    /** @type {Map<string, Map<string, number>>} sheet -> cellRef -> cellXfs index */
    this.cellStyles = new Map();
    /** Cells whose style index changed and must be written on save. */
    this.styledCells = new Set();
    /** Whether styles.xml itself has been rewritten this session. */
    this._stylesDirty = false;
    /**
     * Whether a STRUCTURAL edit (a merge or unmerge) has changed a sheet part.
     * Cell and style edits announce themselves through `dirtyCells`/`styledCells`,
     * but merging empty cells touches neither — yet the file has still changed
     * and the Save button must say so.
     */
    this._structuralDirty = false;
    /** @type {Map<string, Array<object>>} sheet -> merged ranges */
    this.merges = new Map();
    /** @type {Map<string, Array<object>>} sheet -> data-validation rules */
    this.validations = new Map();
    /** @type {Map<string, Array<object>>} sheet -> conditional-formatting rules */
    this.conditionals = new Map();

    for (const { name, part } of this.workbook.sheets()) {
      const xml = this.pkg.text(part);
      this.geometry.set(name, SheetGeometry.fromSheetXml(xml));
      this.cellStyles.set(name, this._readCellStyles(xml));
      this.merges.set(name, readMergedCells(xml));
      this.validations.set(name, readDataValidations(xml));
      this.conditionals.set(name, readConditionalFormatting(xml, this.styles.theme));
    }

    this.drawings = new Map();
    for (const { name, part } of this.workbook.sheets()) {
      this.drawings.set(name, this._readDrawings(part));
    }

    /** @type {Map<string, Map<string, {author: string|null, text: string}>>} sheet -> ref -> note */
    this.comments = new Map();
    for (const { name, part } of this.workbook.sheets()) {
      this.comments.set(name, this._readComments(part));
    }

    this.activeSheet = this.workbook.sheetNames()[0];
    this.selection = Selection.at(0, 0);
    this.scrollX = 0;
    this.scrollY = 0;
    this.viewportWidth = viewportWidth;
    this.viewportHeight = viewportHeight;
    this.editing = null;
    this.dirtyCells = new Set();
    this.history = new History();
    this._editDepth = 0;
    /**
     * The in-sheet clipboard mark: what range was last copied and the text it
     * rendered as. A paste whose text still MATCHES pastes rich — formatting
     * travels, formulas shift; one that does not (copied elsewhere, or the
     * source has since changed) falls back to plain text, which is the honest
     * reading of a clipboard we no longer own. Null, not undefined: this is
     * per-person cursor state and an empty one must restore as empty.
     */
    this.clipboard = null;
    /** The format painter's loaded brush: which range's LOOK was picked up. */
    this.formatBrush = null;
    /** This person's open filter dropdown — a table column's value list. */
    this.filterPanel = null;
    /** Pivot definitions, read lazily and dropped when the parts change. */
    this._pivots = null;
  }

  /**
   * Run a mutation as ONE undo step.
   *
   * A sheet's undo state is NOT the part XML. Edits live in the calculation
   * engine until save() writes them back, so the thing to put back is the INPUTS
   * of the cells that changed — which is both smaller and exact. Pasting a
   * thousand cells snapshots a thousand inputs, not a megabyte of XML.
   *
   * `cells` is the range the edit will touch, captured before it runs.
   */
  _edit(label, group, cells, fn, {
    styles = false, parts = null, structural = false, tracksNewParts = false,
  } = {}) {
    const outermost = this._editDepth === 0;
    // THE protection gate. One gate, here, because every cell-changing
    // gesture already hands `_edit` the cells it will touch — that list is
    // the undo footprint, and the same list is exactly what protection needs
    // to judge. Refusing BEFORE the history record keeps the house rule that
    // a refused edit leaves no undo step.
    if (outermost) {
      const p = this.protection();
      if (p.sheet) {
        if (structural) {
          throw protectionError('This sheet is protected — unprotect it before changing its structure.');
        }
        // A pure formatting edit is governed by the author's formatCells
        // allowance, not by each cell's locked flag — Excel's own split.
        if (!(styles && p.formatCells)) {
          for (const { row, col } of cells) {
            if (this.isCellLocked(row, col)) {
              throw protectionError('This sheet is protected and ' + ref(row, col)
                + ' is locked. Unprotect the sheet to change it.');
            }
          }
        }
      }
    }
    // An edit that CREATES parts — adding a pivot table — has to be able to
    // take them away again, or undo leaves a ghost feature behind with a
    // relationship still pointing at it. The names are diffed after the
    // mutation runs, because before it there is nothing to name.
    const partsBefore = tracksNewParts && outermost ? new Set(this.pkg.partNames()) : null;
    if (outermost) {
      const sheet = this.activeSheet;
      this.history.record({
        label,
        group,
        state: {
          sheet,
          cells: cells.map(({ row, col }) => ({
            row, col, input: this.calc.getInput(sheet, row, col),
          })),
          // A STRUCTURAL edit flushes every pending cell edit into the parts
          // before transforming them, and undo restores the PRE-flush parts —
          // so those pending inputs must travel in the entry to be re-applied
          // on top of the rebuilt calc model. See `_step`.
          structural,
          pending: structural ? this._capturePending() : null,
          // Formatting is undoable too, so the style index travels with the
          // value. Without it, undoing a bold leaves the cell bold and only
          // puts the text back — which reads as the undo having failed.
          styles: cells.map(({ row, col }) => ({
            row, col, index: this._styleIndexAt(sheet, row, col),
          })),
          // Only a FORMAT edit carries styles.xml. Capturing it on every
          // keystroke would copy the whole style table per character, and a
          // value edit cannot change it anyway.
          stylesXml: styles && this.pkg.has('xl/styles.xml') ? this.pkg.text('xl/styles.xml') : null,
          // A STRUCTURAL edit (merge) rewrites the sheet PART, not just cell
          // inputs. Snapshot the whole part so undo can put the mergeCells block
          // — and everything else in it — back exactly, the same way stylesXml
          // travels for a format edit.
          parts: parts ? this.workbook.snapshotParts(parts) : null,
          structuralDirty: this._structuralDirty,
          // Restoring the dirty set matters: after undoing the only edit, the
          // document is no longer unsaved, and the Save button must say so.
          dirty: [...this.dirtyCells],
          styled: [...this.styledCells],
        },
        meta: { selection: this.selection.clone(), sheet },
      });
    }
    this._editDepth += 1;
    try {
      return fn();
    } finally {
      this._editDepth -= 1;
      if (partsBefore) {
        const entry = this.history.past.at(-1);
        const added = {};
        for (const name of this.pkg.partNames()) {
          if (!partsBefore.has(name)) added[name] = this.pkg.text(name);
        }
        // Recorded on the PRE-edit state: undoing back to it removes them,
        // and the mirror `_step` builds carries them forward for redo.
        if (entry && Object.keys(added).length) entry.state.addedParts = added;
      }
    }
  }

  get canUndo() { return this.history.canUndo; }
  get canRedo() { return this.history.canRedo; }

  /** The cells of the current selection, as plain coordinates. */
  _selectedCells() { return [...this.selection.cells()]; }

  undo() { return this._step('undo'); }
  redo() { return this._step('redo'); }

  _step(direction) {
    const previous = this.history[direction === 'undo' ? 'past' : 'future'].at(-1);
    if (!previous) return false;

    // Capture the CURRENT values of exactly the cells the step will overwrite,
    // so the opposite direction can put them back.
    const sheet = previous.state.sheet;
    const current = {
      sheet,
      cells: previous.state.cells.map(({ row, col }) => ({
        row, col, input: this.calc.getInput(sheet, row, col),
      })),
      styles: (previous.state.styles ?? []).map(({ row, col }) => ({
        row, col, index: this._styleIndexAt(sheet, row, col),
      })),
      // Mirror what the step being reversed captured: if it carried a style
      // table, the opposite direction has to carry one back.
      stylesXml: previous.state.stylesXml !== null && previous.state.stylesXml !== undefined
        && this.pkg.has('xl/styles.xml') ? this.pkg.text('xl/styles.xml') : null,
      // Mirror a structural step: if it carried sheet parts, the opposite
      // direction has to carry the CURRENT ones back, snapshotted as they are now.
      parts: previous.state.parts !== null && previous.state.parts !== undefined
        ? this.workbook.snapshotParts(Object.keys(previous.state.parts)) : null,
      structural: Boolean(previous.state.structural),
      // Mirror the created-parts record so the opposite direction can put
      // them back: reading their CURRENT text where they exist (undoing, so
      // they do), falling back to the stored copy where they do not (redoing).
      addedParts: previous.state.addedParts
        ? Object.fromEntries(Object.entries(previous.state.addedParts)
          .map(([n, text]) => [n, this.pkg.has(n) ? this.pkg.text(n) : text]))
        : null,
      // The same coordinates, read in the state being left. In a post-edit
      // state the flushed parts already hold these values, so re-applying them
      // is a no-op; in a pre-edit state they are exactly the pending edits.
      pending: previous.state.structural
        ? (previous.state.pending ?? []).map((p) => ({
          ...p,
          input: this.calc.getInput(p.sheet, p.row, p.col),
          styleIndex: this._styleIndexAt(p.sheet, p.row, p.col),
        }))
        : null,
      structuralDirty: this._structuralDirty,
      dirty: [...this.dirtyCells],
      styled: [...this.styledCells],
    };
    const entry = this.history[direction](current, { selection: this.selection.clone(), sheet: this.activeSheet });
    if (!entry) return false;

    // Parts an edit CREATED. Undoing lands on the state before they existed,
    // so they go; redoing lands on the state after, so they come back. Their
    // content-type overrides ride in the `parts` snapshot of
    // `[Content_Types].xml`, which is restored below — so re-adding here
    // deliberately declares no type of its own rather than guessing one.
    if (entry.state.addedParts) {
      for (const [name, text] of Object.entries(entry.state.addedParts)) {
        if (direction === 'undo') this.pkg.removePart(name);
        else if (!this.pkg.has(name)) this.pkg.addPart(name, text);
      }
      this._pivots = null;
    }

    // A structural step is a different animal: the parts are the truth and
    // everything derived from them — the calc model, the geometry, the style
    // map, the merges — is rebuilt from the restored XML, then the pending
    // edits the flush had absorbed are laid back on top.
    if (entry.state.structural) {
      this.activeSheet = entry.state.sheet;
      this.workbook.restoreParts(entry.state.parts ?? {});
      this._rebuildDerivedState();
      for (const p of entry.state.pending ?? []) {
        this.calc.setCell(p.sheet, p.row, p.col, p.input ?? '');
        this._setStyleIndex(p.sheet, p.row, p.col, p.styleIndex ?? null);
      }
      this.calc.recalculate();
      if (entry.state.structuralDirty !== undefined) this._structuralDirty = entry.state.structuralDirty;
      this.dirtyCells = new Set(entry.state.dirty);
      this.styledCells = new Set(entry.state.styled ?? []);
      if (entry.meta?.selection) this.selection = entry.meta.selection;
      this.editing = null;
      this.ensureVisible();
      return true;
    }

    this.activeSheet = entry.state.sheet;
    for (const { row, col, input } of entry.state.cells) {
      this.calc.setCell(entry.state.sheet, row, col, input ?? '');
    }
    this.calc.recalculate();
    // styles.xml is restored WHOLE rather than per cell: undoing the first bold
    // in a file has to drop the <font> and <xf> that bold appended, or the next
    // one reuses an entry the document no longer references.
    if (entry.state.stylesXml !== null && entry.state.stylesXml !== undefined) {
      this.pkg.write_('xl/styles.xml', entry.state.stylesXml);
      this.styles = readStyles(this.pkg, { BUILTIN_FORMATS });
    }
    for (const { row, col, index } of entry.state.styles ?? []) {
      this._setStyleIndex(entry.state.sheet, row, col, index);
    }
    // A structural part goes back WHOLE, like styles.xml does: the mergeCells
    // block is not per-cell state and cannot be rebuilt from cell inputs. Only
    // the merge map is derived from it, so that is all that is re-read — geometry
    // and cell styles are untouched by a merge and re-reading them here would
    // clobber in-session style edits the part does not carry.
    if (entry.state.parts) {
      this.workbook.restoreParts(entry.state.parts);
      // A restored part may be a pivot's definition or its cache.
      this._pivots = null;
      let namesTouched = false;
      for (const name of Object.keys(entry.state.parts)) {
        // workbook.xml carries the defined names; a restored copy means the
        // calc model's name table is stale and every formula may read differently.
        if (name === this.workbook.mainPart) { namesTouched = true; continue; }
        const sheetName = this._sheetNameForPart(name);
        if (!sheetName) continue;
        const xml = this.pkg.text(name);
        this.merges.set(sheetName, readMergedCells(xml));
        this.validations.set(sheetName, readDataValidations(xml));
        this.conditionals.set(sheetName, readConditionalFormatting(xml, this.styles.theme));
        // Geometry is part-derived too: undoing a resize restores the part,
        // so the widths and heights must be re-read from it or the grid keeps
        // drawing the size the file no longer has.
        this.geometry.set(sheetName, SheetGeometry.fromSheetXml(xml));
        // And so are the drawings: undoing an inserted shape or chart
        // restores the parts, and the cache must stop painting what the
        // file no longer carries.
        this.drawings.set(sheetName, this._readDrawings(name));
      }
      if (namesTouched) {
        this._syncNames();
        this.calc.recalculate();
      }
    }
    if (entry.state.structuralDirty !== undefined) this._structuralDirty = entry.state.structuralDirty;
    this.dirtyCells = new Set(entry.state.dirty);
    this.styledCells = new Set(entry.state.styled ?? []);
    if (entry.meta?.selection) this.selection = entry.meta.selection;
    this.editing = null;
    this.ensureVisible();
    return true;
  }

  static open(buf, opts) { return new SheetView(buf, opts); }

  /**
   * cellRef -> cellXfs index.
   *
   * The INDEX is kept rather than the derived format code, because the same
   * index also carries the font, fill, border and alignment. Deriving one thing
   * and throwing the rest away is what made a formatted workbook render as plain
   * text.
   */
  _readCellStyles(xml) {
    const out = new Map();
    for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>)/g)) {
      const r = /\br="([A-Z]+\d+)"/.exec(m[1]);
      const s = /\bs="(\d+)"/.exec(m[1]);
      if (r && s) out.set(r[1], Number(s[1]));
    }
    return out;
  }

  /** The full style behind a cell, or null where the file said nothing. */
  styleFor(row, col) {
    const index = this.cellStyles.get(this.activeSheet)?.get(ref(row, col));
    return index === undefined ? null : this.styles.byStyleIndex[index] ?? null;
  }

  /** The merge this cell belongs to, if any. */
  mergeAt(row, col) {
    return (this.merges.get(this.activeSheet) ?? []).find(
      (m) => row >= m.top && row <= m.bottom && col >= m.left && col <= m.right,
    ) ?? null;
  }

  /**
   * Charts and pictures anchored to a sheet.
   *
   * A chart from a customer's file is re-drawn in OUR palette rather than
   * approximated in theirs — the original part is preserved untouched, so this
   * is a reading, not a rewrite. Anything we cannot parse comes back with a null
   * scene and a reason, so the shell can show a placeholder that says what it is
   * instead of a blank rectangle.
   */
  _readDrawings(sheetPart) {
    const relsForSheet = this.pkg.rels(sheetPart);
    const drawingRel = relsForSheet.find((r) => String(r.Type).endsWith('/drawing'));
    if (!drawingRel) return [];
    const drawingPart = OoxmlPackage.resolveTarget(sheetPart, drawingRel.Target);
    if (!this.pkg.has(drawingPart)) return [];

    const rels = new Map(this.pkg.rels(drawingPart).map((r) => [r.Id, r.Target]));
    let found = [];
    try {
      found = readSheetDrawings({
        drawingXml: this.pkg.text(drawingPart),
        resolveRelationship: (id) => {
          const target = rels.get(id);
          return target ? OoxmlPackage.resolveTarget(drawingPart, target) : null;
        },
        readPart: (name) => (this.pkg.has(name) ? this.pkg.text(name) : null),
        // pictures need the BYTES, not the text — an image part is binary
        readPartBinary: (name) => (this.pkg.has(name) ? this.pkg.read(name) : null),
        mode: this.mode,
      });
    } catch {
      return [];
    }
    return found.map((d, i) => ({ ...d, id: 'drawing-' + i }));
  }

  /**
   * The notes on a sheet's cells, from the comments part its rels point at.
   *
   * These are the classic cell comments — the red-corner kind. Excel's
   * threaded comments write a legacy comments part too (that is how older
   * Excels see them), so reading the one part shows both. Display only:
   * writing a comment is a different feature, and the part is preserved
   * byte-identically either way.
   */
  _readComments(sheetPart) {
    const rel = this.pkg.rels(sheetPart).find((r) => String(r.Type).endsWith('/comments'));
    if (!rel) return new Map();
    const partName = OoxmlPackage.resolveTarget(sheetPart, rel.Target);
    if (!this.pkg.has(partName)) return new Map();
    const xml = this.pkg.text(partName);
    const authors = [...(/<authors>([\s\S]*?)<\/authors>/.exec(xml)?.[1] ?? '')
      .matchAll(/<author>([\s\S]*?)<\/author>/g)].map((m) => unescXml(m[1]));
    const out = new Map();
    for (const m of xml.matchAll(/<comment\b([^>]*)>([\s\S]*?)<\/comment>/g)) {
      const refMatch = /\bref="([A-Z]+\d+)"/.exec(m[1]);
      if (!refMatch) continue;
      const authorId = Number((/\bauthorId="(\d+)"/.exec(m[1]) ?? [])[1] ?? -1);
      const text = [...m[2].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((t) => unescXml(t[1])).join('');
      // A threaded comment's shadow author is a GUID marker, not a name.
      const author = authors[authorId] && !authors[authorId].startsWith('tc=') ? authors[authorId] : null;
      out.set(refMatch[1], { author, text });
    }
    return out;
  }

  sheetNames() { return this.workbook.sheetNames(); }

  selectSheet(name) {
    if (!this.sheetNames().includes(name)) throw new Error('no such sheet: ' + name);
    this.activeSheet = name;
    this.selection = Selection.at(0, 0);
    this.scrollX = 0;
    this.scrollY = 0;
    return this;
  }

  get geo() { return this.geometry.get(this.activeSheet); }
  get bounds() { return this.calc.usedBounds(this.activeSheet); }

  formatFor(row, col) {
    return this.styleFor(row, col)?.numFmt ?? 'General';
  }

  /** What the grid shows: calculated, then formatted. */
  displayValue(row, col) {
    const value = this.calc.getValue(this.activeSheet, row, col);
    return formatValue(value, this.formatFor(row, col));
  }

  /** What the formula bar shows: exactly what was typed. */
  editValue(row, col) {
    const input = this.calc.getInput(this.activeSheet, row, col);
    return input === '' || input === null || input === undefined ? '' : String(input);
  }

  isFilled(row, col) {
    const c = this.calc.cell(this.activeSheet, row, col);
    return Boolean(c && c.input !== '' && c.input !== null && c.input !== undefined);
  }

  // ---- rendering ---------------------------------------------------------

  /**
   * Everything on screen right now. This is the whole contract with the shell:
   * hand it this object and it can draw a frame.
   */
  render() {
    const geo = this.geo;
    const vp = geo.viewport({
      scrollX: this.scrollX,
      scrollY: this.scrollY,
      width: this.viewportWidth,
      height: this.viewportHeight,
    });

    // Frozen rows and columns ride in EVERY frame, wherever the viewport has
    // scrolled — the client pins them while the rest slides underneath.
    const frozen = this.frozenPane();
    const rowIndices = [];
    for (let r = 0; r < frozen.rows && r < vp.firstRow; r++) rowIndices.push(r);
    for (let r = vp.firstRow; r <= vp.lastRow; r++) rowIndices.push(r);
    const colIndices = [];
    for (let c = 0; c < frozen.cols && c < vp.firstCol; c++) colIndices.push(c);
    for (let c = vp.firstCol; c <= vp.lastCol; c++) colIndices.push(c);

    const cells = [];
    // Conditional-formatting range aggregates, computed once per pass — and
    // the sheet's tables, read once here and reused by the frame below.
    const cfCache = new Map();
    const sheetTables = this.sheetTables();
    for (const row of rowIndices) {
      const height = geo.rowHeight(row);
      if (height === 0) continue;
      for (const col of colIndices) {
        const width = geo.colWidth(col);
        if (width === 0) continue;

        // A merged range is drawn ONCE, by its top-left cell, spanning the whole
        // range. The cells it covers are not empty — they are not there at all,
        // and drawing them would put grid lines through the middle of a title.
        const merge = this.mergeAt(row, col);
        if (merge && (merge.top !== row || merge.left !== col)) continue;

        let style = this.styleFor(row, col);
        const display = this.displayValue(row, col);
        // What conditional formatting adds, layered OVER the cell's own style
        // the way Excel paints it — the rule's fill and font win, the border
        // stays the cell's.
        const cf = this._conditionalStyle(row, col, cfCache);
        if (cf && (cf.fill || cf.font)) {
          const font = cf.font
            ? {
              ...(style?.font ?? {}),
              ...(cf.font.bold ? { bold: true } : {}),
              ...(cf.font.italic ? { italic: true } : {}),
              ...(cf.font.strike ? { strike: true } : {}),
              ...(cf.font.colour ? { colour: cf.font.colour } : {}),
            }
            : style?.font ?? null;
          style = { ...(style ?? {}), font, fill: cf.fill ?? style?.fill ?? null };
        }
        // Table banding, where nothing louder already spoke: the cell's own
        // fill and conditional formatting both beat it, so a table look never
        // hides what the file or a rule chose. Painted in OUR palette from
        // the theme's first accent — the philosophy charts established.
        if (!style?.fill) {
          const look = tableLookAt(sheetTables, row, col, this.styles.theme);
          if (look) {
            style = {
              ...(style ?? {}),
              font: look.font ? { ...(style?.font ?? {}), ...look.font } : style?.font ?? null,
              fill: look.fill ?? style?.fill ?? null,
            };
          }
        }
        const note = this.comments.get(this.activeSheet)?.get(ref(row, col)) ?? null;
        // A styled but empty cell still has to be drawn: a shaded header with no
        // text in it is a real thing, and skipping it leaves a hole in the band.
        const decorated = Boolean(style && (style.fill || style.border)) || Boolean(cf?.bar)
          || Boolean(cf?.icon) || Boolean(note);
        if (display.text === '' && !decorated && !merge && !this.selection.contains(row, col)) continue;

        const spanW = merge
          ? geo.colOffset(merge.right) + geo.colWidth(merge.right) - geo.colOffset(merge.left)
          : width;
        const spanH = merge
          ? geo.rowOffset(merge.bottom) + geo.rowHeight(merge.bottom) - geo.rowOffset(merge.top)
          : height;

        cells.push({
          row,
          col,
          ref: ref(row, col),
          // An icon rule with showValue="0" shows the icon ALONE — the value
          // still calculates, copies and edits; only the paint hides it.
          text: cf?.hideValue ? '' : display.text,
          // The file's own alignment wins over the format's default: a number
          // the author centred is centred, however it is formatted.
          align: style?.align?.horizontal ?? display.align,
          valign: style?.align?.vertical ?? null,
          wrap: Boolean(style?.align?.wrap),
          indent: style?.align?.indent ?? 0,
          colour: display.colour,
          style: style ? { font: style.font, fill: style.fill, border: style.border } : null,
          // A data bar rides beside the style: length as a fraction, so the
          // client needs no numbers to draw it.
          bar: cf?.bar ?? null,
          // An icon-set rule's verdict: which icon of which set. The client
          // owns the glyphs; a set it cannot draw paints nothing.
          icon: cf?.icon ?? null,
          // The cell's comment, for the corner mark and its tooltip.
          note,
          merged: merge ? { ref: merge.ref, rows: merge.bottom - merge.top + 1, cols: merge.right - merge.left + 1 } : null,
          isError: isError(this.calc.getValue(this.activeSheet, row, col)),
          isFormula: this.editValue(row, col).startsWith('='),
          // The formula itself, for a grid showing formulas instead of results.
          ...(this.editValue(row, col).startsWith('=') ? { formula: this.editValue(row, col) } : {}),
          selected: this.selection.contains(row, col),
          active: this.selection.active.row === row && this.selection.active.col === col,
          x: geo.colOffset(col),
          y: geo.rowOffset(row),
          width: spanW,
          height: spanH,
        });
      }
    }

    const columns = [];
    for (const col of colIndices) {
      const width = geo.colWidth(col);
      if (width === 0) continue;
      columns.push({ index: col, name: colName(col), x: geo.colOffset(col), width });
    }
    const rows = [];
    for (const row of rowIndices) {
      const height = geo.rowHeight(row);
      if (height === 0) continue;
      rows.push({ index: row, label: String(row + 1), y: geo.rowOffset(row), height });
    }

    // Drawings are culled like cells are. Rendering an SVG costs real work, so
    // a sheet with forty charts must not re-draw all forty on every keystroke.
    const visibleLeft = this.scrollX;
    const visibleTop = this.scrollY;
    const visibleRight = this.scrollX + this.viewportWidth;
    const visibleBottom = this.scrollY + this.viewportHeight;

    const drawings = (this.drawings.get(this.activeSheet) ?? []).map((d) => {
      const x = d.from ? geo.colOffset(d.from.col) + Math.round((d.from.colOffsetEmu ?? 0) / 9525) : 0;
      const y = d.from ? geo.rowOffset(d.from.row) + Math.round((d.from.rowOffsetEmu ?? 0) / 9525) : 0;
      // A two-cell anchor resizes with its cells, so its size comes from the
      // grid, not from a stored extent.
      const width = d.to
        ? Math.max(80, geo.colOffset(d.to.col) + Math.round((d.to.colOffsetEmu ?? 0) / 9525) - x)
        : Math.round(d.widthPx ?? 320);
      const height = d.to
        ? Math.max(60, geo.rowOffset(d.to.row) + Math.round((d.to.rowOffsetEmu ?? 0) / 9525) - y)
        : Math.round(d.heightPx ?? 240);

      if (x > visibleRight || y > visibleBottom || x + width < visibleLeft || y + height < visibleTop) {
        return null;
      }

      let svg = null;
      let unsupported = null;
      const box = { x: 0, y: 0, width, height };
      try {
        if (d.spec) {
          svg = renderSvg(buildChart({ ...d.spec, width, height, mode: this.mode }));
        } else if (d.descriptor?.kind === 'shape') {
          svg = renderSvg(scene({
            width, height, mode: this.mode, background: 'none',
            title: d.name ?? 'Shape',
            children: [buildShape(d.descriptor, box, { mode: this.mode })],
          }));
        } else if (d.descriptor?.kind === 'picture' && d.descriptor.href) {
          svg = renderSvg(scene({
            width, height, mode: this.mode, background: 'none',
            title: d.name ?? 'Picture',
            children: [buildPicture(d.descriptor, box)],
          }));
        } else {
          unsupported = d.kind === 'chart'
            ? 'this chart could not be read'
            : d.kind === 'image'
              ? 'this picture could not be loaded'
              : 'a ' + d.kind + ' is preserved but not yet drawn';
        }
      } catch (e) {
        unsupported = e.message;
      }
      return { id: d.id, kind: d.kind, name: d.name, x, y, width, height, svg, unsupported };
    }).filter(Boolean);

    const active = this.selection.active;
    return {
      drawings,
      // What the toolbar needs: whether the buttons are live and what they say.
      history: this.history.describe(),
      // What the toolbar needs to LOOK like: which toggles the selection reads as
      // already on, so Bold appears pressed when the selection is bold.
      format: this.formatState(),
      // Whether the selection is already a single merged region, so the Merge
      // button lights up and pressing it unmerges rather than re-merges.
      merged: this.isSelectionMerged(),
      sheet: this.activeSheet,
      sheets: this.sheetNames(),
      // The defined names, for the Name Manager and the name box's jump list.
      names: this.names(),
      // The tables (ListObjects) on this sheet, for styling and the filter
      // UI — plus the sheet-level autofilter as a pseudo-entry named #sheet,
      // so the grid's funnels and panel serve both without a second path.
      tables: (() => {
        const af = this.sheetFilter();
        return af ? [...sheetTables, af] : sheetTables;
      })(),
      // This person's open filter dropdown, if any.
      filterPanel: this.filterPanel,
      // The sheet's stored what-if scenarios, for the Data tab.
      scenarios: this.scenarios().map((s) => ({ name: s.name, comment: s.comment, cells: s.cells.length })),
      // Whether this sheet refuses edits to locked cells, and whether its
      // author sealed that with a password only Excel can lift.
      protection: this.protection(),
      // The data block around the cursor, with its headings — what a New
      // pivot panel offers as fields without the person typing them.
      region: (() => {
        const r = this._currentRegion(active.row, active.col);
        if (r.bottom <= r.top || r.right < r.left) return null;
        const headers = [];
        for (let c = r.left; c <= r.right; c++) {
          headers.push(String(this.displayValue(r.top, c).text ?? ''));
        }
        return headers.every((h) => h) ? { ref: areaRef(r), headers } : null;
      })(),
      // The workbook's pivots — each says where it lives, what it summarises,
      // and why it cannot be refreshed when that is the case.
      pivots: this.pivots().map((p) => ({
        name: p.name,
        sheet: p.sheet,
        ref: p.unsupported ? null : areaRef(p.location),
        source: p.source ? (p.source.sheet + '!' + areaRef(p.source)) : null,
        fields: p.cacheFields.map((f) => f.name),
        rows: p.rowFields.map((f) => p.cacheFields[f]?.name ?? ''),
        cols: p.colFields.map((f) => p.cacheFields[f]?.name ?? ''),
        values: p.dataFields.map((d) => dataFieldLabel(p, d)),
        unsupported: p.unsupported,
      })),
      viewport: vp,
      cells,
      columns,
      rows,
      // The canvas SPREADS: it covers the used range, wherever the selection
      // has gone, and wherever the viewport has scrolled to — plus the margin
      // totalSize adds. Scrolling toward the edge therefore keeps extending
      // the sheet, and clicking or arrowing into empty space always has grid
      // under it, the way a spreadsheet behaves rather than a document.
      //
      // And it never CONTRACTS while the session lives: the extent is a
      // per-sheet high-water mark. It used to follow the viewport back down,
      // so a frame computed behind a scrolled-ahead client shrank the canvas,
      // the browser clamped the scroll position to fit, the clamp fired a
      // scroll event, and that echo was a new action — a feedback loop that
      // idled at one request per second in a background tab, forever (the
      // sink's coalescing hid it at ~8/s in the foreground). A floor cannot
      // clamp anybody; scrolling back simply leaves explored grid behind,
      // which is how a spreadsheet behaves anyway.
      total: this._stickyTotal(geo, {
        maxRow: Math.max(this.bounds.maxRow, this.selection.range.bottom, vp.lastRow),
        maxCol: Math.max(this.bounds.maxCol, this.selection.range.right, vp.lastCol),
      }),
      headerWidth: geo.headerWidth,
      headerHeight: geo.headerHeight,
      // The frozen pane, with its band sizes in pixels for the client's clip.
      frozen: {
        rows: frozen.rows,
        cols: frozen.cols,
        width: frozen.cols ? geo.colOffset(frozen.cols) : 0,
        height: frozen.rows ? geo.rowOffset(frozen.rows) : 0,
      },
      selection: {
        ...this.selection.range,
        ref: this.selection.toString(),
        active: { ...active, ref: ref(active.row, active.col) },
      },
      // The active cell's validation rule, for the in-cell dropdown and the
      // author's input prompt. Null for the unruled majority.
      validation: (() => {
        const rule = this.validationAt(active.row, active.col);
        if (!rule) return null;
        const list = rule.type === 'list' && !rule.suppressDropdown
          ? this.validationOptions(rule, active.row, active.col)
          : null;
        return {
          type: rule.type,
          prompt: rule.prompt,
          promptTitle: rule.promptTitle,
          list: list && list.length ? list : null,
        };
      })(),
      formulaBar: this.editing ? this.editing.draft : this.editValue(active.row, active.col),
      editing: this.editing ? { ...this.editing } : null,
      status: this.statusLine(),
    };
  }

  /** The summary strip a spreadsheet shows for a multi-cell selection. */
  statusLine() {
    if (this.selection.isSingle) return null;
    let count = 0;
    let numeric = 0;
    let sum = 0;
    let min = Infinity;
    let max = -Infinity;
    for (const { row, col } of this.selection.cells({ limit: 50000 })) {
      const v = this.calc.getValue(this.activeSheet, row, col);
      if (v === '' || v === null || v === undefined) continue;
      count += 1;
      if (typeof v === 'number') {
        numeric += 1;
        sum += v;
        min = Math.min(min, v);
        max = Math.max(max, v);
      }
    }
    if (!count) return null;
    return numeric
      ? { count, numeric, sum, average: sum / numeric, min, max }
      : { count, numeric: 0 };
  }

  // ---- interaction -------------------------------------------------------

  select(row, col, { extend = false } = {}) {
    // A cell change ends the run: "type in A1, click B4, type" is two undos.
    this.history.break();
    if (extend) this.selection.extendTo(row, col);
    else this.selection.collapseTo(row, col);
    this.ensureVisible();
    return this;
  }

  /**
   * Navigation is bounded by the SHEET, not by the used range.
   *
   * `usedBounds` is for clipping ranges in formulas — a whole-column `SUM(A:A)`
   * must not walk a million empty rows. Reusing it here stops the cursor dead at
   * the last filled column, so a sheet with data in A:F refuses to let anyone
   * click into G. Empty cells are where new data goes; you have to be able to
   * get to them.
   */
  moveSelection(direction, opts = {}) {
    this.selection.move(direction, {
      ...opts,
      isFilled: (r, c) => this.isFilled(r, c),
      bounds: { maxRow: MAX_ROWS - 1, maxCol: MAX_COLS - 1 },
    });
    this.ensureVisible();
    return this;
  }

  tab(back = false) {
    this.selection.tab(back, { bounds: { maxRow: MAX_ROWS - 1, maxCol: MAX_COLS - 1 } });
    this.ensureVisible();
    return this;
  }
  enterKey(back = false) {
    this.selection.enter(back, { bounds: { maxRow: MAX_ROWS - 1, maxCol: MAX_COLS - 1 } });
    this.ensureVisible();
    return this;
  }

  ensureVisible() {
    const { row, col } = this.selection.active;
    const next = this.geo.scrollToShow({
      row, col,
      scrollX: this.scrollX,
      scrollY: this.scrollY,
      width: this.viewportWidth,
      height: this.viewportHeight,
    });
    this.scrollX = next.scrollX;
    this.scrollY = next.scrollY;
    return this;
  }

  scrollTo(x, y) {
    this.scrollX = Math.max(0, x);
    this.scrollY = Math.max(0, y);
    return this;
  }

  /**
   * Begin editing. `replace` is what typing a character does — it discards the
   * old content — while F2 or a double click keeps it for amendment.
   */
  beginEdit({ replace = false, initial = '' } = {}) {
    const { row, col } = this.selection.active;
    this.editing = {
      row, col,
      ref: ref(row, col),
      draft: replace ? initial : this.editValue(row, col) + initial,
      startedFrom: this.editValue(row, col),
    };
    return this;
  }

  updateDraft(text) {
    if (!this.editing) this.beginEdit({ replace: true });
    this.editing.draft = text;
    return this;
  }

  cancelEdit() {
    this.editing = null;
    return this;
  }

  /** Commit the edit into the calculation model and recalculate dependents. */
  commitEdit({ move = 'down' } = {}) {
    if (!this.editing) return this;
    const { row, col, draft, startedFrom } = this.editing;
    this.editing = null;
    if (draft !== startedFrom) this.setCell(row, col, draft);
    // Shift reverses the direction, as it does everywhere else in a sheet, and
    // 'none' commits where the cursor is — which is what a click away means.
    if (move === 'down') this.enterKey(false);
    else if (move === 'up') this.enterKey(true);
    else if (move === 'right') this.tab(false);
    else if (move === 'left') this.tab(true);
    return this;
  }

  /** Direct set, for paste and programmatic edits. */
  setCell(row, col, input) {
    // A formula entered in a table's data column FILLS the column — Excel's
    // calculated-column behaviour — when doing so overwrites nothing: every
    // other data cell must be empty or already hold this formula shifted to
    // its row. One edit, one undo, and the protection gate judges the whole
    // column because the touched list carries it.
    const fillRows = this._calculatedColumnRows(row, col, input);
    if (fillRows.length) {
      return this._edit('edit ' + ref(row, col), null,
        [{ row, col }, ...fillRows.map((r) => ({ row: r, col }))], () => {
          this._setCell(row, col, input);
          for (const r of fillRows) this._setCell(r, col, shiftFormula(input, r - row, 0));
        });
    }
    // Group by cell: retyping the same cell repeatedly is one undo, moving to
    // another cell starts a new one. That is Excel's behaviour and the reason
    // an undo does not walk back through every intermediate value.
    return this._edit('edit ' + ref(row, col), 'cell:' + this.activeSheet + '!' + ref(row, col),
      [{ row, col }], () => this._setCell(row, col, input));
  }

  /**
   * The other data rows a table-column formula should fill, or [] when it
   * should not: not a formula, not in a table's data band, or the column
   * holds something the fill would overwrite. Rows already holding the same
   * shifted formula are left alone — refilling them would dirty cells that
   * are not changing.
   */
  /**
   * The per-sheet canvas extent, monotonic for the life of the view.
   *
   * Pixel sizes still come fresh from the geometry — a resized column changes
   * the TOTAL, never the extent — so only the row/col indices are held.
   * Structural deletes leave the floor where it was; that is the point, not a
   * leak: a shrinking total is the one thing that can move a client's scroll
   * position without the person touching anything.
   */
  _stickyTotal(geo, want) {
    if (!this._extents) this._extents = new Map();
    const held = this._extents.get(this.activeSheet) ?? { maxRow: 0, maxCol: 0 };
    const merged = {
      maxRow: Math.max(want.maxRow, held.maxRow),
      maxCol: Math.max(want.maxCol, held.maxCol),
    };
    this._extents.set(this.activeSheet, merged);
    return geo.totalSize(merged);
  }

  _calculatedColumnRows(row, col, input) {
    if (typeof input !== 'string' || !input.startsWith('=')) return [];
    const t = this.sheetTables().find((x) => col >= x.left && col <= x.right
      && row >= x.top + x.headerRows && row <= x.bottom - x.totalsRows);
    if (!t) return [];
    const rows = [];
    for (let r = t.top + t.headerRows; r <= t.bottom - t.totalsRows; r++) {
      if (r === row) continue;
      const current = this.editValue(r, col);
      if (current === '') rows.push(r);
      else if (current !== shiftFormula(input, r - row, 0)) return [];
    }
    return rows;
  }

  _setCell(row, col, input) {
    this.calc.setCell(this.activeSheet, row, col, coerceInput(input));
    this.calc.recalculate();
    this.dirtyCells.add(this.activeSheet + '!' + ref(row, col));
    return this;
  }

  clearSelection() {
    const cells = this._selectedCells();
    return this._edit('clear', null, cells, () => {
      for (const { row, col } of cells) {
        if (this.isFilled(row, col)) this._setCell(row, col, '');
      }
      return this;
    });
  }

  // ---- merged cells ------------------------------------------------------

  /** The part backing a sheet, inverted from a part name. */
  _sheetNameForPart(part) {
    return this.workbook.sheets().find((s) => s.part === part)?.name ?? null;
  }

  /**
   * Re-read the merge map for a sheet from the part as it stands NOW.
   *
   * The part is loaded and dirty at this point — its new `<mergeCells>` block
   * lives in the SheetPart, not yet in the package — so `snapshotParts` renders
   * it rather than reading stale package text.
   */
  _refreshMerges(sheetName) {
    const part = this.workbook.partNameFor(sheetName);
    const xml = this.workbook.snapshotParts([part])[part];
    this.merges.set(sheetName, readMergedCells(xml));
  }

  /**
   * Whether the selection sits inside a single merged region — what a toolbar
   * reads to show the Merge button as already pressed. A merged cell clicked on
   * its own (a single-cell selection at the anchor) counts, as does the whole
   * region selected; a selection straddling a merge and open cells does not.
   */
  isSelectionMerged() {
    const s = this.selection.range;
    return (this.merges.get(this.activeSheet) ?? []).some(
      (m) => m.top <= s.top && m.bottom >= s.bottom && m.left <= s.left && m.right >= s.right,
    );
  }

  /**
   * Merge the selection rectangle into one region anchored at its top-left.
   *
   * Excel semantics: only the top-left cell's value survives; every other cell
   * in the rectangle is cleared. A single cell is not a region, so it is a
   * no-op. A selection that would cut an existing merge in half is refused — a
   * merge fully inside the new one is absorbed, but a partial overlap would have
   * to split a region and that is never what the user meant.
   */
  mergeSelection() {
    const r = this.selection.range;
    if (r.top === r.bottom && r.left === r.right) return this; // one cell is not a merge

    const existing = this.merges.get(this.activeSheet) ?? [];
    const contained = [];
    for (const m of existing) {
      const overlaps = !(m.right < r.left || m.left > r.right || m.bottom < r.top || m.top > r.bottom);
      if (!overlaps) continue;
      const inside = m.top >= r.top && m.bottom <= r.bottom && m.left >= r.left && m.right <= r.right;
      if (!inside) throw new Error('cannot merge across part of an existing merged region');
      contained.push(m);
    }

    const mergeRef = ref(r.top, r.left) + ':' + ref(r.bottom, r.right);
    const part = this.workbook.partNameFor(this.activeSheet);
    // Everything except the anchor is cleared, so the undo must cover them all.
    const cells = [];
    for (let row = r.top; row <= r.bottom; row++) {
      for (let col = r.left; col <= r.right; col++) {
        if (row === r.top && col === r.left) continue;
        cells.push({ row, col });
      }
    }

    return this._edit('merge', null, cells, () => {
      for (const m of contained) this.workbook.removeMerge(this.activeSheet, m.ref);
      this.workbook.addMerge(this.activeSheet, mergeRef);
      for (const { row, col } of cells) {
        if (this.isFilled(row, col)) this._setCell(row, col, '');
      }
      this._structuralDirty = true;
      this._refreshMerges(this.activeSheet);
      return this;
    }, { parts: [part] });
  }

  /**
   * Remove every merged region the selection touches.
   *
   * Unmerging restores no values — Excel does not un-clear the cells a merge
   * emptied, and neither do we — so nothing but the part travels in the undo.
   */
  unmergeSelection() {
    const r = this.selection.range;
    const touched = (this.merges.get(this.activeSheet) ?? []).filter(
      (m) => !(m.right < r.left || m.left > r.right || m.bottom < r.top || m.top > r.bottom),
    );
    if (!touched.length) return this;

    const part = this.workbook.partNameFor(this.activeSheet);
    return this._edit('unmerge', null, [], () => {
      for (const m of touched) this.workbook.removeMerge(this.activeSheet, m.ref);
      this._structuralDirty = true;
      this._refreshMerges(this.activeSheet);
      return this;
    }, { parts: [part] });
  }

  // ---- sheet structure: resize, insert, delete ---------------------------

  /**
   * Set one column's width in pixels. One undo step; the part carries the
   * `<cols>` record so undo restores the XML byte-for-byte, and the geometry
   * map is updated in place so the very next frame draws the new width.
   */
  setColWidth(col, widthPx) {
    if (!Number.isInteger(col) || col < 0) throw new Error('bad column index: ' + col);
    const px = Math.max(MIN_COL_WIDTH_PX, Math.round(Number(widthPx) || 0));
    const part = this.workbook.partNameFor(this.activeSheet);
    return this._edit('resize column', null, [], () => {
      this.workbook.setColWidthChars(this.activeSheet, col, pixelsToCharWidth(px));
      this.geo.colWidths.set(col, px);
      this._structuralDirty = true;
      return this;
    }, { parts: [part] });
  }

  /** Set one row's height in pixels. The mirror of `setColWidth`. */
  setRowHeight(row, heightPx) {
    if (!Number.isInteger(row) || row < 0) throw new Error('bad row index: ' + row);
    const px = Math.max(MIN_ROW_HEIGHT_PX, Math.round(Number(heightPx) || 0));
    const part = this.workbook.partNameFor(this.activeSheet);
    return this._edit('resize row', null, [], () => {
      this.workbook.setRowHeightPoints(this.activeSheet, row, pixelsToPoints(px));
      this.geo.rowHeights.set(row, px);
      this._structuralDirty = true;
      return this;
    }, { parts: [part] });
  }

  insertRows(at, count = 1) { return this._structural('insert rows', 'row', 'insert', at, count); }
  deleteRows(at, count = 1) { return this._structural('delete rows', 'row', 'delete', at, count); }
  insertCols(at, count = 1) { return this._structural('insert columns', 'col', 'insert', at, count); }
  deleteCols(at, count = 1) { return this._structural('delete columns', 'col', 'delete', at, count); }

  /**
   * One structural edit on the active sheet, as ONE undo step.
   *
   * The order is the whole design. Pending cell and style edits live in the
   * calc model until save; a structural edit renumbers the PART, so the parts
   * must first be brought up to truth (the flush), then transformed, then
   * everything derived from them rebuilt — the calc model the same way the
   * constructor builds it. The undo entry snapshots every sheet part plus
   * workbook.xml (defined names) BEFORE the flush, and carries the pending
   * edits separately so undo can lay them back on top of the restored parts.
   */
  _structural(label, axis, op, at, count) {
    if (!Number.isInteger(at) || at < 0) throw new Error('bad index for ' + label + ': ' + at);
    if (!Number.isInteger(count) || count < 1) throw new Error('bad count for ' + label + ': ' + count);
    const sheet = this.activeSheet;

    // Refuse BEFORE recording history: a refused edit must leave no undo step.
    if (op === 'delete') {
      const split = this.workbook.findMergeSplit(sheet, axis, at, count);
      if (split) {
        throw new Error('cannot delete ' + (axis === 'row' ? 'rows' : 'columns')
          + ': that would split a merged cell (' + split + ')');
      }
    }

    const parts = [...new Set([
      ...this.workbook.sheets().map((s) => s.part),
      this.workbook.mainPart,
    ])];

    return this._edit(label, null, [], () => {
      this._flushPendingEdits();
      if (axis === 'row' && op === 'insert') this.workbook.insertRows(sheet, at, count);
      else if (axis === 'row') this.workbook.deleteRows(sheet, at, count);
      else if (op === 'insert') this.workbook.insertCols(sheet, at, count);
      else this.workbook.deleteCols(sheet, at, count);

      // The parts now hold everything, including what the flush moved in, so
      // the per-cell dirty sets have nothing left to say — the structural flag
      // alone keeps the document reading as unsaved.
      this.dirtyCells.clear();
      this.styledCells.clear();
      this._structuralDirty = true;
      this._rebuildDerivedState();
      return this;
    }, { parts, structural: true });
  }

  /** Every pending cell edit and style edit, with enough to put them back. */
  _capturePending() {
    const out = [];
    const seen = new Set();
    for (const key of [...this.dirtyCells, ...this.styledCells]) {
      if (seen.has(key)) continue;
      seen.add(key);
      const [sheetName, cellRef] = key.split('!');
      const [row, col] = parseRefPair(cellRef);
      out.push({
        sheet: sheetName, row, col,
        input: this.calc.getInput(sheetName, row, col),
        styleIndex: this._styleIndexAt(sheetName, row, col),
      });
    }
    return out;
  }

  /** Write pending calc-level edits into the workbook parts, as save() would. */
  _flushPendingEdits() {
    for (const key of this.dirtyCells) {
      const [sheetName, cellRef] = key.split('!');
      const input = this.calc.getInput(sheetName, ...parseRefPair(cellRef));
      this.workbook.setCell(sheetName, cellRef, input === '' ? null : input);
    }
    for (const key of this.styledCells) {
      const [sheetName, cellRef] = key.split('!');
      const [row, col] = parseRefPair(cellRef);
      this.workbook.setCellStyle(sheetName, cellRef, this._styleIndexAt(sheetName, row, col));
    }
    return this;
  }

  /**
   * Rebuild everything derived from the sheet parts, exactly as the
   * constructor does: the calc model (which re-parses every formula, so a
   * text-level adjustment becomes real), the geometry, the style map and the
   * merges. Loaded parts are read through `snapshotParts` because their edits
   * may not have been flushed into the package yet.
   */
  _rebuildDerivedState() {
    const { sheet } = toSpreadsheet(this.workbook, { now: this._now });
    this.calc = sheet;
    this.calc.recalculate();
    this._pivots = null;
    for (const { name, part } of this.workbook.sheets()) {
      const xml = this.workbook.snapshotParts([part])[part];
      this.geometry.set(name, SheetGeometry.fromSheetXml(xml));
      this.cellStyles.set(name, this._readCellStyles(xml));
      this.merges.set(name, readMergedCells(xml));
      this.validations.set(name, readDataValidations(xml));
      this.conditionals.set(name, readConditionalFormatting(xml, this.styles.theme));
      this.comments.set(name, this._readComments(part));
    }
    return this;
  }

  /** Tab-separated text of the selection, for copy. */
  copyText() {
    const r = this.selection.range;
    const lines = [];
    for (let row = r.top; row <= r.bottom; row++) {
      const line = [];
      for (let col = r.left; col <= r.right; col++) line.push(this.displayValue(row, col).text);
      lines.push(line.join('\t'));
    }
    return lines.join('\n');
  }

  /** Record what a copy meant, so a matching paste can be rich. */
  markClipboard() {
    this.clipboard = {
      sheet: this.activeSheet,
      range: { ...this.selection.range },
      text: this.copyText(),
    };
    return this;
  }

  /**
   * The selection as an HTML table — the flavour OTHER applications read.
   * Inline styles carry the look Excel and Sheets both honour on paste:
   * weight, slant, colour, fill and alignment. The text is the DISPLAY text,
   * because the receiving application has no number formats of ours.
   */
  copyHtml() {
    const r = this.selection.range;
    const rows = [];
    for (let row = r.top; row <= r.bottom; row++) {
      const cells = [];
      for (let col = r.left; col <= r.right; col++) {
        const display = this.displayValue(row, col);
        const style = this.styleFor(row, col);
        const css = [];
        if (style?.font?.bold) css.push('font-weight:bold');
        if (style?.font?.italic) css.push('font-style:italic');
        if (style?.font?.colour) css.push('color:' + style.font.colour);
        if (style?.fill?.colour) css.push('background:' + style.fill.colour);
        if (display.align === 'right') css.push('text-align:right');
        cells.push('<td' + (css.length ? ' style="' + css.join(';') + '"' : '') + '>'
          + escapeHtml(display.text) + '</td>');
      }
      rows.push('<tr>' + cells.join('') + '</tr>');
    }
    return '<table>' + rows.join('') + '</table>';
  }

  /**
   * Paste at the active cell. Three roads, best first: our own copy pastes
   * RICH (formulas shift, styles travel); another application's HTML table
   * pastes by CELL, which splits more honestly than tab-separated text — a
   * value containing a tab or a newline stays one cell; anything else is
   * plain text split on tabs and lines.
   */
  pasteText(text, html) {
    const normalized = String(text ?? '').replace(/\r\n?/g, '\n');
    const clip = this.clipboard;
    if (clip && clip.text === normalized && this.sheetNames().includes(clip.sheet)) {
      return this._pasteCells(clip);
    }

    const { row, col } = this.selection.active;
    const grid = html ? parseHtmlTable(html) : null;
    if (grid) {
      const cells = [];
      grid.forEach((line, r) => line.forEach((_, c) => cells.push({ row: row + r, col: col + c })));
      const anyFormat = grid.some((line) => line.some((cell) => cell.format));
      if (anyFormat) this._ensureStylesPart();
      return this._edit('paste', null, cells, () => {
        grid.forEach((line, r) => line.forEach((cell, c) => this._setCell(row + r, col + c, cell.value)));
        // The formatting the foreign markup declared — bold, colours, fills,
        // alignment — lands the way setFormat lands it: per-cell xf deltas
        // appended to styles.xml. An UNDO restores each cell's old style
        // index; the appended xfs linger unused, which the append-never-
        // rewrite discipline makes harmless.
        if (anyFormat) {
          let xml = this.pkg.text('xl/styles.xml');
          let touched = false;
          grid.forEach((line, r) => line.forEach((cell, c) => {
            if (!cell.format) return;
            const base = this._styleIndexAt(this.activeSheet, row + r, col + c);
            const out = applyFormat(xml, base, cell.format);
            xml = out.xml;
            if (out.index !== base) {
              this._setStyleIndex(this.activeSheet, row + r, col + c, out.index);
              touched = true;
            }
          }));
          if (touched) {
            this.pkg.write_('xl/styles.xml', xml);
            this.styles = readStyles(this.pkg, { BUILTIN_FORMATS });
            this._stylesDirty = true;
          }
        }
        this.selection.collapseTo(row, col);
        this.selection.extendTo(row + grid.length - 1, col + Math.max(...grid.map((l) => l.length)) - 1);
        return this;
      });
    }

    const lines = normalized.split('\n');
    // The undo has to cover every cell the paste lands on, so the target range
    // is computed before anything is written.
    const cells = [];
    lines.forEach((line, r) => {
      line.split('\t').forEach((_, c) => cells.push({ row: row + r, col: col + c }));
    });
    return this._edit('paste', null, cells, () => this._pasteText(text, row, col, lines));
  }

  /**
   * The rich paste: inputs, with formulas SHIFTED by where they land, and the
   * source cells' style indices. Everything is read BEFORE anything is
   * written, so pasting a range onto a spot that overlaps it cannot read its
   * own output.
   */
  _pasteCells(clip) {
    const src = clip.range;
    const anchor = this.selection.active;
    const dr = anchor.row - src.top;
    const dc = anchor.col - src.left;

    const writes = [];
    for (let r = src.top; r <= src.bottom; r++) {
      for (let c = src.left; c <= src.right; c++) {
        const input = this.calc.getInput(clip.sheet, r, c);
        writes.push({
          row: r + dr,
          col: c + dc,
          input: typeof input === 'string' && input.startsWith('=') ? shiftFormula(input, dr, dc) : input,
          styleIndex: this._styleIndexAt(clip.sheet, r, c),
        });
      }
    }
    if (writes.some((w) => w.row < 0 || w.col < 0)) throw new Error('that paste would fall off the sheet');

    return this._edit('paste', null, writes.map(({ row, col }) => ({ row, col })), () => {
      for (const w of writes) {
        this._setCell(w.row, w.col, w.input ?? '');
        if ((this._styleIndexAt(this.activeSheet, w.row, w.col) ?? null) !== (w.styleIndex ?? null)) {
          this._setStyleIndex(this.activeSheet, w.row, w.col, w.styleIndex ?? null);
        }
      }
      this.selection.collapseTo(anchor.row, anchor.col);
      this.selection.extendTo(src.bottom + dr, src.right + dc);
      return this;
    });
  }

  _pasteText(text, row, col, lines) {
    lines.forEach((line, r) => {
      line.split('\t').forEach((value, c) => this._setCell(row + r, col + c, value));
    });
    const lastRow = row + lines.length - 1;
    const lastCol = col + Math.max(...lines.map((l) => l.split('\t').length)) - 1;
    this.selection.collapseTo(row, col);
    this.selection.extendTo(lastRow, lastCol);
    return this;
  }

  // ---- AutoSum -----------------------------------------------------------

  _isNumberAt(row, col) {
    return typeof this.calc.getValue(this.activeSheet, row, col) === 'number' && this.isFilled(row, col);
  }

  /**
   * AutoSum, with Excel's inference.
   *
   * A single cell looks UP for a run of numbers first, then LEFT, and writes
   * `=FN(range)` over the run it found; nothing adjacent to sum is a no-op
   * rather than a `=SUM()` that evaluates to a baffling 0. A multi-cell
   * selection writes one formula BELOW each selected column that contains a
   * number — all of it one undo step.
   */
  autoSum(fnName = 'SUM') {
    const fn = String(fnName).toUpperCase();
    if (!['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].includes(fn)) {
      throw new Error('autosum does not take ' + fnName);
    }

    const r = this.selection.range;
    if (r.top !== r.bottom || r.left !== r.right) {
      const targets = [];
      for (let col = r.left; col <= r.right; col++) {
        let hasNumber = false;
        for (let row = r.top; row <= r.bottom; row++) {
          if (this._isNumberAt(row, col)) { hasNumber = true; break; }
        }
        if (hasNumber) {
          targets.push({
            row: r.bottom + 1, col,
            formula: '=' + fn + '(' + ref(r.top, col) + ':' + ref(r.bottom, col) + ')',
          });
        }
      }
      if (!targets.length) return this;
      return this._edit('autosum', null, targets.map(({ row, col }) => ({ row, col })), () => {
        for (const t of targets) this._setCell(t.row, t.col, t.formula);
        this.selection.collapseTo(r.bottom + 1, targets[0].col);
        this.selection.extendTo(r.bottom + 1, targets[targets.length - 1].col);
        return this;
      });
    }

    const { row, col } = this.selection.active;
    let start = null;
    let end = null;
    let targetCol = true;
    if (row > 0 && this._isNumberAt(row - 1, col)) {
      end = row - 1;
      start = end;
      while (start > 0 && this._isNumberAt(start - 1, col)) start -= 1;
    } else if (col > 0 && this._isNumberAt(row, col - 1)) {
      targetCol = false;
      end = col - 1;
      start = end;
      while (start > 0 && this._isNumberAt(row, start - 1)) start -= 1;
    } else {
      return this;
    }

    const range = targetCol
      ? ref(start, col) + ':' + ref(end, col)
      : ref(row, start) + ':' + ref(row, end);
    return this.setCell(row, col, '=' + fn + '(' + range + ')');
  }

  // ---- whole columns and rows --------------------------------------------

  /**
   * Select a whole column, bounded by the used range rather than the
   * million-row sheet: what "the column" means to every operation we offer
   * (format, clear, sort) is the part of it that exists, and a bounded
   * selection keeps those operations bounded too.
   */
  selectColumn(col, { extend = false } = {}) {
    this.history.break();
    const maxRow = Math.max(this.bounds.maxRow, this.selection.range.bottom);
    const anchorCol = extend ? this.selection.anchor.col : col;
    this.selection.collapseTo(0, anchorCol);
    this.selection.extendTo(maxRow, col);
    return this;
  }

  selectRow(row, { extend = false } = {}) {
    this.history.break();
    const maxCol = Math.max(this.bounds.maxCol, this.selection.range.right);
    const anchorRow = extend ? this.selection.anchor.row : row;
    this.selection.collapseTo(anchorRow, 0);
    this.selection.extendTo(row, maxCol);
    return this;
  }

  // ---- sorting -----------------------------------------------------------

  /** The contiguous block of data around one cell — Excel's "current region". */
  _currentRegion(row, col) {
    const { maxRow, maxCol } = this.bounds;
    const rowHasFill = (r, left, right) => {
      for (let c = left; c <= right; c++) if (this.isFilled(r, c)) return true;
      return false;
    };
    const colHasFill = (c, top, bottom) => {
      for (let r = top; r <= bottom; r++) if (this.isFilled(r, c)) return true;
      return false;
    };
    let top = row;
    let bottom = row;
    let left = col;
    let right = col;
    let grew = true;
    while (grew) {
      grew = false;
      if (top > 0 && rowHasFill(top - 1, left, right)) { top -= 1; grew = true; }
      if (bottom < maxRow && rowHasFill(bottom + 1, left, right)) { bottom += 1; grew = true; }
      if (left > 0 && colHasFill(left - 1, top, bottom)) { left -= 1; grew = true; }
      if (right < maxCol && colHasFill(right + 1, top, bottom)) { right += 1; grew = true; }
    }
    return { top, bottom, left, right };
  }

  /**
   * Whether a range's first row reads as a HEADER: every cell in it is
   * non-blank text, and at least one column turns numeric right below it.
   * Only consulted when the range was auto-expanded — a person who selected
   * rows explicitly meant those rows.
   */
  _looksLikeHeader(range) {
    if (range.top === range.bottom) return false;
    let numericBelow = false;
    for (let c = range.left; c <= range.right; c++) {
      const head = this.calc.getValue(this.activeSheet, range.top, c);
      if (typeof head !== 'string' || head === '') return false;
      if (typeof this.calc.getValue(this.activeSheet, range.top + 1, c) === 'number') numericBelow = true;
    }
    return numericBelow;
  }

  /**
   * Sort the selection's rows by the active cell's column.
   *
   * A single cell first expands to its current region and skips a detected
   * header row; a real multi-cell selection is sorted exactly as chosen.
   * Ordering is the spreadsheet's: numbers, then text (case-insensitive),
   * then logicals, then errors — and blanks last in BOTH directions. The
   * sort is stable, formulas shift by how far their row moved, styles move
   * with their cells, and the whole thing is one undo step.
   */
  sortSelection({ ascending = true } = {}) {
    let range = { ...this.selection.range };
    const keyCol = Math.min(Math.max(this.selection.active.col, range.left), range.right);
    if (range.top === range.bottom && range.left === range.right) {
      range = this._currentRegion(range.top, range.left);
      if (this._looksLikeHeader(range)) range = { ...range, top: range.top + 1 };
    }
    if (range.top === range.bottom) return this;

    const rows = [];
    for (let r = range.top; r <= range.bottom; r++) {
      const cells = [];
      for (let c = range.left; c <= range.right; c++) {
        cells.push({ input: this.calc.getInput(this.activeSheet, r, c), styleIndex: this._styleIndexAt(this.activeSheet, r, c) });
      }
      rows.push({ from: r, cells, key: this.calc.getValue(this.activeSheet, r, keyCol) });
    }

    const kind = (v) => {
      if (v === '' || v === null || v === undefined) return 4; // blanks last, always
      if (typeof v === 'number') return 0;
      if (typeof v === 'string') return 1;
      if (typeof v === 'boolean') return 2;
      return 3; // errors
    };
    const dir = ascending ? 1 : -1;
    const order = rows.map((rec, i) => ({ rec, i })).sort((a, b) => {
      const ka = kind(a.rec.key);
      const kb = kind(b.rec.key);
      if (ka === 4 || kb === 4) return ka === kb ? a.i - b.i : ka - kb; // blanks sink regardless of direction
      if (ka !== kb) return (ka - kb) * dir;
      let cmp = 0;
      if (ka === 0) cmp = a.rec.key - b.rec.key;
      else if (ka === 1) {
        const x = a.rec.key.toUpperCase();
        const y = b.rec.key.toUpperCase();
        cmp = x < y ? -1 : x > y ? 1 : 0;
      } else if (ka === 2) cmp = (a.rec.key ? 1 : 0) - (b.rec.key ? 1 : 0);
      if (cmp !== 0) return cmp * dir;
      return a.i - b.i; // stable
    });

    if (order.every(({ rec }, i) => rec.from === range.top + i)) return this;

    const touched = [];
    for (let r = range.top; r <= range.bottom; r++) {
      for (let c = range.left; c <= range.right; c++) touched.push({ row: r, col: c });
    }
    return this._edit('sort', null, touched, () => {
      order.forEach(({ rec }, i) => {
        const to = range.top + i;
        const dr = to - rec.from;
        rec.cells.forEach((cell, ci) => {
          const col = range.left + ci;
          const input = typeof cell.input === 'string' && cell.input.startsWith('=')
            ? shiftFormula(cell.input, dr, 0) : cell.input;
          this._setCell(to, col, input ?? '');
          if ((this._styleIndexAt(this.activeSheet, to, col) ?? null) !== (cell.styleIndex ?? null)) {
            this._setStyleIndex(this.activeSheet, to, col, cell.styleIndex ?? null);
          }
        });
      });
      this.selection.collapseTo(range.top, range.left);
      this.selection.extendTo(range.bottom, range.right);
      return this;
    });
  }

  // ---- autofill ----------------------------------------------------------

  /**
   * Fill from the current selection into `target` — the fill handle's drop.
   *
   * `target` must contain the selection and extend it along exactly ONE axis;
   * anything else is a refused gesture, not a guess. Per lane (a column when
   * filling vertically, a row when horizontally):
   *
   *   - sources walking a month or weekday name list -> the list continues,
   *     wrapping, keeping the source's case (Jan -> Feb; MON -> TUE)
   *   - one DATE-formatted source -> the next days (a lone number copies;
   *     the format is what says the cell means a date)
   *   - dates on the same day of consecutive months (or month-ends) -> the
   *     months continue, clamped as Excel clamps (Jan 31 -> Feb 28)
   *   - all sources numeric, two or more  -> linear series, continuing the
   *     average step (1,2 -> 3,4; 10,20 -> 30,40) — which also covers
   *     constant-day date series like weekly meetings
   *   - one source, text ending in a number -> the number counts (Item 1 ->
   *     Item 2, Item 3), the way every list in a spreadsheet gets made
   *   - anything else -> repeat the sources cyclically, formulas SHIFTING
   *     their relative references to where they land — which is what makes
   *     dragging `=B2*1.17` down a column do what forty years of muscle
   *     memory expects
   *
   * Styles travel cyclically in every case. One undo step.
   */
  fill(target) {
    const t = {
      top: Number(target?.top), left: Number(target?.left),
      bottom: Number(target?.bottom), right: Number(target?.right),
    };
    for (const k of ['top', 'left', 'bottom', 'right']) {
      if (!Number.isInteger(t[k]) || t[k] < 0) throw new Error('bad fill target: ' + k);
    }
    const s = this.selection.range;
    const sameCols = t.left === s.left && t.right === s.right;
    const sameRows = t.top === s.top && t.bottom === s.bottom;
    let direction = null;
    if (sameCols && t.top === s.top && t.bottom > s.bottom) direction = 'down';
    else if (sameCols && t.bottom === s.bottom && t.top < s.top) direction = 'up';
    else if (sameRows && t.left === s.left && t.right > s.right) direction = 'right';
    else if (sameRows && t.right === s.right && t.left < s.left) direction = 'left';
    if (!direction) {
      if (t.top === s.top && t.bottom === s.bottom && t.left === s.left && t.right === s.right) return this;
      throw new Error('a fill extends the selection along one axis');
    }

    const vertical = direction === 'down' || direction === 'up';
    const outward = direction === 'down' || direction === 'right' ? 1 : -1;
    const lanes = [];
    const laneKeys = vertical
      ? Array.from({ length: s.right - s.left + 1 }, (_, i) => s.left + i)
      : Array.from({ length: s.bottom - s.top + 1 }, (_, i) => s.top + i);

    for (const lane of laneKeys) {
      // Sources ordered in the DIRECTION of the fill, so "the last one" is
      // always the one nearest the new cells and a series continues from it.
      const positions = [];
      const from = vertical ? s.top : s.left;
      const to = vertical ? s.bottom : s.right;
      for (let p = from; p <= to; p++) positions.push(p);
      if (outward < 0) positions.reverse();

      lanes.push({
        lane,
        sources: positions.map((p) => {
          const row = vertical ? p : lane;
          const col = vertical ? lane : p;
          return { row, col, input: this.calc.getInput(this.activeSheet, row, col), styleIndex: this._styleIndexAt(this.activeSheet, row, col) };
        }),
      });
    }

    // The new cells, from the selection's edge outward.
    const start = vertical
      ? (outward > 0 ? s.bottom + 1 : s.top - 1)
      : (outward > 0 ? s.right + 1 : s.left - 1);
    const end = vertical
      ? (outward > 0 ? t.bottom : t.top)
      : (outward > 0 ? t.right : t.left);
    const count = Math.abs(end - start) + 1;

    const writes = [];
    for (const { lane, sources } of lanes) {
      const n = sources.length;
      const numbers = sources.map((c) => c.input);
      const allNumeric = numbers.every((v) => typeof v === 'number');
      const step = allNumeric && n > 1 ? (numbers[n - 1] - numbers[0]) / (n - 1) : 0;
      const textNum = n === 1 && typeof sources[0].input === 'string' && !sources[0].input.startsWith('=')
        ? /^(.*?)(\d+)$/.exec(sources[0].input) : null;
      // Date lanes are their own series kinds. A single date steps by a day
      // (a single NUMBER copies — the format is what says "this is a date"),
      // and a same-day-of-month walk steps by months, which no linear step
      // over unequal months can express. Constant-day date series (weekly)
      // fall through to the ordinary numeric step, which already gets them.
      const allDates = allNumeric
        && sources.every((c) => isDateFormat(this.formatFor(c.row, c.col)));
      const monthly = allDates && n > 1 ? monthSeries(numbers) : null;
      const named = listSeries(numbers);

      for (let k = 1; k <= count; k++) {
        const pos = start + (k - 1) * outward;
        const row = vertical ? pos : lane;
        const col = vertical ? lane : pos;
        const source = sources[(k - 1) % n];
        let input;
        if (named) input = named(k);
        else if (allDates && n === 1) input = numbers[0] + k;
        else if (monthly) input = monthly(k);
        else if (allNumeric && n > 1) input = numbers[n - 1] + step * k;
        else if (textNum) input = textNum[1] + (Number(textNum[2]) + k);
        else if (typeof source.input === 'string' && source.input.startsWith('=')) {
          input = shiftFormula(source.input, row - source.row, col - source.col);
        } else input = source.input;
        writes.push({ row, col, input, styleIndex: source.styleIndex });
      }
    }

    return this._edit('fill', null, writes.map(({ row, col }) => ({ row, col })), () => {
      for (const w of writes) {
        this._setCell(w.row, w.col, w.input ?? '');
        if ((this._styleIndexAt(this.activeSheet, w.row, w.col) ?? null) !== (w.styleIndex ?? null)) {
          this._setStyleIndex(this.activeSheet, w.row, w.col, w.styleIndex ?? null);
        }
      }
      this.selection.collapseTo(t.top, t.left);
      this.selection.extendTo(t.bottom, t.right);
      return this;
    });
  }

  // ---- the format painter -------------------------------------------------

  /** Pick up the selection's formatting, to be painted somewhere else. */
  markFormatBrush() {
    this.formatBrush = { sheet: this.activeSheet, range: { ...this.selection.range } };
    return this;
  }

  /**
   * Paint the loaded brush onto the current selection, the source pattern
   * TILING across a larger target the way Excel's painter does. Only style
   * INDICES move — nothing is appended to styles.xml, so this is cheap and
   * exactly reproduces the source's look. One undo step; the brush stays
   * loaded (double-click-the-painter semantics are the client's choice).
   */
  paintFormat() {
    const brush = this.formatBrush;
    if (!brush || !this.sheetNames().includes(brush.sheet)) return this;
    const src = brush.range;
    const srcRows = src.bottom - src.top + 1;
    const srcCols = src.right - src.left + 1;
    const target = this.selection.range;

    const writes = [];
    for (let r = target.top; r <= target.bottom; r++) {
      for (let c = target.left; c <= target.right; c++) {
        const sr = src.top + ((r - target.top) % srcRows);
        const sc = src.left + ((c - target.left) % srcCols);
        writes.push({ row: r, col: c, styleIndex: this._styleIndexAt(brush.sheet, sr, sc) });
      }
    }
    if (writes.every((w) => (this._styleIndexAt(this.activeSheet, w.row, w.col) ?? null) === (w.styleIndex ?? null))) {
      return this;
    }
    return this._edit('paint format', null, writes.map(({ row, col }) => ({ row, col })), () => {
      for (const w of writes) {
        if ((this._styleIndexAt(this.activeSheet, w.row, w.col) ?? null) !== (w.styleIndex ?? null)) {
          this._setStyleIndex(this.activeSheet, w.row, w.col, w.styleIndex ?? null);
        }
      }
      return this;
    });
  }

  // ---- what-if -----------------------------------------------------------

  /**
   * Goal Seek: find the input value for one cell that makes a formula cell
   * reach a target — "what does the price have to be for the total to hit
   * 10,000". The whole thing runs against the calculation engine's own
   * dependency graph, trying candidate values with a secant search and
   * restoring the original before committing the answer as ONE ordinary,
   * undoable edit. Failure restores everything and says so; it never leaves
   * a half-tried value in the sheet.
   */
  goalSeek({ set, to, by }) {
    const parseTarget = (text, name) => {
      const clean = String(text ?? '').trim().toUpperCase().replace(/\$/g, '');
      if (!/^[A-Z]{1,3}\d{1,7}$/.test(clean)) throw new Error('Goal Seek needs a cell for "' + name + '", like B4');
      return parseRefPair(clean);
    };
    const [setRow, setCol] = parseTarget(set, 'set cell');
    const [byRow, byCol] = parseTarget(by, 'by changing');
    const goal = Number(to);
    if (!Number.isFinite(goal)) throw new Error('Goal Seek needs a number to aim for');

    const sheet = this.activeSheet;
    const setInput = this.calc.getInput(sheet, setRow, setCol);
    if (typeof setInput !== 'string' || !setInput.startsWith('=')) {
      throw new Error('The cell to set must hold a formula — otherwise just type the value');
    }
    const byInput = this.calc.getInput(sheet, byRow, byCol);
    if (typeof byInput === 'string' && byInput.startsWith('=')) {
      throw new Error('The changing cell must hold a plain value, not a formula');
    }

    const miss = (x) => {
      this.calc.setCell(sheet, byRow, byCol, x);
      this.calc.recalculate();
      const v = this.calc.getValue(sheet, setRow, setCol);
      return typeof v === 'number' ? v - goal : NaN;
    };
    const restore = () => {
      this.calc.setCell(sheet, byRow, byCol, byInput ?? '');
      this.calc.recalculate();
    };

    const tolerance = 1e-7 * Math.max(1, Math.abs(goal));
    let x0 = typeof byInput === 'number' ? byInput : 0;
    let found = null;
    try {
      let f0 = miss(x0);
      if (Number.isNaN(f0)) throw new Error('That formula does not produce a number');
      if (Math.abs(f0) <= tolerance) found = x0;
      let x1 = x0 !== 0 ? x0 * 1.1 : 1;
      let f1 = found === null ? miss(x1) : f0;
      for (let i = 0; found === null && i < 100; i++) {
        if (Number.isNaN(f1)) throw new Error('That formula does not produce a number');
        if (Math.abs(f1) <= tolerance) { found = x1; break; }
        if (f1 === f0) {
          // A flat secant: the target does not respond to this cell here.
          // Step further out before giving up — the response may start away
          // from zero (a capped or floored formula).
          x1 = x1 === 0 ? 1 : x1 * 2 + 1;
          f1 = miss(x1);
          continue;
        }
        const x2 = x1 - f1 * ((x1 - x0) / (f1 - f0));
        if (!Number.isFinite(x2)) break;
        x0 = x1; f0 = f1;
        x1 = x2; f1 = miss(x1);
      }
    } finally {
      restore();
    }
    if (found === null) {
      throw new Error('Goal Seek could not find a value that reaches ' + goal);
    }

    // Snap float noise so the sheet shows 5 rather than 4.999999999999999.
    const rounded = Number(found.toPrecision(12));
    if (Math.abs(miss(rounded)) <= tolerance) found = rounded;
    restore();
    this.setCell(byRow, byCol, found);
    return { value: found };
  }

  /**
   * A Data Table: one formula, evaluated across a grid of substituted inputs.
   *
   * The third of Excel's what-if trio, and the one that needs the engine to
   * do something no formula can express — SUBSTITUTION. The mechanism is
   * Goal Seek's, used differently: put a candidate in the input cell,
   * recalculate, read the answer, and put the original back. One pass per
   * input value rather than per output cell, so an n x m table costs n
   * recalculations, not n*m.
   *
   * The three shapes Excel supports, chosen by which inputs are named:
   *   - column input only: candidates run DOWN the range's first column and
   *     the formulas sit across its top row;
   *   - row input only: candidates run ACROSS the top row, formulas down the
   *     first column;
   *   - both: the corner holds the one formula, candidates on both edges.
   */
  dataTable({ range, rowInput, colInput }) {
    const r = {
      top: Number(range?.top), left: Number(range?.left),
      bottom: Number(range?.bottom), right: Number(range?.right),
    };
    for (const k of ['top', 'left', 'bottom', 'right']) {
      if (!Number.isInteger(r[k]) || r[k] < 0) throw new Error('bad data table range: ' + k);
    }
    if (r.bottom <= r.top || r.right <= r.left) {
      throw new Error('a data table needs at least two rows and two columns');
    }

    const sheet = this.activeSheet;
    const parseInput = (text, label) => {
      if (text === undefined || text === null || String(text).trim() === '') return null;
      const clean = String(text).trim().toUpperCase().replace(/\$/g, '');
      if (!/^[A-Z]{1,3}\d{1,7}$/.test(clean)) {
        throw new Error('the ' + label + ' must be a cell, like B4');
      }
      const [row, col] = parseRefPair(clean);
      const input = this.calc.getInput(sheet, row, col);
      if (typeof input === 'string' && input.startsWith('=')) {
        throw new Error('the ' + label + ' must hold a plain value, not a formula');
      }
      return { row, col, original: input };
    };
    const rowCell = parseInput(rowInput, 'row input cell');
    const colCell = parseInput(colInput, 'column input cell');
    if (!rowCell && !colCell) throw new Error('a data table needs a row input cell, a column input cell, or both');

    const twoWay = Boolean(rowCell && colCell);
    // Where the candidates are, and which cells hold the formulas being run.
    const downValues = [];
    for (let row = r.top + 1; row <= r.bottom; row++) downValues.push({ row, value: this.calc.getValue(sheet, row, r.left) });
    const acrossValues = [];
    for (let col = r.left + 1; col <= r.right; col++) acrossValues.push({ col, value: this.calc.getValue(sheet, r.top, col) });

    let formulaCells;
    if (twoWay) {
      const corner = this.calc.getInput(sheet, r.top, r.left);
      if (typeof corner !== 'string' || !corner.startsWith('=')) {
        throw new Error('a two-way data table needs its formula in the top-left cell of the range');
      }
      formulaCells = [{ row: r.top, col: r.left }];
    } else if (colCell) {
      formulaCells = acrossValues
        .filter(({ col }) => String(this.calc.getInput(sheet, r.top, col) ?? '').startsWith('='))
        .map(({ col }) => ({ row: r.top, col }));
      if (!formulaCells.length) throw new Error('put the formula in the top row, to the right of the input column');
    } else {
      formulaCells = downValues
        .filter(({ row }) => String(this.calc.getInput(sheet, row, r.left) ?? '').startsWith('='))
        .map(({ row }) => ({ row, col: r.left }));
      if (!formulaCells.length) throw new Error('put the formula in the left column, below the input row');
    }

    const setInput = (cell, value) => {
      if (cell) this.calc.setCell(sheet, cell.row, cell.col, value);
    };
    const restore = () => {
      if (rowCell) this.calc.setCell(sheet, rowCell.row, rowCell.col, rowCell.original ?? '');
      if (colCell) this.calc.setCell(sheet, colCell.row, colCell.col, colCell.original ?? '');
      this.calc.recalculate();
    };

    const writes = [];
    try {
      if (twoWay) {
        for (const down of downValues) {
          setInput(colCell, down.value);
          for (const across of acrossValues) {
            setInput(rowCell, across.value);
            this.calc.recalculate();
            writes.push({
              row: down.row,
              col: across.col,
              value: this.calc.getValue(sheet, formulaCells[0].row, formulaCells[0].col),
            });
          }
        }
      } else if (colCell) {
        for (const down of downValues) {
          setInput(colCell, down.value);
          this.calc.recalculate();
          for (const f of formulaCells) {
            writes.push({ row: down.row, col: f.col, value: this.calc.getValue(sheet, f.row, f.col) });
          }
        }
      } else {
        for (const across of acrossValues) {
          setInput(rowCell, across.value);
          this.calc.recalculate();
          for (const f of formulaCells) {
            writes.push({ row: f.row, col: across.col, value: this.calc.getValue(sheet, f.row, f.col) });
          }
        }
      }
    } finally {
      // Whatever happened, the sheet goes back to the values it had. A
      // half-substituted model is a wrong model that looks authoritative.
      restore();
    }

    // The {=TABLE()} marker — `<f t="dataTable">` on the interior's master
    // cell — is what makes the table LIVE in Excel instead of a snapshot.
    // Written only when it would be true: the computed cells must fill the
    // whole interior rectangle (Excel's dialog always does; ours skips
    // non-formula columns), and the master's value must be a number, because
    // a marker shares its cell with a cached <v> and our text cache is an
    // inline string no formula cell may carry. Anything else stays a plain
    // snapshot, exactly what this wrote before the marker existed.
    const interior = { top: r.top + 1, left: r.left + 1, bottom: r.bottom, right: r.right };
    const fullInterior = twoWay
      || (colCell
        ? formulaCells.length === interior.right - interior.left + 1
        : formulaCells.length === interior.bottom - interior.top + 1);
    const master = writes.find((w) => w.row === interior.top && w.col === interior.left);
    let markerAttrs = null;
    if (fullInterior && master && typeof master.value === 'number' && Number.isFinite(master.value)) {
      const refText = ref(interior.top, interior.left) + ':' + ref(interior.bottom, interior.right);
      markerAttrs = twoWay
        ? ' t="dataTable" ref="' + refText + '" dt2D="1" dtr="0"'
          + ' r1="' + ref(rowCell.row, rowCell.col) + '" r2="' + ref(colCell.row, colCell.col) + '"'
        : ' t="dataTable" ref="' + refText + '" dt2D="0" dtr="' + (rowCell ? '1' : '0') + '"'
          + ' r1="' + (rowCell ? ref(rowCell.row, rowCell.col) : ref(colCell.row, colCell.col)) + '"';
    }

    // A parts edit in the structural shape — flush, write the values AND the
    // marker into the sheet part, rebuild — because the marker must live in
    // the FILE (it is the record Excel reads), and the save-time flush of
    // per-cell edits would overwrite it with a plain value. Structural also
    // matches Excel's own rule that part of a data table cannot be edited.
    const sheetPartName = this.workbook.partNameFor(sheet);
    this._edit('data table', null, [], () => {
      this._flushPendingEdits();
      // A candidate that made the formula ERROR writes the error's name as
      // text — a person reading the table needs to see WHICH input failed,
      // and a blank reads as "zero-ish", which is worse than wrong.
      for (const w of writes) {
        this.workbook.setCell(sheet, ref(w.row, w.col), isError(w.value) ? w.value.type : w.value);
      }
      if (markerAttrs) {
        this.workbook.setCellFormulaMarker(sheet, ref(interior.top, interior.left), markerAttrs);
        // Excel recomputes the marked table on open; the hint costs nothing
        // and covers a consumer that would otherwise trust a stale cache.
        this.workbook.setFullCalcOnLoad();
      }
      this.dirtyCells.clear();
      this.styledCells.clear();
      this._structuralDirty = true;
      this._rebuildDerivedState();
      return this;
    }, { parts: [sheetPartName, this.workbook.mainPart], structural: true });
    return { cells: writes.length, live: Boolean(markerAttrs) };
  }

  // ---- authoring conditional formatting -----------------------------------

  /**
   * Write one conditional-formatting rule over the selection — the other
   * half of a feature that could only read. The rule XML is exactly what
   * Excel writes for the same dialog choices; the look presets are Excel's
   * own three (their classic dxf colours), so a rule authored here reads
   * identically there. One undoable parts edit; painting picks it up the
   * moment the conditionals cache re-reads the part.
   *
   * @param {object} spec
   *   {kind:'cellIs', operator, values:[v1,v2?], look} — greaterThan,
   *     lessThan, between, equal
   *   {kind:'containsText', text, look}
   *   {kind:'duplicates', look}
   *   {kind:'colorScale'} — the classic 3-stop red→yellow→green
   *   {kind:'dataBar'}
   */
  addConditionalRule(spec = {}) {
    if (this.protection().sheet) {
      throw protectionError('This sheet is protected — unprotect it before changing its rules.');
    }
    const looks = {
      red: { fontColour: '#9c0006', fill: '#ffc7ce' },
      yellow: { fontColour: '#9c5700', fill: '#ffeb9c' },
      green: { fontColour: '#006100', fill: '#c6efce' },
    };
    const r = this.selection.range;
    const sqref = ref(r.top, r.left)
      + (r.top === r.bottom && r.left === r.right ? '' : ':' + ref(r.bottom, r.right));
    const priority = Math.max(0,
      ...(this.conditionals.get(this.activeSheet) ?? []).map((x) => x.priority)) + 1;

    // A formula VALUE: numbers ride plain, anything else rides quoted — the
    // same coercion the cell grammar applies to a typed comparison.
    const lit = (v) => {
      const n = Number(v);
      return String(v).trim() !== '' && Number.isFinite(n)
        ? String(n)
        : '"' + String(v ?? '').replace(/"/g, '""') + '"';
    };
    const escAttr = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

    const needsLook = spec.kind === 'cellIs' || spec.kind === 'containsText' || spec.kind === 'duplicates';
    const look = needsLook ? looks[spec.look] ?? looks.red : null;

    let rule;
    const build = (dxfId) => {
      const dxf = dxfId === null ? '' : ' dxfId="' + dxfId + '"';
      switch (spec.kind) {
        case 'cellIs': {
          const op = String(spec.operator ?? '');
          if (!['greaterThan', 'lessThan', 'between', 'equal'].includes(op)) {
            throw new Error('"' + op + '" is not a rule this dialog writes');
          }
          const values = (spec.values ?? []).slice(0, op === 'between' ? 2 : 1);
          if (!values.length || (op === 'between' && values.length < 2)) {
            throw new Error('the rule needs ' + (op === 'between' ? 'two values' : 'a value'));
          }
          return '<cfRule type="cellIs"' + dxf + ' priority="' + priority + '" operator="' + op + '">'
            + values.map((v) => '<formula>' + lit(v) + '</formula>').join('')
            + '</cfRule>';
        }
        case 'containsText': {
          const text = String(spec.text ?? '').trim();
          if (!text) throw new Error('the rule needs the text to look for');
          // The formula Excel writes beside the attribute, anchored to the
          // range's own top-left — it shifts per cell like any rule formula.
          return '<cfRule type="containsText"' + dxf + ' priority="' + priority
            + '" operator="containsText" text="' + escAttr(text) + '">'
            + '<formula>NOT(ISERROR(SEARCH(' + lit(text) + ',' + ref(r.top, r.left) + ')))</formula>'
            + '</cfRule>';
        }
        case 'duplicates':
          return '<cfRule type="duplicateValues"' + dxf + ' priority="' + priority + '"/>';
        case 'colorScale':
          return '<cfRule type="colorScale" priority="' + priority + '"><colorScale>'
            + '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>'
            + '<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/>'
            + '</colorScale></cfRule>';
        case 'dataBar':
          return '<cfRule type="dataBar" priority="' + priority + '"><dataBar>'
            + '<cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/>'
            + '</dataBar></cfRule>';
        default:
          throw new Error('"' + String(spec.kind ?? '') + '" is not a rule this dialog writes');
      }
    };
    // Validation runs BEFORE the edit records — a refusal leaves no undo step.
    rule = build(0);

    if (needsLook) this._ensureStylesPart();
    const sheetPartName = this.workbook.partNameFor(this.activeSheet);
    this._edit('conditional format', null, [], () => {
      let dxfId = null;
      if (needsLook) {
        const out = ensureDxf(this.pkg.text('xl/styles.xml'), look);
        if (out.xml !== this.pkg.text('xl/styles.xml')) this.pkg.write_('xl/styles.xml', out.xml);
        dxfId = out.index;
        this.styles = readStyles(this.pkg, { BUILTIN_FORMATS });
        this._stylesDirty = true;
      }
      this.workbook.addConditionalFormatting(this.activeSheet, sqref, build(dxfId));
      const xml = this.workbook.snapshotParts([sheetPartName])[sheetPartName];
      this.conditionals.set(this.activeSheet, readConditionalFormatting(xml, this.styles.theme));
      this._structuralDirty = true;
      return this;
    }, { parts: [sheetPartName], styles: needsLook });
    return this;
  }

  /**
   * Clear rules: the blocks living entirely inside the selection, or every
   * block on the sheet. Removing a RANGE from a block it partly covers would
   * mean rewriting sqref arithmetic Excel itself gets subtle about — the two
   * scopes offered are the two a person actually reaches for.
   */
  clearConditionalRules({ all = false } = {}) {
    if (this.protection().sheet) {
      throw protectionError('This sheet is protected — unprotect it before changing its rules.');
    }
    const sel = this.selection.range;
    const inside = (sqref) => String(sqref).split(/\s+/).every((token) => {
      const [a, b] = token.split(':');
      try {
        const p1 = parseRefPair(a.replace(/\$/g, ''));
        const p2 = b ? parseRefPair(b.replace(/\$/g, '')) : p1;
        return Math.min(p1[0], p2[0]) >= sel.top && Math.max(p1[0], p2[0]) <= sel.bottom
          && Math.min(p1[1], p2[1]) >= sel.left && Math.max(p1[1], p2[1]) <= sel.right;
      } catch {
        return false;
      }
    });
    const sheetPartName = this.workbook.partNameFor(this.activeSheet);
    let removed = 0;
    this._edit('clear rules', null, [], () => {
      removed = this.workbook.removeConditionalFormattings(this.activeSheet,
        all ? () => true : inside);
      const xml = this.workbook.snapshotParts([sheetPartName])[sheetPartName];
      this.conditionals.set(this.activeSheet, readConditionalFormatting(xml, this.styles.theme));
      this._structuralDirty = true;
      return this;
    }, { parts: [sheetPartName] });
    return removed;
  }

  // ---- authoring data validation ------------------------------------------

  /**
   * Write one validation rule over the selection — the dialog half of a
   * feature that could only enforce. The entry is exactly what Excel's
   * dialog writes; enforcement and the in-cell dropdown pick it up the
   * moment the cache re-reads the part, because they always read the file.
   *
   * @param {object} spec
   *   {kind:'list', items:['a','b',...]}         — the in-cell dropdown
   *   {kind:'list', source:'$D$1:$D$9'}          — range-fed
   *   {kind:'whole'|'decimal', operator, values:[v1,v2?]}
   *   plus optional errorTitle/error — the words the refusal speaks in.
   */
  addValidationRule(spec = {}) {
    if (this.protection().sheet) {
      throw protectionError('This sheet is protected — unprotect it before changing its rules.');
    }
    const r = this.selection.range;
    const sqref = ref(r.top, r.left)
      + (r.top === r.bottom && r.left === r.right ? '' : ':' + ref(r.bottom, r.right));
    const escAttr = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
    const escText = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');

    let type;
    let operator = null;
    let f1;
    let f2 = null;
    if (spec.kind === 'list') {
      type = 'list';
      if (spec.source) {
        f1 = String(spec.source).trim();
        if (!/^\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?$/i.test(f1)) {
          throw new Error('the list source must be a range, like $D$1:$D$9');
        }
      } else {
        const items = (spec.items ?? []).map((s) => String(s).trim()).filter(Boolean);
        if (!items.length) throw new Error('the list needs its choices — comma-separated, or a range');
        if (items.some((s) => s.includes('"'))) throw new Error('list choices cannot contain quotes');
        f1 = '"' + items.join(',') + '"';
      }
    } else if (spec.kind === 'whole' || spec.kind === 'decimal') {
      type = spec.kind;
      operator = String(spec.operator ?? 'between');
      if (!['between', 'greaterThanOrEqual', 'lessThanOrEqual', 'greaterThan', 'lessThan', 'equal'].includes(operator)) {
        throw new Error('"' + operator + '" is not an operator this dialog writes');
      }
      const values = (spec.values ?? []).map(Number);
      const need = operator === 'between' ? 2 : 1;
      if (values.length < need || values.slice(0, need).some((n) => !Number.isFinite(n))) {
        throw new Error('the rule needs ' + (need === 2 ? 'two numbers' : 'a number'));
      }
      f1 = String(values[0]);
      if (need === 2) f2 = String(values[1]);
    } else {
      throw new Error('"' + String(spec.kind ?? '') + '" is not a rule this dialog writes');
    }

    const entry = '<dataValidation type="' + type + '"'
      + (operator && operator !== 'between' ? ' operator="' + operator + '"' : '')
      + ' allowBlank="1" showInputMessage="1" showErrorMessage="1"'
      + (spec.errorTitle ? ' errorTitle="' + escAttr(spec.errorTitle) + '"' : '')
      + (spec.error ? ' error="' + escAttr(spec.error) + '"' : '')
      + ' sqref="' + sqref + '">'
      + '<formula1>' + escText(f1) + '</formula1>'
      + (f2 !== null ? '<formula2>' + escText(f2) + '</formula2>' : '')
      + '</dataValidation>';

    const sheetPartName = this.workbook.partNameFor(this.activeSheet);
    this._edit('data validation', null, [], () => {
      this.workbook.addDataValidation(this.activeSheet, entry);
      const xml = this.workbook.snapshotParts([sheetPartName])[sheetPartName];
      this.validations.set(this.activeSheet, readDataValidations(xml));
      this._structuralDirty = true;
      return this;
    }, { parts: [sheetPartName] });
    return this;
  }

  /** Clear validations: entries inside the selection, or the whole sheet's. */
  clearValidationRules({ all = false } = {}) {
    if (this.protection().sheet) {
      throw protectionError('This sheet is protected — unprotect it before changing its rules.');
    }
    const sel = this.selection.range;
    const inside = (sqref) => String(sqref).split(/\s+/).every((token) => {
      const [a, b] = token.split(':');
      try {
        const p1 = parseRefPair(a.replace(/\$/g, ''));
        const p2 = b ? parseRefPair(b.replace(/\$/g, '')) : p1;
        return Math.min(p1[0], p2[0]) >= sel.top && Math.max(p1[0], p2[0]) <= sel.bottom
          && Math.min(p1[1], p2[1]) >= sel.left && Math.max(p1[1], p2[1]) <= sel.right;
      } catch {
        return false;
      }
    });
    const sheetPartName = this.workbook.partNameFor(this.activeSheet);
    let removed = 0;
    this._edit('clear validation', null, [], () => {
      removed = this.workbook.removeDataValidations(this.activeSheet, all ? () => true : inside);
      const xml = this.workbook.snapshotParts([sheetPartName])[sheetPartName];
      this.validations.set(this.activeSheet, readDataValidations(xml));
      this._structuralDirty = true;
      return this;
    }, { parts: [sheetPartName] });
    return removed;
  }

  // ---- inserting drawings -------------------------------------------------

  /** The parts an insert must snapshot: what EXISTS now, so undo restores it. */
  _drawingEditParts() {
    const sheetPartName = this.workbook.partNameFor(this.activeSheet);
    const parts = [sheetPartName, '[Content_Types].xml'];
    const relsName = OoxmlPackage.relsPathFor(sheetPartName);
    if (this.pkg.has(relsName)) parts.push(relsName);
    const drawingRel = this.pkg.rels(sheetPartName).find((r) => String(r.Type).endsWith('/drawing'));
    if (drawingRel) {
      const drawingPart = OoxmlPackage.resolveTarget(sheetPartName, drawingRel.Target);
      if (this.pkg.has(drawingPart)) {
        parts.push(drawingPart);
        const drawingRels = OoxmlPackage.relsPathFor(drawingPart);
        if (this.pkg.has(drawingRels)) parts.push(drawingRels);
      }
    }
    return { sheetPartName, parts };
  }

  /**
   * Insert a preset shape, anchored over the selection — a multi-cell
   * selection is its box, a single cell gets a sensible default. The anchor
   * XML is the creator's own writer (`drawingAnchorXml`), the shapes the
   * Office gate proved real Excel accepts; the paint path draws it the
   * moment the drawings cache refreshes, because rendering existing shapes
   * has worked all along.
   */
  insertShape({ geometry = 'rect', text = '' } = {}) {
    if (!SUPPORTED_GEOMETRY.includes(geometry)) {
      throw new Error('"' + geometry + '" is not a shape this editor draws — pick one of the presets');
    }
    if (this.protection().sheet) {
      throw protectionError('This sheet is protected — unprotect it before inserting objects.');
    }
    const sel = this.selection.range;
    const single = sel.top === sel.bottom && sel.left === sel.right;
    const from = { row: sel.top, col: sel.left };
    const to = single
      ? { row: sel.top + 6, col: sel.left + 3 }
      : { row: sel.bottom + 1, col: sel.right + 1 };
    const { sheetPartName, parts } = this._drawingEditParts();
    this._edit('insert shape', null, [], () => {
      const drawingPart = this.workbook.ensureSheetDrawing(this.activeSheet);
      this.workbook.appendDrawingAnchor(drawingPart, (id) => drawingAnchorXml({
        kind: 'shape',
        id,
        name: 'Shape ' + id,
        geometry,
        fill: 'accent1',
        text: String(text ?? '').trim() || undefined,
        from,
        to,
      }, () => null));
      this.drawings.set(this.activeSheet, this._readDrawings(sheetPartName));
      this._structuralDirty = true;
    }, { parts, tracksNewParts: true });
    return this;
  }

  /**
   * Insert a clustered column chart over the data region around the cursor:
   * headers in the first row name the series, the first column is the
   * categories, the numbers beside them are the values. The chart part is
   * the creator's own writer (`chartPartXml`) — real cell references plus
   * cached values that match the sheet, so Excel re-plots it live and every
   * non-calculating reader paints the numbers it was given.
   */
  insertChart({ title = '', kind = 'column' } = {}) {
    if (!CHART_KINDS.includes(kind)) {
      throw new Error('"' + kind + '" is not a chart kind this editor writes — '
        + CHART_KINDS.join(', '));
    }
    if (this.protection().sheet) {
      throw protectionError('This sheet is protected — unprotect it before inserting objects.');
    }
    const active = this.selection.active;
    const r = this._currentRegion(active.row, active.col);
    if (r.bottom - r.top < 1 || r.right - r.left < 1) {
      throw new Error('Put the cursor in a block of data — headers across the top, '
        + 'categories down the first column, numbers beside them.');
    }
    const sheet = this.activeSheet;
    const q = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : "'" + sheet.replace(/'/g, "''") + "'";
    const abs = (row, col) => '$' + colName(col) + '$' + (row + 1);
    const colRef = (col, top, bottom) => q + '!' + abs(top, col) + ':' + abs(bottom, col);

    const categories = {
      ref: colRef(r.left, r.top + 1, r.bottom),
      values: [],
    };
    for (let row = r.top + 1; row <= r.bottom; row++) {
      categories.values.push(this.displayValue(row, r.left).text);
    }
    const series = [];
    for (let col = r.left + 1; col <= r.right; col++) {
      const values = [];
      for (let row = r.top + 1; row <= r.bottom; row++) {
        const v = this.calc.getValue(sheet, row, col);
        values.push(typeof v === 'number' && Number.isFinite(v) ? v : null);
      }
      series.push({
        name: this.displayValue(r.top, col).text || colName(col),
        nameRef: q + '!' + abs(r.top, col),
        ref: colRef(col, r.top + 1, r.bottom),
        values,
      });
    }

    const from = { row: r.top, col: r.right + 2 };
    const to = { row: r.top + 15, col: r.right + 9 };
    const { sheetPartName, parts } = this._drawingEditParts();
    this._edit('insert chart', null, [], () => {
      const drawingPart = this.workbook.ensureSheetDrawing(this.activeSheet);
      const n = this.pkg.nextPartNumber('xl/charts/', 'chart');
      const chartPart = 'xl/charts/chart' + n + '.xml';
      this.pkg.addPart(chartPart,
        chartPartXml({ kind, title: String(title ?? '').trim() || undefined, categories, series }),
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
      const relId = this.pkg.addRelationshipTo(drawingPart,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
        '../charts/chart' + n + '.xml');
      this.workbook.appendDrawingAnchor(drawingPart, (id) => drawingAnchorXml({
        kind: 'chart',
        id,
        name: String(title ?? '').trim() || 'Chart ' + id,
        from,
        to,
      }, () => relId));
      this.drawings.set(this.activeSheet, this._readDrawings(sheetPartName));
      this._structuralDirty = true;
    }, { parts, tracksNewParts: true });
    return this;
  }

  // ---- defined names ------------------------------------------------------

  /**
   * The workbook's defined names, with the parsed target where it is a plain
   * range. A name whose ref we cannot parse (a formula, an external book) is
   * still listed — it is real and deletable — with a null target.
   */
  names() {
    return this.workbook.definedNames().map(({ name, ref: target }) => ({
      name, ref: target, target: parseDefinedNameRange(target),
    }));
  }

  /** The tables on the active sheet, parsed, for the frame and the filter UI. */
  sheetTables() {
    return this.workbook.tables()
      .filter((t) => t.sheet === this.activeSheet)
      .map((t) => {
        const [a, b] = String(t.ref).split(':');
        const [r1, c1] = parseRefPair(a.replace(/\$/g, ''));
        const [r2, c2] = b ? parseRefPair(b.replace(/\$/g, '')) : [r1, c1];
        return {
          name: t.name,
          ref: t.ref,
          top: Math.min(r1, r2),
          left: Math.min(c1, c2),
          bottom: Math.max(r1, r2),
          right: Math.max(c1, c2),
          headerRows: t.headerRowCount,
          totalsRows: t.totalsRowCount,
          columns: t.columns,
          styleName: t.styleName,
          showRowStripes: t.showRowStripes,
          // Which columns are filtered right now, for the funnel marks.
          filtered: [...this.workbook.tableFilters(t.part).entries()]
            .filter(([, vals]) => vals === null || vals.length)
            .map(([colId]) => colId),
        };
      });
  }

  /** Re-derive the calc model's name table from the workbook part. */
  _syncNames() {
    this.calc.names.clear();
    for (const { name, ref: target } of this.workbook.definedNames()) {
      const parsed = parseDefinedNameRange(target);
      if (parsed) this.calc.defineName(name, parsed);
    }
  }

  /** The current selection as a sheet-qualified absolute ref: `Data!$A$1:$B$3`. */
  _selectionRefText() {
    const r = this.selection.range;
    const sheet = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(this.activeSheet)
      ? this.activeSheet
      : "'" + this.activeSheet.replace(/'/g, "''") + "'";
    const abs = (row, col) => '$' + colName(col) + '$' + (row + 1);
    const range = r.top === r.bottom && r.left === r.right
      ? abs(r.top, r.left)
      : abs(r.top, r.left) + ':' + abs(r.bottom, r.right);
    return sheet + '!' + range;
  }

  /**
   * Name the current selection (or an explicit ref) — Excel's "define name".
   *
   * A defined name is workbook STRUCTURE, not cell state: it goes into
   * workbook.xml immediately, so undo travels as a snapshot of that one part,
   * the same way a merge travels its sheet part.
   */
  defineName(name, refText) {
    const trimmed = String(name ?? '').trim();
    if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(trimmed) || trimmed.length > 255) {
      throw new Error('"' + trimmed + '" is not a valid name — start with a letter, no spaces');
    }
    if (/^[A-Za-z]{1,3}\d+$/.test(trimmed)) {
      throw new Error('"' + trimmed + '" reads as a cell reference — pick another name');
    }
    const target = refText === undefined ? this._selectionRefText() : String(refText).trim();
    if (!parseDefinedNameRange(target)) {
      throw new Error('"' + target + '" is not a range a name can point at — use Sheet!$A$1:$B$3');
    }
    this._edit('define name', null, [], () => {
      this.workbook.setDefinedName(trimmed, target);
      this._syncNames();
      // Empty dirty set means a FULL pass, which is what a new name needs:
      // every #NAME? that was waiting for it comes alive.
      this.calc.recalculate();
      this._structuralDirty = true;
    }, { parts: [this.workbook.mainPart] });
    return this;
  }

  /** Remove a defined name. Cells using it show #NAME? on the next pass. */
  deleteName(name) {
    const trimmed = String(name ?? '').trim();
    if (!this.workbook.definedNames().some((n) => n.name === trimmed)) {
      throw new Error('no defined name "' + trimmed + '"');
    }
    this._edit('delete name', null, [], () => {
      this.workbook.deleteDefinedName(trimmed);
      this._syncNames();
      this.calc.recalculate();
      this._structuralDirty = true;
    }, { parts: [this.workbook.mainPart] });
    return this;
  }

  // ---- data validation ----------------------------------------------------

  /** The validation rule governing a cell, or null. First rule wins, as Excel. */
  validationAt(row, col) {
    return (this.validations.get(this.activeSheet) ?? []).find((rule) =>
      rule.ranges.some((r) => row >= r.top && row <= r.bottom && col >= r.left && col <= r.right),
    ) ?? null;
  }

  /**
   * Evaluate a validation formula against the live model. A bare range (or a
   * defined name holding one) is fetched WHOLE through the resolver — the
   * evaluator's top level collapses a lone range to its first cell, which is
   * right for a cell formula and wrong for a list's members.
   */
  _validationValue(formulaText, row, col) {
    if (formulaText === null || formulaText === undefined) return null;
    const resolver = this.calc.resolver();
    try {
      const ast = parse('=' + formulaText);
      let target = null;
      if (ast.type === 'range') target = { sheet: ast.sheet ?? ast.start.sheet, start: ast.start, end: ast.end };
      else if (ast.type === 'name') {
        const resolved = resolver.getName(ast.name);
        if (resolved && resolved.start && resolved.end) target = resolved;
      }
      if (target) {
        const sheet = target.sheet ?? this.activeSheet;
        const used = resolver.usedBounds(sheet);
        return resolver.getRange(sheet, {
          row: Math.min(target.start.row, target.end.row),
          col: Math.min(target.start.col, target.end.col),
        }, {
          row: Math.min(Math.max(target.start.row, target.end.row), used.maxRow),
          col: Math.min(Math.max(target.start.col, target.end.col), used.maxCol),
        });
      }
    } catch {
      return null;
    }
    return calculate('=' + formulaText, resolver, { sheet: this.activeSheet, row, col });
  }

  /**
   * A list rule's members, as display text. Inline members arrive quoted
   * (`"Yes,No"`), a range arrives as its values; either way the dropdown and
   * the membership check read the same array.
   */
  validationOptions(rule, row, col) {
    if (rule.type !== 'list' || rule.formula1 === null) return [];
    const resolved = this._validationValue(rule.formula1, row, col);
    if (isError(resolved) || resolved === null) return [];
    if (Array.isArray(resolved)) {
      return resolved.flat(Infinity).filter((v) => v !== '' && v !== null && v !== undefined).map(String);
    }
    return String(resolved).split(',').map((s) => s.trim()).filter(Boolean);
  }

  /**
   * Would this input break the cell's validation rule? Null when it may go in;
   * otherwise the message to show, preferring the rule author's own words.
   *
   * Scope is deliberately Excel's: DIRECT entry is checked, while paste and
   * fill overwrite rules silently — surprising, but matching it means a
   * workbook behaves identically in both editors. `custom` rules would need
   * the candidate value substituted into the model before evaluation, so they
   * are preserved and not enforced, and only `stop` blocks — the other two
   * error styles are dialogues this server cannot hold.
   */
  checkValidation(row, col, rawInput) {
    const rule = this.validationAt(row, col);
    if (!rule || rule.errorStyle !== 'stop' || rule.type === 'custom' || rule.type === 'none') return null;

    const raw = String(rawInput ?? '');
    if (raw.trim() === '') return rule.allowBlank ? null : this._validationMessage(rule, 'this cell may not be left empty');

    let value;
    if (raw.startsWith('=')) {
      value = calculate(raw, this.calc.resolver(), { sheet: this.activeSheet, row, col });
      // A formula that errors shows its error in the cell — that is its own,
      // louder signal, and refusing the entry would hide what went wrong.
      if (isError(value)) return null;
    } else {
      value = coerceInput(raw);
    }

    if (rule.type === 'list') {
      const options = this.validationOptions(rule, row, col);
      if (!options.length) return null; // an unresolvable list cannot judge
      const text = String(value);
      if (options.some((o) => o.localeCompare(text, undefined, { sensitivity: 'accent' }) === 0)) return null;
      return this._validationMessage(rule, 'pick one of: ' + options.join(', '));
    }

    const n1 = Number(this._validationValue(rule.formula1, row, col));
    const n2 = Number(this._validationValue(rule.formula2, row, col));

    if (rule.type === 'textLength') {
      return this._checkNumeric(rule, String(value).length, n1, n2, 'text of length');
    }

    // whole, decimal, date and time all judge a NUMBER — dates and times are
    // serials by the time they are values.
    if (typeof value !== 'number') {
      return this._validationMessage(rule, 'this cell needs ' + NUMERIC_KIND[rule.type]);
    }
    if (rule.type === 'whole' && !Number.isInteger(value)) {
      return this._validationMessage(rule, 'this cell needs a whole number');
    }
    return this._checkNumeric(rule, value, n1, n2, NUMERIC_KIND[rule.type]);
  }

  _checkNumeric(rule, actual, n1, n2, kind) {
    if (!Number.isFinite(n1)) return null; // a bound we cannot resolve cannot judge
    const ok = {
      between: () => Number.isFinite(n2) && actual >= Math.min(n1, n2) && actual <= Math.max(n1, n2),
      notBetween: () => Number.isFinite(n2) && (actual < Math.min(n1, n2) || actual > Math.max(n1, n2)),
      equal: () => actual === n1,
      notEqual: () => actual !== n1,
      greaterThan: () => actual > n1,
      lessThan: () => actual < n1,
      greaterThanOrEqual: () => actual >= n1,
      lessThanOrEqual: () => actual <= n1,
    }[rule.operator];
    if (!ok || ok()) return null;
    const bounds = {
      between: 'between ' + n1 + ' and ' + n2,
      notBetween: 'outside ' + n1 + ' to ' + n2,
      equal: 'equal to ' + n1,
      notEqual: 'anything but ' + n1,
      greaterThan: 'greater than ' + n1,
      lessThan: 'less than ' + n1,
      greaterThanOrEqual: 'at least ' + n1,
      lessThanOrEqual: 'at most ' + n1,
    }[rule.operator];
    return this._validationMessage(rule, 'this cell needs ' + kind + ' ' + bounds);
  }

  /** The rule author's words when they wrote any; a built sentence otherwise. */
  _validationMessage(rule, fallback) {
    const text = rule.error ?? fallback;
    return rule.errorTitle ? rule.errorTitle + ': ' + text : text;
  }

  // ---- conditional formatting ---------------------------------------------

  /**
   * What conditional formatting says about one cell: `{font, fill, bar}` or
   * null. Rules run in priority order; the first to claim each aspect keeps
   * it, and `stopIfTrue` ends the walk — the spec's own cascade.
   *
   * `cache` is per render pass: a rule's range aggregates (values, min, max,
   * percentiles, duplicate counts) are computed once for the first visible
   * cell that needs them, not once per cell.
   */
  _conditionalStyle(row, col, cache) {
    const rules = this.conditionals.get(this.activeSheet) ?? [];
    if (!rules.length) return null;
    const out = { font: null, fill: null, bar: null, icon: null, hideValue: false };
    let any = false;
    for (const rule of rules) {
      const range = rule.ranges.find((r) => row >= r.top && row <= r.bottom && col >= r.left && col <= r.right);
      if (!range) continue;
      const hit = this._cfEvaluate(rule, range, row, col, cache);
      if (!hit) continue;
      any = true;
      if (hit.bar && !out.bar) out.bar = hit.bar;
      if (hit.fill && !out.fill) out.fill = hit.fill;
      if (hit.font && !out.font) out.font = hit.font;
      if (hit.icon && !out.icon) { out.icon = hit.icon; out.hideValue = Boolean(hit.hideValue); }
      if (rule.stopIfTrue) break;
    }
    return any ? out : null;
  }

  /** The numeric values of a rule's whole sqref, with the stats rules ask for. */
  _cfStats(rule, cache) {
    let stats = cache.get(rule);
    if (stats) return stats;
    const numbers = [];
    const textCounts = new Map();
    for (const r of rule.ranges) {
      const bottom = Math.min(r.bottom, r.top + 65535); // a whole-column rule must not walk a million rows
      for (let row = r.top; row <= bottom; row++) {
        for (let col = r.left; col <= r.right; col++) {
          const v = this.calc.getValue(this.activeSheet, row, col);
          if (typeof v === 'number') numbers.push(v);
          if (v !== '' && v !== null && v !== undefined && !isError(v)) {
            const key = String(v).toLowerCase();
            textCounts.set(key, (textCounts.get(key) ?? 0) + 1);
          }
        }
      }
    }
    const sorted = [...numbers].sort((a, b) => a - b);
    stats = {
      numbers,
      sorted,
      min: sorted.length ? sorted[0] : 0,
      max: sorted.length ? sorted[sorted.length - 1] : 0,
      average: numbers.length ? numbers.reduce((s, v) => s + v, 0) / numbers.length : 0,
      textCounts,
    };
    cache.set(rule, stats);
    return stats;
  }

  /** A cfRule formula, shifted to the asking cell exactly as Excel copies it. */
  _cfFormula(text, range, row, col) {
    if (text === null || text === undefined || text === '') return null;
    // Formulas in a rule are written relative to the range's TOP-LEFT cell;
    // for any other cell they shift like a copied formula, $ anchors and all.
    const shifted = shiftFormula('=' + text, row - range.top, col - range.left);
    return calculate(shifted, this.calc.resolver(), { sheet: this.activeSheet, row, col });
  }

  /** Resolve a colour-scale/data-bar stop to a number over the rule's values. */
  _cfStop(cfvo, stats, range, row, col) {
    switch (cfvo?.type) {
      case 'min': return stats.min;
      case 'max': return stats.max;
      case 'num': return Number(this._cfFormula(cfvo.val, range, row, col));
      case 'formula': return Number(this._cfFormula(cfvo.val, range, row, col));
      case 'percent': return stats.min + (Number(cfvo.val) / 100) * (stats.max - stats.min);
      case 'percentile': {
        const p = Number(cfvo.val) / 100;
        if (!stats.sorted.length) return 0;
        const at = p * (stats.sorted.length - 1);
        const lo = Math.floor(at);
        const hi = Math.ceil(at);
        return stats.sorted[lo] + (stats.sorted[hi] - stats.sorted[lo]) * (at - lo);
      }
      default: return NaN;
    }
  }

  /** Does this rule fire for this cell, and with what look? */
  _cfEvaluate(rule, range, row, col, cache) {
    const value = this.calc.getValue(this.activeSheet, row, col);
    const dxf = rule.dxfId !== null ? (this.styles.dxfs ?? [])[rule.dxfId] ?? null : null;
    const claim = () => ({ font: dxf?.font ?? null, fill: dxf?.fill ?? null, bar: null });

    switch (rule.type) {
      case 'cellIs': {
        const f1 = this._cfFormula(rule.formulas[0], range, row, col);
        if (f1 === null || isError(f1)) return null;
        const cmp = (a, b) => {
          const c = compareValues(a, b);
          return isError(c) ? null : c;
        };
        const c1 = cmp(value, f1);
        if (c1 === null) return null;
        let fired = false;
        switch (rule.operator) {
          case 'greaterThan': fired = c1 > 0; break;
          case 'lessThan': fired = c1 < 0; break;
          case 'greaterThanOrEqual': fired = c1 >= 0; break;
          case 'lessThanOrEqual': fired = c1 <= 0; break;
          case 'equal': fired = c1 === 0; break;
          case 'notEqual': fired = c1 !== 0; break;
          case 'between': case 'notBetween': {
            const f2 = this._cfFormula(rule.formulas[1], range, row, col);
            if (f2 === null || isError(f2)) return null;
            const c2 = cmp(value, f2);
            if (c2 === null) return null;
            const inside = (c1 >= 0 && c2 <= 0) || (c1 <= 0 && c2 >= 0);
            fired = rule.operator === 'between' ? inside : !inside;
            break;
          }
          default: return null;
        }
        return fired ? claim() : null;
      }

      case 'expression': {
        const result = this._cfFormula(rule.formulas[0], range, row, col);
        const fired = result === true || (typeof result === 'number' && result !== 0);
        return fired ? claim() : null;
      }

      case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith': {
        const needle = String(rule.text ?? '').toLowerCase();
        if (!needle) return null;
        const hay = String(value ?? '').toLowerCase();
        const found = rule.type === 'beginsWith' ? hay.startsWith(needle)
          : rule.type === 'endsWith' ? hay.endsWith(needle)
            : hay.includes(needle);
        const fired = rule.type === 'notContainsText' ? !found && hay !== '' : found;
        return fired ? claim() : null;
      }

      case 'containsBlanks': return String(value ?? '').trim() === '' ? claim() : null;
      case 'notContainsBlanks': return String(value ?? '').trim() !== '' ? claim() : null;
      case 'containsErrors': return isError(value) ? claim() : null;
      case 'notContainsErrors': return !isError(value) ? claim() : null;

      case 'duplicateValues': case 'uniqueValues': {
        if (value === '' || value === null || value === undefined) return null;
        const count = this._cfStats(rule, cache).textCounts.get(String(value).toLowerCase()) ?? 0;
        const fired = rule.type === 'duplicateValues' ? count > 1 : count === 1;
        return fired ? claim() : null;
      }

      case 'top10': {
        if (typeof value !== 'number') return null;
        const stats = this._cfStats(rule, cache);
        const n = rule.percent
          ? Math.max(1, Math.round((rule.rank / 100) * stats.sorted.length))
          : rule.rank;
        const fired = rule.bottom
          ? value <= (stats.sorted[Math.min(n, stats.sorted.length) - 1] ?? -Infinity)
          : value >= (stats.sorted[Math.max(0, stats.sorted.length - n)] ?? Infinity);
        return fired ? claim() : null;
      }

      case 'aboveAverage': {
        if (typeof value !== 'number') return null;
        const { average } = this._cfStats(rule, cache);
        const fired = rule.aboveAverage
          ? (rule.equalAverage ? value >= average : value > average)
          : (rule.equalAverage ? value <= average : value < average);
        return fired ? claim() : null;
      }

      case 'colorScale': {
        if (typeof value !== 'number' || rule.colours.length < 2) return null;
        const stats = this._cfStats(rule, cache);
        const stops = rule.cfvos.map((v) => this._cfStop(v, stats, range, row, col));
        if (stops.some((s) => !Number.isFinite(s))) return null;
        let colour;
        if (stops.length >= 3 && rule.colours.length >= 3) {
          colour = value <= stops[1]
            ? lerpColour(rule.colours[0], rule.colours[1], fraction(value, stops[0], stops[1]))
            : lerpColour(rule.colours[1], rule.colours[2], fraction(value, stops[1], stops[2]));
        } else {
          colour = lerpColour(rule.colours[0], rule.colours[1], fraction(value, stops[0], stops[stops.length - 1]));
        }
        return colour ? { font: null, fill: { pattern: 'solid', colour }, bar: null } : null;
      }

      case 'dataBar': {
        if (typeof value !== 'number') return null;
        const stats = this._cfStats(rule, cache);
        const lo = this._cfStop(rule.cfvos[0] ?? { type: 'min' }, stats, range, row, col);
        const hi = this._cfStop(rule.cfvos[1] ?? { type: 'max' }, stats, range, row, col);
        if (!Number.isFinite(lo) || !Number.isFinite(hi)) return null;
        return {
          font: null, fill: null,
          bar: { fraction: fraction(value, lo, hi), colour: rule.colours[0] ?? '#638ec6' },
        };
      }

      case 'iconSet': {
        if (typeof value !== 'number' || rule.cfvos.length < 2) return null;
        const stats = this._cfStats(rule, cache);
        const stops = rule.cfvos.map((v) => this._cfStop(v, stats, range, row, col));
        if (stops.some((s) => !Number.isFinite(s))) return null;
        // Ascending thresholds: the highest stop the value reaches names the
        // icon. cfvo[0] is the floor (usually percent 0), so index 0 is what
        // a value below every other stop earns.
        const n = rule.cfvos.length;
        let index = 0;
        for (let i = 1; i < stops.length; i++) if (value >= stops[i]) index = i;
        if (rule.iconReverse) index = (n - 1) - index;
        return {
          font: null,
          fill: null,
          bar: null,
          icon: { set: rule.iconSet, index, count: n },
          hideValue: !rule.iconShowValue,
        };
      }

      case 'timePeriod': {
        if (typeof value !== 'number') return null;
        // The same UTC day-arithmetic TODAY() uses, so a rule and the
        // function it usually guards agree about what day it is. Excel's
        // weeks for these rules run Sunday to Saturday.
        const todaySerial = Math.floor(dateToSerial(this._now ? this._now() : new Date()));
        const d = Math.floor(value);
        const weekStart = todaySerial - serialToDate(todaySerial).getUTCDay();
        const monthOf = (serial, shift = 0) => {
          const at = serialToDate(serial);
          return at.getUTCFullYear() * 12 + at.getUTCMonth() + shift;
        };
        let fired;
        switch (rule.timePeriod) {
          case 'today': fired = d === todaySerial; break;
          case 'yesterday': fired = d === todaySerial - 1; break;
          case 'tomorrow': fired = d === todaySerial + 1; break;
          case 'last7Days': fired = d <= todaySerial && d > todaySerial - 7; break;
          case 'thisWeek': fired = d >= weekStart && d < weekStart + 7; break;
          case 'lastWeek': fired = d >= weekStart - 7 && d < weekStart; break;
          case 'nextWeek': fired = d >= weekStart + 7 && d < weekStart + 14; break;
          case 'thisMonth': fired = monthOf(d) === monthOf(todaySerial); break;
          case 'lastMonth': fired = monthOf(d) === monthOf(todaySerial, -1); break;
          case 'nextMonth': fired = monthOf(d) === monthOf(todaySerial, 1); break;
          // A period this editor does not know paints nothing rather than
          // a wrong verdict; the rule rides through untouched either way.
          default: return null;
        }
        return fired ? claim() : null;
      }

      default: return null;
    }
  }

  // ---- pivot tables -------------------------------------------------------

  /**
   * Every pivot in the workbook, with the reason where one cannot be
   * recomputed. Cached per workbook state: reading walks several parts, and
   * `render()` asks on every frame.
   */
  pivots() {
    if (!this._pivots) this._pivots = readPivots(this.workbook);
    return this._pivots;
  }

  /**
   * Create a pivot table over a source range and render it at once.
   *
   * The source defaults to the current selection expanded to its region, the
   * way Insert > PivotTable behaves — a person who selected one cell in a
   * data block means the block.
   */
  createPivot({ name, source, target, rowFields = [], colFields = [], dataFields = [] }) {
    const area = source
      ? { ...source, sheet: source.sheet ?? this.activeSheet }
      : {
        ...this._currentRegion(this.selection.active.row, this.selection.active.col),
        sheet: this.activeSheet,
      };
    const where = target ?? {
      sheet: this.activeSheet,
      row: area.bottom + 2,
      col: area.left,
    };
    if (!this.sheetNames().includes(where.sheet)) throw new Error('no sheet "' + where.sheet + '"');

    const parts = [...new Set([
      '[Content_Types].xml',
      this.workbook.mainPart,
      this.workbook.partNameFor(where.sheet),
      this.workbook.partNameFor(area.sheet),
      OoxmlPackage.relsPathFor(this.workbook.mainPart),
      OoxmlPackage.relsPathFor(this.workbook.partNameFor(where.sheet)),
    ].filter((p) => this.pkg.has(p)))];

    const readCell = (sheet, row, col) => this.calc.getValue(sheet, row, col);
    // Refuse BEFORE recording history, the same rule a structural edit
    // follows: a rejected gesture must leave no undo step behind. Planning
    // is where every refusal lives, and it touches nothing.
    const plan = planPivot(this.workbook, {
      name, source: area, target: where, rowFields, colFields, dataFields, readCell,
    });

    let shape = null;
    // STRUCTURAL, for the reason that mechanism exists: this edit creates
    // parts and writes cells on a sheet that may not be the active one, and
    // structural undo restores the parts and rebuilds everything derived
    // from them — which is precisely "put the file back as it was".
    this._edit('create pivot', null, [], () => {
      this._flushPendingEdits();
      const created = createPivot(this.workbook, {
        name, source: area, target: where, rowFields, colFields, dataFields, readCell,
      }, plan);
      const grid = computePivot(created, readCell);
      for (const c of grid.cells) {
        this.workbook.setCell(where.sheet, ref(c.row, c.col), c.value === '' ? null : c.value);
      }
      updatePivotLocation(this.workbook, created, grid.area, {
        headerRows: grid.firstHeaderRow,
        dataRow: grid.firstDataRow,
        dataCol: grid.firstDataCol,
      });
      this.dirtyCells.clear();
      this.styledCells.clear();
      this._structuralDirty = true;
      this._rebuildDerivedState();
      shape = { rows: grid.height, cols: grid.width, ref: areaRef(grid.area), sheet: where.sheet };
      return this;
    }, { parts, structural: true, tracksNewParts: true });
    return shape;
  }

  /**
   * Recompute a pivot from the LIVE source values and rewrite its rectangle.
   *
   * This is the gesture that was missing: a pivot's numbers are ordinary
   * cells, so editing the source used to leave last week's totals sitting
   * there looking authoritative. One undo step covers the whole rectangle,
   * cells the pivot no longer covers are cleared, and the stored location
   * follows when the shape changes — which is what Excel's own refresh does.
   */
  refreshPivot(name) {
    const pivot = this.pivots().find((p) => p.name === String(name ?? '').trim());
    if (!pivot) throw new Error('no pivot table "' + String(name ?? '').trim() + '"');
    if (pivot.unsupported) {
      throw new Error('"' + pivot.name + '" cannot be refreshed here: ' + pivot.unsupported);
    }
    if (pivot.sheet !== this.activeSheet) {
      throw new Error('"' + pivot.name + '" is on sheet ' + pivot.sheet);
    }

    const grid = computePivot(pivot, (sheet, row, col) => this.calc.getValue(sheet, row, col));

    // Every cell the pivot touches, before and after — the before half is
    // what makes a shrinking pivot clear up after itself, and both halves
    // together are the undo footprint.
    const touched = new Map();
    const mark = (row, col) => touched.set(row + ':' + col, { row, col });
    const old = pivot.location;
    for (let r = old.top; r <= old.bottom; r++) for (let c = old.left; c <= old.right; c++) mark(r, c);
    for (let r = grid.area.top; r <= grid.area.bottom; r++) {
      for (let c = grid.area.left; c <= grid.area.right; c++) mark(r, c);
    }

    const parts = [...new Set([
      this.workbook.partNameFor(this.activeSheet),
      pivot.part,
      ...(pivot.cachePart ? [pivot.cachePart] : []),
    ])];

    this._edit('refresh ' + pivot.name, null, [...touched.values()], () => {
      const wanted = new Map(grid.cells.map((c) => [c.row + ':' + c.col, c.value]));
      for (const { row, col } of touched.values()) {
        const value = wanted.has(row + ':' + col) ? wanted.get(row + ':' + col) : '';
        this._setCell(row, col, value === '' ? '' : value);
      }
      updatePivotLocation(this.workbook, pivot, grid.area, {
        headerRows: grid.firstHeaderRow,
        dataRow: grid.firstDataRow,
        dataCol: grid.firstDataCol,
      });
      this._pivots = null;
      this._structuralDirty = true;
      return this;
    }, { parts });
    return { rows: grid.height, cols: grid.width };
  }

  // ---- sheet protection ---------------------------------------------------

  /** The active sheet's protection, digested for the frame and the gate. */
  protection() {
    return this.workbook.sheetProtection(this.activeSheet);
  }

  /** Locked is the DEFAULT; only a style that says otherwise unlocks a cell. */
  isCellLocked(row, col) {
    const style = this.styleFor(row, col);
    return style ? style.locked !== false : true;
  }

  /**
   * Protect the active sheet, so locked cells refuse edits.
   *
   * Passwords are deliberately not offered. Matching Excel's password hash
   * derivation byte-for-byte is real work that buys nothing here —
   * protection in this editor is a guard rail against accidental edits, and
   * a file that needs a cryptographic gate keeps the one Excel wrote.
   */
  protect() {
    if (this.protection().sheet) return this;
    this._edit('protect sheet', null, [], () => {
      this.workbook.setSheetProtection(this.activeSheet, true);
      this._structuralDirty = true;
    }, { parts: [this.workbook.partNameFor(this.activeSheet)] });
    return this;
  }

  /** Thaw the sheet — unless its author set a password we must respect. */
  unprotect() {
    const p = this.protection();
    if (!p.sheet) return this;
    if (p.hasPassword) {
      throw protectionError('This sheet’s protection has a password. Open the file in Excel to remove it.');
    }
    this._edit('unprotect sheet', null, [], () => {
      this.workbook.setSheetProtection(this.activeSheet, null);
      this._structuralDirty = true;
    }, { parts: [this.workbook.partNameFor(this.activeSheet)] });
    return this;
  }

  // ---- frozen panes -------------------------------------------------------

  /** The frozen pane on the active sheet, always a {rows, cols} pair. */
  frozenPane() {
    return this.workbook.frozenPane(this.activeSheet) ?? { rows: 0, cols: 0 };
  }

  /**
   * Freeze the top `rows` rows and left `cols` columns — 0 and 0 thaws.
   * Stored as the sheet view's `<pane>` element, exactly where Excel keeps
   * it, so the file scrolls the same way wherever it opens next.
   */
  freezePanes(rows, cols) {
    const r = Number(rows);
    const c = Number(cols);
    if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0 || r > 128 || c > 64) {
      throw new Error('freeze takes small, whole counts of rows and columns');
    }
    this._edit(r || c ? 'freeze panes' : 'unfreeze panes', null, [], () => {
      this.workbook.setFrozenPane(this.activeSheet, r, c);
      this._structuralDirty = true;
    }, { parts: [this.workbook.partNameFor(this.activeSheet)] });
    return this;
  }

  // ---- scenarios ----------------------------------------------------------

  /** The active sheet's scenarios, for the What-if menu. */
  scenarios() {
    return this.workbook.scenarios(this.activeSheet);
  }

  /**
   * Capture the CURRENT values of the selected cells as a named scenario —
   * Excel's "Add scenario" with the changing cells already holding the
   * values this scenario should remember. Stored in the sheet part, so the
   * file carries its what-if cases wherever it goes.
   */
  defineScenario(name, comment) {
    const trimmed = String(name ?? '').trim();
    if (!trimmed || trimmed.length > 255) throw new Error('a scenario needs a name');
    const cells = this._selectedCells();
    if (cells.length > 32) throw new Error('a scenario captures at most 32 changing cells');
    for (const { row, col } of cells) {
      const input = this.calc.getInput(this.activeSheet, row, col);
      if (typeof input === 'string' && input.startsWith('=')) {
        throw new Error('changing cells hold plain values, not formulas — ' + ref(row, col) + ' has one');
      }
    }
    this._edit('define scenario', null, [], () => {
      this.workbook.setScenario(this.activeSheet, {
        name: trimmed,
        comment: comment ? String(comment) : null,
        cells: cells.map(({ row, col }) => ({
          ref: ref(row, col),
          value: String(this.calc.getInput(this.activeSheet, row, col) ?? ''),
        })),
      });
      this._structuralDirty = true;
    }, { parts: [this.workbook.partNameFor(this.activeSheet)] });
    return this;
  }

  /** Apply a scenario: its stored inputs land as ONE ordinary, undoable edit. */
  showScenario(name) {
    const s = this.scenarios().find((x) => x.name === String(name ?? '').trim());
    if (!s) throw new Error('no scenario "' + String(name ?? '').trim() + '"');
    const targets = s.cells.map((c) => {
      const [row, col] = parseRefPair(c.ref);
      return { row, col, value: c.value };
    });
    return this._edit('scenario: ' + s.name, null, targets.map(({ row, col }) => ({ row, col })), () => {
      for (const t of targets) this._setCell(t.row, t.col, t.value);
      return this;
    });
  }

  deleteScenario(name) {
    const trimmed = String(name ?? '').trim();
    if (!this.scenarios().some((s) => s.name === trimmed)) {
      throw new Error('no scenario "' + trimmed + '"');
    }
    this._edit('delete scenario', null, [], () => {
      this.workbook.deleteScenario(this.activeSheet, trimmed);
      this._structuralDirty = true;
    }, { parts: [this.workbook.partNameFor(this.activeSheet)] });
    return this;
  }

  // ---- table filters ------------------------------------------------------

  /** A table on the active sheet by name, with its range parsed. */
  _tableByName(name) {
    const t = this.workbook.tables().find((x) => x.sheet === this.activeSheet
      && String(x.name).toLowerCase() === String(name ?? '').toLowerCase());
    if (!t) throw new Error('no table "' + String(name ?? '') + '" on this sheet');
    const [a, b] = String(t.ref).split(':');
    const [r1, c1] = parseRefPair(a.replace(/\$/g, ''));
    const [r2, c2] = b ? parseRefPair(b.replace(/\$/g, '')) : [r1, c1];
    return {
      ...t,
      top: Math.min(r1, r2),
      left: Math.min(c1, c2),
      bottom: Math.max(r1, r2),
      right: Math.max(c1, c2),
    };
  }

  _tableColumnIndex(t, columnName) {
    const idx = t.columns.findIndex((c) => String(c).toLowerCase() === String(columnName ?? '').toLowerCase());
    if (idx < 0) throw new Error('no column "' + String(columnName ?? '') + '" in table ' + t.name);
    return idx;
  }

  /**
   * A filterable thing by name: a ListObject, or `#sheet` — the sheet-level
   * autofilter Excel's Data › Filter toggles over a plain range. The two
   * store their state in different parts (the table part, the sheet part),
   * so the target carries its own read/write closures and the parts an edit
   * must snapshot; everything downstream is then one code path.
   */
  _filterTarget(name) {
    if (name === SHEET_FILTER) {
      const af = this.workbook.sheetAutoFilter(this.activeSheet);
      if (!af) throw new Error('this sheet has no filter — Data › Filter turns one on');
      const columns = [];
      for (let c = af.left; c <= af.right; c++) {
        columns.push(this.displayValue(af.top, c).text || colName(c));
      }
      const sheetPartName = this.workbook.partNameFor(this.activeSheet);
      return {
        name: SHEET_FILTER,
        top: af.top,
        left: af.left,
        bottom: af.bottom,
        right: af.right,
        headerRowCount: 1,
        totalsRowCount: 0,
        columns,
        filters: () => this.workbook.sheetAutoFilter(this.activeSheet).filters,
        setFilter: (colId, values) => this.workbook.setSheetFilter(this.activeSheet, colId, values),
        parts: [sheetPartName],
      };
    }
    const t = this._tableByName(name);
    return {
      ...t,
      filters: () => this.workbook.tableFilters(t.part),
      setFilter: (colId, values) => this.workbook.setTableFilter(t.part, colId, values),
      parts: [this.workbook.partNameFor(this.activeSheet), t.part],
    };
  }

  /**
   * What a filter dropdown offers: the column's distinct display texts, and
   * which currently pass. `selected: null` means no filter — everything shows.
   */
  filterOptions(tableName, columnName) {
    const t = this._filterTarget(tableName);
    const idx = this._tableColumnIndex(t, columnName);
    const col = t.left + idx;
    const seen = new Set();
    const values = [];
    const top = t.top + t.headerRowCount;
    const bottom = t.bottom - t.totalsRowCount;
    for (let r = top; r <= bottom && values.length < 1000; r++) {
      const text = this.displayValue(r, col).text;
      if (!seen.has(text)) {
        seen.add(text);
        values.push(text);
      }
    }
    values.sort((x, y) => x.localeCompare(y));
    const selected = t.filters().get(idx);
    return { table: t.name, column: t.columns[idx], values, selected: selected === undefined ? null : selected };
  }

  /**
   * Open (or, with no table, close) this person's filter dropdown. Cursor
   * state, not a write: two people can browse different columns' lists.
   */
  openFilterPanel(tableName, columnName) {
    this.filterPanel = tableName === null || tableName === undefined
      ? null
      : this.filterOptions(tableName, columnName);
    return this;
  }

  /**
   * Apply one column's value filter — or clear it with null. What Excel
   * stores is what this writes: the filterColumn state in the table part and
   * `hidden` on the sheet rows the WHOLE table's filters exclude, so the
   * file opens identically filtered elsewhere. Matching is on display text,
   * as Excel matches.
   */
  applyFilter(tableName, columnName, values) {
    const t = this._filterTarget(tableName);
    const idx = this._tableColumnIndex(t, columnName);
    const list = values === null || values === undefined ? null : [...values].map(String);
    const sheetPartName = this.workbook.partNameFor(this.activeSheet);
    this._edit('filter', null, [], () => {
      t.setFilter(idx, list);
      const filters = t.filters();
      const top = t.top + t.headerRowCount;
      const bottom = t.bottom - t.totalsRowCount;
      const hidden = new Set();
      for (let r = top; r <= bottom; r++) {
        for (const [cid, vals] of filters) {
          if (!vals || !vals.length) continue;
          if (!vals.includes(this.displayValue(r, t.left + cid).text)) {
            hidden.add(r);
            break;
          }
        }
      }
      this.workbook.setRowsHidden(this.activeSheet, top, bottom, hidden);
      // Row heights are part-derived; the grid must collapse what just hid.
      const xml = this.workbook.snapshotParts([sheetPartName])[sheetPartName];
      this.geometry.set(this.activeSheet, SheetGeometry.fromSheetXml(xml));
      this._structuralDirty = true;
    }, { parts: t.parts });
    this.filterPanel = null;
    return this;
  }

  /**
   * Excel's Data › Filter: toggle the sheet-level autofilter over the data
   * region around the cursor. Turning it OFF also unhides whatever its
   * filters had hidden — the funnels leave, the rows come back, exactly as
   * Excel behaves. One undoable edit either way.
   */
  toggleAutoFilter() {
    const sheetPartName = this.workbook.partNameFor(this.activeSheet);
    const existing = this.workbook.sheetAutoFilter(this.activeSheet);
    this._edit('filter', null, [], () => {
      if (existing) {
        this.workbook.setSheetAutoFilter(this.activeSheet, null);
        this.workbook.setRowsHidden(this.activeSheet, existing.top + 1, existing.bottom, new Set());
      } else {
        const active = this.selection.active;
        const region = this._currentRegion(active.row, active.col);
        if (region.top === region.bottom && region.left === region.right
            && !this.isFilled(region.top, region.left)) {
          throw new Error('Select a cell inside your data first — a filter needs a range.');
        }
        this.workbook.setSheetAutoFilter(this.activeSheet,
          ref(region.top, region.left) + ':' + ref(region.bottom, region.right));
      }
      const xml = this.workbook.snapshotParts([sheetPartName])[sheetPartName];
      this.geometry.set(this.activeSheet, SheetGeometry.fromSheetXml(xml));
      this._structuralDirty = true;
    }, { parts: [sheetPartName] });
    this.filterPanel = null;
    return this;
  }

  /**
   * The sheet-level autofilter as the frame carries it: the same shape as a
   * table entry, named `#sheet` (a name no ListObject can hold), so the
   * grid's funnels, panel and actions serve both without a second code path.
   */
  sheetFilter() {
    const af = this.workbook.sheetAutoFilter(this.activeSheet);
    if (!af) return null;
    const columns = [];
    for (let c = af.left; c <= af.right; c++) {
      columns.push(this.displayValue(af.top, c).text || colName(c));
    }
    return {
      name: SHEET_FILTER,
      ref: af.ref,
      top: af.top,
      left: af.left,
      bottom: af.bottom,
      right: af.right,
      headerRows: 1,
      totalsRows: 0,
      columns,
      filtered: [...af.filters.entries()]
        .filter(([, vals]) => vals === null || vals.length)
        .map(([colId]) => colId),
    };
  }

  /** Jump to what a name points at — the name box's second job. */
  gotoName(name) {
    const wanted = String(name ?? '').trim().toUpperCase();
    const entry = this.names().find((n) => n.name.toUpperCase() === wanted);
    if (!entry) throw new Error('no defined name "' + String(name ?? '').trim() + '"');
    if (!entry.target) throw new Error('"' + entry.name + '" does not point at a range');
    const { sheet, start, end } = entry.target;
    if (sheet && sheet !== this.activeSheet) this.selectSheet(sheet);
    this.select(start.row, start.col);
    if (end.row !== start.row || end.col !== start.col) this.select(end.row, end.col, { extend: true });
    return this;
  }

  // ---- find and replace --------------------------------------------------

  /** Whether one cell matches the search text, against what is SHOWN and what was TYPED. */
  _matches(row, col, needle) {
    if (!this.isFilled(row, col)) return false;
    const q = String(needle).toLowerCase();
    return this.displayValue(row, col).text.toLowerCase().includes(q)
      || this.editValue(row, col).toLowerCase().includes(q);
  }

  /**
   * Move the selection to the next match after the active cell, row-major,
   * wrapping past the end of the used range. Returns whether anything matched.
   * A search moves the cursor and nothing else, so it is not an edit.
   */
  findNext(text) {
    const needle = String(text ?? '');
    if (!needle) return false;
    const { maxRow, maxCol } = this.bounds;
    const width = maxCol + 1;
    const total = (maxRow + 1) * width;
    const from = this.selection.active.row * width + this.selection.active.col;
    for (let step = 1; step <= total; step++) {
      const at = (from + step) % total;
      const row = Math.floor(at / width);
      const col = at % width;
      if (this._matches(row, col, needle)) {
        this.history.break();
        this.selection.collapseTo(row, col);
        this.ensureVisible();
        return true;
      }
    }
    return false;
  }

  /** Case-insensitive replacement of every occurrence within one input. */
  static _replaceIn(input, find, replaceWith) {
    const pattern = new RegExp(String(find).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    return String(input).replace(pattern, () => replaceWith);
  }

  /**
   * Replace in the active cell if it matches — else in the next match — and
   * move on to the one after, Excel's Replace button. Returns whether a
   * replacement happened.
   */
  replaceNext(find, replaceWith) {
    const needle = String(find ?? '');
    if (!needle) return false;
    const { row, col } = this.selection.active;
    if (!this._matches(row, col, needle) && !this.findNext(needle)) return false;
    const at = this.selection.active;
    const input = this.editValue(at.row, at.col);
    const next = SheetView._replaceIn(input, needle, String(replaceWith ?? ''));
    if (next !== input) this.setCell(at.row, at.col, next);
    this.findNext(needle);
    return true;
  }

  /**
   * Replace everywhere in the used range, as ONE undo step.
   * Replacement happens in the INPUT, so a formula's text is edited and
   * re-parsed rather than its result being overwritten.
   * @returns {number} how many cells changed
   */
  replaceAll(find, replaceWith) {
    const needle = String(find ?? '');
    if (!needle) return 0;
    const { maxRow, maxCol } = this.bounds;
    const hits = [];
    for (let row = 0; row <= maxRow; row++) {
      for (let col = 0; col <= maxCol; col++) {
        if (!this._matches(row, col, needle)) continue;
        const input = this.editValue(row, col);
        const next = SheetView._replaceIn(input, needle, String(replaceWith ?? ''));
        if (next !== input) hits.push({ row, col, next });
      }
    }
    if (!hits.length) return 0;
    this._edit('replace all', null, hits.map(({ row, col }) => ({ row, col })), () => {
      for (const h of hits) this._setCell(h.row, h.col, h.next);
      return this;
    });
    return hits.length;
  }

  // ---- persistence -------------------------------------------------------

  // Formatting counts. A person who bolds a heading and walks away must see
  // an unsaved marker, or the change is lost with no warning that it existed.
  get isDirty() { return this.dirtyCells.size > 0 || this.styledCells.size > 0 || this._structuralDirty; }

  /**
   * Write edits back through the preserving path.
   *
   * Literals go through `Workbook.setCell`; formulas are written and then given
   * their calculated cache, so the saved file is correct without needing Excel
   * to recalculate it.
   */
  // ── cell formatting ───────────────────────────────────────────────────────

  /**
   * A workbook with no `styles.xml` still has to be formattable.
   *
   * Excel always writes one, but a file generated without any styling has no
   * reason to — and "you cannot make this bold because nothing in it was ever
   * styled" is not an explanation anybody would accept. This creates the
   * minimum a consumer expects: the four tables, the mandatory second fill
   * (gray125, which the schema requires even though nothing uses it), and one
   * default `xf` for index 0 so existing unstyled cells keep meaning what they
   * meant.
   *
   * The part needs its content-type override AND a relationship from
   * workbook.xml, or Excel opens the file and silently ignores the styles.
   */
  _ensureStylesPart() {
    if (this.pkg.has('xl/styles.xml')) return;

    this.pkg.addPart(
      'xl/styles.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
      + '<fills count="2"><fill><patternFill patternType="none"/></fill>'
      + '<fill><patternFill patternType="gray125"/></fill></fills>'
      + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
      + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
      + '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
      + '</styleSheet>',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
    );

    const relsPart = 'xl/_rels/workbook.xml.rels';
    if (this.pkg.has(relsPart)) {
      const rels = this.pkg.text(relsPart);
      if (!/Target="styles\.xml"/.test(rels)) {
        // A fresh id, because reusing one silently repoints an existing part.
        const used = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
        const id = 'rId' + (used.length ? Math.max(...used) + 1 : 1);
        this.pkg.write_(relsPart, rels.replace(
          '</Relationships>',
          '<Relationship Id="' + id + '" Type="http://schemas.openxmlformats.org/'
          + 'officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        ));
      }
    }
    this.styles = readStyles(this.pkg, { BUILTIN_FORMATS });
  }

  /** This cell's style index, or null where the file said nothing. */
  _styleIndexAt(sheetName, row, col) {
    const index = this.cellStyles.get(sheetName)?.get(ref(row, col));
    return index === undefined ? null : index;
  }

  _setStyleIndex(sheetName, row, col, index) {
    let map = this.cellStyles.get(sheetName);
    if (!map) { map = new Map(); this.cellStyles.set(sheetName, map); }
    if (index === null || index === undefined) map.delete(ref(row, col));
    else map.set(ref(row, col), index);
    this.styledCells.add(sheetName + '!' + ref(row, col));
  }

  /**
   * Whether the whole selection already reads as bold / italic / etc.
   *
   * "Already" is the whole selection, not the active cell: a toolbar that lit
   * up because the corner cell happened to be bold would turn the rest of the
   * selection OFF on the next click, which is the opposite of what pressing it
   * looks like it will do.
   */
  formatState() {
    const xml = this.pkg.has('xl/styles.xml') ? this.pkg.text('xl/styles.xml') : '';
    // Reading every cell of a whole-column selection would parse the style table
    // a million times for a toolbar highlight. Past a sane size the active cell
    // answers for the selection, which is what a user reads it as anyway.
    const SAMPLE = 512;
    const selected = this._selectedCells();
    const cells = selected.length > SAMPLE ? [this.selection.active] : selected;
    if (!xml || !cells.length) {
      return {
        bold: false, italic: false, underline: false, strike: false,
        align: null, valign: null, wrap: false, numberFormat: 'General',
        locked: true,
      };
    }
    let all = null;
    for (const { row, col } of cells) {
      const at = formatOf(xml, this._styleIndexAt(this.activeSheet, row, col) ?? 0);
      if (all === null) { all = { ...at }; continue; }
      // wrap reconciles like the boolean toggles: a selection where any cell does
      // not wrap reads as false, so pressing the button turns the whole run on.
      for (const k of ['bold', 'italic', 'underline', 'strike', 'wrap']) if (!at[k]) all[k] = false;
      if (all.align !== at.align) all.align = null;
      // valign reconciles like horizontal align: a selection carrying two
      // different vertical alignments has none to show, so it reads null.
      if (all.valign !== at.valign) all.valign = null;
      // A selection spanning two different formats has no single one to show, so
      // a toolbar reads null and lights nothing rather than the corner cell's.
      if (all.numberFormat !== at.numberFormat) all.numberFormat = null;
      // A mixed locked/unlocked selection reads LOCKED, so the button offers
      // to unlock the lot — the gesture a person mixing them actually wants.
      if (at.locked) all.locked = true;
    }
    return all;
  }

  /**
   * Apply a formatting delta to the selection.
   *
   * Toggles are resolved ONCE for the whole selection rather than per cell, so
   * a mixed selection goes uniformly bold on the first press instead of
   * inverting cell by cell into a checkerboard.
   */
  setFormat(delta) {
    this._ensureStylesPart();
    const cells = this._selectedCells();
    if (!cells.length) return this;

    const state = this.formatState();
    const resolved = {};
    for (const k of ['bold', 'italic', 'underline', 'strike']) {
      if (delta[k] === 'toggle') resolved[k] = !state[k];
      else if (delta[k] !== undefined) resolved[k] = Boolean(delta[k]);
    }
    if (delta.align !== undefined) {
      resolved.align = delta.align === state.align ? null : delta.align;
    }
    // Values rather than toggles: a family, a size, a colour, a fill and a set
    // of border edges are SET to what was asked for, or to null to clear. There
    // is no sensible "toggle Georgia", and pressing a colour twice should leave
    // the colour rather than removing it — which is what a toggle would do.
    for (const k of ['fontName', 'fontSize', 'fontColour', 'fill', 'border', 'numberFormat', 'valign', 'wrap', 'locked']) {
      if (delta[k] !== undefined) resolved[k] = delta[k];
    }
    if (!Object.keys(resolved).length) return this;

    return this._edit('format', 'format:' + this.activeSheet + ':' + this.selection.toString(), cells, () => {
      let xml = this.pkg.text('xl/styles.xml');
      let touched = false;
      for (const { row, col } of cells) {
        const base = this._styleIndexAt(this.activeSheet, row, col);
        const out = applyFormat(xml, base, resolved);
        xml = out.xml;
        if (out.index !== base) { this._setStyleIndex(this.activeSheet, row, col, out.index); touched = true; }
      }
      if (touched) {
        this.pkg.write_('xl/styles.xml', xml);
        this.styles = readStyles(this.pkg, { BUILTIN_FORMATS });
        this._stylesDirty = true;
      }
      return this;
    }, { styles: true });
  }

  save() {
    for (const key of this.dirtyCells) {
      const [sheetName, cellRef] = key.split('!');
      const input = this.calc.getInput(sheetName, ...parseRefPair(cellRef));
      this.workbook.setCell(sheetName, cellRef, input === '' ? null : input);
    }

    /*
     * Re-cache EVERY formula, not only the ones the user touched.
     *
     * Editing B2 changes the value of C2 without C2 being edited. Caching only
     * the dirty set leaves C2's stored result describing the old B2 — and since
     * a consumer trusts the cache until it recalculates, the file would open
     * showing a number that was never true. The dependency graph already knows
     * the new value; this is just asking it.
     *
     * Cheap because writeCachedValue is idempotent: a cache that already agrees
     * is left alone and marks nothing modified.
     */
    // The live spills, so each anchor's `<f>` is written in the ARRAY form
    // (`t="array"` with its spilled range) — what makes the file mean the
    // same thing in an Excel without dynamic arrays. A blocked #SPILL!
    // anchor writes the plain formula: its range is a wish, not a fact.
    const spillAnchors = new Map();
    for (const s of (this.calc.spills ?? new Map()).values()) {
      if (s.blocked || (s.h <= 1 && s.w <= 1)) continue;
      spillAnchors.set(s.sheet + '|' + s.top + '|' + s.left,
        ref(s.top, s.left) + ':' + ref(s.top + s.h - 1, s.left + s.w - 1));
    }
    for (const [sheetName, cells] of this.calc.sheets) {
      for (const cell of cells.values()) {
        if (!cell.formula) continue;
        writeCachedValue(this.workbook, sheetName, ref(cell.row, cell.col), cell.formula, cell.value,
          { arrayRef: spillAnchors.get(sheetName + '|' + cell.row + '|' + cell.col) ?? null });
      }
    }
    // Ghost cells cache no values, so the next consumer must compute them.
    if (spillAnchors.size) this.workbook.setFullCalcOnLoad();

    for (const key of this.styledCells) {
      const [sheetName, cellRef] = key.split('!');
      const [row, col] = parseRefPair(cellRef);
      this.workbook.setCellStyle(sheetName, cellRef, this._styleIndexAt(sheetName, row, col));
    }

    const out = this.workbook.save();
    this.dirtyCells.clear();
    this.styledCells.clear();
    // Structural changes are in the saved bytes now too. Leaving this set —
    // as it was until 2026-08-31 — kept the Save button lit forever after
    // saving a merge, a resize or an unprotect.
    this._structuralDirty = false;
    return out;
  }

  /**
   * The same bytes save() produces, WITHOUT claiming the work is saved.
   *
   * For the autosave draft: a draft flush must not consume the dirty
   * accounting, or one background tick darkens everyone's Save button and
   * the next real save answers "unchanged" over work that is only in the
   * draft slot. The workbook writes save() performs are idempotent, so
   * re-arming the tracking sets afterwards costs nothing but honesty.
   */
  serialize() {
    const dirty = [...this.dirtyCells];
    const styled = [...this.styledCells];
    const structural = this._structuralDirty;
    const out = this.save();
    for (const key of dirty) this.dirtyCells.add(key);
    for (const key of styled) this.styledCells.add(key);
    this._structuralDirty = structural;
    return out;
  }
}

/** `B7` -> [6, 1] */
function parseRefPair(cellRef) {
  const m = /^([A-Z]+)(\d+)$/.exec(cellRef);
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return [Number(m[2]) - 1, col - 1];
}

/**
 * What a typed string means.
 *
 * A spreadsheet reads `42` as a number and `042` as… also a number, but reads
 * a leading apostrophe as "keep this as text", which is how people stop
 * part numbers and phone numbers being mangled.
 */
/**
 * The name lists a drag continues through — months and weekdays, long and
 * short. What Excel calls custom lists; these four are the built-in ones
 * every teacher's worksheet and every report header leans on.
 */
const FILL_LISTS = [
  ['January', 'February', 'March', 'April', 'May', 'June', 'July',
    'August', 'September', 'October', 'November', 'December'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
  ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
];

/**
 * Do a lane's sources walk one of the name lists with a constant step?
 * Returns a continuation function or null. The source's case pattern is
 * kept — JAN drags to FEB, jan to feb.
 */
function listSeries(inputs) {
  if (!inputs.length || !inputs.every((v) => typeof v === 'string' && v)) return null;
  for (const list of FILL_LISTS) {
    const indices = inputs.map((v) => list.findIndex((e) => e.toLowerCase() === v.toLowerCase()));
    if (indices.some((i) => i < 0)) continue;
    let step = 1;
    if (indices.length > 1) {
      step = indices[1] - indices[0];
      if (step === 0) continue;
      if (!indices.every((v, i) => i === 0 || v - indices[i - 1] === step)) continue;
    }
    const last = indices[indices.length - 1];
    const shape = inputs[0] === inputs[0].toUpperCase() ? 'upper'
      : inputs[0] === inputs[0].toLowerCase() ? 'lower' : 'canonical';
    return (k) => {
      const entry = list[((last + step * k) % list.length + list.length) % list.length];
      return shape === 'upper' ? entry.toUpperCase() : shape === 'lower' ? entry.toLowerCase() : entry;
    };
  }
  return null;
}

/**
 * Do a lane's date serials step by MONTHS — same day each month, or each
 * month's end? Linear day arithmetic cannot say "the 15th of every month"
 * (months are unequal), so this is its own series kind, clamped the way
 * Excel clamps: Jan 31 + one month is Feb 28.
 */
function monthSeries(serials) {
  if (serials.length < 2) return null;
  const dates = serials.map((s) => serialToDate(s));
  if (dates.some((d) => !(d instanceof Date) || Number.isNaN(d.getTime()))) return null;
  const months = dates.map((d) => d.getUTCFullYear() * 12 + d.getUTCMonth());
  const step = months[1] - months[0];
  if (step === 0) return null;
  if (!months.every((m, i) => i === 0 || m - months[i - 1] === step)) return null;
  const lastDay = (y, m) => new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  const day = dates[0].getUTCDate();
  const sameDay = dates.every((d) => d.getUTCDate() === Math.min(day, lastDay(d.getUTCFullYear(), d.getUTCMonth())));
  const monthEnd = dates.every((d) => d.getUTCDate() === lastDay(d.getUTCFullYear(), d.getUTCMonth()));
  if (!sameDay && !monthEnd) return null;
  const base = months[months.length - 1];
  return (k) => {
    const m = base + step * k;
    const year = Math.floor(m / 12);
    const month = ((m % 12) + 12) % 12;
    const wanted = !sameDay && monthEnd ? lastDay(year, month) : Math.min(day, lastDay(year, month));
    return dateToSerial(new Date(Date.UTC(year, month, wanted)));
  };
}

const escapeHtml = (s) => String(s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/**
 * The first <table> in clipboard HTML, as a grid of cell texts. A regex
 * walk, not a DOM: the shapes Excel, Sheets and browsers write are regular
 * enough, and anything that defeats it simply falls back to the plain-text
 * paste — degraded, never wrong.
 */
/** A CSS colour as '#rrggbb', or null for anything this cannot say honestly. */
function cssColour(raw) {
  const s = String(raw ?? '').trim().toLowerCase();
  const hex6 = /^#([0-9a-f]{6})$/.exec(s);
  if (hex6) return '#' + hex6[1];
  const hex3 = /^#([0-9a-f]{3})$/.exec(s);
  if (hex3) return '#' + [...hex3[1]].map((ch) => ch + ch).join('');
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(s);
  if (rgb) {
    return '#' + [rgb[1], rgb[2], rgb[3]]
      .map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0')).join('');
  }
  return null;
}

/**
 * What a foreign cell's markup says about its LOOK — the half of an inbound
 * paste that used to be dropped. Excel and Sheets put the truth on the td
 * (inline style, sometimes bgcolor) with the older tag forms inside; both
 * are read, and anything this cannot express honestly is left unread rather
 * than guessed. White backgrounds are noise, not a choice, and left/top are
 * the defaults nothing needs to say.
 */
function tdFormat(attrs, inner) {
  const f = {};
  const style = /style="([^"]*)"/i.exec(attrs)?.[1] ?? '';
  if (/<(?:b|strong)\b/i.test(inner) || /font-weight\s*:\s*(?:bold|[7-9]00)/i.test(style)) f.bold = true;
  if (/<(?:i|em)\b/i.test(inner) || /font-style\s*:\s*italic/i.test(style)) f.italic = true;
  if (/<u\b/i.test(inner) || /text-decoration[^;"]*underline/i.test(style)) f.underline = true;
  if (/<(?:s|strike|del)\b/i.test(inner) || /text-decoration[^;"]*line-through/i.test(style)) f.strike = true;
  const colour = cssColour(/(?:^|;)\s*color\s*:\s*([^;]+)/i.exec(style)?.[1])
    ?? cssColour(/<font\b[^>]*\bcolor="([^"]+)"/i.exec(inner)?.[1]);
  if (colour && colour !== '#000000') f.fontColour = colour;
  const bg = cssColour(/background(?:-color)?\s*:\s*([^;]+)/i.exec(style)?.[1])
    ?? cssColour(/\bbgcolor="([^"]+)"/i.exec(attrs)?.[1]);
  if (bg && bg !== '#ffffff') f.fill = bg;
  const align = /text-align\s*:\s*(left|center|right)/i.exec(style)?.[1]?.toLowerCase();
  if (align === 'center' || align === 'right') f.align = align;
  return Object.keys(f).length ? f : null;
}

function parseHtmlTable(html) {
  const table = /<table\b[^>]*>([\s\S]*?)<\/table>/i.exec(String(html));
  if (!table) return null;
  const rows = [];
  for (const tr of table[1].matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/gi)) {
    const line = [];
    for (const td of tr[1].matchAll(/<t[dh]\b([^>]*)>([\s\S]*?)<\/t[dh]>/gi)) {
      line.push({
        value: td[2]
          .replace(/<br\s*\/?>/gi, '\n')
          .replace(/<[^>]+>/g, '')
          .replace(/&nbsp;/gi, ' ')
          .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
          .replace(/&#0*39;/g, "'").replace(/&amp;/g, '&')
          .trim(),
        format: tdFormat(td[1], td[2]),
      });
    }
    if (line.length) rows.push(line);
  }
  return rows.length ? rows : null;
}

/**
 * The look a table gives one of its cells, or null where it gives none.
 *
 * A table that declares NO tableStyleInfo stays plain — the author asked for
 * that. One that names a style is painted in OUR palette (the theme's first
 * accent), not an imitation of Excel's named gallery: the same decision the
 * chart renderer made, and for the same reason — a consistent surface beats
 * a poor forgery.
 */
function tableLookAt(tables, row, col, theme) {
  for (const t of tables) {
    if (row < t.top || row > t.bottom || col < t.left || col > t.right) continue;
    if (!t.styleName) return null;
    const accent = theme?.[4] ?? '#1f5f8b';
    if (t.headerRows && row <= t.top + t.headerRows - 1) {
      return { fill: { pattern: 'solid', colour: accent }, font: { bold: true, colour: '#ffffff' } };
    }
    if (t.totalsRows && row >= t.bottom - t.totalsRows + 1) {
      return { fill: { pattern: 'solid', colour: applyTint(accent, 0.75) }, font: { bold: true } };
    }
    if (t.showRowStripes && (row - (t.top + t.headerRows)) % 2 === 1) {
      return { fill: { pattern: 'solid', colour: applyTint(accent, 0.9) } };
    }
    return null;
  }
  return null;
}

/**
 * A protection refusal: an ordinary Error wearing a marker, so the session
 * layer can tell "the sheet said no, in a sentence for the person" apart
 * from a genuine fault it must not swallow.
 */
function protectionError(message) {
  const e = new Error(message);
  e.protection = true;
  return e;
}

/** Where a value sits between two bounds, clamped to [0, 1]. */
function fraction(value, lo, hi) {
  if (hi === lo) return value >= hi ? 1 : 0;
  return Math.max(0, Math.min(1, (value - lo) / (hi - lo)));
}

/** Linear blend of two #rrggbb colours. */
function lerpColour(a, b, t) {
  const pa = /^#([0-9a-f]{6})/i.exec(String(a ?? ''));
  const pb = /^#([0-9a-f]{6})/i.exec(String(b ?? ''));
  if (!pa || !pb) return a ?? b ?? null;
  const ca = parseInt(pa[1], 16);
  const cb = parseInt(pb[1], 16);
  const mix = (shift) => Math.round(((ca >> shift) & 255) + (((cb >> shift) & 255) - ((ca >> shift) & 255)) * t);
  return '#' + ((mix(16) << 16) | (mix(8) << 8) | mix(0)).toString(16).padStart(6, '0');
}

/** What a numeric validation type asks for, in the refusal sentence. */
const NUMERIC_KIND = {
  whole: 'a whole number',
  decimal: 'a number',
  date: 'a date',
  time: 'a time',
  textLength: 'text of a permitted length',
};

export function coerceInput(text) {
  if (text === null || text === undefined) return '';
  const s = String(text);
  if (s === '') return '';
  if (s.startsWith("'")) return s.slice(1);
  if (s.startsWith('=')) return s;
  const trimmed = s.trim();
  if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(trimmed)) {
    const n = Number(trimmed);
    if (Number.isFinite(n)) return n;
  }
  if (/^(TRUE|FALSE)$/i.test(trimmed)) return trimmed.toUpperCase() === 'TRUE';
  return s;
}

export { OoxmlPackage };
