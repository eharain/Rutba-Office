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
  createPivot, planPivot, parseArea as parsePivotArea,
  pivotChartXml, readPivotChart, pivotSourceName, PIVOT_CHART_KINDS,
  sourceValues, ensureSharedItems, setPivotItemsShown, itemLabel, pivotKey,
} from '@rutba/ooxml/pivot';
import {
  readSlicers, addSlicer, removeSlicer, slicerAnchorXml, writeSlicerCacheItems, setSlicerProps as writeSlicerProps,
} from '@rutba/ooxml/slicers';
import {
  isError, shiftFormula, calculate, parse, compareValues, serialToDate, dateToSerial, FormulaEvaluation,
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
  group as outlineGroup, ungroup as outlineUngroup, toggleGroup, showLevel, detail as outlineDetail,
  clearOutline, outlineFrame, selectionAxis, subtotal, removeSubtotals, listFields,
} from './outline.js';
import { advancedFilter, clearAdvancedFilter, filterNames } from './advanced-filter.js';
import { paginate, planBands, pageSetup, readPageSetup, parseArea, PAPER, PX_PER_MM } from './print.js';
import { inferProgram, runProgram } from './flash-fill.js';
import { consolidate as consolidateRanges, lastConsolidation, consolidateRefText } from './consolidate.js';
import { createForecastSheet, forecastDefaults, forecastPreview } from './forecast-sheet.js';
import {
  readSheetDrawings, buildChart, buildShape, buildPicture, renderSvg, scene,
  SUPPORTED_GEOMETRY, childElements, xfrmOf, anchorBody,
} from '@rutba/drawing';
import { drawingAnchorXml, chartPartXml } from '@rutba/ooxml/build';
import { newGuid } from '@rutba/ooxml/workbook';
import { History } from '@rutba/editing';
import {
  readWorkbookDesign, designedThemeXml, stylesFollowingFonts, THEME_PART, THEME_TYPE, THEME_REL,
} from './themes.js';

/**
 * Rows and columns drawn beyond the viewport's edges, so a scroll shows
 * cells before the next frame lands. A wheel notch moves a hundred pixels —
 * five default rows — and a person scrolls up as briskly as down, so the
 * rows are generous on both sides; columns move a notch at a time.
 */
const OVERSCAN_ROWS = 10;
const OVERSCAN_COLS = 3;

/** Sheet XML beyond which cell styles are read on demand rather than mapped up front. */
const LAZY_STYLES_XML = 64 * 1024 * 1024;

/**
 * The pseudo-name the sheet-level autofilter answers to in the filter UI.
 * A ListObject's name must start with a letter or underscore, so no real
 * table can collide with it.
 */
export const SHEET_FILTER = '#sheet';

/** Excel's words for a wrong password, on a sheet, a workbook or a range. */
export const WRONG_PASSWORD = 'The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.';

/** Excel's words for a sheet added, deleted, renamed, moved or hidden in a locked workbook. */
export const WORKBOOK_LOCKED = 'Workbook is protected and cannot be changed.';

/**
 * The chart kinds the shared writer takes — Studio's vocabulary, the
 * convergence the roadmap asked for. Lock-step with `chartPartXml`'s kind
 * branch; the ribbon offers exactly this list.
 */
export const CHART_KINDS = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter'];

/** Insert → Scatter's three looks, as `chartPartXml` writes them. */
export const SCATTER_STYLES = ['markers', 'lines', 'smooth'];

/**
 * Error Checking: the plain-English reason behind each error type, keyed by
 * the error text the engine stores on the cell. `#CIRCULAR!` is not a real
 * Excel error type — it is how this engine settles a cell it could not order
 * (see `Spreadsheet#recalculate`) — but it reads the same to a person as any
 * other error, so it gets a row too. Anything else (`#SPILL!`, `#CALC!`)
 * falls back to a generic reason rather than going unexplained.
 */
const ERROR_KIND = {
  '#DIV/0!': { kind: 'div0', reason: 'Divides by zero' },
  '#REF!': { kind: 'ref', reason: 'Refers to a cell that was deleted' },
  '#NAME?': { kind: 'name', reason: 'Uses a name that is not defined' },
  '#VALUE!': { kind: 'value', reason: 'Uses the wrong kind of value' },
  '#N/A': { kind: 'na', reason: 'A lookup found nothing' },
  '#NUM!': { kind: 'num', reason: 'A number that cannot be represented' },
  '#NULL!': { kind: 'null', reason: 'Ranges that do not intersect' },
  '#CIRCULAR!': { kind: 'circular', reason: 'Refers to itself, directly or through other cells' },
};
const DEFAULT_ERROR_KIND = { kind: 'error', reason: 'The formula could not be calculated' };

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
    // Formulas → Calculation Options, as the file says: a workbook saved in
    // manual mode opens in manual mode. It is calculated once on opening
    // all the same — the engine holds no values until it has.
    this.calc.manual = this.workbook.calcMode() === 'manual';

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
    /** Sheets whose cell styles are read from the part on demand — see the constructor. */
    this._lazyStyles = new Set();


    for (const { name, part } of this.workbook.sheets()) {
      const xml = this.pkg.text(part);
      this.geometry.set(name, SheetGeometry.fromSheetXml(xml));
      // A map of every cell's style is fine at two million cells and
      // impossible at eighteen: a JavaScript Map stops at 16.7 million
      // entries, and a 46 MB sample-data workbook has more. Past a size the
      // styles are read from the part as cells are asked for, and only the
      // cells written to are kept here.
      if (xml.length > LAZY_STYLES_XML) {
        this._lazyStyles.add(name);
        this.cellStyles.set(name, new Map());
      } else {
        this.cellStyles.set(name, this._readCellStyles(xml));
      }
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
    /** sheet -> { suffix, map: ref -> link }, read when a sheet is first drawn and again whenever its tail changes (an edit, an undo). */
    this.links = new Map();
    /** sheet -> { xml, map: ref -> note }, read again whenever the comments part changes (a note put on, an undo). */
    this.notes = new Map();

    // A workbook with no sheet in it is not a workbook. Without this the view
    // carried an undefined active sheet all the way to the first frame and
    // threw there — "Cannot read properties of undefined (reading
    // 'viewport')" — which reaches a person as a window that never draws
    // rather than as a file that could not be read. A flipped byte in the
    // part that lists the sheets is enough to produce one; six of six hundred
    // damaged workbooks did (tools/fuzz-open.js).
    // The first sheet showing: a hidden first sheet is not where Excel opens.
    const hiddenAtOpen = new Set(this.workbook.hiddenSheets());
    this.activeSheet = this.workbook.sheetNames().find((n) => !hiddenAtOpen.has(n)) ?? this.workbook.sheetNames()[0];
    if (!this.activeSheet) throw new Error('not a workbook: it lists no sheets');
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
    styles = false, parts = null, structural = false, tracksNewParts = false, sheetGate = true,
  } = {}) {
    const outermost = this._editDepth === 0;
    // What is read once per state of the workbook (the slicers' panels) is
    // read again after any edit.
    if (outermost) this._editStamp = (this._editStamp ?? 0) + 1;
    // THE protection gate. One gate, here, because every cell-changing
    // gesture already hands `_edit` the cells it will touch — that list is
    // the undo footprint, and the same list is exactly what protection needs
    // to judge. Refusing BEFORE the history record keeps the house rule that
    // a refused edit leaves no undo step. A change to the workbook's
    // structure (a sheet added, moved, hidden) is not the sheet's to refuse
    // — the workbook's protection judges it — so it passes `sheetGate: false`.
    if (outermost && sheetGate) {
      const p = this.protection();
      if (p.sheet) {
        if (structural) {
          throw protectionError('This sheet is protected — unprotect it before changing its structure.');
        }
        // A pure formatting edit is governed by the author's formatCells
        // allowance, not by each cell's locked flag — Excel's own split.
        if (!(styles && p.formatCells)) {
          // Review → Allow Edit Ranges: a locked cell in a range the author
          // left editable may change — at once when the range has no
          // password, once it has been unlocked when it has.
          const ranges = cells.length && !styles ? this._editRangeIndex() : null;
          for (const { row, col } of cells) {
            if (!this.isCellLocked(row, col)) continue;
            const open = ranges ? this._rangeVerdict(ranges, row, col) : null;
            if (open === true) continue;
            if (open) {
              throw protectionError('The range "' + open + '" is protected by a password — unlock it to change ' + ref(row, col) + '.', { range: open });
            }
            throw protectionError('This sheet is protected and ' + ref(row, col)
              + ' is locked. Unprotect the sheet to change it.');
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
      if (outermost) this._editStamp += 1;
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
    this._editStamp = (this._editStamp ?? 0) + 1;

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
        // A part the edit made and a later read parsed (a new sheet drawn)
        // goes from the parsed set too, or the next save writes it back.
        if (direction === 'undo') { this.pkg.removePart(name); this.workbook._loaded.delete(name); }
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
        // workbook.xml also carries the calculation mode an undo may put back.
        this.calc.manual = this.workbook.calcMode() === 'manual';
        this.calc.recalculate();
      }
      this._afterPartsRestored(Object.keys(entry.state.parts));
    }
    if (entry.state.structuralDirty !== undefined) this._structuralDirty = entry.state.structuralDirty;
    this.dirtyCells = new Set(entry.state.dirty);
    this.styledCells = new Set(entry.state.styled ?? []);
    if (entry.meta?.selection) this.selection = entry.meta.selection;
    this.editing = null;
    this.ensureVisible();
    return true;
  }

  /**
   * What is derived from parts other than a sheet's, read again once an undo
   * or redo has put those parts back: a chart, a slicer or a drawing redraws
   * every sheet's drawings; the theme or the style table re-reads the styles.
   */
  _afterPartsRestored(names) {
    if (names.some((n) => /^xl\/(theme|styles)/.test(n))) {
      this.styles = readStyles(this.pkg, { BUILTIN_FORMATS });
      for (const { name, part } of this.workbook.sheets()) {
        this.conditionals.set(name, readConditionalFormatting(this.workbook.snapshotParts([part])[part], this.styles.theme));
      }
    }
    if (names.some((n) => /^xl\/(charts|drawings|slicers|slicerCaches|pivotTables|pivotCache|theme)\//.test(n) || n === this.workbook.mainPart)) {
      for (const { name, part } of this.workbook.sheets()) this.drawings.set(name, this._readDrawings(part));
      this._pivots = null;
    }
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
    const index = this._styleIndexAt(this.activeSheet, row, col);
    return index === null ? null : this.styles.byStyleIndex[index] ?? null;
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
    // Each drawing is known by its own id where that is unique on the sheet,
    // so a selection survives Bring Forward; by its place otherwise.
    const count = new Map();
    for (const d of found) count.set(d.nvId, (count.get(d.nvId) ?? 0) + 1);
    return found.map((d, i) => ({
      ...d,
      index: i,
      id: d.nvId && count.get(d.nvId) === 1 ? 'd' + d.nvId : 'drawing-' + i,
      // A PivotChart names the pivot it is drawn from.
      pivot: d.kind === 'chart' && d.part && this.pkg.has(d.part) ? (readPivotChart(this.pkg.text(d.part))?.name ?? null) : null,
    }));
  }

  /**
   * Where a drawing sits on the sheet, in pixels: its top-left from the
   * anchor's `from` marker, its size from the `to` marker (a two-cell
   * anchor resizes with its cells) or its own extent.
   */
  _drawingBox(d, geo = this.geo) {
    const x = d.from ? geo.colOffset(d.from.col) + Math.round((d.from.colOffsetEmu ?? 0) / 9525) : 0;
    const y = d.from ? geo.rowOffset(d.from.row) + Math.round((d.from.rowOffsetEmu ?? 0) / 9525) : 0;
    // "Move but don't size with cells" (a slicer's way): its own extent
    // holds while rows under it are hidden by a filter.
    if ((d.editAs === 'oneCell' || d.editAs === 'absolute') && d.xfrmPx) {
      return { x, y, width: d.xfrmPx.width, height: d.xfrmPx.height };
    }
    // A chart smaller than a postcard is a chart nobody can read, so it
    // gets a floor; a shape is whatever size its author drew — a connector
    // one pixel wide and a column tall was being widened to eighty.
    const floor = d.kind === 'chart' ? { w: 80, h: 60 } : { w: 1, h: 1 };
    const width = d.to
      ? Math.max(floor.w, geo.colOffset(d.to.col) + Math.round((d.to.colOffsetEmu ?? 0) / 9525) - x)
      : Math.round(d.widthPx ?? 320);
    const height = d.to
      ? Math.max(floor.h, geo.rowOffset(d.to.row) + Math.round((d.to.rowOffsetEmu ?? 0) / 9525) - y)
      : Math.round(d.heightPx ?? 240);
    return { x, y, width, height };
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
      const shadow = Boolean(authors[authorId] && authors[authorId].startsWith('tc='));
      const author = authors[authorId] && !shadow ? authors[authorId] : null;
      out.set(refMatch[1], shadow ? { author, text, shadow } : { author, text });
    }
    return out;
  }

  sheetNames() { return this.workbook.sheetNames(); }

  /**
   * The hyperlinks of a sheet by cell, from its tail. Read lazily — a sheet
   * that is never shown is never parsed for them — and keyed on the tail's
   * text, so an edit or an undo that changes the block is seen next frame.
   */
  _links(sheetName = this.activeSheet) {
    const { part } = this.workbook._sheetPart(sheetName);
    const cached = this.links.get(sheetName);
    if (cached && cached.suffix === part.suffix) return cached.map;
    const map = new Map(this.workbook.hyperlinks(sheetName).map((h) => [h.ref, h]));
    this.links.set(sheetName, { suffix: part.suffix, map });
    return map;
  }

  /**
   * The notes of a sheet by cell, read again whenever the comments part
   * changes — a note put on, an undo — the way the links follow the tail.
   */
  _notes(sheetName = this.activeSheet) {
    const sheetPartName = this.workbook.partNameFor(sheetName);
    const rel = this.pkg.rels(sheetPartName).find((r) => String(r.Type).endsWith('/comments'));
    const partName = rel ? OoxmlPackage.resolveTarget(sheetPartName, rel.Target) : null;
    const xml = partName && this.pkg.has(partName) ? this.pkg.text(partName) : '';
    const threads = this._threads(sheetName);
    const cached = this.notes.get(sheetName);
    if (cached && cached.xml === xml && cached.threads === threads) return cached.map;
    const map = this._readComments(sheetPartName);
    // A threaded comment's classic shadow is not a note: the thread itself
    // is drawn. A shadow whose thread has gone (a tool that dropped the
    // threaded part) is all there is, and shows as the note it reads as.
    for (const [at, note] of map) if (note.shadow && threads.has(at)) map.delete(at);
    this.notes.set(sheetName, { xml, map, threads });
    return map;
  }

  /**
   * Review → comments: the threads of a sheet by cell, read again whenever
   * its threaded-comments part or the person list changes — a comment put
   * on, an undo — the way the notes follow their part.
   */
  _threads(sheetName = this.activeSheet) {
    const sheetPartName = this.workbook.partNameFor(sheetName);
    const rel = this.pkg.rels(sheetPartName).find((r) => String(r.Type).endsWith('/threadedComment'));
    const partName = rel ? OoxmlPackage.resolveTarget(sheetPartName, rel.Target) : null;
    const xml = partName && this.pkg.has(partName) ? this.pkg.text(partName) : '';
    const personsPart = this.workbook._personsPart();
    const people = personsPart ? this.pkg.text(personsPart) : '';
    if (!this._threadCache) this._threadCache = new Map();
    const cached = this._threadCache.get(sheetName);
    if (cached && cached.xml === xml && cached.people === people) return cached.map;
    const map = new Map(xml ? this.workbook.threads(sheetName).map((t) => [t.ref, t]) : []);
    this._threadCache.set(sheetName, { xml, people, map });
    return map;
  }

  /** The parts a comment lives in, for the undo record — a note's, the thread's, and the people's. */
  _threadParts() {
    const parts = [...this._noteParts(), '[Content_Types].xml', this.workbook.mainPart];
    const sheetPartName = this.workbook.partNameFor(this.activeSheet);
    for (const rel of this.pkg.rels(sheetPartName)) {
      if (!String(rel.Type).endsWith('/threadedComment')) continue;
      const name = OoxmlPackage.resolveTarget(sheetPartName, rel.Target);
      if (this.pkg.has(name)) parts.push(name);
    }
    const mainRels = OoxmlPackage.relsPathFor(this.workbook.mainPart);
    if (this.pkg.has(mainRels)) parts.push(mainRels);
    const persons = this.workbook._personsPart();
    if (persons) parts.push(persons);
    return [...new Set(parts)];
  }

  /**
   * Review → New Comment: a comment on a cell — a new thread, or a reply
   * when the cell already has one — signed by `author`, dated now. Written
   * as Excel writes one: the threaded part, the person, and the classic
   * shadow with its box. One undo step. Returns the comment's id.
   */
  addComment({ row = this.selection.active.row, col = this.selection.active.col, text = '', author = '', date = new Date() } = {}) {
    const at = ref(row, col);
    // Refused before the history hears of it, so a refusal leaves no undo step.
    if (!String(text ?? '').trim()) throw new Error('a comment needs some words');
    if (this._notes().has(at)) throw new Error('This cell has a note. Delete the note, or edit it, before starting a comment here.');
    let id = null;
    this._edit('comment', null, [], () => {
      id = this.workbook.addThreadedComment(this.activeSheet, at, { author, text, date });
      this._structuralDirty = true;
    }, { parts: this._threadParts(), tracksNewParts: true });
    return id;
  }

  /** A comment's words changed. One undo step. */
  editComment({ id, text }) {
    this._edit('edit comment', null, [], () => {
      this.workbook.editThreadedComment(this.activeSheet, id, text);
      this._structuralDirty = true;
    }, { parts: this._threadParts() });
    return this;
  }

  /** A reply taken out, or — the first comment — its thread. One undo step. */
  deleteComment({ id }) {
    let gone = false;
    this._edit('delete comment', null, [], () => {
      gone = this.workbook.deleteThreadedComment(this.activeSheet, id);
      this._structuralDirty = true;
    }, { parts: this._threadParts() });
    return gone;
  }

  /** Review → Delete: the whole thread on a cell. */
  deleteThread({ row = this.selection.active.row, col = this.selection.active.col } = {}) {
    const thread = this._threads().get(ref(row, col));
    if (!thread) throw new Error('There is no comment on ' + ref(row, col) + ' to delete.');
    return this.deleteComment({ id: thread.id });
  }

  /** Resolve a cell's thread, or reopen it. One undo step. */
  resolveComment({ row = this.selection.active.row, col = this.selection.active.col, done = true } = {}) {
    this._edit(done ? 'resolve thread' : 'reopen thread', null, [], () => {
      this.workbook.setThreadDone(this.activeSheet, ref(row, col), done);
      this._structuralDirty = true;
    }, { parts: this._threadParts() });
    return this;
  }

  /**
   * Every thread in the workbook, sheet by sheet in tab order and cell by
   * cell down each sheet — the Comments pane's list, and the order Previous
   * and Next walk.
   */
  allThreads() {
    const out = [];
    for (const sheet of this.sheetNames()) {
      const list = [...this._threads(sheet).values()].map((t) => {
        const [r, c] = parseRefPair(t.ref);
        return { ...t, sheet, row: r, col: c };
      });
      list.sort((a, b) => a.row - b.row || a.col - b.col);
      out.push(...list);
    }
    return out;
  }

  /**
   * Review → Previous / Next: the thread before or after the active cell,
   * across the sheets and round again, selected. Returns its ref, or null
   * when the workbook has none.
   */
  stepComment(direction = 'next') {
    const all = this.allThreads();
    if (!all.length) return null;
    const sheets = this.sheetNames();
    const here = { s: sheets.indexOf(this.activeSheet), row: this.selection.active.row, col: this.selection.active.col };
    const key = (t) => [sheets.indexOf(t.sheet), t.row, t.col];
    const cmp = (a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2];
    const at = [here.s, here.row, here.col];
    const pick = direction === 'prev'
      ? [...all].reverse().find((t) => cmp(key(t), at) < 0) ?? all[all.length - 1]
      : all.find((t) => cmp(key(t), at) > 0) ?? all[0];
    if (pick.sheet !== this.activeSheet) this.selectSheet(pick.sheet);
    this.select(pick.row, pick.col);
    return pick.ref;
  }

  /** The parts a note lives in, for the undo record: the sheet, its rels, the comments part and the VML box. */
  _noteParts() {
    const sheetPartName = this.workbook.partNameFor(this.activeSheet);
    const parts = [sheetPartName];
    const relsName = OoxmlPackage.relsPathFor(sheetPartName);
    if (this.pkg.has(relsName)) parts.push(relsName);
    for (const rel of this.pkg.rels(sheetPartName)) {
      if (!/\/(comments|vmlDrawing)$/.test(String(rel.Type))) continue;
      const name = OoxmlPackage.resolveTarget(sheetPartName, rel.Target);
      if (this.pkg.has(name)) parts.push(name);
    }
    return parts;
  }

  /** A note on a cell — the words and who wrote them — as Excel keeps one. One undo step. */
  setNote({ row, col, author = '', text = '' }) {
    const at = ref(row, col);
    // A cell holds a note or a comment thread, never both — Excel's rule,
    // and the thread's classic shadow is the note this would overwrite.
    if (this._threads().has(at)) throw new Error('This cell has a comment. Reply to it, or delete it, before adding a note.');
    return this._edit('note', null, [], () => {
      this.workbook.setComment(this.activeSheet, at, { author, text });
      return this;
    }, { parts: this._noteParts(), tracksNewParts: true });
  }

  /**
   * Format the selection — or, from one cell, the block of data round it —
   * as a table: a header row, banded rows if asked, and a style Excel knows
   * by name, written as Excel keeps a table (the part, the rels, the
   * sheet's `<tableParts>`). The header row is the range's first row; a
   * header cell with nothing in it is given the column's name, as Excel
   * does, so the file and the sheet agree. Painted at once, in the
   * renderer's own palette. One undo step. A range that already holds a
   * table is refused.
   */
  formatAsTable({ style = 'TableStyleMedium2', stripes = true } = {}) {
    const sel = this.selection.range;
    const single = sel.top === sel.bottom && sel.left === sel.right;
    const r = single ? { ...this._currentRegion(sel.top, sel.left) } : { top: sel.top, bottom: sel.bottom, left: sel.left, right: sel.right };
    if (r.bottom < r.top + 1) r.bottom = r.top + 1;
    const clash = this.sheetTables().some((t) => !(r.bottom < t.top || r.top > t.bottom || r.right < t.left || r.left > t.right));
    if (clash) throw new Error('That range already has a table in it.');
    const headers = [];
    for (let c = r.left; c <= r.right; c++) headers.push(String(this.displayValue(r.top, c).text ?? '').trim());
    const sheetPartName = this.workbook.partNameFor(this.activeSheet);
    const parts = [sheetPartName, '[Content_Types].xml'];
    const relsName = OoxmlPackage.relsPathFor(sheetPartName);
    if (this.pkg.has(relsName)) parts.push(relsName);
    // Every header cell may be renamed — an empty one is given its column's
    // name, a repeated one its unique name (Qty, Qty2) — so all of them are
    // the edit's cells.
    const touched = headers.map((_, i) => ({ row: r.top, col: r.left + i }));
    return this._edit('format as table', null, touched, () => {
      const written = this.workbook.addTable(this.activeSheet, areaRef(r), { style, stripes, headerNames: headers });
      written.columns.forEach((name, i) => { if (name !== headers[i]) this._setCell(r.top, r.left + i, name); });
      return this;
    }, { parts, tracksNewParts: true });
  }

  /**
   * A picture at the active cell, as Excel keeps one: the bytes become a
   * media part, the sheet's drawing part (made if it has none) gets a
   * one-cell anchor at the cell pointing at them, at the size asked for.
   * Drawn at once over the cells. One undo step, the parts going with it.
   */
  insertPicture({ name = 'Picture', contentType, data, widthPx, heightPx } = {}) {
    const ext = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'webp' }[contentType];
    if (!ext) throw new Error('unsupported image type: ' + contentType + ' (png, jpeg, gif, bmp or webp)');
    const bytes = Buffer.isBuffer(data) ? data : data instanceof Uint8Array ? Buffer.from(data) : Buffer.from(String(data ?? ''), 'base64');
    if (!bytes.length) throw new Error('the picture has no bytes');
    const w = Math.round(Number(widthPx));
    const h = Math.round(Number(heightPx));
    if (!(w > 0) || !(h > 0)) throw new Error('a picture needs a positive width and height');
    const { sheetPartName, parts } = this._drawingEditParts();
    const { row, col } = this.selection.active;
    return this._edit('insert picture', null, [], () => {
      const drawingPart = this.workbook.ensureSheetDrawing(this.activeSheet);
      // The package's part numbering counts XML parts; a media part has the
      // picture's own extension, so the next free number is found here.
      let n = 1;
      while (this.pkg.partNames().some((p) => p.startsWith('xl/media/image' + n + '.'))) n += 1;
      const mediaName = 'xl/media/image' + n + '.' + ext;
      this.pkg.ensureDefault(ext, contentType);
      this.pkg.addPart(mediaName, bytes);
      const rId = this.pkg.addRelationshipTo(drawingPart, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', '../media/image' + n + '.' + ext);
      this.workbook.appendDrawingAnchor(drawingPart, (id) => drawingAnchorXml({
        kind: 'picture', id, name: String(name || 'Picture'), from: { col, row }, widthPx: w, heightPx: h,
      }, () => rId));
      this.drawings.set(this.activeSheet, this._readDrawings(sheetPartName));
      this._structuralDirty = true;
      return this;
    }, { parts, tracksNewParts: true });
  }

  removeNote({ row, col }) {
    const at = ref(row, col);
    return this._edit('remove note', null, [], () => {
      this.workbook.removeComment(this.activeSheet, at);
      return this;
    }, { parts: this._noteParts() });
  }

  /**
   * A hyperlink on a cell: an address that opens outside the suite, or a
   * place in this workbook that is gone to. Both are the file's own
   * hyperlinks block, so Excel sees them too.
   */
  setHyperlink({ row, col, href = null, location = null, tooltip = null }) {
    const at = ref(row, col);
    const sheetPartName = this.workbook.partNameFor(this.activeSheet);
    const parts = [sheetPartName];
    const relsName = OoxmlPackage.relsPathFor(sheetPartName);
    if (this.pkg.has(relsName)) parts.push(relsName);
    return this._edit('hyperlink', null, [], () => {
      this.workbook.setHyperlink(this.activeSheet, at, { href: href || null, location: href ? null : (location || null), tooltip: tooltip || null });
      return this;
    }, { parts, tracksNewParts: true });
  }

  removeHyperlink({ row, col }) {
    const at = ref(row, col);
    const sheetPartName = this.workbook.partNameFor(this.activeSheet);
    return this._edit('remove hyperlink', null, [], () => {
      this.workbook.removeHyperlink(this.activeSheet, at);
      return this;
    }, { parts: [sheetPartName] });
  }

  /** The next free "SheetN", as Excel names a new tab. */
  _freshSheetName() {
    const taken = new Set(this.sheetNames().map((n) => n.toLowerCase()));
    let n = taken.size + 1;
    while (taken.has('sheet' + n)) n += 1;
    return 'Sheet' + n;
  }

  /**
   * A new sheet at the end of the tabs, and the view on it. One undo step:
   * undoing takes the part out again.
   */
  addSheet(name) {
    this._structureGate();
    const proposed = String(name ?? '').trim() || this._freshSheetName();
    this.workbook._checkSheetName(proposed);
    this._edit('add sheet', null, [], () => {
      this.workbook.addSheet(proposed);
      this._rebuildDerivedState();
      this._structuralDirty = true;
      return this;
    }, { parts: [this.workbook.mainPart, 'xl/_rels/workbook.xml.rels'], tracksNewParts: true, structural: true, sheetGate: false });
    this.selectSheet(proposed);
    return proposed;
  }

  /**
   * Rename a sheet: the tab, the names that point into it and every formula
   * that reads it. One undo step over every part it touched.
   */
  renameSheet(from, to) {
    this._structureGate();
    const clean = this.workbook._checkSheetName(to, { except: from });
    if (clean === from) return this;
    const parts = [this.workbook.mainPart, ...this.workbook.sheets().map((s) => s.part)];
    this._edit('rename sheet', null, [], () => {
      this.workbook.renameSheet(from, clean);
      if (this.activeSheet === from) this.activeSheet = clean;
      for (const map of [this.links, this.notes, this.geometry, this.cellStyles, this.merges, this.validations, this.conditionals]) {
        if (map.has(from)) { map.set(clean, map.get(from)); map.delete(from); }
      }
      this._rebuildDerivedState();
      this._structuralDirty = true;
      return this;
    }, { parts, structural: true, sheetGate: false });
    return this;
  }

  /**
   * Delete a sheet. Not undoable — the part is gone from the package, as in
   * Excel, which says so before it does it; the window asks first.
   */
  removeSheet(name) {
    this._structureGate();
    const hidden = new Set(this.hiddenSheets());
    if (!hidden.has(name) && this.sheetNames().filter((n) => !hidden.has(n)).length <= 1 && this.sheetNames().length > 1) {
      throw new Error('A workbook must contain at least one visible worksheet.');
    }
    this.workbook.removeSheet(name);
    if (this.activeSheet === name) this.activeSheet = this.sheetNames().find((n) => !hidden.has(n)) ?? this.sheetNames()[0];
    this.selection = Selection.at(0, 0);
    for (const map of [this.links, this.notes, this.geometry, this.cellStyles, this.merges, this.validations, this.conditionals]) map.delete(name);
    this._rebuildDerivedState();
    this.history = new History();
    this._structuralDirty = true;
    return this;
  }

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
  /**
   * The frame the window draws. In View → Page Layout the window scrolls
   * over paper: its scroll is turned into the sheet's before the grid frame
   * is built, and every position in the frame is then put where it lies on
   * its page (see `_pageLayout`).
   */
  render() {
    if (this.viewMode() !== 'pageLayout') return this._renderGrid();
    const layout = this._pageLayout();
    const paper = { x: this.scrollX, y: this.scrollY };
    this.scrollX = layout.toSheetX(paper.x);
    this.scrollY = layout.toSheetY(paper.y);
    this._inPageLayout = true;
    let frame;
    try {
      frame = this._renderGrid();
    } finally {
      this.scrollX = paper.x;
      this.scrollY = paper.y;
      this._inPageLayout = false;
    }
    // The layout reaches a little past what has been seen, and grows as the
    // window scrolls toward its end — the way the grid's own canvas spreads.
    const vp = frame.viewport;
    const ext = this._plExtent?.sheet === this.activeSheet ? this._plExtent : { sheet: this.activeSheet, rows: 0, cols: 0 };
    if (vp.lastRow > layout.rowsTo - 60 || vp.lastCol > layout.colsTo - 6) {
      this._plExtent = { sheet: this.activeSheet, rows: Math.max(ext.rows, vp.lastRow + 400), cols: Math.max(ext.cols, vp.lastCol + 20) };
    }
    return pageLayoutFrame(frame, layout, this.geo, {
      x: paper.x, y: paper.y, width: this.viewportWidth, height: this.viewportHeight,
    });
  }

  _renderGrid() {
    const geo = this.geo;
    const vp = geo.viewport({
      scrollX: this.scrollX,
      scrollY: this.scrollY,
      width: this.viewportWidth,
      height: this.viewportHeight,
      overscan: OVERSCAN_ROWS,
      overscanCols: OVERSCAN_COLS,
    });

    // Frozen rows and columns ride in EVERY frame, wherever the viewport has
    // scrolled — the client pins them while the rest slides underneath.
    // Page Layout draws the sheet as paper: no pane is frozen or split there.
    const frozen = this._inPageLayout ? { rows: 0, cols: 0 } : this.frozenPane();
    let rowIndices = [];
    for (let r = 0; r < frozen.rows && r < vp.firstRow; r++) rowIndices.push(r);
    for (let r = vp.firstRow; r <= vp.lastRow; r++) rowIndices.push(r);
    let colIndices = [];
    for (let c = 0; c < frozen.cols && c < vp.firstCol; c++) colIndices.push(c);
    for (let c = vp.firstCol; c <= vp.lastCol; c++) colIndices.push(c);

    // A split window's top and left panes scroll on their own: their rows
    // and columns ride in every frame too, from wherever each pane is.
    const split = this._inPageLayout ? null : this.splitPane();
    let splitFrame = null;
    if (split) {
      const rowsTop = [];
      const colsLeft = [];
      for (let r = split.top, y = 0; split.height && r < MAX_ROWS && y < split.height && rowsTop.length < 400; r++) {
        rowsTop.push(r);
        y += geo.rowHeight(r);
      }
      for (let c = split.left, x = 0; split.width && c < MAX_COLS && x < split.width && colsLeft.length < 200; c++) {
        colsLeft.push(c);
        x += geo.colWidth(c);
      }
      if (rowsTop.length) rowIndices = [...new Set([...rowsTop, ...rowIndices])].sort((a, b) => a - b);
      if (colsLeft.length) colIndices = [...new Set([...colsLeft, ...colIndices])].sort((a, b) => a - b);
      splitFrame = {
        ...split,
        topY: geo.rowOffset(split.top),
        leftX: geo.colOffset(split.left),
        rows: rowsTop,
        cols: colsLeft,
      };
    }

    const cells = [];
    // Conditional-formatting range aggregates, computed once per pass — and
    // the sheet's tables, read once here and reused by the frame below.
    const cfCache = new Map();
    const sheetTables = this.sheetTables();
    const links = this._links();
    const notes = this._notes();
    const threads = this._threads();
    // A sparkline cell is normally blank, and a blank cell that draws
    // nothing else is otherwise left out of the frame entirely — so its ref
    // is read once here, the way a note's or a link's is, to keep it in.
    const sparklineCells = new Set();
    for (const g of this.sparklineGroups()) {
      for (const s of g.sparklines) sparklineCells.add(String(s.at).split(':')[0]);
    }
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
        const note = notes.get(ref(row, col)) ?? null;
        const link = links.get(ref(row, col)) ?? null;
        const thread = threads.get(ref(row, col)) ?? null;
        // A styled but empty cell still has to be drawn: a shaded header with no
        // text in it is a real thing, and skipping it leaves a hole in the band.
        const decorated = Boolean(style && (style.fill || style.border)) || Boolean(cf?.bar)
          || Boolean(cf?.icon) || Boolean(note) || Boolean(link) || Boolean(thread) || sparklineCells.has(ref(row, col));
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
          // Excel's textRotation, for the painter: 1..90 anticlockwise,
          // 91..180 clockwise less 90, 255 stacked upright, 0 horizontal.
          rotation: style?.align?.rotation ?? 0,
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
          // The cell's hyperlink, for the underline, the tip and the follow.
          link,
          // The cell's comment thread, for Excel's purple corner and its tip:
          // who started it, its first words, how many replies, whether resolved.
          thread: thread ? {
            id: thread.id,
            done: thread.done,
            replies: thread.comments.length - 1,
            author: thread.comments[0]?.author ?? '',
            text: thread.comments[0]?.text ?? '',
          } : null,
          merged: merge ? { ref: merge.ref, rows: merge.bottom - merge.top + 1, cols: merge.right - merge.left + 1 } : null,
          isError: isError(this.calc.getValue(this.activeSheet, row, col)),
          // A number never spills into its neighbours (Excel shows #### instead);
          // text does. The painter needs to know which it is drawing.
          numeric: typeof this.calc.getValue(this.activeSheet, row, col) === 'number',
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

    const palette = this._themeAccents();
    const drawings = (this.drawings.get(this.activeSheet) ?? []).map((d) => {
      const { x, y, width, height } = this._drawingBox(d);

      if (x > visibleRight || y > visibleBottom || x + width < visibleLeft || y + height < visibleTop) {
        return null;
      }

      let svg = null;
      let unsupported = null;
      let slicer = null;
      let members = null;
      const box = { x: 0, y: 0, width, height };
      try {
        if (d.kind === 'slicer') {
          // A slicer is a panel of buttons the window draws itself, from
          // the state its table or pivot is in now.
          slicer = this.slicerState(d.slicerName) ?? { name: d.slicerName, caption: d.slicerName, items: [], broken: 'its slicer part is missing' };
        } else if (d.spec) {
          svg = renderSvg(buildChart({ ...d.spec, width, height, mode: this.mode, palette }));
        } else if (d.descriptor?.kind === 'shape') {
          svg = renderSvg(scene({
            width, height, mode: this.mode, background: 'none',
            title: d.name ?? 'Shape',
            children: [buildShape(d.descriptor, box, { mode: this.mode, palette })],
          }));
        } else if (d.descriptor?.kind === 'picture' && d.descriptor.href) {
          svg = renderSvg(scene({
            width, height, mode: this.mode, background: 'none',
            title: d.name ?? 'Picture',
            children: [buildPicture(d.descriptor, box)],
          }));
        } else if (d.descriptor?.kind === 'group') {
          // A group's members, each laid out in the group's box by its place
          // in the group's own space, and drawn as it would be on its own.
          members = d.descriptor.children.map((c, i) => {
            const b = {
              x: Math.round(c.frac.x * width), y: Math.round(c.frac.y * height),
              width: Math.max(1, Math.round(c.frac.w * width)), height: Math.max(1, Math.round(c.frac.h * height)),
            };
            const own = { x: 0, y: 0, width: b.width, height: b.height };
            let child = null;
            if (c.kind === 'chart' && c.spec) child = renderSvg(buildChart({ ...c.spec, width: b.width, height: b.height, mode: this.mode, palette }));
            else if (c.kind === 'shape' && c.descriptor) child = renderSvg(scene({ width: b.width, height: b.height, mode: this.mode, background: 'none', title: c.name ?? 'Shape', children: [buildShape(c.descriptor, own, { mode: this.mode, palette })] }));
            else if (c.kind === 'image' && c.descriptor?.href) child = renderSvg(scene({ width: b.width, height: b.height, mode: this.mode, background: 'none', title: c.name ?? 'Picture', children: [buildPicture(c.descriptor, own)] }));
            return { key: i, ...b, svg: child, kind: c.kind };
          });
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
      return {
        id: d.id, kind: d.kind, name: d.name, x, y, width, height, svg, unsupported,
        anchor: d.from ? { row: d.from.row, col: d.from.col } : null,
        index: d.index, hidden: Boolean(d.hidden), pivot: d.pivot ?? null,
        // The turn and flips of what turns (a shape and a picture draw theirs
        // in their own SVG; a group's the window applies to the whole).
        rot: d.descriptor?.rotation ?? 0, flipH: Boolean(d.descriptor?.flipH), flipV: Boolean(d.descriptor?.flipV),
        ...(slicer ? { slicer } : {}),
        ...(members ? { members } : {}),
      };
    }).filter(Boolean);

    const active = this.selection.active;
    return {
      drawings,
      // Page Layout → Themes: the theme the workbook wears now, for the galleries to tick.
      design: this.workbookDesign(),
      // The Normal style's face, which a cell with no font of its own is
      // written in — once the workbook wears a theme of its own (the rule
      // charts follow); otherwise the window's own face, as ever.
      defaultFont: palette ? (this.styles.byStyleIndex[0]?.font?.family ?? this.workbookDesign().fonts.minor) : null,
      // Every drawing on the sheet, seen or not, in drawing order — what the
      // Selection Pane lists (front first, as Excel's does).
      objects: (this.drawings.get(this.activeSheet) ?? []).map((d) => ({
        id: d.id, kind: d.kind, name: d.kind === 'slicer' ? d.slicerName : d.name, hidden: Boolean(d.hidden), index: d.index,
      })),
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
      // The active cell's hyperlink, for the status bar and the Link dialog,
      // and how many notes this sheet carries.
      link: links.get(ref(active.row, active.col)) ?? null,
      notes: notes.size,
      // The active cell's note, for the status bar and the Note dialog.
      note: notes.get(ref(active.row, active.col)) ?? null,
      // Review → comments: how many threads this sheet has, the active
      // cell's whole thread for its card, and — while the Comments pane is
      // open — every thread in the workbook for the list.
      threads: threads.size,
      thread: threads.get(ref(active.row, active.col)) ?? null,
      comments: this.commentsOpen ? this.allThreads() : null,
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
      // Review → Protect Workbook, the tabs Hide has put away, the sheet's
      // Allow Edit Ranges, and — when the active cell is in a password range
      // not yet unlocked — that range's title, so the window asks first.
      workbookProtection: this.workbookProtection(),
      hiddenSheets: this.hiddenSheets(),
      editRanges: this.editRanges(),
      rangeLock: this.rangeLockAt(active.row, active.col),
      // View → Custom Views — the list, and Excel's reason when it is greyed —
      // and Page Layout → Background, the picture's part (the window fetches
      // its bytes once), and View → Ruler for Page Layout.
      customViews: this.customViews(),
      customViewsBlocked: this.customViewsBlocked(),
      background: this.sheetBackground(),
      showRuler: this.showRuler(),
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
      // The outline: per axis its depth and the groups near the viewport,
      // with their pixel extents, for the gutter's brackets and buttons.
      // Null when nothing on the sheet is grouped.
      outline: this._inPageLayout ? null : outlineFrame(geo, vp, frozen),
      // View → Split: the two (or four) panes' sizes, where the top and left
      // ones have scrolled to, and the rows and columns they show; null
      // when the window is not split.
      split: splitFrame,
      // View → Normal / Page Break Preview, and — in the preview — the
      // pages as the printer will cut them: the printed area, each page's
      // box and number, and each break, put by hand or by the paper.
      viewMode: this.viewMode(),
      pageBreaks: this.viewMode() === 'pageBreakPreview' ? this.pageBreakPreview({ viewport: vp }) : null,
      // The frozen pane, with its band sizes in pixels for the client's clip.
      frozen: {
        rows: frozen.rows,
        cols: frozen.cols,
        width: frozen.cols ? geo.colOffset(frozen.cols) : 0,
        height: frozen.rows ? geo.rowOffset(frozen.rows) : 0,
      },
      selection: this._selectionFields(active),
      validation: this._validationFields(active),
      formulaBar: this.editing ? this.editing.draft : this.editValue(active.row, active.col),
      editing: this.editing ? { ...this.editing } : null,
      status: this.statusLine(),
      // Formulas → Calculation Options, and whether the status bar says "Calculate".
      calc: this.calcState(),
    };
  }

  _selectionFields(active) {
    return {
      ...this.selection.range,
      ref: this.selection.toString(),
      active: { ...active, ref: ref(active.row, active.col) },
      // Every rectangle when Ctrl+click has added some; the grid paints them all.
      ranges: this.selection.isMultiple ? this.selection.allRanges : null,
      // Whole rows or whole columns, for Group: 'row', 'col', or null when
      // the selection is neither and Excel would ask which.
      whole: selectionAxis(this),
    };
  }

  /**
   * The active cell's validation rule, for the in-cell dropdown and the
   * author's input prompt. Null for the unruled majority.
   */
  _validationFields(active) {
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
  }

  /**
   * What a selection move changes when nothing scrolls: the selection, the
   * active cell's format, validation and formula-bar text, the status strip
   * and the canvas extent. The cells on screen are the same cells, so the
   * frame — a quarter of a megabyte on a wide sheet, a hundred milliseconds
   * to build — need not be built or sent. An arrow key on an eighteen-
   * million-cell workbook took that long, and a held key queued one per
   * repeat.
   */
  selectionFrame() {
    const geo = this.geo;
    const vp = geo.viewport({ scrollX: this.scrollX, scrollY: this.scrollY, width: this.viewportWidth, height: this.viewportHeight });
    const active = this.selection.active;
    return {
      selection: this._selectionFields(active),
      format: this.formatState(),
      validation: this._validationFields(active),
      formulaBar: this.editing ? this.editing.draft : this.editValue(active.row, active.col),
      status: this.statusLine(),
      // What belongs to the active cell rides with it: its link, its note,
      // its comment thread (the card a comment opens follows the cell).
      link: this._links().get(ref(active.row, active.col)) ?? null,
      note: this._notes().get(ref(active.row, active.col)) ?? null,
      thread: this._threads().get(ref(active.row, active.col)) ?? null,
      rangeLock: this.rangeLockAt(active.row, active.col),
      total: this.viewMode() === 'pageLayout' ? this._pageLayout().total : this._stickyTotal(geo, {
        maxRow: Math.max(this.bounds.maxRow, this.selection.range.bottom, vp.lastRow),
        maxCol: Math.max(this.bounds.maxCol, this.selection.range.right, vp.lastCol),
      }),
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

  select(row, col, { extend = false, add = false } = {}) {
    // A cell change ends the run: "type in A1, click B4, type" is two undos.
    this.history.break();
    if (add) this.selection.beginAnother(row, col);
    else if (extend) this.selection.extendTo(row, col);
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
    if (this.viewMode() === 'pageLayout') {
      // On paper: the active cell's place on its page, kept in the window.
      const layout = this._pageLayout();
      const y = layout.mapY(row);
      const x = layout.mapX(col);
      const h = this.geo.rowHeight(row);
      const w = this.geo.colWidth(col);
      if (y < this.scrollY) this.scrollY = Math.max(0, y - 12);
      else if (y + h > this.scrollY + this.viewportHeight) this.scrollY = y + h - this.viewportHeight + 12;
      if (x < this.scrollX) this.scrollX = Math.max(0, x - 12);
      else if (x + w > this.scrollX + this.viewportWidth) this.scrollX = x + w - this.viewportWidth + 12;
      return this;
    }
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
    // A cell in a password range, on a protected sheet, asks for the
    // range's password before an edit starts — Excel's Unlock Range.
    const lock = this.rangeLockAt(row, col);
    if (lock) {
      throw protectionError('The range "' + lock + '" is protected by a password — unlock it to change ' + ref(row, col) + '.', { range: lock });
    }
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
    this.calc.manual = this.workbook.calcMode() === 'manual';
    this._pivots = null;
    for (const { name, part } of this.workbook.sheets()) {
      const xml = this.workbook.snapshotParts([part])[part];
      this.geometry.set(name, SheetGeometry.fromSheetXml(xml));
      // A sheet whose styles are read on demand keeps reading them on demand:
      // rebuilding the whole map here is the 16.7-million-entry map the lazy
      // path exists to avoid, and it would drop the cleared-style marks with it.
      this.cellStyles.set(name, this._lazyStyles.has(name) ? new Map() : this._readCellStyles(xml));
      this.merges.set(name, readMergedCells(xml));
      this.validations.set(name, readDataValidations(xml));
      this.conditionals.set(name, readConditionalFormatting(xml, this.styles.theme));
      this.comments.set(name, this._readComments(part));
      // Charts drawn from a pivot, slicers and anything else a structural
      // step put back are read again with the rest.
      this.drawings.set(name, this._readDrawings(part));
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

  // ---- Formula auditing --------------------------------------------------

  /**
   * Trace Precedents / Trace Dependents: the cells a formula reads, or the
   * formulas that read a cell, on this sheet — as ranges to draw arrows to.
   * Names and structured references are counted but not drawn; references
   * to other sheets are counted as `elsewhere`.
   *
   * @param {number} row
   * @param {number} col
   * @param {'precedents'|'dependents'} kind
   * @returns {{ kind: string, at: { row: number, col: number }, arrows: Array<{ top: number, left: number, bottom: number, right: number }>, elsewhere: number }}
   */
  traceOf(row, col, kind = 'precedents') {
    const sheet = this.activeSheet;
    const arrows = [];
    let elsewhere = 0;
    const onThisSheet = (dep) => !dep.sheet || dep.sheet === sheet;
    const rangeOf = (dep) => (dep.type === 'cell'
      ? { top: dep.row, left: dep.col, bottom: dep.row, right: dep.col }
      : dep.type === 'range'
        ? { top: Math.min(dep.start.row, dep.end.row), left: Math.min(dep.start.col, dep.end.col), bottom: Math.max(dep.start.row, dep.end.row), right: Math.max(dep.start.col, dep.end.col) }
        : null);
    if (kind === 'dependents') {
      for (const c of this.calc.sheets.get(sheet)?.values() ?? []) {
        if (!c.deps || c.tombstone) continue;
        const reads = c.deps.some((dep) => {
          if (!onThisSheet(dep)) return false;
          const r = rangeOf(dep);
          return r && row >= r.top && row <= r.bottom && col >= r.left && col <= r.right;
        });
        if (reads) arrows.push({ top: c.row, left: c.col, bottom: c.row, right: c.col });
      }
    } else {
      const cell = this.calc.cell(sheet, row, col);
      for (const dep of cell?.deps ?? []) {
        if (!onThisSheet(dep)) { elsewhere += 1; continue; }
        const r = rangeOf(dep);
        if (!r) { elsewhere += 1; continue; }
        // A whole column or row is drawn as far as the sheet has anything.
        const { maxRow, maxCol } = this.bounds;
        arrows.push({ top: r.top, left: r.left, bottom: Math.min(r.bottom, maxRow), right: Math.min(r.right, maxCol) });
      }
    }
    return { kind, at: { row, col }, arrows, elsewhere };
  }

  /**
   * Error Checking: every cell whose calculated value is an error — divide
   * by zero, a broken reference, an undefined name and the rest — plus every
   * cell the recalc engine settled as part of a circular reference. One pass
   * over the cells the calc model already holds (a formula, or a value
   * someone typed); nothing is recalculated.
   *
   * @param {{ sheet?: string, all?: boolean }} [opts] `all` walks every sheet,
   *   in workbook order, instead of just `sheet` (the active sheet by default).
   * @returns {Array<{ ref: string, row: number, col: number, sheet: string,
   *   formula: string|null, value: string, kind: string, reason: string }>}
   *   In row-major order per sheet.
   */
  errorCells({ sheet = this.activeSheet, all = false } = {}) {
    const sheetNames = all ? this.workbook.sheetNames() : [sheet];
    const out = [];
    for (const sheetName of sheetNames) {
      const cells = this.calc.sheets.get(sheetName);
      if (!cells) continue;
      const found = [];
      for (const cell of cells.values()) {
        if (cell.tombstone || !isError(cell.value)) continue;
        const { kind, reason } = ERROR_KIND[cell.value.type] ?? DEFAULT_ERROR_KIND;
        found.push({
          ref: ref(cell.row, cell.col),
          row: cell.row,
          col: cell.col,
          sheet: sheetName,
          formula: cell.formula ?? null,
          value: cell.value.type,
          kind,
          reason,
        });
      }
      found.sort((a, b) => (a.row - b.row) || (a.col - b.col));
      out.push(...found);
    }
    return out;
  }

  // ---- Data tools --------------------------------------------------------

  /**
   * Text to Columns: each cell of ONE selected column split on a delimiter,
   * the pieces written into the cells to its right (overwriting what is
   * there, as Excel warns and does). One undo step.
   *
   * @param {{ delimiter?: 'comma'|'tab'|'semicolon'|'space'|string }} [spec]
   * @returns {{ rows: number, columns: number }} rows split and the widest split
   */
  textToColumns({ delimiter = 'comma' } = {}) {
    const sep = { comma: ',', tab: '\t', semicolon: ';', space: ' ' }[delimiter] ?? String(delimiter || ',');
    const r = this.selection.range;
    if (r.left !== r.right) throw new Error('Text to Columns takes one column at a time: select the cells of one column');
    const col = r.left;
    const splits = [];
    for (let row = r.top; row <= r.bottom; row++) {
      const input = this.calc.getInput(this.activeSheet, row, col);
      if (input == null || input === '' || String(input).startsWith('=')) continue;
      const parts = sep === ' ' ? String(input).trim().split(/\s+/) : String(input).split(sep).map((s) => s.trim());
      if (parts.length > 1) splits.push({ row, parts });
    }
    if (!splits.length) return { rows: 0, columns: 0 };
    const widest = Math.max(...splits.map((s) => s.parts.length));
    const cells = [];
    for (const s of splits) for (let k = 0; k < widest; k++) cells.push({ row: s.row, col: col + k });
    this._edit('textToColumns', null, cells, () => {
      for (const s of splits) s.parts.forEach((piece, k) => this._setCell(s.row, col + k, piece));
      this.selection.collapseTo(r.top, col);
      this.selection.extendTo(r.bottom, col + widest - 1);
      return this;
    });
    return { rows: splits.length, columns: widest };
  }

  /**
   * Remove Duplicates: the selected range (or the block of data round the
   * cell) loses every row that repeats an earlier one, cell for cell across
   * its columns; the rows that stay close up from the top and the rest of
   * the range is cleared, as Excel does it. A header row, when the block
   * has one, is left alone. One undo step.
   *
   * @returns {{ removed: number, kept: number, range: string }}
   */
  removeDuplicates() {
    const sel = this.selection.range;
    const single = sel.top === sel.bottom && sel.left === sel.right;
    const r = single ? this._currentRegion(sel.top, sel.left) : sel;
    // A header is only inferred for the block found round the cell: rows a person selected by hand are all data.
    const hasHeader = single && this._looksLikeHeader(r);
    const firstData = r.top + (hasHeader ? 1 : 0);
    const inputsOf = (row) => {
      const out = [];
      for (let c = r.left; c <= r.right; c++) out.push(this.calc.getInput(this.activeSheet, row, c) ?? '');
      return out;
    };
    const seen = new Set();
    const kept = [];
    let removed = 0;
    for (let row = firstData; row <= r.bottom; row++) {
      const inputs = inputsOf(row);
      // A row with nothing in it is neither kept nor a duplicate: it drops out when the rest close up.
      if (inputs.every((v) => String(v).trim() === '')) continue;
      const key = JSON.stringify(inputs.map((v) => String(v).trim().toLowerCase()));
      if (seen.has(key)) { removed += 1; continue; }
      seen.add(key);
      kept.push(inputs);
    }
    const range = ref(r.top, r.left) + ':' + ref(r.bottom, r.right);
    if (!removed) return { removed: 0, kept: kept.length, range };
    const cells = [];
    for (let row = firstData; row <= r.bottom; row++) for (let c = r.left; c <= r.right; c++) cells.push({ row, col: c });
    this._edit('removeDuplicates', null, cells, () => {
      for (let i = 0; i < r.bottom - firstData + 1; i++) {
        const inputs = kept[i];
        for (let c = r.left; c <= r.right; c++) {
          const value = inputs ? inputs[c - r.left] : '';
          this._setCell(firstData + i, c, value === '' ? null : value);
        }
      }
      this.selection.collapseTo(r.top, r.left);
      this.selection.extendTo(r.bottom, r.right);
      return this;
    });
    return { removed, kept: kept.length, range };
  }

  // ---- the outline and Subtotal (see outline.js) --------------------------

  /**
   * Data → Group: the selected rows (whole columns: columns) a level
   * deeper. `axis` and the span can be given; by default they are read
   * off the selection, as Shift+Alt+Right reads them.
   */
  group({ axis = selectionAxis(this) ?? 'row', from, to } = {}) {
    const r = this.selection.range;
    return outlineGroup(this, axis, from ?? (axis === 'row' ? r.top : r.left), to ?? (axis === 'row' ? r.bottom : r.right));
  }

  /** Data → Ungroup: the mirror of `group`. */
  ungroup({ axis = selectionAxis(this) ?? 'row', from, to } = {}) {
    const r = this.selection.range;
    return outlineUngroup(this, axis, from ?? (axis === 'row' ? r.top : r.left), to ?? (axis === 'row' ? r.bottom : r.right));
  }

  /** The + or − beside one group, named by its level and first row (or column). */
  toggleOutlineGroup({ axis = 'row', level, start }) { return toggleGroup(this, axis, level, start); }

  /** The outline's level buttons: show level n, fold everything deeper. */
  showOutlineLevel({ axis = 'row', level }) { return showLevel(this, axis, level); }

  /** Data → Show Detail / Hide Detail, at the active cell. */
  showDetail() { return outlineDetail(this, true); }
  hideDetail() { return outlineDetail(this, false); }

  /** Data → Ungroup → Clear Outline. */
  clearOutline() { return clearOutline(this); }

  /** Data → Subtotal, as the dialog specifies it. */
  subtotal(spec) { return subtotal(this, spec); }

  /** Subtotal → Remove All. */
  removeSubtotals() { return removeSubtotals(this); }

  /** The list round the cell and its columns by header — what the Subtotal and Advanced Filter dialogs offer. */
  listFields() { return { ...listFields(this), filter: filterNames(this) }; }

  /** Data → Advanced (see advanced-filter.js). */
  advancedFilter(spec) { return advancedFilter(this, spec); }

  /** Data → Clear: the rows an advanced filter hid, shown again. */
  clearAdvancedFilter() { return clearAdvancedFilter(this); }

  /**
   * Data → Flash Fill (Ctrl+E): the active cell's column filled from the
   * examples typed in it, by the transformation that makes every one of
   * them out of the block's other columns (see flash-fill.js). Nothing is
   * filled when no transformation makes them all. One undo step.
   *
   * @returns {{ filled: number, examples: number }}
   */
  flashFill() {
    const sheet = this.activeSheet;
    const { row, col } = this.selection.active;
    const region = this._currentRegion(row, col);
    const sourceCols = [];
    for (let c = region.left; c <= region.right; c++) if (c !== col) sourceCols.push(c);
    if (!sourceCols.length) throw new Error('Flash Fill works beside a block of data — type an example in the column next to it');
    const sourcesOf = (r) => sourceCols.map((c) => {
      const text = String(this.displayValue(r, c).text ?? '');
      if (text === '') return null;
      const v = this.calc.getValue(sheet, r, c);
      return { text, serial: typeof v === 'number' && isDateFormat(this.formatFor(r, c)) ? v : undefined };
    });
    const typed = (r) => {
      const input = String(this.calc.getInput(sheet, r, col) ?? '');
      return input !== '' && !input.startsWith('=');
    };
    const examples = [];
    const targets = [];
    for (let r = region.top; r <= region.bottom; r++) {
      if (typed(r)) examples.push(r);
      else if (String(this.calc.getInput(sheet, r, col) ?? '') === '' && sourcesOf(r).some(Boolean)) targets.push(r);
    }
    if (!examples.length) throw new Error('Type an example of what you want in this column first, then Flash Fill');
    if (!targets.length) throw new Error('Every row of the column is already filled');
    const infer = (rows) => inferProgram(rows.map((r) => ({ sources: sourcesOf(r), output: String(this.displayValue(r, col).text ?? '') })));
    // The column's heading reads as an example it is not: without it, then.
    let program = infer(examples);
    let used = examples.length;
    if (!program && examples[0] === region.top && examples.length > 1) {
      program = infer(examples.slice(1));
      used -= 1;
    }
    if (!program) throw new Error('Flash Fill found no pattern that makes every example — type another example and try again');
    const results = [];
    for (const r of targets) {
      const value = runProgram(program, sourcesOf(r));
      if (value) results.push({ row: r, value });
    }
    if (!results.length) throw new Error('Flash Fill found the pattern but no other row has the parts it needs');
    this._edit('flash fill', null, results.map(({ row: r }) => ({ row: r, col })), () => {
      for (const { row: r, value } of results) this._setCell(r, col, value);
      return this;
    });
    return { filled: results.length, examples: used };
  }

  /** The style index a label wears bold, made from the one it wears now. */
  _boldIndex(base) {
    this._ensureStylesPart();
    const xml = this.pkg.text('xl/styles.xml');
    const out = applyFormat(xml, base ?? null, { bold: true });
    if (out.xml !== xml) {
      this.pkg.write_('xl/styles.xml', out.xml);
      this.styles = readStyles(this.pkg, { BUILTIN_FORMATS });
      this._stylesDirty = true;
    }
    return out.index;
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
  selectColumn(col, { extend = false, add = false } = {}) {
    this.history.break();
    const maxRow = Math.max(this.bounds.maxRow, this.selection.range.bottom);
    const anchorCol = extend ? this.selection.anchor.col : col;
    if (add) this.selection.beginAnother(0, col);
    else this.selection.collapseTo(0, anchorCol);
    this.selection.extendTo(maxRow, col);
    return this;
  }

  selectRow(row, { extend = false, add = false } = {}) {
    this.history.break();
    const maxCol = Math.max(this.bounds.maxCol, this.selection.range.right);
    const anchorRow = extend ? this.selection.anchor.row : row;
    if (add) this.selection.beginAnother(row, 0);
    else this.selection.collapseTo(anchorRow, 0);
    this.selection.extendTo(row, maxCol);
    return this;
  }

  // ---- sorting -----------------------------------------------------------

  /**
   * The data block around the cursor with its headings — what a pivot panel
   * offers as fields without the person typing them — or null where the
   * cursor is not in a block with a heading row.
   *
   * Asked for, never sent with a frame: growing the region walks every row
   * of the block it finds, which on a sixty-thousand-row table is a third of
   * a second, and four seconds the first time while the lazy sheet reads its
   * rows. The frame carried it on every scroll step, and nothing read it.
   */
  regionAround(row = this.selection.active.row, col = this.selection.active.col) {
    const r = this._currentRegion(row, col);
    if (r.bottom <= r.top || r.right < r.left) return null;
    const headers = [];
    for (let c = r.left; c <= r.right; c++) {
      headers.push(String(this.displayValue(r.top, c).text ?? ''));
    }
    return headers.every((h) => h) ? { ref: areaRef(r), headers } : null;
  }

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
  sortSelection({ ascending = true, keys = null } = {}) {
    let range = { ...this.selection.range };
    const keyCol = Math.min(Math.max(this.selection.active.col, range.left), range.right);
    if (range.top === range.bottom && range.left === range.right) {
      range = this._currentRegion(range.top, range.left);
      if (this._looksLikeHeader(range)) range = { ...range, top: range.top + 1 };
    }
    if (range.top === range.bottom) return this;
    // Several keys, each its own way, as Excel's Sort dialog takes them:
    // the first decides, the next breaks its ties. One key is the active
    // cell's column, the way the A→Z and Z→A buttons have always sorted.
    const keyList = (Array.isArray(keys) && keys.length ? keys : [{ col: keyCol, ascending }])
      .map((k) => ({ col: Math.min(Math.max(Number(k.col), range.left), range.right), dir: k.ascending === false ? -1 : 1 }));

    const rows = [];
    for (let r = range.top; r <= range.bottom; r++) {
      const cells = [];
      for (let c = range.left; c <= range.right; c++) {
        cells.push({ input: this.calc.getInput(this.activeSheet, r, c), styleIndex: this._styleIndexAt(this.activeSheet, r, c) });
      }
      rows.push({ from: r, cells, keys: keyList.map((k) => this.calc.getValue(this.activeSheet, r, k.col)) });
    }

    const kind = (v) => {
      if (v === '' || v === null || v === undefined) return 4; // blanks last, always
      if (typeof v === 'number') return 0;
      if (typeof v === 'string') return 1;
      if (typeof v === 'boolean') return 2;
      return 3; // errors
    };
    const compareKey = (va, vb, dir) => {
      const ka = kind(va);
      const kb = kind(vb);
      if (ka === 4 || kb === 4) return ka === kb ? 0 : ka - kb; // blanks sink regardless of direction
      if (ka !== kb) return (ka - kb) * dir;
      let cmp = 0;
      if (ka === 0) cmp = va - vb;
      else if (ka === 1) {
        const x = va.toUpperCase();
        const y = vb.toUpperCase();
        cmp = x < y ? -1 : x > y ? 1 : 0;
      } else if (ka === 2) cmp = (va ? 1 : 0) - (vb ? 1 : 0);
      return cmp * dir;
    };
    const order = rows.map((rec, i) => ({ rec, i })).sort((a, b) => {
      for (let k = 0; k < keyList.length; k++) {
        const cmp = compareKey(a.rec.keys[k], b.rec.keys[k], keyList[k].dir);
        if (cmp !== 0) return cmp;
      }
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
      this.calc.recalculate({ force: true });
      const v = this.calc.getValue(sheet, setRow, setCol);
      return typeof v === 'number' ? v - goal : NaN;
    };
    const restore = () => {
      this.calc.setCell(sheet, byRow, byCol, byInput ?? '');
      this.calc.recalculate({ force: true });
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
      this.calc.recalculate({ force: true });
    };

    const writes = [];
    try {
      if (twoWay) {
        for (const down of downValues) {
          setInput(colCell, down.value);
          for (const across of acrossValues) {
            setInput(rowCell, across.value);
            this.calc.recalculate({ force: true });
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
          this.calc.recalculate({ force: true });
          for (const f of formulaCells) {
            writes.push({ row: down.row, col: f.col, value: this.calc.getValue(sheet, f.row, f.col) });
          }
        }
      } else {
        for (const across of acrossValues) {
          setInput(rowCell, across.value);
          this.calc.recalculate({ force: true });
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

  // ---- sparklines ----------------------------------------------------------

  /** The sparkline groups on a sheet, as the engine reads them (or wrote them). */
  sparklineGroups(sheetName = this.activeSheet) {
    return this.workbook.sparklineGroups(sheetName);
  }

  /**
   * Insert → Sparklines: a line or column spark from the numbers in `data`,
   * one per cell of `at` — see the engine for the shapes this accepts. One
   * undo step, the sheet part travelling with it the way a validation rule's
   * does.
   */
  addSparklines({ type, data, at }) {
    if (this.protection().sheet) {
      throw protectionError('This sheet is protected — unprotect it before adding a sparkline.');
    }
    const sheetPartName = this.workbook.partNameFor(this.activeSheet);
    this._edit('sparkline', null, [], () => {
      this.workbook.addSparklines(this.activeSheet, { type, data, at });
      this._structuralDirty = true;
      return this;
    }, { parts: [sheetPartName] });
    return this;
  }

  /** Remove Sparkline: every sparkline whose cell is in `at` (a cell or a range). Returns how many went. */
  removeSparklines(at) {
    if (this.protection().sheet) {
      throw protectionError('This sheet is protected — unprotect it before removing a sparkline.');
    }
    const sheetPartName = this.workbook.partNameFor(this.activeSheet);
    let removed = 0;
    this._edit('remove sparkline', null, [], () => {
      removed = this.workbook.removeSparklines(this.activeSheet, at);
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
  insertChart({ title = '', kind = 'column', scatterStyle = 'markers' } = {}) {
    if (kind === 'scatter' && !SCATTER_STYLES.includes(scatterStyle)) {
      throw new Error('"' + scatterStyle + '" is not a scatter look — ' + SCATTER_STYLES.join(', '));
    }
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

    // A scatter plots numbers against numbers: its first column is the X
    // values when they are all numbers (Excel's reading); a first column of
    // words is labels, and the points go 1, 2, 3… across as Excel's do.
    const firstIsX = kind === 'scatter' && (() => {
      for (let row = r.top + 1; row <= r.bottom; row++) {
        const v = this.calc.getValue(sheet, row, r.left);
        if (typeof v !== 'number' && v !== '' && v !== null) return false;
      }
      return true;
    })();
    const categories = {
      ref: colRef(r.left, r.top + 1, r.bottom),
      values: [],
    };
    for (let row = r.top + 1; row <= r.bottom; row++) {
      if (kind === 'scatter') {
        const v = this.calc.getValue(sheet, row, r.left);
        categories.values.push(typeof v === 'number' && Number.isFinite(v) ? v : null);
      } else {
        categories.values.push(this.displayValue(row, r.left).text);
      }
    }
    if (kind === 'scatter' && r.right - r.left > 3) {
      throw new Error('A scatter chart here plots up to three series — select a block of an X column and up to three Y columns.');
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
        chartPartXml({ kind, title: String(title ?? '').trim() || undefined, categories: kind === 'scatter' && !firstIsX ? null : categories, series, scatterStyle }),
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

  /**
   * Formulas → Create from Selection: a name for each column of the
   * selection (or the block round the cell), from its header cell, pointing
   * at the rows below it. A header that is not a valid name is made one the
   * way Excel does — spaces and punctuation to underscores, a leading digit
   * given an underscore, a cell-like name a trailing one; an empty header is
   * skipped.
   *
   * @returns {{ made: string[], skipped: string[], range: string }}
   */
  namesFromSelection() {
    const sel = this.selection.range;
    const single = sel.top === sel.bottom && sel.left === sel.right;
    const r = single ? this._currentRegion(sel.top, sel.left) : sel;
    if (r.top === r.bottom) throw new Error('Create from Selection needs a header row with at least one row of data under it');
    const sheet = this.activeSheet;
    const quoted = /[^A-Za-z0-9_]/.test(sheet) ? "'" + sheet.replace(/'/g, "''") + "'" : sheet;
    const made = [];
    const skipped = [];
    for (let c = r.left; c <= r.right; c++) {
      const head = String(this.displayValue(r.top, c).text ?? '').trim();
      let name = head.replace(/[^A-Za-z0-9_.]+/g, '_').replace(/^_+|_+$/g, '');
      if (!name) { skipped.push(colName(c)); continue; }
      if (!/^[A-Za-z_]/.test(name)) name = '_' + name;
      if (/^[A-Za-z]{1,3}\d+$/.test(name)) name = name + '_';
      const target = quoted + '!$' + colName(c) + '$' + (r.top + 2) + ':$' + colName(c) + '$' + (r.bottom + 1);
      try {
        this.defineName(name, target);
        made.push(name);
      } catch {
        skipped.push(name);
      }
    }
    return { made, skipped, range: ref(r.top, r.left) + ':' + ref(r.bottom, r.right) };
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
  createPivot({ name, source, target, rowFields = [], colFields = [], dataFields = [], chart = null, fileName = null }) {
    // The dialog says the source as text — `Sheet1!$A$1:$E$40`, or `A1:E40`
    // on this sheet — and the values as `{ name, fn }`; the engine's own
    // callers pass the area and `{ field, subtotal }`. Both are understood.
    const parsed = typeof source === 'string' ? parsePivotArea(source.trim(), this.activeSheet) : source;
    if (typeof source === 'string' && source.trim() && !parsed) throw new Error('"' + source + '" is not a range — write it like Sheet1!A1:E40');
    const area = parsed
      ? { ...parsed, sheet: parsed.sheet ?? this.activeSheet }
      : {
        ...this._currentRegion(this.selection.active.row, this.selection.active.col),
        sheet: this.activeSheet,
      };
    const FNS = { SUM: 'sum', COUNT: 'count', AVERAGE: 'average', MAX: 'max', MIN: 'min', PRODUCT: 'product' };
    const values = dataFields.map((d) => (d.field !== undefined
      ? d
      : { field: d.name, subtotal: FNS[String(d.fn ?? 'SUM').toUpperCase()] ?? String(d.fn ?? 'sum') }));
    // A pivot nobody named is PivotTable1, 2… as Excel names them; one named
    // empty on purpose is still refused.
    if (name === undefined || name === null) {
      const taken = new Set(this.pivots().map((p) => p.name.toLowerCase()));
      let n = 1;
      while (taken.has('pivottable' + n)) n += 1;
      name = 'PivotTable' + n;
    }
    const where = target ?? {
      sheet: this.activeSheet === area.sheet ? area.sheet : this.activeSheet,
      row: this.activeSheet === area.sheet ? area.bottom + 2 : this.selection.active.row,
      col: this.activeSheet === area.sheet ? area.left : this.selection.active.col,
    };
    if (!this.sheetNames().includes(where.sheet)) throw new Error('no sheet "' + where.sheet + '"');
    if (chart && !PIVOT_CHART_KINDS.includes(chart.kind ?? 'column')) {
      throw new Error('"' + chart.kind + '" is not a PivotChart kind — ' + PIVOT_CHART_KINDS.join(', '));
    }

    const parts = [...new Set([
      '[Content_Types].xml',
      this.workbook.mainPart,
      this.workbook.partNameFor(where.sheet),
      this.workbook.partNameFor(area.sheet),
      OoxmlPackage.relsPathFor(this.workbook.mainPart),
      OoxmlPackage.relsPathFor(this.workbook.partNameFor(where.sheet)),
      ...(chart ? this._drawingPartsOf(where.sheet) : []),
    ].filter((p) => this.pkg.has(p)))];

    const readCell = (sheet, row, col) => this.calc.getValue(sheet, row, col);
    // Refuse BEFORE recording history, the same rule a structural edit
    // follows: a rejected gesture must leave no undo step behind. Planning
    // is where every refusal lives, and it touches nothing.
    const plan = planPivot(this.workbook, {
      name, source: area, target: where, rowFields, colFields, dataFields: values, readCell,
    });

    let shape = null;
    // STRUCTURAL, for the reason that mechanism exists: this edit creates
    // parts and writes cells on a sheet that may not be the active one, and
    // structural undo restores the parts and rebuilds everything derived
    // from them — which is precisely "put the file back as it was".
    this._edit(chart ? 'create PivotChart' : 'create pivot', null, [], () => {
      this._flushPendingEdits();
      const created = createPivot(this.workbook, {
        name, source: area, target: where, rowFields, colFields, dataFields: values, readCell,
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
      // Columns too narrow for the labels are widened, as Excel's pivot
      // widens its own ("Grand Total" is not cut to "Grand Tota").
      const geo = this.geometry.get(where.sheet);
      for (let c = grid.area.left; c <= grid.area.right; c++) {
        const longest = Math.max(0, ...grid.cells.filter((x) => x.col === c).map((x) => String(x.value ?? '').length));
        const need = longest * 7 + 14;
        if (geo && need > geo.colWidth(c)) this.workbook.setColWidthChars(where.sheet, c, pixelsToCharWidth(Math.min(need, 320)));
      }
      if (chart) this._addPivotChart(created, grid, { ...chart, fileName });
      this.dirtyCells.clear();
      this.styledCells.clear();
      this._structuralDirty = true;
      this._rebuildDerivedState();
      shape = { name, rows: grid.height, cols: grid.width, ref: areaRef(grid.area), sheet: where.sheet };
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
   * Every PivotChart drawn from the pivot follows it in the same step.
   * With no name, every pivot on this sheet is refreshed (Refresh All).
   */
  refreshPivot(name) {
    if (name === undefined || name === null || name === '') return this.refreshAllPivots();
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

    const charts = this._pivotChartParts(pivot);
    const parts = [...new Set([
      this.workbook.partNameFor(this.activeSheet),
      pivot.part,
      ...(pivot.cachePart ? [pivot.cachePart] : []),
      ...charts,
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
      this._updatePivotCharts(pivot, grid);
      this._pivots = null;
      this._structuralDirty = true;
      return this;
    }, { parts });
    return { rows: grid.height, cols: grid.width };
  }

  /** Data → Refresh All: every pivot on this sheet that can be recomputed, and its charts. */
  refreshAllPivots() {
    const names = this.pivots().filter((p) => p.sheet === this.activeSheet && !p.unsupported).map((p) => p.name);
    for (const n of names) this.refreshPivot(n);
    return { refreshed: names.length };
  }

  /** The pivot whose rectangle holds a cell, or null. */
  pivotAt(row, col, sheet = this.activeSheet) {
    return this.pivots().find((p) => p.sheet === sheet && !p.unsupported
      && row >= p.location.top && row <= p.location.bottom && col >= p.location.left && col <= p.location.right) ?? null;
  }

  // ---- PivotCharts ----------------------------------------------------------

  /** The sheet's drawing part and its rels, where they exist: what a drawing edit snapshots. */
  _drawingPartsOf(sheet) {
    const sheetPartName = this.workbook.partNameFor(sheet);
    const out = [sheetPartName, '[Content_Types].xml', OoxmlPackage.relsPathFor(sheetPartName)];
    const rel = this.pkg.rels(sheetPartName).find((r) => String(r.Type).endsWith('/drawing'));
    if (rel) {
      const part = OoxmlPackage.resolveTarget(sheetPartName, rel.Target);
      out.push(part, OoxmlPackage.relsPathFor(part));
    }
    return out.filter((p) => this.pkg.has(p));
  }

  /** Every chart part drawn from a pivot (by its `c:pivotSource`), with the sheets that show them. */
  _pivotChartParts(pivot) {
    const out = [];
    for (const name of this.pkg.partNames()) {
      if (!/^xl\/charts\/chart\d+\.xml$/.test(name)) continue;
      const info = readPivotChart(this.pkg.text(name));
      if (info && info.name === pivot.name && info.sheet === pivot.sheet) out.push(name);
    }
    return out;
  }

  /** A chart bound to a pivot, written into the pivot's sheet beside it. */
  _addPivotChart(pivot, grid, { kind = 'column', title = '', fileName = null } = {}) {
    const sheet = pivot.sheet;
    const drawingPart = this.workbook.ensureSheetDrawing(sheet);
    const n = this.pkg.nextPartNumber('xl/charts/', 'chart');
    const chartPart = 'xl/charts/chart' + n + '.xml';
    this.pkg.addPart(chartPart, pivotChartXml({
      sourceName: pivotSourceName(fileName || this.fileName, pivot),
      kind, title, sheet, chart: grid.chart,
    }), 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
    const relId = this.pkg.addRelationshipTo(drawingPart,
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
      '../charts/chart' + n + '.xml');
    const area = grid.area;
    let chartId = 0;
    this.workbook.appendDrawingAnchor(drawingPart, (id) => {
      chartId = id;
      return drawingAnchorXml({
        kind: 'chart', id, name: 'Chart ' + id,
        from: { row: area.top, col: area.right + 2 },
        to: { row: area.top + 15, col: area.right + 9 },
      }, () => relId);
    });
    this.drawings.set(sheet, this._readDrawings(this.workbook.partNameFor(sheet)));
    return { chartPart, id: chartId };
  }

  /** Every chart drawn from this pivot rewritten from its grid: the same kind and title, today's figures. */
  _updatePivotCharts(pivot, grid) {
    let changed = false;
    for (const part of this._pivotChartParts(pivot)) {
      const info = readPivotChart(this.pkg.text(part));
      this.pkg.write_(part, pivotChartXml({
        sourceName: info.source, kind: info.kind, title: info.title, sheet: pivot.sheet, chart: grid.chart,
      }));
      changed = true;
    }
    if (changed) for (const { name, part } of this.workbook.sheets()) this.drawings.set(name, this._readDrawings(part));
    return changed;
  }

  /**
   * Insert → PivotChart on a pivot: a chart bound to the pivot the cursor is
   * in, drawn from its current values beside it. (On data, the window asks
   * for a pivot first — `createPivot` with a `chart` makes both.)
   */
  insertPivotChart({ kind = 'column', title = '', fileName = null } = {}) {
    if (!PIVOT_CHART_KINDS.includes(kind)) throw new Error('"' + kind + '" is not a PivotChart kind — ' + PIVOT_CHART_KINDS.join(', '));
    if (this.protection().sheet) throw protectionError('This sheet is protected — unprotect it before inserting objects.');
    const active = this.selection.active;
    const pivot = this.pivotAt(active.row, active.col);
    if (!pivot) throw new Error('Put the cursor in a pivot table first — or make one with Insert → PivotChart on the data.');
    const grid = computePivot(pivot, (sheet, row, col) => this.calc.getValue(sheet, row, col));
    const { parts } = this._drawingEditParts();
    let made = null;
    this._edit('insert PivotChart', null, [], () => {
      made = this._addPivotChart(pivot, grid, { kind, title, fileName });
      this._structuralDirty = true;
    }, { parts, tracksNewParts: true });
    return made;
  }

  // ---- slicers --------------------------------------------------------------

  /** The workbook's slicers, read once per edit. */
  _slicerDefs() {
    if (!this._slicerCache || this._slicerCache.stamp !== this._editStamp) {
      let list = [];
      try { list = readSlicers(this.workbook); } catch { list = []; }
      this._slicerCache = { stamp: this._editStamp, list, states: new Map() };
    }
    return this._slicerCache.list;
  }

  /** The table whose rectangle holds a cell on this sheet, with its column ids. */
  _tableAt(row, col) {
    for (const t of this.workbook.tables()) {
      if (t.sheet !== this.activeSheet) continue;
      const box = this._tableByName(t.name);
      if (row >= box.top && row <= box.bottom && col >= box.left && col <= box.right) return box;
    }
    return null;
  }

  /** A table's column ids in order (`tableColumn id`), which is what a slicer cache names. */
  _tableColumnIds(tablePart) {
    return [...this.pkg.text(tablePart).matchAll(/<tableColumn\b([^>]*?)\/?>/g)]
      .map((m) => Number(/\bid="(\d+)"/.exec(m[1])?.[1] ?? 0));
  }

  /**
   * What Insert → Slicer offers at the cursor: the table or pivot it is in
   * and the fields a slicer can be made for, those that already have one
   * marked — or the reason there is nothing.
   */
  slicerSources() {
    const { row, col } = this.selection.active;
    const pivot = this.pivotAt(row, col);
    const defs = this._slicerDefs();
    if (pivot) {
      const has = new Set(defs.filter((s) => s.cache?.kind === 'pivot' && s.cache.pivots.some((p) => p.name === pivot.name)).map((s) => s.cache.sourceName));
      return { kind: 'pivot', name: pivot.name, fields: pivot.cacheFields.map((f) => ({ name: f.name, has: has.has(f.name) })) };
    }
    const table = this._tableAt(row, col);
    if (table) {
      const ids = this._tableColumnIds(table.part);
      const has = new Set(defs.filter((s) => s.cache?.kind === 'table' && s.cache.table.id === table.id).map((s) => s.cache.table.column));
      return { kind: 'table', name: table.name, fields: table.columns.map((c, i) => ({ name: c, has: has.has(ids[i]) })) };
    }
    return { kind: null, name: null, fields: [], reason: 'Put the cursor in a table or a pivot table — a slicer filters one of them.' };
  }

  /**
   * Insert → Slicer: one panel per field, each a button per item, cascaded
   * beside the table or pivot as Excel places them. One undo step for all.
   */
  insertSlicers({ fields = [] } = {}) {
    if (this.protection().sheet) throw protectionError('This sheet is protected — unprotect it before inserting objects.');
    const src = this.slicerSources();
    if (!src.kind) throw new Error(src.reason);
    const wanted = [...new Set(fields.map(String))].filter((f) => src.fields.some((x) => x.name.toLowerCase() === f.toLowerCase()));
    if (!wanted.length) throw new Error('Tick at least one field to make a slicer for.');
    const { row, col } = this.selection.active;
    const pivot = src.kind === 'pivot' ? this.pivotAt(row, col) : null;
    const table = src.kind === 'table' ? this._tableAt(row, col) : null;
    const box = pivot ? pivot.location : table;
    const sheet = this.activeSheet;
    const geo = this.geo;
    const extra = pivot ? [pivot.part, pivot.cachePart, this._recordsPartOf(pivot)] : [];
    const parts = [...new Set([
      ...this._drawingPartsOf(sheet), this.workbook.mainPart, OoxmlPackage.relsPathFor(this.workbook.mainPart),
      ...extra,
    ].filter((p) => p && this.pkg.has(p)))];
    const made = [];
    this._edit('insert slicer', null, [], () => {
      const drawingPart = this.workbook.ensureSheetDrawing(sheet);
      const startY = geo.rowOffset(box.top);
      // Beside the table or pivot, and clear of what is drawn there already
      // (a PivotChart sits where a slicer would otherwise go).
      let startX = geo.colOffset(box.right + 2);
      const band = { top: startY, bottom: startY + 252 };
      for (const d of this.drawings.get(sheet) ?? []) {
        const b = this._drawingBox(d, geo);
        if (b.y < band.bottom && b.y + b.height > band.top && b.x + b.width > startX - 8) startX = Math.max(startX, b.x + b.width + 16);
      }
      wanted.forEach((field, i) => {
        const def = { sheet, kind: src.kind, field: src.fields.find((x) => x.name.toLowerCase() === field.toLowerCase()).name };
        if (pivot) {
          const fld = pivot.cacheFields.findIndex((f) => f.name === def.field);
          const readCell = (s, r, c) => this.calc.getValue(s, r, c);
          const live = sourceValues(pivot, readCell, fld);
          const items = ensureSharedItems(this.workbook, pivot, fld, live);
          const hidden = pivot.hidden.get(fld) ?? new Set();
          def.pivot = pivot;
          def.pivotTabId = this.workbook.sheetIdOf(pivot.sheet);
          def.items = items.map((v, x) => ({ x, selected: !hidden.has(x), noData: !live.some((l) => pivotKey(l) === pivotKey(v)) }));
        } else {
          const idx = table.columns.findIndex((c) => c === def.field);
          def.table = { id: table.id, column: this._tableColumnIds(table.part)[idx] || idx + 1 };
        }
        const added = addSlicer(this.workbook, def);
        // Excel's panel, 1.92 by 2.64 inches; the next one beside it rather
        // than stacked over it, so every caption shows.
        const x = startX + i * (184 + 12);
        const y = startY;
        const w = 184;
        const h = 252;
        const from = this._markerAt(x, y);
        const to = this._markerAt(x + w, y + h);
        this.workbook.appendDrawingAnchor(drawingPart, (id) => slicerAnchorXml({
          id, name: added.name, kind: src.kind, from, to,
          offset: { x: Math.round(x * 9525), y: Math.round(y * 9525) }, size: { cx: w * 9525, cy: h * 9525 },
        }));
        made.push(added.name);
      });
      this.drawings.set(sheet, this._readDrawings(this.workbook.partNameFor(sheet)));
      this._pivots = null;
      this._structuralDirty = true;
    }, { parts, tracksNewParts: true });
    return made;
  }

  /** The records part behind a pivot's cache, where there is one. */
  _recordsPartOf(pivot) {
    if (!pivot?.cachePart) return null;
    const rel = this.pkg.rels(pivot.cachePart).find((r) => String(r.Type).endsWith('/pivotCacheRecords'));
    return rel ? OoxmlPackage.resolveTarget(pivot.cachePart, rel.Target) : null;
  }

  /** A pixel point on the sheet as an anchor marker: the cell it is in, and how far in. */
  _markerAt(x, y) {
    const geo = this.geo;
    const col = Math.max(0, geo.colAt(Math.max(0, x)));
    const row = Math.max(0, geo.rowAt(Math.max(0, y)));
    return {
      col, colOff: Math.max(0, Math.round((x - geo.colOffset(col)) * 9525)),
      row, rowOff: Math.max(0, Math.round((y - geo.rowOffset(row)) * 9525)),
    };
  }

  /**
   * A slicer as its panel shows it: the caption, a button per item — selected
   * or not, with data under the other filters or not — whether it filters
   * anything, and what it slices.
   */
  slicerState(name) {
    this._slicerDefs();
    const memo = this._slicerCache.states;
    if (memo.has(name)) return memo.get(name);
    const def = this._slicerDefs().find((s) => s.sheet === this.activeSheet && s.name === name);
    let out;
    try {
      out = def ? this._computeSlicer(def) : null;
    } catch (e) {
      out = { name, caption: def?.caption ?? name, items: [], broken: e.message };
    }
    memo.set(name, out);
    return out;
  }

  _computeSlicer(def) {
    const base = { name: def.name, caption: def.caption, columns: def.columns, rowHeightPx: Math.max(18, Math.round(def.rowHeight / 9525)), showCaption: def.showCaption };
    const cache = def.cache;
    if (!cache) return { ...base, items: [], broken: 'its slicer cache is missing' };
    const order = (a, b) => (a.hasData === b.hasData ? 0 : a.hasData ? -1 : 1);
    if (cache.kind === 'table') {
      const t = this.workbook.tables().find((x) => x.id === cache.table.id);
      if (!t) return { ...base, items: [], broken: 'its table is gone' };
      const box = this._tableByName(t.name);
      const ids = this._tableColumnIds(t.part);
      const idx = Math.max(0, ids.indexOf(cache.table.column));
      const filters = this.workbook.tableFilters(t.part);
      const chosen = filters.get(idx);
      const top = box.top + box.headerRowCount;
      const bottom = box.bottom - box.totalsRowCount;
      const all = new Map();
      const live = new Set();
      for (let r = top; r <= bottom; r++) {
        const text = this.displayValue(r, box.left + idx).text;
        const label = text === '' ? '(blank)' : text;
        if (!all.has(label)) all.set(label, this.calc.getValue(t.sheet, r, box.left + idx));
        let passes = true;
        for (const [cid, vals] of filters) {
          if (cid === idx || !vals || !vals.length) continue;
          if (!vals.includes(this.displayValue(r, box.left + cid).text)) { passes = false; break; }
        }
        if (passes) live.add(label);
      }
      const items = [...all.entries()]
        .sort((a, b) => (a[0] === '(blank)') - (b[0] === '(blank)') || slicerCompare(a[1], b[1]) || a[0].localeCompare(b[0]))
        .map(([label]) => ({ label, selected: !chosen || !chosen.length ? true : chosen.includes(label === '(blank)' ? '' : label), hasData: live.has(label) }));
      items.sort(order);
      return { ...base, kind: 'table', source: t.name, field: t.columns[idx], items, filtered: Boolean(chosen && chosen.length) };
    }
    const pivotName = cache.pivots[0]?.name;
    const pivotSheet = this.workbook.sheetWithId(cache.pivots[0]?.tabId);
    const pivot = this.pivots().find((p) => p.name === pivotName && (!pivotSheet || p.sheet === pivotSheet));
    if (!pivot || pivot.unsupported) return { ...base, items: [], broken: pivot ? pivot.unsupported : 'its pivot table is gone' };
    const fld = pivot.cacheFields.findIndex((f) => f.name === cache.sourceName);
    if (fld < 0) return { ...base, items: [], broken: 'its field is gone from the pivot' };
    const readCell = (s, r, c) => this.calc.getValue(s, r, c);
    const liveValues = sourceValues(pivot, readCell, fld);
    const hiddenX = pivot.hidden.get(fld) ?? new Set();
    const cacheItems = pivot.cacheFields[fld].items ?? [];
    const hiddenKeys = new Set([...hiddenX].map((x) => pivotKey(cacheItems[x])));
    // With data: the values rows still show once every OTHER field's hidden items are out.
    const src = pivot.source;
    const others = [...pivot.hidden.entries()].filter(([f]) => f !== fld)
      .map(([f, set]) => [f, new Set([...set].map((x) => pivotKey(pivot.cacheFields[f]?.items?.[x])))]);
    const withData = new Set();
    for (let r = src.top + 1; r <= src.bottom; r++) {
      let ok = true;
      for (const [f, keys] of others) {
        if (keys.has(pivotKey(readCell(src.sheet, r, src.left + f) ?? ''))) { ok = false; break; }
      }
      if (ok) withData.add(pivotKey(readCell(src.sheet, r, src.left + fld) ?? ''));
    }
    const items = [...liveValues]
      .sort((a, b) => ((a === '') - (b === '')) || slicerCompare(a, b))
      .map((v) => ({ label: itemLabel(v), selected: !hiddenKeys.has(pivotKey(v)), hasData: withData.has(pivotKey(v)) }));
    items.sort(order);
    return { ...base, kind: 'pivot', source: pivot.name, field: cache.sourceName, items, filtered: items.some((i) => !i.selected) };
  }

  /**
   * Click a slicer's buttons: the items to show, by their labels — or null
   * (Clear Filter) for every one. A table's autoFilter takes the choice and
   * the rows it leaves out hide; a pivot hides the items in its field, its
   * slicer cache keeps the states, and it is refreshed with its charts.
   */
  setSlicerSelection({ name, values = null }) {
    const def = this._slicerDefs().find((s) => s.sheet === this.activeSheet && s.name === name);
    if (!def) throw new Error('no slicer "' + name + '" on this sheet');
    const state = this.slicerState(name);
    if (state.broken) throw new Error('This slicer cannot filter: ' + state.broken + '.');
    const labels = values === null || values === undefined ? null : [...new Set(values.map(String))];
    if (labels && !labels.length) throw new Error('A slicer keeps at least one item selected — Clear Filter shows them all.');
    const everything = !labels || state.items.every((i) => labels.includes(i.label));
    if (state.kind === 'table') {
      const t = this.workbook.tables().find((x) => x.id === def.cache.table.id);
      return this.applyFilter(t.name, state.field, everything ? null : labels.map((l) => (l === '(blank)' ? '' : l)));
    }
    const pivot = this.pivots().find((p) => p.name === state.source);
    const fld = pivot.cacheFields.findIndex((f) => f.name === state.field);
    const readCell = (s, r, c) => this.calc.getValue(s, r, c);
    const liveValues = sourceValues(pivot, readCell, fld);
    const keep = everything ? null : liveValues.filter((v) => labels.includes(itemLabel(v)));
    // Every slicer cache on this field of this pivot keeps the same states.
    const caches = [...new Set(this._slicerDefs()
      .filter((s) => s.cache?.kind === 'pivot' && s.cache.sourceName === state.field && s.cache.pivots.some((p) => p.name === pivot.name))
      .map((s) => s.cache.part))];
    const parts = [...new Set([
      this.workbook.partNameFor(pivot.sheet), pivot.part, pivot.cachePart, this._recordsPartOf(pivot),
      ...caches, ...this._pivotChartParts(pivot).flatMap((p) => [p]),

    ].filter((p) => p && this.pkg.has(p)))];
    this._edit('slicer ' + name, null, [], () => {
      this._flushPendingEdits();
      const { items, hidden } = setPivotItemsShown(this.workbook, pivot, fld, keep, readCell);
      const liveKeys = new Set(liveValues.map(pivotKey));
      for (const part of caches) {
        writeSlicerCacheItems(this.workbook, part, items.map((v, x) => ({ x, selected: !hidden.has(x), noData: !liveKeys.has(pivotKey(v)) })));
      }
      this._writePivotGrid(pivot);
      this.dirtyCells.clear();
      this.styledCells.clear();
      this._structuralDirty = true;
      this._rebuildDerivedState();
    }, { parts, structural: true });
    return this;
  }

  /** A pivot's rectangle rewritten from its definition, in the parts (a structural edit's way), and its charts. */
  _writePivotGrid(pivot) {
    const readCell = (s, r, c) => this.calc.getValue(s, r, c);
    const grid = computePivot(pivot, readCell);
    const old = pivot.location;
    const wanted = new Map(grid.cells.map((c) => [c.row + ':' + c.col, c.value]));
    const clear = (r, c) => { if (!wanted.has(r + ':' + c)) this.workbook.setCell(pivot.sheet, ref(r, c), null); };
    for (let r = old.top; r <= old.bottom; r++) for (let c = old.left; c <= old.right; c++) clear(r, c);
    for (const c of grid.cells) this.workbook.setCell(pivot.sheet, ref(c.row, c.col), c.value === '' ? null : c.value);
    updatePivotLocation(this.workbook, pivot, grid.area, {
      headerRows: grid.firstHeaderRow, dataRow: grid.firstDataRow, dataCol: grid.firstDataCol,
    });
    this._updatePivotCharts(pivot, grid);
    this._pivots = null;
    return grid;
  }

  /** Slicer → Caption and Columns, from the slicer's own menu. */
  setSlicerProps({ name, caption, columns }) {
    const def = this._slicerDefs().find((s) => s.sheet === this.activeSheet && s.name === name);
    if (!def) throw new Error('no slicer "' + name + '" on this sheet');
    const cols = columns === undefined ? undefined : Math.max(1, Math.min(20, Math.round(Number(columns) || 1)));
    this._edit('slicer settings', null, [], () => {
      writeSlicerProps(this.workbook, this.activeSheet, name, { caption: caption === undefined ? undefined : String(caption), columns: cols });
      this._structuralDirty = true;
    }, { parts: [def.part] });
    return this;
  }

  // ---- drawings: move, resize, delete -----------------------------------------

  /** The active sheet's drawing part, or null. */
  _drawingPart(sheet = this.activeSheet) {
    const sheetPart = this.workbook.partNameFor(sheet);
    const rel = this.pkg.rels(sheetPart).find((r) => String(r.Type).endsWith('/drawing'));
    if (!rel) return null;
    const part = OoxmlPackage.resolveTarget(sheetPart, rel.Target);
    return this.pkg.has(part) ? part : null;
  }

  /** A drawing of the active sheet by its frame id, with its place in the part. */
  _drawingById(id) {
    const list = this.drawings.get(this.activeSheet) ?? [];
    const d = list.find((x) => x.id === String(id));
    if (!d) throw new Error('no drawing "' + id + '" on this sheet');
    return d;
  }

  /** Rewrite drawing anchors of the active sheet in one undo step: `edit(spans, xml)` returns the new part. */
  _editAnchors(label, edit, { extraParts = [] } = {}) {
    const part = this._drawingPart();
    if (!part) throw new Error('this sheet has no drawings');
    const { sheetPartName, parts } = this._drawingEditParts();
    let result = null;
    this._edit(label, null, [], () => {
      const xml = this.pkg.text(part);
      const spans = anchorSpans(xml);
      const out = edit(spans, xml);
      result = out.result ?? null;
      if (out.xml !== xml) this.pkg.write_(part, out.xml);
      this.drawings.set(this.activeSheet, this._readDrawings(sheetPartName));
      this._structuralDirty = true;
    }, { parts: [...new Set([...parts, ...extraParts.filter((p) => this.pkg.has(p))])], tracksNewParts: true });
    return result;
  }

  /**
   * A drawing moved or resized by hand: its box in sheet pixels. The anchor
   * follows the cells — a two-cell anchor's `from` and `to` markers, a
   * one-cell anchor's `from` and extent — and the shape's own `a:xfrm`
   * (where it states one) agrees, so Excel puts it back exactly here.
   */
  setDrawingBox({ id, x, y, width, height }) {
    const d = this._drawingById(id);
    if (this.protection().sheet) throw protectionError('This sheet is protected — unprotect it before moving objects.');
    const box = {
      x: Math.max(0, Math.round(Number(x) || 0)),
      y: Math.max(0, Math.round(Number(y) || 0)),
      width: Math.max(8, Math.round(Number(width) || 8)),
      height: Math.max(8, Math.round(Number(height) || 8)),
    };
    return this._editAnchors('move ' + (d.name || d.kind), (spans, xml) => {
      const span = spans[d.index];
      const next = anchorWithBox(span.xml, box, (px, py) => this._markerAt(px, py));
      return { xml: xml.slice(0, span.start) + next + xml.slice(span.end), result: box };
    });
  }

  /**
   * Delete drawings — the Delete key on a selected picture, shape, chart or
   * slicer. A slicer takes its `<slicer>`, and its cache when nothing else
   * uses it; a chart its part; the relationships they leave unused go too.
   */
  deleteDrawings({ ids = [] } = {}) {
    if (this.protection().sheet) throw protectionError('This sheet is protected — unprotect it before deleting objects.');
    const list = (this.drawings.get(this.activeSheet) ?? []).filter((d) => ids.map(String).includes(d.id));
    if (!list.length) return 0;
    const drawingPart = this._drawingPart();
    const slicers = list.filter((d) => d.kind === 'slicer');
    const extra = [this.workbook.mainPart, OoxmlPackage.relsPathFor(this.workbook.mainPart), OoxmlPackage.relsPathFor(drawingPart)];
    for (const s of slicers) {
      const def = this._slicerDefs().find((x) => x.sheet === this.activeSheet && x.name === s.slicerName);
      if (def) extra.push(def.part, ...(def.cache ? [def.cache.part] : []));
    }
    return this._editAnchors('delete ' + (list.length === 1 ? (list[0].name || list[0].kind) : list.length + ' objects'), (spans, xml) => {
      const gone = new Set(list.map((d) => d.index));
      let out = '';
      let at = 0;
      spans.forEach((s, i) => {
        out += xml.slice(at, s.start);
        if (!gone.has(i)) out += s.xml;
        at = s.end;
      });
      out += xml.slice(at);
      // Relationships no anchor names any more — a chart's, a picture's.
      const relsPath = OoxmlPackage.relsPathFor(drawingPart);
      if (this.pkg.has(relsPath)) {
        let rels = this.pkg.text(relsPath);
        for (const r of this.pkg.rels(drawingPart)) {
          if (new RegExp('"' + r.Id + '"').test(out)) continue;
          if (!/\/(chart|image)$/.test(String(r.Type))) continue;
          rels = rels.replace(new RegExp('<Relationship\\b[^>]*Id="' + r.Id + '"[^>]*/>'), '');
        }
        this.pkg.write_(relsPath, rels);
      }
      for (const s of slicers) {
        try { removeSlicer(this.workbook, this.activeSheet, s.slicerName); } catch { /* a slicer part already gone */ }
      }
      return { xml: out, result: list.length };
    }, { extraParts: extra });
  }
  // ---- Arrange: order, align, rotate, group, the Selection Pane ---------------

  /** Drawings of the active sheet by frame id, in drawing order, refusing ids that are not there. */
  _drawingsById(ids) {
    const want = [...new Set((ids ?? []).map(String))];
    const list = (this.drawings.get(this.activeSheet) ?? []).filter((d) => want.includes(d.id));
    if (!list.length) throw new Error('Select a picture, shape, chart or slicer first.');
    return list.sort((a, b) => a.index - b.index);
  }

  /**
   * Page Layout → Bring Forward / Send Backward / to Front / to Back: the
   * drawing part draws its anchors in order, so the order IS the layering.
   * Several picked move together, keeping their own order among themselves.
   */
  reorderDrawings({ ids, to = 'forward' }) {
    const picked = this._drawingsById(ids);
    if (!['forward', 'backward', 'front', 'back'].includes(to)) throw new Error('"' + to + '" is not a way to reorder');
    return this._editAnchors(to === 'front' ? 'bring to front' : to === 'back' ? 'send to back' : to === 'forward' ? 'bring forward' : 'send backward', (spans, xml) => {
      const order = spans.map((s, i) => i);
      const mine = new Set(picked.map((d) => d.index));
      let next;
      if (to === 'front') next = [...order.filter((i) => !mine.has(i)), ...order.filter((i) => mine.has(i))];
      else if (to === 'back') next = [...order.filter((i) => mine.has(i)), ...order.filter((i) => !mine.has(i))];
      else {
        next = [...order];
        const step = to === 'forward' ? 1 : -1;
        const seq = to === 'forward' ? [...next].reverse() : [...next];
        for (const i of seq) {
          if (!mine.has(i)) continue;
          const at = next.indexOf(i);
          const j = at + step;
          if (j < 0 || j >= next.length || mine.has(next[j])) continue;
          [next[at], next[j]] = [next[j], next[at]];
        }
      }
      if (next.every((v, k) => v === k)) return { xml, result: false };
      const head = xml.slice(0, spans[0].start);
      const tail = xml.slice(spans[spans.length - 1].end);
      // Whatever sat between anchors (nothing, as a rule) stays with the head.
      return { xml: head + next.map((i) => spans[i].xml).join('') + tail, result: true };
    });
  }

  /** Several drawings at new boxes, one undo step: a move of several picked at once. */
  setDrawingBoxes({ boxes = [] } = {}) {
    if (this.protection().sheet) throw protectionError('This sheet is protected — unprotect it before moving objects.');
    const list = boxes.map((b) => ({ d: this._drawingById(b.id), box: {
      x: Math.max(0, Math.round(Number(b.x) || 0)), y: Math.max(0, Math.round(Number(b.y) || 0)),
      width: Math.max(8, Math.round(Number(b.width) || 8)), height: Math.max(8, Math.round(Number(b.height) || 8)),
    } }));
    if (!list.length) return 0;
    return this._editAnchors(list.length === 1 ? 'move ' + (list[0].d.name || list[0].d.kind) : 'move ' + list.length + ' objects', (spans, xml) => {
      const byIndex = new Map(list.map((l) => [l.d.index, l.box]));
      return { xml: rebuild(spans, xml, (s, i) => (byIndex.has(i) ? anchorWithBox(s.xml, byIndex.get(i), (px, py) => this._markerAt(px, py)) : s.xml)), result: list.length };
    });
  }

  /**
   * Page Layout → Align: the picked drawings lined up on the left, centre,
   * right, top, middle or bottom of their combined bounds; Distribute puts
   * equal gaps between three or more.
   */
  alignDrawings({ ids, edge }) {
    const picked = this._drawingsById(ids);
    if (picked.length < 2) throw new Error('Select two or more objects to align.');
    const boxes = picked.map((d) => ({ d, b: this._drawingBox(d) }));
    const left = Math.min(...boxes.map(({ b }) => b.x));
    const top = Math.min(...boxes.map(({ b }) => b.y));
    const right = Math.max(...boxes.map(({ b }) => b.x + b.width));
    const bottom = Math.max(...boxes.map(({ b }) => b.y + b.height));
    const out = boxes.map(({ d, b }) => {
      const n = { id: d.id, ...b };
      if (edge === 'left') n.x = left;
      else if (edge === 'center') n.x = Math.round((left + right) / 2 - b.width / 2);
      else if (edge === 'right') n.x = right - b.width;
      else if (edge === 'top') n.y = top;
      else if (edge === 'middle') n.y = Math.round((top + bottom) / 2 - b.height / 2);
      else if (edge === 'bottom') n.y = bottom - b.height;
      else throw new Error('"' + edge + '" is not an edge to align to');
      return n;
    });
    return this.setDrawingBoxes({ boxes: out });
  }

  distributeDrawings({ ids, axis = 'horizontal' }) {
    const picked = this._drawingsById(ids);
    if (picked.length < 3) throw new Error('Select three or more objects to distribute.');
    const h = axis === 'horizontal';
    const items = picked.map((d) => ({ d, b: this._drawingBox(d) })).sort((a, b) => (h ? a.b.x - b.b.x : a.b.y - b.b.y));
    const start = h ? items[0].b.x : items[0].b.y;
    const last = items[items.length - 1].b;
    const end = h ? last.x + last.width : last.y + last.height;
    const total = items.reduce((n, { b }) => n + (h ? b.width : b.height), 0);
    const gap = (end - start - total) / (items.length - 1);
    let at = start;
    const out = items.map(({ d, b }) => {
      const n = { id: d.id, ...b, ...(h ? { x: Math.round(at) } : { y: Math.round(at) }) };
      at += (h ? b.width : b.height) + gap;
      return n;
    });
    return this.setDrawingBoxes({ boxes: out });
  }

  /**
   * Page Layout → Rotate: Right 90°, Left 90°, Flip Vertical, Flip
   * Horizontal — or a turn to an angle (the rotation handle). Written on the
   * drawing's `a:xfrm` (`rot`, `flipH`, `flipV`) as Excel writes them; a
   * chart or a slicer does not turn, in Excel or here.
   */
  rotateDrawings({ ids, by = 0, to = null, flip = null }) {
    const picked = this._drawingsById(ids).filter((d) => ['shape', 'image', 'group'].includes(d.kind));
    if (!picked.length) throw new Error('Charts and slicers do not turn — select a picture, a shape or a group.');
    const label = flip ? 'flip ' + flip : to !== null ? 'rotate' : by > 0 ? 'rotate right' : 'rotate left';
    return this._editAnchors(label, (spans, xml) => {
      const mine = new Map(picked.map((d) => [d.index, d]));
      return {
        xml: rebuild(spans, xml, (s, i) => {
          if (!mine.has(i)) return s.xml;
          const box = this._drawingBox(mine.get(i));
          return withTransform(s.xml, box, (t) => {
            if (flip === 'horizontal') t.flipH = !t.flipH;
            else if (flip === 'vertical') t.flipV = !t.flipV;
            else t.rot = ((((to !== null ? Number(to) : t.rot + Number(by)) % 360) + 360) % 360);
            return t;
          });
        }),
        result: picked.length,
      };
    });
  }

  /**
   * Page Layout → Group: the picked drawings gathered into one `xdr:grpSp`
   * in one anchor over their combined bounds, each keeping its place (its
   * xfrm in the group's space, which is the sheet's EMU), at the topmost
   * member's place in the drawing order. A slicer is not grouped.
   */
  groupDrawings({ ids }) {
    const picked = this._drawingsById(ids);
    if (picked.length < 2) throw new Error('Select two or more objects to group.');
    if (picked.some((d) => d.kind === 'slicer')) throw new Error('A slicer is not grouped with other objects here — take it out of the selection.');
    const boxes = picked.map((d) => this._drawingBox(d));
    const union = {
      x: Math.min(...boxes.map((b) => b.x)), y: Math.min(...boxes.map((b) => b.y)),
      right: Math.max(...boxes.map((b) => b.x + b.width)), bottom: Math.max(...boxes.map((b) => b.y + b.height)),
    };
    const box = { x: union.x, y: union.y, width: union.right - union.x, height: union.bottom - union.y };
    let groupId = null;
    this._editAnchors('group', (spans, xml) => {
      const ids2 = [...xml.matchAll(/<([\w]+:)?cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[2]));
      groupId = (ids2.length ? Math.max(...ids2) : 1) + 1;
      const p = spans[0].prefix;
      const emu = (v) => Math.round(v * 9525);
      const members = picked.map((d, k) => {
        const body = anchorBody(spans[d.index].xml);
        return withTransform(body.xml, boxes[k], (t) => t, { absolute: true });
      }).join('');
      const off = '<a:off x="' + emu(box.x) + '" y="' + emu(box.y) + '"/><a:ext cx="' + emu(box.width) + '" cy="' + emu(box.height) + '"/>';
      const grp = '<' + p + 'grpSp><' + p + 'nvGrpSpPr><' + p + 'cNvPr id="' + groupId + '" name="Group ' + groupId + '"/><' + p + 'cNvGrpSpPr/></' + p + 'nvGrpSpPr>'
        + '<' + p + 'grpSpPr><a:xfrm>' + off + off.replace('<a:off', '<a:chOff').replace('<a:ext', '<a:chExt') + '</a:xfrm></' + p + 'grpSpPr>'
        + members + '</' + p + 'grpSp>';
      const from = this._markerAt(box.x, box.y);
      const to = this._markerAt(box.x + box.width, box.y + box.height);
      const marker = (tag, m) => '<' + p + tag + '><' + p + 'col>' + m.col + '</' + p + 'col><' + p + 'colOff>' + m.colOff + '</' + p + 'colOff><' + p + 'row>' + m.row + '</' + p + 'row><' + p + 'rowOff>' + m.rowOff + '</' + p + 'rowOff></' + p + tag + '>';
      const anchor = '<' + p + 'twoCellAnchor>' + marker('from', from) + marker('to', to) + grp + '<' + p + 'clientData/></' + p + 'twoCellAnchor>';
      const mine = new Set(picked.map((d) => d.index));
      const topmost = picked[picked.length - 1].index;
      return { xml: rebuild(spans, xml, (s, i) => (i === topmost ? anchor : mine.has(i) ? '' : s.xml)), result: groupId };
    });
    return 'd' + groupId;
  }

  /** Page Layout → Ungroup: a group's members back on the sheet, each in its own anchor, where they were drawn. */
  ungroupDrawings({ ids }) {
    const groups = this._drawingsById(ids).filter((d) => d.kind === 'group');
    if (!groups.length) throw new Error('Select a group to ungroup.');
    return this._editAnchors('ungroup', (spans, xml) => {
      const byIndex = new Map(groups.map((g) => [g.index, g]));
      return {
        xml: rebuild(spans, xml, (s, i) => {
          if (!byIndex.has(i)) return s.xml;
          const gbox = this._drawingBox(byIndex.get(i));
          const body = anchorBody(s.xml);
          const own = xfrmOf(/<([\w]+:)?grpSpPr\b[\s\S]*?<\/([\w]+:)?grpSpPr>/.exec(body.xml)?.[0] ?? '') ?? { x: 0, y: 0, cx: 1, cy: 1 };
          const space = { x: own.chX ?? own.x, y: own.chY ?? own.y, cx: own.chCx || own.cx || 1, cy: own.chCy || own.cy || 1 };
          const inner = body.xml.replace(/^<[^>]*>/, '').replace(/<\/[^>]*>$/, '');
          const p = s.prefix;
          const marker = (tag, m) => '<' + p + tag + '><' + p + 'col>' + m.col + '</' + p + 'col><' + p + 'colOff>' + m.colOff + '</' + p + 'colOff><' + p + 'row>' + m.row + '</' + p + 'row><' + p + 'rowOff>' + m.rowOff + '</' + p + 'rowOff></' + p + tag + '>';
          return childElements(inner).filter((el) => /^(sp|pic|grpSp|graphicFrame|cxnSp)$/.test(el.name)).map((el) => {
            const b = xfrmOf(el.name === 'grpSp' ? (/<([\w]+:)?grpSpPr\b[\s\S]*?<\/([\w]+:)?grpSpPr>/.exec(el.xml)?.[0] ?? '') : el.xml) ?? { x: space.x, y: space.y, cx: space.cx, cy: space.cy };
            const px = {
              x: Math.round(gbox.x + ((b.x - space.x) / space.cx) * gbox.width),
              y: Math.round(gbox.y + ((b.y - space.y) / space.cy) * gbox.height),
              width: Math.max(1, Math.round((b.cx / space.cx) * gbox.width)),
              height: Math.max(1, Math.round((b.cy / space.cy) * gbox.height)),
            };
            const placed = withTransform(el.xml, px, (t) => t, { absolute: true });
            return '<' + p + 'twoCellAnchor>' + marker('from', this._markerAt(px.x, px.y)) + marker('to', this._markerAt(px.x + px.width, px.y + px.height))
              + placed + '<' + p + 'clientData/></' + p + 'twoCellAnchor>';
          }).join('');
        }),
        result: groups.length,
      };
    });
  }

  /** The Selection Pane's eye: a drawing hidden (cNvPr hidden="1") or shown. */
  setDrawingHidden({ id, hidden = true }) {
    const d = this._drawingById(id);
    return this._editAnchors((hidden ? 'hide ' : 'show ') + (d.name || d.kind), (spans, xml) => ({
      xml: rebuild(spans, xml, (s, i) => (i === d.index ? withNvAttr(s.xml, 'hidden', hidden ? '1' : null) : s.xml)),
      result: true,
    }));
  }

  /** The Selection Pane's Show All / Hide All. */
  setAllDrawingsHidden({ hidden = false }) {
    if (!(this.drawings.get(this.activeSheet) ?? []).length) return false;
    return this._editAnchors(hidden ? 'hide all' : 'show all', (spans, xml) => ({
      xml: rebuild(spans, xml, (s) => withNvAttr(s.xml, 'hidden', hidden ? '1' : null)),
      result: true,
    }));
  }

  /** The Selection Pane's rename: the drawing's cNvPr name (a slicer keeps its own name, which its part knows it by). */
  renameDrawing({ id, name }) {
    const d = this._drawingById(id);
    const text = String(name ?? '').trim();
    if (!text) throw new Error('A name cannot be empty.');
    if (d.kind === 'slicer') return this.setSlicerProps({ name: d.slicerName, caption: text });
    return this._editAnchors('rename', (spans, xml) => ({
      xml: rebuild(spans, xml, (s, i) => (i === d.index ? withNvAttr(s.xml, 'name', text) : s.xml)),
      result: true,
    }));
  }

  // ---- Page Layout → Themes, Colours, Fonts, Effects ---------------------------

  /** The workbook's theme as the Page Layout tab shows it; read once per edit. */
  workbookDesign() {
    if (!this._designCache || this._designCache.stamp !== this._editStamp) {
      this._designCache = { stamp: this._editStamp, design: readWorkbookDesign(this.pkg) };
    }
    return this._designCache.design;
  }

  /**
   * The accents charts and theme-styled shapes are drawn in: the theme's own
   * once the workbook wears one other than Office's — a theme picked here, or
   * a file's own — and the suite's validated palette otherwise, as ever.
   */
  _themeAccents() {
    const d = this.workbookDesign();
    if (!d.exists || /^Office/i.test(d.colorName || d.name || '')) return null;
    return [1, 2, 3, 4, 5, 6].map((n) => '#' + String(d.colors['accent' + n]).toLowerCase());
  }

  /**
   * Page Layout → Themes (`{ theme }`), Colours (`{ colors, name }`), Fonts
   * (`{ fonts, name }`) and Effects (`{ effects }`): the theme part rewritten
   * — made, with its relationship and content type, when the workbook has
   * none — and the style table's theme fonts and the Normal style's font
   * following the new faces. One undo step; the cells, tables, charts and
   * shapes that take their look from the theme follow at once.
   */
  setWorkbookTheme(spec = {}) {
    const had = this.pkg.has(THEME_PART);
    const before = this.workbookDesign();
    // Refused before anything is recorded: a bad palette leaves no step.
    const xml = designedThemeXml(this.pkg, spec);
    const parts = [THEME_PART, 'xl/styles.xml', '[Content_Types].xml', OoxmlPackage.relsPathFor(this.workbook.mainPart)]
      .filter((p) => this.pkg.has(p));
    this._edit(spec.theme ? 'theme' : spec.colors ? 'theme colours' : spec.fonts ? 'theme fonts' : 'theme effects', null, [], () => {
      if (had) this.pkg.write_(THEME_PART, xml);
      else {
        this.pkg.addPart(THEME_PART, xml, THEME_TYPE);
        this.pkg.addRelationshipTo(this.workbook.mainPart, THEME_REL, 'theme/theme1.xml');
      }
      this._designCache = null;
      const after = readWorkbookDesign(this.pkg);
      const facesMoved = !had || after.fonts.major !== before.fonts.major || after.fonts.minor !== before.fonts.minor;
      // A workbook with no style table gets one, so its default font can follow.
      if (facesMoved) this._ensureStylesPart();
      if (this.pkg.has('xl/styles.xml') && facesMoved) {
        const styles = this.pkg.text('xl/styles.xml');
        const next = stylesFollowingFonts(styles, after.fonts, before.fonts);
        if (next !== styles) this.pkg.write_('xl/styles.xml', next);
      }
      this._afterPartsRestored([THEME_PART, 'xl/styles.xml']);
      this._structuralDirty = true;
    }, { parts, tracksNewParts: true, sheetGate: false });
    this._designCache = null;
    return this.workbookDesign();
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
   * Protect the active sheet, so locked cells refuse edits — with a
   * password when one is given, hashed as Excel hashes one (SHA-512, a
   * fresh salt, 100,000 rounds), so Excel asks for the same password.
   */
  protect({ password = '' } = {}) {
    if (this.protection().sheet) return this;
    this._edit('protect sheet', null, [], () => {
      this.workbook.setSheetProtection(this.activeSheet, true, { password });
      this._structuralDirty = true;
    }, { parts: [this.workbook.partNameFor(this.activeSheet)] });
    return this;
  }

  /**
   * Thaw the sheet. A sheet protected with a password — this editor's or
   * Excel's, the modern hash or the legacy 16-bit one — asks for it, and
   * refuses a wrong one with Excel's sentence.
   */
  unprotect({ password = '' } = {}) {
    const p = this.protection();
    if (!p.sheet) return this;
    if (p.hasPassword && !this.workbook.checkSheetPassword(this.activeSheet, password)) {
      throw protectionError(password ? WRONG_PASSWORD : 'This sheet’s protection has a password — type it to unprotect the sheet.', { needsPassword: 'sheet' });
    }
    this._edit('unprotect sheet', null, [], () => {
      this.workbook.setSheetProtection(this.activeSheet, null);
      this._unlockedRanges?.clear();
      this._structuralDirty = true;
    }, { parts: [this.workbook.partNameFor(this.activeSheet)] });
    return this;
  }

  // ---- Data → Consolidate, Data → Forecast Sheet --------------------------------

  /**
   * Data → Consolidate at the active cell (see consolidate.js): `fn`, the
   * references, labels in the top row and left column, links to the source.
   */
  consolidate(spec = {}) {
    return consolidateRanges(this, spec);
  }

  /**
   * What the Consolidate dialog opens on: the sheet's last consolidation, as
   * Excel remembers it, and the selection as a reference to add.
   */
  consolidateInfo() {
    const r = this.selection.range;
    return {
      last: lastConsolidation(this),
      selection: consolidateRefText({ sheet: this.activeSheet, ...r }),
    };
  }

  /** What the Forecast Sheet dialog opens on (see forecast-sheet.js). */
  forecastInfo() {
    return forecastDefaults(this);
  }

  /** The Forecast Sheet dialog's preview: the chart as the sheet will draw it, or why not. */
  forecastPreview(spec = {}, size = {}) {
    return forecastPreview(this, spec, size);
  }

  /** Data → Forecast Sheet: a new sheet with the forecast's table and chart. */
  forecastSheet(spec = {}) {
    this._structureGate();
    return createForecastSheet(this, spec);
  }

  // ---- workbook protection ---------------------------------------------------

  /** Review → Protect Workbook: `{ structure, windows, hasPassword }`. */
  workbookProtection() {
    return this.workbook.workbookProtection();
  }

  /**
   * Lock the workbook's structure: while it is locked no sheet is added,
   * deleted, renamed, moved, hidden or shown again. The password, when one
   * is given, is hashed as Excel hashes it and written as
   * `workbookAlgorithmName` / `workbookHashValue` / `workbookSaltValue` /
   * `workbookSpinCount`. One undo step, as Protect Sheet is.
   */
  protectWorkbook({ password = '' } = {}) {
    if (this.workbookProtection().structure) return this;
    this._edit('protect workbook', null, [], () => {
      this.workbook.setWorkbookProtection({ structure: true, password: password || null });
      this._structuralDirty = true;
    }, { parts: [this.workbook.mainPart], sheetGate: false });
    return this;
  }

  /** Unlock the structure; a workbook locked with a password wants it. */
  unprotectWorkbook({ password = '' } = {}) {
    const p = this.workbookProtection();
    if (!p.structure && !p.windows) return this;
    if (p.hasPassword && !this.workbook.checkWorkbookPassword(password)) {
      throw protectionError(password ? WRONG_PASSWORD : 'The workbook’s protection has a password — type it to unprotect the workbook.', { needsPassword: 'workbook' });
    }
    this._edit('unprotect workbook', null, [], () => {
      this.workbook.setWorkbookProtection(null);
      this._structuralDirty = true;
    }, { parts: [this.workbook.mainPart], sheetGate: false });
    return this;
  }

  /** Refuse a change to the workbook's structure while it is locked, in Excel's words. */
  _structureGate() {
    if (this.workbookProtection().structure) throw protectionError(WORKBOOK_LOCKED);
  }

  /** The sheets put away by Hide, by name. */
  hiddenSheets() {
    return this.workbook.hiddenSheets();
  }

  /**
   * Hide a sheet, as its tab's Hide does. The last sheet showing cannot go
   * (Excel's rule: a workbook keeps one visible sheet); the view moves to
   * the next sheet showing.
   */
  hideSheet(name = this.activeSheet) {
    this._structureGate();
    if (!this.sheetNames().includes(name)) throw new Error('no such sheet: ' + name);
    const hidden = new Set(this.hiddenSheets());
    if (hidden.has(name)) return this;
    const showing = this.sheetNames().filter((n) => !hidden.has(n));
    if (showing.length <= 1) throw new Error('A workbook must contain at least one visible worksheet.');
    this._edit('hide sheet', null, [], () => {
      this.workbook.setSheetHidden(name, true);
      this._structuralDirty = true;
    }, { parts: [this.workbook.mainPart], sheetGate: false });
    if (this.activeSheet === name) {
      const i = showing.indexOf(name);
      this.selectSheet(showing[i + 1] ?? showing[i - 1]);
    }
    return this;
  }

  /** Show a hidden sheet again — Unhide — and go to it. */
  unhideSheet(name) {
    this._structureGate();
    if (!this.hiddenSheets().includes(name)) return this;
    this._edit('unhide sheet', null, [], () => {
      this.workbook.setSheetHidden(name, false);
      this._structuralDirty = true;
    }, { parts: [this.workbook.mainPart], sheetGate: false });
    this.selectSheet(name);
    return this;
  }

  /** Move a sheet in the tab order — its tab's Move Left and Move Right. */
  moveSheet(name, toIndex) {
    this._structureGate();
    if (!this.sheetNames().includes(name)) throw new Error('no such sheet: ' + name);
    const was = this.activeSheet;
    this._edit('move sheet', null, [], () => {
      this.workbook.moveSheet(name, toIndex);
      this._rebuildDerivedState();
      this._structuralDirty = true;
    }, { parts: [this.workbook.mainPart], structural: true, sheetGate: false });
    this.activeSheet = was;
    return this;
  }

  // ---- Allow Edit Ranges --------------------------------------------------------

  /**
   * Review → Allow Edit Ranges: the active sheet's ranges that stay
   * editable when it is protected — `{ title, ref, hasPassword }` each.
   */
  editRanges() {
    return this.workbook.protectedRanges(this.activeSheet).map((r) => ({ title: r.title, ref: r.sqref, hasPassword: r.hasPassword }));
  }

  /** The active sheet's ranges parsed for the gate, kept until the sheet's tail changes. */
  _editRangeIndex() {
    const sheet = this.activeSheet;
    const block = this.workbook._sheetPart(sheet).part.tailElement('protectedRanges') || '';
    const hit = this._rangeCache;
    if (hit && hit.sheet === sheet && hit.block === block) return hit.list;
    const list = this.workbook.protectedRanges(sheet).map((r) => ({
      title: r.title,
      hasPassword: r.hasPassword,
      areas: String(r.sqref).split(/\s+/).map((a) => parseArea(a)).filter(Boolean),
    }));
    this._rangeCache = { sheet, block, list };
    return list;
  }

  /**
   * May the gate let a locked cell change? True when it is in a range with
   * no password, or one unlocked this session; the title of the range whose
   * password it wants; null when it is in no range.
   */
  _rangeVerdict(ranges, row, col) {
    let wants = null;
    for (const r of ranges) {
      if (!r.areas.some((a) => row >= a.top && row <= a.bottom && col >= a.left && col <= a.right)) continue;
      if (!r.hasPassword || this._unlockedRanges?.has(this.activeSheet + '|' + r.title)) return true;
      wants = wants ?? r.title;
    }
    return wants;
  }

  /**
   * What the active cell says about typing into it: on a protected sheet,
   * a locked cell in a password range not yet unlocked names that range —
   * the window asks for its password before an edit starts, as Excel's
   * Unlock Range does. Null otherwise.
   */
  rangeLockAt(row = this.selection.active.row, col = this.selection.active.col) {
    if (!this.protection().sheet || !this.isCellLocked(row, col)) return null;
    const v = this._rangeVerdict(this._editRangeIndex(), row, col);
    return typeof v === 'string' ? v : null;
  }

  /** Normalise the cells of an edit range: A1-style areas, spaces between them. */
  _cleanRangeRef(text) {
    const areas = String(text ?? '').replace(/\$/g, '').replace(/^=/, '').split(/[\s,;]+/).filter(Boolean);
    if (!areas.length) throw new Error('Say which cells the range covers, such as B2:D10.');
    const out = [];
    for (const a of areas) {
      const bare = a.replace(/^.*!/, '');
      const r = parseArea(bare);
      if (!r) throw new Error('"' + a + '" is not a range of cells.');
      out.push(r.top === r.bottom && r.left === r.right ? ref(r.top, r.left) : ref(r.top, r.left) + ':' + ref(r.bottom, r.right));
    }
    return out.join(' ');
  }

  /** Excel greys New, Modify and Delete while the sheet is protected; so does this. */
  _rangesGate() {
    if (this.protection().sheet) throw protectionError('The sheet is protected — unprotect it to change the ranges that stay editable.');
  }

  /**
   * Add or change an edit range. `was` names the range being changed
   * (Modify); a `password` of null keeps the one it has, '' takes it off.
   */
  setEditRange({ was = null, title, ref: cells, password = null }) {
    this._rangesGate();
    const name = String(title ?? '').trim();
    if (!name) throw new Error('A range needs a title.');
    if (/[^\p{L}\p{N}_ .-]/u.test(name) || name.length > 255) throw new Error('A range title is letters, digits, spaces, dots, hyphens and underscores.');
    const sqref = this._cleanRangeRef(cells);
    const list = this.workbook.protectedRanges(this.activeSheet);
    if (list.some((r) => r.title.toLowerCase() === name.toLowerCase() && r.title !== was)) throw new Error('There is already a range called "' + name + '".');
    const next = list.map((r) => ({ title: r.title, sqref: r.sqref, hash: r.hash }));
    const entry = { title: name, sqref, password: password === null ? undefined : password };
    const at = was === null ? -1 : next.findIndex((r) => r.title === was);
    if (at >= 0) {
      entry.hash = next[at].hash;
      next[at] = entry;
    } else {
      next.push(entry);
    }
    this._edit(at >= 0 ? 'modify edit range' : 'new edit range', null, [], () => {
      this.workbook.setProtectedRanges(this.activeSheet, next);
      this._structuralDirty = true;
    }, { parts: [this.workbook.partNameFor(this.activeSheet)] });
    return this;
  }

  /** Delete an edit range. */
  deleteEditRange(title) {
    this._rangesGate();
    const list = this.workbook.protectedRanges(this.activeSheet);
    if (!list.some((r) => r.title === title)) return this;
    this._edit('delete edit range', null, [], () => {
      this.workbook.setProtectedRanges(this.activeSheet, list.filter((r) => r.title !== title).map((r) => ({ title: r.title, sqref: r.sqref, hash: r.hash })));
      this._structuralDirty = true;
    }, { parts: [this.workbook.partNameFor(this.activeSheet)] });
    return this;
  }

  /**
   * Unlock Range: the password of an edit range, typed once — right, and
   * the range's cells take edits until the workbook is closed; wrong, and
   * Excel's sentence.
   */
  unlockRange(title, password) {
    if (!this.workbook.checkRangePassword(this.activeSheet, title, password)) throw protectionError(WRONG_PASSWORD, { needsPassword: 'range', range: title });
    if (!this._unlockedRanges) this._unlockedRanges = new Set();
    this._unlockedRanges.add(this.activeSheet + '|' + title);
    return this;
  }

  // ---- calculation ----------------------------------------------------------

  /** Formulas → Calculation Options: 'auto', 'autoNoTable' or 'manual', as the workbook keeps it. */
  calcMode() {
    return this.workbook.calcMode();
  }

  /**
   * Set the calculation mode, written to `workbook.xml` where Excel keeps
   * it. Going back to automatic calculates everything that was waiting, as
   * Excel does the moment the option is chosen. One undo step.
   *
   * "Automatic except for data tables" calculates exactly as automatic here:
   * a data table this editor makes holds its answers as values, so there is
   * nothing of one for a recalculation to redo.
   */
  setCalcMode(mode) {
    if (!['auto', 'autoNoTable', 'manual'].includes(mode)) throw new Error('calculation is auto, autoNoTable or manual');
    if (mode === this.calcMode()) return this;
    this._edit('calculation options', null, [], () => {
      this.workbook.setCalcMode(mode);
      this.calc.manual = mode === 'manual';
      if (!this.calc.manual) this.calc.recalculate({ force: true });
      this._structuralDirty = true;
    }, { parts: [this.workbook.mainPart] });
    return this;
  }

  /**
   * Calculate Now (F9) — every formula an edit has reached, on every sheet —
   * or, with `sheet`, Calculate Sheet (Shift+F9), this sheet's alone.
   * `full` works out every formula whatever changed (Ctrl+Alt+F9). Returns
   * how many formulas were worked out.
   */
  calculate({ scope = 'workbook' } = {}) {
    const result = this.calc.recalculate({
      force: true,
      all: scope === 'full',
      sheet: scope === 'sheet' ? this.activeSheet : null,
    });
    return result.calculated;
  }

  /** What the status bar says about calculation: the mode, and whether "Calculate" shows. */
  calcState() {
    return { mode: this.calcMode(), pending: this.calc.calculationPending };
  }

  /**
   * Formulas → Evaluate Formula: the active cell's formula (or `row`/`col`'s)
   * after a list of presses — 'evaluate', 'stepIn', 'stepOut', 'restart' —
   * replayed from the start, so the window keeps only the presses.
   */
  evaluateFormula({ row = this.selection.active.row, col = this.selection.active.col, actions = [] } = {}) {
    const resolver = this.calc.resolver();
    const getFormula = (sheet, r, c) => {
      const input = this.calc.getInput(sheet, r, c);
      return typeof input === 'string' && input.startsWith('=') ? input : null;
    };
    const at = { sheet: this.activeSheet, row, col };
    if (!getFormula(at.sheet, row, col)) {
      throw new Error('Select a cell with a formula in it, then Evaluate Formula steps through it.');
    }
    return FormulaEvaluation.replay(resolver, at, actions, { getFormula }).state();
  }

  // ---- views: Normal, Page Break Preview, Split -------------------------------

  /** View → Normal or Page Break Preview, as the sheet's view keeps it. */
  viewMode() {
    return this.workbook.sheetViewMode(this.activeSheet);
  }

  /**
   * Switch the sheet's view, written to `<sheetView view>` as Excel keeps
   * it. How a sheet is looked at is not an edit: like Excel's, it is kept
   * in the file and never on the undo list.
   */
  setViewMode(mode) {
    if (!['normal', 'pageBreakPreview', 'pageLayout'].includes(mode)) throw new Error('a sheet is shown normal, as a page break preview or as its page layout');
    if (mode === this.viewMode()) return this;
    this.workbook.setSheetViewMode(this.activeSheet, mode);
    // Each view scrolls over its own canvas — cells, or paper — so the
    // window starts the new one at its top.
    this.scrollX = 0;
    this.scrollY = 0;
    this._structuralDirty = true;
    return this;
  }

  // ---- View → Custom Views ---------------------------------------------------

  /**
   * Why Custom Views cannot be used, or null. Excel greys the command in a
   * workbook that has a table (a ListObject) anywhere in it; so does this.
   */
  customViewsBlocked() {
    return this.pkg.partNames().some((p) => /^xl\/tables\/[^/]+\.xml$/.test(p))
      ? 'not available in a workbook that contains a table, as in Excel'
      : null;
  }

  /** The workbook's custom views: name, whether each keeps print settings and hidden rows, and its sheet. */
  customViews() {
    return this.workbook.customWorkbookViews().map((v) => ({
      name: v.name, printSettings: v.printSettings, hiddenRowCol: v.hiddenRowCol, sheet: this.workbook.sheetWithId(v.activeSheetId),
    }));
  }

  /**
   * View → Custom Views → Add: the way the workbook looks now, kept under a
   * name, written as Excel writes it — a `<customWorkbookView>` with a GUID
   * in workbook.xml, a `<customSheetView>` of that GUID in every sheet
   * (its zoom, selection, first cell shown and view; with print settings
   * its margins, setup, header and footer and breaks; with hidden rows,
   * columns and filter settings its autofilter), and Excel's hidden
   * `Z_<GUID>_.wvu.Rows`, `.wvu.Cols`, `.wvu.PrintArea`,
   * `.wvu.PrintTitles` and `.wvu.FilterData` names for what it hides and
   * prints. A view of the same name is replaced, as Excel offers to. Not an
   * edit that undoes: like Excel's, it clears the undo list.
   */
  addCustomView({ name, printSettings = true, hiddenRowCol = true, zoom = 1, windowWidth = 1440, windowHeight = 900 } = {}) {
    const blocked = this.customViewsBlocked();
    if (blocked) throw new Error('Custom Views are ' + blocked + '.');
    const clean = String(name ?? '').trim();
    if (!clean) throw new Error('A custom view needs a name.');
    if (clean.length > 255) throw new Error('A custom view name is 255 characters at most.');
    if (this.customViews().some((v) => v.name.toLowerCase() === clean.toLowerCase())) this.deleteCustomView(this.customViews().find((v) => v.name.toLowerCase() === clean.toLowerCase()).name);
    const guid = newGuid();
    const names = this.sheetNames();
    names.forEach((sheet, index) => {
      const xml = this._customSheetViewXml(sheet, index, guid, { printSettings, hiddenRowCol, zoom });
      const map = this.workbook.customSheetViews(sheet);
      map.set(guid, xml);
      this.workbook.setCustomSheetViews(sheet, map);
    });
    const list = this.workbook.customWorkbookViews();
    list.push({
      xml: '<customWorkbookView name="' + escapeXml(clean) + '" guid="' + guid + '"'
        + (printSettings ? '' : ' includePrintSettings="0"')
        + (hiddenRowCol ? '' : ' includeHiddenRowCol="0"')
        + ' maximized="1" xWindow="-8" yWindow="-8" windowWidth="' + Math.max(1, Math.round(windowWidth)) + '" windowHeight="' + Math.max(1, Math.round(windowHeight))
        + '" activeSheetId="' + this.workbook.sheetIdOf(this.activeSheet) + '"/>',
    });
    this.workbook.setCustomWorkbookViews(list);
    this.history = new History();
    this._structuralDirty = true;
    return clean;
  }

  /** Excel's name for what a custom view keeps of a sheet: `Z_<GUID, underscores>_.wvu.<kind>`. */
  _wvuName(guid, kind) {
    return 'Z_' + String(guid).replace(/[{}]/g, '').replace(/-/g, '_') + '_.wvu.' + kind;
  }

  _customSheetViewXml(sheet, index, guid, { printSettings, hiddenRowCol, zoom }) {
    const wb = this.workbook;
    const { part } = wb._sheetPart(sheet);
    const geo = this.geometry.get(sheet);
    const own = part.sheetViewAttrs();
    const active = sheet === this.activeSheet;
    const quoted = /[^A-Za-z0-9_]/.test(sheet) || /^\d/.test(sheet) ? "'" + sheet.replace(/'/g, "''") + "'" : sheet;
    const scale = active ? Math.round((Number(zoom) || 1) * 100) : Number(own.zoomScale || 100);
    const mode = wb.sheetViewMode(sheet);
    let topLeft = own.topLeftCell || null;
    let selection = (/<selection\b[^>]*\/>/.exec(part.prefix) ?? [null])[0];
    if (active) {
      const scroll = mode === 'pageLayout'
        ? { x: this._pageLayout().toSheetX(this.scrollX), y: this._pageLayout().toSheetY(this.scrollY) }
        : { x: this.scrollX, y: this.scrollY };
      topLeft = ref(geo.rowAt(scroll.y), geo.colAt(scroll.x));
      const r = this.selection.range;
      const a = this.selection.active;
      const sqref = r.top === r.bottom && r.left === r.right ? ref(r.top, r.left) : ref(r.top, r.left) + ':' + ref(r.bottom, r.right);
      selection = '<selection activeCell="' + ref(a.row, a.col) + '" sqref="' + sqref + '"/>';
    }
    const hiddenRows = [...(geo?.hiddenRows ?? [])].sort((x, y) => x - y);
    const hiddenCols = [...(geo?.hiddenCols ?? [])].sort((x, y) => x - y);
    const filter = part.autoFilterXml();
    const setup = readPageSetup(this, sheet);
    const hidden = wb.hiddenSheets().includes(sheet);
    const attrsText = ' guid="' + guid + '"'
      + (scale !== 100 ? ' scale="' + scale + '"' : '')
      + (hiddenRowCol && filter ? ' filter="1" showAutoFilter="1"' : '')
      + (hiddenRowCol && hiddenRows.length ? ' hiddenRows="1"' : '')
      + (hiddenRowCol && hiddenCols.length ? ' hiddenColumns="1"' : '')
      + (printSettings && setup.area ? ' printArea="1"' : '')
      + (printSettings && setup.fit !== 'none' ? ' fitToPage="1"' : '')
      + (hidden ? ' state="hidden"' : '')
      + (mode !== 'normal' ? ' view="' + mode + '"' : '')
      + (topLeft && topLeft !== 'A1' ? ' topLeftCell="' + topLeft + '"' : '');
    const kids = [];
    if (selection) kids.push(selection);
    if (printSettings) {
      for (const tag of ['rowBreaks', 'colBreaks', 'pageMargins', 'printOptions', 'pageSetup', 'headerFooter']) {
        const el = part.tailElement(tag);
        if (el) kids.push(el);
      }
    }
    if (hiddenRowCol && filter) kids.push(filter);
    // What it hides and prints, in Excel's hidden names scoped to the sheet.
    const runs = (list, fmt) => {
      const out = [];
      for (let k = 0; k < list.length; k++) {
        let end = k;
        while (end + 1 < list.length && list[end + 1] === list[end] + 1) end += 1;
        out.push(quoted + '!' + fmt(list[k]) + ':' + fmt(list[end]));
        k = end;
      }
      return out.join(',');
    };
    const scope = { localSheetId: index, hidden: 1 };
    if (hiddenRowCol && hiddenRows.length) wb.setDefinedName(this._wvuName(guid, 'Rows'), runs(hiddenRows, (r) => '$' + (r + 1)), scope);
    if (hiddenRowCol && hiddenCols.length) wb.setDefinedName(this._wvuName(guid, 'Cols'), runs(hiddenCols, (c) => '$' + colName(c)), scope);
    if (hiddenRowCol && filter) {
      const at = /\bref="([^"]+)"/.exec(filter)?.[1];
      if (at) wb.setDefinedName(this._wvuName(guid, 'FilterData'), quoted + '!' + at.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2'), scope);
    }
    if (printSettings && setup.area) wb.setDefinedName(this._wvuName(guid, 'PrintArea'), quoted + '!' + setup.area.toUpperCase().replace(/([A-Z]+)(\d+)/g, '$$$1$$$2'), scope);
    if (printSettings && setup.repeatRows > 0) wb.setDefinedName(this._wvuName(guid, 'PrintTitles'), quoted + '!$1:$' + setup.repeatRows, scope);
    return '<customSheetView' + attrsText + (kids.length ? '>' + kids.join('') + '</customSheetView>' : '/>');
  }

  /**
   * View → Custom Views → Show: the workbook as the view kept it — its
   * sheet, and on every sheet the zoom, selection, first cell shown and
   * view, with the print settings and the hidden rows, columns and filter
   * when the view kept them. Answers the sheet and the zoom, which the
   * window's own. Not undoable, as in Excel.
   */
  showCustomView(name) {
    const blocked = this.customViewsBlocked();
    if (blocked) throw new Error('Custom Views are ' + blocked + '.');
    const view = this.workbook.customWorkbookViews().find((v) => v.name === name);
    if (!view) throw new Error('No custom view is called "' + name + '".');
    const wb = this.workbook;
    const defined = wb.definedNames();
    const nameFor = (kind, index) => defined.find((d) => d.name === this._wvuName(view.guid, kind) && attrsOfText(d.attrsStr).localSheetId === String(index))?.ref ?? null;
    const target = wb.sheetWithId(view.activeSheetId) ?? this.activeSheet;
    let zoom = 1;
    let selection = null;
    let topLeft = null;
    this.sheetNames().forEach((sheet, index) => {
      const xml = wb.customSheetViews(sheet).get(view.guid);
      if (!xml) return;
      const a = attrsOfText(/<customSheetView\b([^>]*?)\/?>/.exec(xml)[1]);
      const { part } = wb._sheetPart(sheet);
      if (view.hiddenRowCol) {
        const hiddenRows = new Set();
        const hiddenCols = new Set();
        for (const piece of String(nameFor('Rows', index) ?? '').split(',')) {
          const m = /\$?(\d+):\$?(\d+)\s*$/.exec(piece);
          if (m) for (let r = Number(m[1]) - 1; r <= Number(m[2]) - 1; r++) hiddenRows.add(r);
        }
        for (const piece of String(nameFor('Cols', index) ?? '').split(',')) {
          const m = /\$?([A-Z]+):\$?([A-Z]+)\s*$/i.exec(piece);
          if (m) for (let c = colIndexOf(m[1]); c <= colIndexOf(m[2]); c++) hiddenCols.add(c);
        }
        const geo = this.geometry.get(sheet);
        for (const r of new Set([...(geo?.hiddenRows ?? []), ...hiddenRows])) part.setRowOutline(r, { hidden: hiddenRows.has(r) });
        for (const c of new Set([...(geo?.hiddenCols ?? []), ...hiddenCols])) part.setColOutline(c, { hidden: hiddenCols.has(c) });
        part.setAutoFilter((/<autoFilter\b[^>]*(?:\/>|>[\s\S]*?<\/autoFilter>)/.exec(xml) ?? [null])[0]);
        part.dirty = true;
      }
      if (view.printSettings) {
        for (const tag of ['printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks']) {
          part.setTailElement(tag, (new RegExp('<' + tag + '\\b[^>]*(?:/>|>[\\s\\S]*?</' + tag + '>)').exec(xml) ?? [null])[0]);
        }
        const area = nameFor('PrintArea', index);
        const titles = nameFor('PrintTitles', index);
        wb.setDefinedName('_xlnm.Print_Area', area, { localSheetId: index });
        wb.setDefinedName('_xlnm.Print_Titles', titles, { localSheetId: index });
      }
      wb.setSheetViewMode(sheet, a.view === 'pageLayout' || a.view === 'pageBreakPreview' ? a.view : 'normal');
      if (a.state === 'hidden' || a.state === 'veryHidden') {
        if (sheet !== target) wb.setSheetHidden(sheet, true);
      } else if (wb.hiddenSheets().includes(sheet)) {
        wb.setSheetHidden(sheet, false);
      }
      if (sheet === target) {
        zoom = (Number(a.scale) || 100) / 100;
        selection = attrsOfText((/<selection\b([^>]*)\/?>/.exec(xml) ?? ['', ''])[1]);
        topLeft = a.topLeftCell || 'A1';
      }
    });
    this._rebuildDerivedState();
    this.activeSheet = target;
    this.selection = Selection.at(0, 0);
    try {
      if (selection?.sqref) {
        const [first] = String(selection.sqref).split(/\s+/);
        const [p, q] = first.split(':');
        const [r1, c1] = parseRefPair(p);
        const [r2, c2] = q ? parseRefPair(q) : [r1, c1];
        const [ar, ac] = selection.activeCell ? parseRefPair(selection.activeCell) : [r1, c1];
        const sel = Selection.at(r1, c1);
        sel.extendTo(r2, c2);
        sel.active = { row: ar, col: ac };
        this.selection = sel;
      }
    } catch { /* a view whose selection cannot be read starts at A1 */ }
    let [tr, tc] = [0, 0];
    try { [tr, tc] = parseRefPair(topLeft || 'A1'); } catch { /* A1 */ }
    if (this.viewMode() === 'pageLayout') {
      const L = this._pageLayout();
      this.scrollX = Math.max(0, L.pageX(0) - PAGE_LAYOUT_EDGE);
      this.scrollY = Math.max(0, L.mapY(tr) - L.margins.top - L.head - PAGE_LAYOUT_EDGE);
    } else {
      this.scrollX = this.geo.colOffset(tc);
      this.scrollY = this.geo.rowOffset(tr);
    }
    this.history = new History();
    this._structuralDirty = true;
    return { sheet: target, zoom };
  }

  /** View → Custom Views → Delete: the view, its sheet views and its hidden names. */
  deleteCustomView(name) {
    const list = this.workbook.customWorkbookViews();
    const view = list.find((v) => v.name === name);
    if (!view) return this;
    this.workbook.setCustomWorkbookViews(list.filter((v) => v !== view));
    for (const sheet of this.sheetNames()) {
      const map = this.workbook.customSheetViews(sheet);
      if (map.delete(view.guid)) this.workbook.setCustomSheetViews(sheet, map);
    }
    const prefix = this._wvuName(view.guid, '');
    for (const d of this.workbook.definedNames()) {
      if (d.name.startsWith(prefix)) this.workbook.deleteDefinedName(d.name);
    }
    this.history = new History();
    this._structuralDirty = true;
    return this;
  }

  // ---- Page Layout → Background ---------------------------------------------------

  /** The active sheet's background picture's part, or null. */
  sheetBackground() {
    return this.workbook.sheetBackground(this.activeSheet);
  }

  /**
   * Page Layout → Background: a picture tiled behind the cells, written as
   * Excel writes it (`<picture r:id>` and the image part). Excel neither
   * prints it nor undoes it — the undo list is cleared, as Excel clears it.
   */
  setBackground({ contentType, data } = {}) {
    const ext = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'webp' }[contentType];
    if (!ext) throw new Error('A background is a PNG, JPEG, GIF, BMP or WebP picture.');
    const bytes = Buffer.isBuffer(data) ? data : data instanceof Uint8Array ? Buffer.from(data) : Buffer.from(String(data ?? ''), 'base64');
    if (!bytes.length) throw new Error('the picture has no bytes');
    const part = this.workbook.setSheetBackground(this.activeSheet, bytes, ext, contentType);
    this.history = new History();
    this._structuralDirty = true;
    return part;
  }

  /** Page Layout → Delete Background. */
  deleteBackground() {
    if (!this.workbook.removeSheetBackground(this.activeSheet)) return this;
    this.history = new History();
    this._structuralDirty = true;
    return this;
  }

  /** View → Ruler, shown in Page Layout: `<sheetView showRuler>`, on unless it says "0". */
  showRuler() {
    const v = this.workbook._sheetPart(this.activeSheet).part.sheetViewAttrs().showRuler;
    return !(v === '0' || v === 'false');
  }

  setShowRuler(on) {
    if (Boolean(on) === this.showRuler()) return this;
    this.workbook._sheetPart(this.activeSheet).part.setSheetViewAttr('showRuler', on ? null : '0');
    this._structuralDirty = true;
    return this;
  }

  /**
   * View → Page Layout: the sheet laid out as the printer will cut it — the
   * same page setup and the same bands (`planBands`, which the print's own
   * `paginate` crosses into pages) — each band of rows and of columns a
   * sheet of paper with its margins, header and footer, the pages in a grid
   * with a gap between them. Past the printed range the same cutting goes
   * on, as blank pages to type onto, as Excel's does. The bands are cached
   * until the setup, the sizes or the reach change.
   *
   * Positions are the sheet's own pixels at 100%: a page is the paper's size
   * over the print's scale, so the cells keep their size and a page scaled
   * to fit holds as many of them as the printed page will.
   */
  _pageLayout() {
    const sheet = this.activeSheet;
    const geo = this.geo;
    const setup = pageSetup(readPageSetup(this, sheet));
    const bounds = this.calc.usedBounds(sheet);
    const printRange = parseArea(setup.area) || { top: 0, left: 0, bottom: bounds.maxRow, right: bounds.maxCol };
    const ext = this._plExtent?.sheet === sheet ? this._plExtent : { rows: 0, cols: 0 };
    const rowsTo = Math.min(MAX_ROWS - 1, Math.max(printRange.bottom + 60, this.selection.range.bottom + 60, ext.rows, 199));
    const colsTo = Math.min(MAX_COLS - 1, Math.max(printRange.right + 12, this.selection.range.right + 12, ext.cols, 29));
    const ruler = this.showRuler();
    const key = [sheet, JSON.stringify(setup), JSON.stringify(printRange), rowsTo, colsTo, ruler].join('|');
    const rowsIndex = geo._rowIndex();
    const colsIndex = geo._colIndex();
    const hit = this._plCache;
    if (hit && hit.key === key && hit.geo === geo && hit.rows === rowsIndex && hit.cols === colsIndex) return hit.layout;
    // The print's own cut, for its scale and its page count…
    const printed = planBands({ geo, range: printRange, setup });
    // …then the same cut carried on past it, at that scale.
    const laid = planBands({ geo, range: { top: 0, left: 0, bottom: rowsTo, right: colsTo }, setup: { ...setup, fit: 'none', scale: printed.scale } });
    const layout = buildPageLayout(geo, laid, setup, {
      rowsTo, colsTo, ruler,
      printedCols: printed.colBands.length,
      printedRows: printed.rowBands.length,
      printCount: printed.colBands.length * printed.rowBands.length,
    });
    this._plCache = { key, geo, rows: rowsIndex, cols: colsIndex, layout };
    return layout;
  }

  /**
   * View → Page Break Preview: where the printer will cut the sheet, from
   * the same page setup and the same pagination the print uses — the area
   * printed, each page's box with its number, and the breaks between them,
   * each marked as put by hand (a manual break) or by the paper. Pages and
   * breaks are those near `viewport` (all of them without one); `count` is
   * every page.
   */
  pageBreakPreview({ viewport = null } = {}) {
    const sheet = this.activeSheet;
    const geo = this.geo;
    const setup = pageSetup(readPageSetup(this, sheet));
    const bounds = this.calc.usedBounds(sheet);
    const range = parseArea(setup.area) || { top: 0, left: 0, bottom: bounds.maxRow, right: bounds.maxCol };
    // The plan is arithmetic over every row of the printed range; it is
    // kept until the setup, the range or a size changes, so a scroll step
    // on a long sheet pays nothing for it.
    const key = sheet + '|' + JSON.stringify(setup) + '|' + JSON.stringify(range);
    const rowsIndex = geo._rowIndex();
    const colsIndex = geo._colIndex();
    const hit = this._breakPlan;
    let laid;
    if (hit && hit.key === key && hit.geo === geo && hit.rows === rowsIndex && hit.cols === colsIndex) {
      laid = hit.laid;
    } else {
      laid = layOutPages(geo, range, setup, paginate({ geo, range, setup, maxPages: 20000 }));
      this._breakPlan = { key, geo, rows: rowsIndex, cols: colsIndex, laid };
    }
    const near = viewport
      ? (x, y, w, h) => !(x > this.scrollX + this.viewportWidth + 200 || y > this.scrollY + this.viewportHeight + 200 || x + w < this.scrollX - 200 || y + h < this.scrollY - 200)
      : () => true;
    const { area } = laid;
    return {
      area,
      pages: laid.pages.filter((p) => near(p.x, p.y, p.width, p.height)),
      rows: laid.rows.filter((b) => near(area.x, b.y, area.width, 1)),
      cols: laid.cols.filter((b) => near(b.x, area.y, 1, area.height)),
      count: laid.count,
      truncated: laid.truncated,
    };
  }

  /**
   * View → Split: the window's panes — `width` and `height` of the left
   * and top panes in pixels (0 when the window is split one way only), and
   * `top` / `left`, the first row and column those panes show — or null.
   * Read from the sheet's `<pane>`, where xSplit and ySplit are twentieths
   * of a point from the window's edge, the headings included; where the top
   * and left panes have been scrolled to is this window's, as it is Excel's.
   */
  splitPane() {
    const s = this.workbook.splitPane(this.activeSheet);
    if (!s) return null;
    const geo = this.geo;
    const px = (twips) => Math.round(twips / 15);
    const width = s.xSplit ? Math.max(0, px(s.xSplit) - geo.headerWidth) : 0;
    const height = s.ySplit ? Math.max(0, px(s.ySplit) - geo.headerHeight) : 0;
    if (!width && !height) return null;
    let top = 0;
    let left = 0;
    if (s.viewTopLeftCell) {
      try { [top, left] = parseRefPair(s.viewTopLeftCell); } catch { /* a bad ref starts at A1 */ }
    }
    const scrolled = this._splitScroll?.get(this.activeSheet);
    return { width, height, top: scrolled?.top ?? top, left: scrolled?.left ?? left };
  }

  /**
   * Split the window at a cell, as View → Split does at the active cell:
   * the panes divide above and left of it, one way only when it is in the
   * first row or column showing, and in four at the middle of the window
   * when it is the top-left cell showing.
   */
  splitAt({ row = this.selection.active.row, col = this.selection.active.col } = {}) {
    const geo = this.geo;
    const firstRow = geo.rowAt(this.scrollY);
    const firstCol = geo.colAt(this.scrollX);
    let height = row > firstRow ? geo.rowOffset(row) - geo.rowOffset(firstRow) : 0;
    let width = col > firstCol ? geo.colOffset(col) - geo.colOffset(firstCol) : 0;
    if (height >= this.viewportHeight - 10) height = 0;
    if (width >= this.viewportWidth - 10) width = 0;
    if (!height && !width) {
      height = Math.round(this.viewportHeight / 2);
      width = Math.round(this.viewportWidth / 2);
    }
    return this.setSplit({ width, height, top: firstRow, left: firstCol });
  }

  /**
   * The split bars moved: each pane's size snapped to the nearest row or
   * column edge. A bar dragged to its edge of the window takes that split
   * away; with neither left, the window is whole again. Not an edit: kept
   * in the file, never on the undo list, as Excel keeps it.
   */
  setSplit({ width = 0, height = 0, top, left } = {}) {
    const geo = this.geo;
    const current = this.splitPane();
    const t = Math.max(0, Math.round(top ?? current?.top ?? geo.rowAt(this.scrollY)));
    const l = Math.max(0, Math.round(left ?? current?.left ?? geo.colAt(this.scrollX)));
    const snap = (size, start, offset, sizeOf, max) => {
      if (!(size > 6)) return 0;
      const target = offset(start) + size;
      let i = start;
      while (i < max && offset(i) + sizeOf(i) <= target) i += 1;
      const before = offset(i) - offset(start);
      const after = offset(i) + sizeOf(i) - offset(start);
      const snapped = Math.abs(target - offset(start) - before) <= Math.abs(after - (target - offset(start))) ? before : after;
      return snapped > 6 ? snapped : 0;
    };
    const h = snap(Number(height) || 0, t, (r) => geo.rowOffset(r), (r) => geo.rowHeight(r), MAX_ROWS - 1);
    const w = snap(Number(width) || 0, l, (c) => geo.colOffset(c), (c) => geo.colWidth(c), MAX_COLS - 1);
    if (!this._splitScroll) this._splitScroll = new Map();
    this._splitScroll.delete(this.activeSheet);
    if (!h && !w) return this.removeSplit();
    const bottomRight = ref(geo.rowAt(this.scrollY + h), geo.colAt(this.scrollX + w));
    this.workbook.setSplitPane(this.activeSheet, {
      xSplit: w ? (w + geo.headerWidth) * 15 : 0,
      ySplit: h ? (h + geo.headerHeight) * 15 : 0,
      topLeftCell: bottomRight,
      viewTopLeftCell: ref(t, l),
    });
    this._structuralDirty = true;
    return this;
  }

  /** View → Split pressed again: the window whole. */
  removeSplit() {
    if (!this.workbook.splitPane(this.activeSheet)) return this;
    this.workbook.setSplitPane(this.activeSheet, null);
    this._splitScroll?.delete(this.activeSheet);
    this._structuralDirty = true;
    return this;
  }

  /** View → Split as a toggle: split at the active cell, or whole again. */
  toggleSplit() {
    return this.splitPane() ? this.removeSplit() : this.splitAt();
  }

  /** The top or left pane scrolled by whole rows or columns, never before the first. */
  scrollSplit({ rows = 0, cols = 0 } = {}) {
    const s = this.splitPane();
    if (!s) return this;
    if (!this._splitScroll) this._splitScroll = new Map();
    this._splitScroll.set(this.activeSheet, {
      top: Math.max(0, Math.min(MAX_ROWS - 1, s.top + Math.round(Number(rows) || 0))),
      left: Math.max(0, Math.min(MAX_COLS - 1, s.left + Math.round(Number(cols) || 0))),
    });
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
    const own = this.cellStyles.get(sheetName)?.get(ref(row, col));
    if (own !== undefined) return own;
    if (this._lazyStyles.has(sheetName)) {
      // The part's own cell, read as asked for; a cleared style is a null
      // in the map above, so the file's does not show through.
      const cell = this.workbook._sheetPart(sheetName).part.getCell(row, col);
      return cell?.style != null && cell.style !== '' ? Number(cell.style) : null;
    }
    return null;
  }

  _setStyleIndex(sheetName, row, col, index) {
    let map = this.cellStyles.get(sheetName);
    if (!map) { map = new Map(); this.cellStyles.set(sheetName, map); }
    if (index === null || index === undefined) {
      if (this._lazyStyles.has(sheetName)) map.set(ref(row, col), null);
      else map.delete(ref(row, col));
    } else map.set(ref(row, col), index);
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
        align: null, valign: null, wrap: false, indent: 0, rotation: 0, numberFormat: 'General',
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
      // Indent and rotation the same way: two values read as none, so the
      // orientation menu marks nothing and the indent arrows step each cell.
      if (all.indent !== at.indent) all.indent = null;
      if (all.rotation !== at.rotation) all.rotation = null;
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
    for (const k of ['fontName', 'fontSize', 'fontColour', 'fill', 'border', 'numberFormat', 'valign', 'wrap', 'indent', 'indentBy', 'rotation', 'locked']) {
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

  save({ recalc = true } = {}) {
    // Excel recalculates a workbook in manual mode before saving it, so the
    // file never keeps an answer the sheet had not caught up with.
    // A draft kept for recovery (serialize) leaves the waiting alone.
    if (recalc && this.calc.calculationPending) this.calc.recalculate({ force: true });
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
    const out = this.save({ recalc: false });
    for (const key of dirty) this.dirtyCells.add(key);
    for (const key of styled) this.styledCells.add(key);
    this._structuralDirty = structural;
    return out;
  }
}

/** Page Layout's paper: the space round the pages, and between them. */
export const PAGE_LAYOUT_EDGE = 36;
export const PAGE_LAYOUT_GAP = 36;

/**
 * View → Page Layout, as geometry: the pages' size and margins in the
 * sheet's pixels, where each band of rows and columns starts and ends, and
 * the maps from a row or column to its place on paper and back from a
 * scroll position on paper to the sheet's.
 */
function buildPageLayout(geo, laid, setup, info) {
  const s = laid.scale || 1;
  const paper = PAPER[setup.paper] || PAPER.A4;
  const landscape = setup.orientation === 'landscape';
  const mm = (v) => (Number(v) || 0) * PX_PER_MM / s;
  const pageW = mm(landscape ? paper.height : paper.width);
  const pageH = mm(landscape ? paper.width : paper.height);
  const margins = { top: mm(setup.margins.top), right: mm(setup.margins.right), bottom: mm(setup.margins.bottom), left: mm(setup.margins.left) };
  // The running header and footer take a line of the printed area each, as they do on paper.
  const head = setup.header ? 22 / s : 0;
  const foot = setup.footer ? 22 / s : 0;
  const bandsOf = (bands, first) => bands.map((b, i) => ({
    start: i === 0 ? first : (b[0]?.index ?? first),
    end: b.length ? b[b.length - 1].index : first,
  }));
  const cols = bandsOf(laid.colBands, 0);
  const rows = bandsOf(laid.rowBands, 0);
  // Print titles belong to the first band of rows: they are its top.
  if (laid.titleRows?.length && rows.length) rows[0].start = 0;
  const E = PAGE_LAYOUT_EDGE;
  const G = PAGE_LAYOUT_GAP;
  const pageX = (i) => E + i * (pageW + G);
  const pageY = (j) => E + j * (pageH + G);
  const find = (bands, index) => {
    let lo = 0;
    let hi = bands.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (bands[mid].start <= index) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  };
  const contentX = (i) => pageX(i) + margins.left;
  const contentY = (j) => pageY(j) + margins.top + head;
  const mapX = (col) => { const i = find(cols, col); return contentX(i) + geo.colOffset(col) - geo.colOffset(cols[i].start); };
  const mapY = (row) => { const j = find(rows, row); return contentY(j) + geo.rowOffset(row) - geo.rowOffset(rows[j].start); };
  const spanX = (i) => geo.colOffset(cols[i].end) + geo.colWidth(cols[i].end) - geo.colOffset(cols[i].start);
  const spanY = (j) => geo.rowOffset(rows[j].end) + geo.rowHeight(rows[j].end) - geo.rowOffset(rows[j].start);
  const band = (p, first, size, n) => Math.max(0, Math.min(n - 1, Math.floor((p - first) / size)));
  const toSheetX = (px) => {
    const i = band(px, E, pageW + G, cols.length);
    return geo.colOffset(cols[i].start) + Math.max(0, Math.min(spanX(i), px - contentX(i)));
  };
  const toSheetY = (py) => {
    const j = band(py, E, pageH + G, rows.length);
    return geo.rowOffset(rows[j].start) + Math.max(0, Math.min(spanY(j), py - contentY(j)));
  };
  return {
    scale: s, pageW, pageH, margins, head, foot, cols, rows, pageX, pageY, contentX, contentY, spanX, spanY,
    mapX, mapY, toSheetX, toSheetY,
    rowsTo: info.rowsTo, colsTo: info.colsTo, ruler: info.ruler,
    printedCols: info.printedCols, printedRows: info.printedRows, printCount: info.printCount,
    order: setup.order === 'across' ? 'across' : 'down',
    header: setup.header || '', footer: setup.footer || '',
    total: { width: Math.ceil(E + cols.length * (pageW + G)), height: Math.ceil(E + rows.length * (pageH + G)) },
  };
}

/**
 * The grid's frame put onto paper: every row, column, cell and drawing moved
 * to its place on its page, the canvas the size of the pages, and the pages
 * in sight listed — each with its number as printed (none past the printed
 * range: those are Excel's blank pages to type onto), its printable box and
 * where its header and footer go.
 */
function pageLayoutFrame(frame, L, geo, sight) {
  // What lies past the last page laid out is left for the next frame, when the layout has grown to it.
  const lastRow = L.rows[L.rows.length - 1].end;
  const lastCol = L.cols[L.cols.length - 1].end;
  frame.cells = frame.cells.filter((c) => c.row <= lastRow && c.col <= lastCol);
  frame.columns = frame.columns.filter((c) => c.index <= lastCol);
  frame.rows = frame.rows.filter((r) => r.index <= lastRow);
  for (const c of frame.cells) { c.x = L.mapX(c.col); c.y = L.mapY(c.row); }
  for (const c of frame.columns) c.x = L.mapX(c.index);
  for (const r of frame.rows) r.y = L.mapY(r.index);
  for (const d of frame.drawings) {
    if (!d.anchor) continue;
    d.x = L.mapX(d.anchor.col) + (d.x - geo.colOffset(d.anchor.col));
    d.y = L.mapY(d.anchor.row) + (d.y - geo.rowOffset(d.anchor.row));
  }
  const E = PAGE_LAYOUT_EDGE;
  const G = PAGE_LAYOUT_GAP;
  const firstI = Math.max(0, Math.floor((sight.x - E) / (L.pageW + G)) - 1);
  const lastI = Math.min(L.cols.length - 1, Math.floor((sight.x + sight.width - E) / (L.pageW + G)) + 1);
  const firstJ = Math.max(0, Math.floor((sight.y - E) / (L.pageH + G)) - 1);
  const lastJ = Math.min(L.rows.length - 1, Math.floor((sight.y + sight.height - E) / (L.pageH + G)) + 1);
  const pages = [];
  for (let i = firstI; i <= lastI; i++) {
    for (let j = firstJ; j <= lastJ; j++) {
      const printed = i < L.printedCols && j < L.printedRows;
      const n = printed ? (L.order === 'down' ? i * L.printedRows + j + 1 : j * L.printedCols + i + 1) : null;
      pages.push({
        n, blank: !printed, col: i, row: j,
        x: L.pageX(i), y: L.pageY(j), width: L.pageW, height: L.pageH,
        // The printable box: inside the margins, the header's line and the footer's included.
        box: { x: L.contentX(i), y: L.pageY(j) + L.margins.top, width: L.pageW - L.margins.left - L.margins.right, height: L.pageH - L.margins.top - L.margins.bottom },
        first: { row: L.rows[j].start, col: L.cols[i].start },
      });
    }
  }
  return {
    ...frame,
    total: L.total,
    frozen: { rows: 0, cols: 0, width: 0, height: 0 },
    split: null,
    pageBreaks: null,
    pageLayout: {
      pages,
      count: L.printCount,
      scale: L.scale,
      pageWidth: L.pageW,
      pageHeight: L.pageH,
      margins: L.margins,
      head: L.head,
      foot: L.foot,
      header: L.header,
      footer: L.footer,
      ruler: L.ruler,
    },
  };
}

/**
 * Page Break Preview's layout of a pagination: the printed area, every
 * page's box in the sheet's pixels with its number, and every break between
 * the bands, marked as put by hand or by the paper. Worked out once per
 * plan — a long sheet has a thousand pages, and a frame only filters them.
 */
function layOutPages(geo, range, setup, plan) {
  const manualRows = new Set((setup.rowBreaks || []).map(Number));
  const manualCols = new Set((setup.colBreaks || []).map(Number));
  const right = (c) => geo.colOffset(c) + geo.colWidth(c);
  const bottom = (r) => geo.rowOffset(r) + geo.rowHeight(r);
  // The bands, by where each starts: the pages are their crossings, numbered
  // in the plan's order — down, then across.
  const colStarts = new Set();
  const rowStarts = new Set();
  for (const page of plan.pages) {
    if (page.cols.length) colStarts.add(page.cols[0].index);
    if (page.rows.length) rowStarts.add(page.rows[0].index);
  }
  const firstBand = Math.min(...rowStarts);
  const area = {
    top: range.top, left: range.left, bottom: range.bottom, right: range.right,
    x: geo.colOffset(range.left), y: geo.rowOffset(range.top),
    width: right(range.right) - geo.colOffset(range.left), height: bottom(range.bottom) - geo.rowOffset(range.top),
  };
  const pages = [];
  plan.pages.forEach((page, i) => {
    if (!page.cols.length) return;
    // The first band down starts at the top of the range: print titles are its rows too.
    const firstRow = !page.rows.length || page.rows[0].index === firstBand ? range.top : page.rows[0].index;
    const lastRow = page.rows.length ? page.rows[page.rows.length - 1].index : range.bottom;
    const x = geo.colOffset(page.cols[0].index);
    const y = geo.rowOffset(firstRow);
    pages.push({ n: i + 1, x, y, width: right(page.cols[page.cols.length - 1].index) - x, height: bottom(lastRow) - y });
  });
  const rows = [...rowStarts].sort((a, b) => a - b).slice(1).map((index) => ({ index, y: geo.rowOffset(index), manual: manualRows.has(index) }));
  const cols = [...colStarts].sort((a, b) => a - b).slice(1).map((index) => ({ index, x: geo.colOffset(index), manual: manualCols.has(index) }));
  return { area, pages, rows, cols, count: plan.pages.length, truncated: plan.truncated };
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
/** Attributes of an element's text, for the custom views' XML. */
function attrsOfText(text) {
  const out = {};
  for (const m of String(text || '').matchAll(/([\w:.-]+)="([^"]*)"/g)) out[m[1]] = m[2];
  return out;
}

/** A column's letters to its 0-based index. */
function colIndexOf(letters) {
  return [...String(letters).toUpperCase()].reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
}

/** Text for an XML attribute. */
function escapeXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function protectionError(message, extra = {}) {
  const e = new Error(message);
  e.protection = true;
  Object.assign(e, extra);
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

/**
 * The anchors of a drawing part in document order — which is drawing order —
 * each with where it starts and ends, so one can be rewritten in place.
 */
export function anchorSpans(xml) {
  const out = [];
  const re = /<([\w]+:)?(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>[\s\S]*?<\/([\w]+:)?\2>/g;
  let m;
  while ((m = re.exec(xml))) out.push({ start: m.index, end: m.index + m[0].length, xml: m[0], type: m[2], prefix: m[1] ?? '' });
  return out;
}

/**
 * One anchor put at a new box (sheet pixels): its markers — `from` and `to`
 * for a two-cell anchor, `from` and its extent for a one-cell one, position
 * and extent for an absolute one — and the first `a:xfrm` inside, which a
 * shape, a picture or a group states and Excel checks against the anchor.
 */
export function anchorWithBox(anchorXml, box, markerAt) {
  const p = /^<([\w]+:)?/.exec(anchorXml)?.[1] ?? '';
  const emu = (px) => Math.round(px * 9525);
  const marker = (tag, m) => '<' + p + tag + '><' + p + 'col>' + m.col + '</' + p + 'col><' + p + 'colOff>' + m.colOff + '</' + p + 'colOff>'
    + '<' + p + 'row>' + m.row + '</' + p + 'row><' + p + 'rowOff>' + m.rowOff + '</' + p + 'rowOff></' + p + tag + '>';
  const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const el = (tag) => new RegExp('<' + esc(p + tag) + '>[\\s\\S]*?</' + esc(p + tag) + '>');
  let out = anchorXml;
  const from = markerAt(box.x, box.y);
  const to = markerAt(box.x + box.width, box.y + box.height);
  if (el('from').test(out)) out = out.replace(el('from'), () => marker('from', from));
  if (el('to').test(out)) out = out.replace(el('to'), () => marker('to', to));
  // The anchor's own extent and position (a one-cell or absolute anchor): the
  // direct children, not the `a:ext` inside a shape's xfrm.
  out = out.replace(new RegExp('(</' + esc(p + 'from') + '>|^<' + esc(p) + 'absoluteAnchor\\b[^>]*>)(\\s*<' + esc(p) + 'pos\\b[^>]*/>)?(\\s*<' + esc(p) + 'ext\\b[^>]*/>)?'), (whole, head, pos, ext) => head
    + (pos ? '<' + p + 'pos x="' + emu(box.x) + '" y="' + emu(box.y) + '"/>' : '')
    + (ext ? '<' + p + 'ext cx="' + emu(box.width) + '" cy="' + emu(box.height) + '"/>' : ''));
  // The first xfrm inside states the same box.
  out = out.replace(/(<a:xfrm\b[^>]*>\s*)<a:off\b[^>]*\/>(\s*)<a:ext\b[^>]*\/>/, (whole, head, gap) => {
    if (/<a:off x="0" y="0"\/>\s*<a:ext cx="0" cy="0"\/>/.test(whole)) return whole; // a chart frame's placeholder
    return head + '<a:off x="' + emu(box.x) + '" y="' + emu(box.y) + '"/>' + gap + '<a:ext cx="' + emu(box.width) + '" cy="' + emu(box.height) + '"/>';
  });
  return out;
}

/** A drawing part with each anchor replaced by what `fn(span, i)` makes of it — what lies between them kept. */
function rebuild(spans, xml, fn) {
  let out = '';
  let at = 0;
  spans.forEach((s, i) => {
    out += xml.slice(at, s.start) + fn(s, i);
    at = s.end;
  });
  return out + xml.slice(at);
}

/** One attribute of an anchor's (first) cNvPr set, or taken off with null. */
function withNvAttr(anchorXml, name, value) {
  return anchorXml.replace(/<([\w]+:)?cNvPr\b([^>]*?)(\/?)>/, (tag, p, attrsText, close) => {
    let a = attrsText.replace(new RegExp('\\s+' + name + '="[^"]*"'), '');
    if (value !== null && value !== undefined) a += ' ' + name + '="' + escapeXml(String(value)) + '"';
    return '<' + (p ?? '') + 'cNvPr' + a + close + '>';
  });
}

/**
 * A drawing's own transform rewritten — its `a:xfrm` in `spPr` (a shape, a
 * picture), in `grpSpPr` (a group), or a graphic frame's own `xdr:xfrm`:
 * `mutate({ rot, flipH, flipV })` says the turn and the flips; the offset
 * and extent are kept, or written from `box` (sheet pixels) where the
 * drawing states none or `absolute` asks for them (a member put into or
 * out of a group). A group keeps its child space.
 */
function withTransform(xml, box, mutate, { absolute = false } = {}) {
  const emu = (v) => Math.round(v * 9525);
  const candidates = [
    /<([\w]+:)?(spPr|grpSpPr)\b[^>]*?\/>/.exec(xml),
    /<([\w]+:)?(spPr|grpSpPr)\b[^>]*>[\s\S]*?<\/([\w]+:)?(spPr|grpSpPr)>/.exec(xml),
    /<([\w]+:)?graphicFrame\b/.test(xml) ? /<([\w]+:)?xfrm\b[^>]*>[\s\S]*?<\/([\w]+:)?xfrm>/.exec(xml) : null,
  ].filter(Boolean).sort((a, b) => a.index - b.index);
  const block = candidates[0];
  if (!block) return xml;
  const frameXfrm = /xfrm\b/.test(block[0].slice(0, 12)) && !/spPr/.test(block[0].slice(0, 16));
  const current = /<a:xfrm\b([^>]*?)(?:\/>|>([\s\S]*?)<\/a:xfrm>)/.exec(block[0]);
  const src = frameXfrm ? /<([\w]+:)?xfrm\b([^>]*)>([\s\S]*?)<\/([\w]+:)?xfrm>/.exec(block[0]) : null;
  const attrText = frameXfrm ? src[2] : current ? current[1] : '';
  const body = frameXfrm ? src[3] : current ? (current[2] ?? '') : '';
  const attr = (n) => new RegExp('\\b' + n + '="([^"]*)"').exec(attrText)?.[1];
  const t = mutate({ rot: attr('rot') ? Number(attr('rot')) / 60000 : 0, flipH: attr('flipH') === '1', flipV: attr('flipV') === '1' });
  const off = /<a:off\b[^>]*\/>/.exec(body)?.[0];
  const ext = /<a:ext\b[^>]*\/>/.exec(body)?.[0];
  const ch = (body.match(/<a:ch(Off|Ext)\b[^>]*\/>/g) ?? []).join('');
  const offXml = !absolute && off ? off : '<a:off x="' + emu(box.x) + '" y="' + emu(box.y) + '"/>';
  const extXml = !absolute && ext && !/cx="0" cy="0"/.test(ext) ? ext : '<a:ext cx="' + emu(box.width) + '" cy="' + emu(box.height) + '"/>';
  const rot = Math.round((Number(t.rot) || 0) * 60000);
  const head = (rot ? ' rot="' + rot + '"' : '') + (t.flipH ? ' flipH="1"' : '') + (t.flipV ? ' flipV="1"' : '');
  if (frameXfrm) {
    const p = src[1] ?? '';
    const next = '<' + p + 'xfrm' + head + '>' + offXml + extXml + '</' + p + 'xfrm>';
    return xml.slice(0, block.index) + next + xml.slice(block.index + block[0].length);
  }
  const xfrm = '<a:xfrm' + head + '>' + offXml + extXml + ch + '</a:xfrm>';
  let nextBlock;
  if (current) nextBlock = block[0].replace(current[0], xfrm);
  else if (/\/>$/.test(block[0])) nextBlock = block[0].replace(/\s*\/>$/, '>' + xfrm + '</' + (block[1] ?? '') + block[2] + '>');
  else nextBlock = block[0].replace(/^<[^>]*>/, (m) => m + xfrm);
  return xml.slice(0, block.index) + nextBlock + xml.slice(block.index + block[0].length);
}

/** Items in a slicer's order: numbers by size, then words as a person sorts them. */
function slicerCompare(a, b) {
  const na = typeof a === 'number';
  const nb = typeof b === 'number';
  if (na && nb) return a - b;
  if (na !== nb) return na ? -1 : 1;
  return String(a ?? '').localeCompare(String(b ?? ''), undefined, { numeric: true, sensitivity: 'base' });
}

export { OoxmlPackage };
