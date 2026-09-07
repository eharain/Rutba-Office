/**
 * The capability register — what "all features are done" actually enumerates.
 *
 * "Finish when all features are done" is only a plan if the feature list is
 * written down, so this is it: the OOXML feature areas that matter, each with an
 * honest state and the part it lives in.
 *
 * States, in increasing order of what they cost us to reach:
 *
 *   preserve  survives a round trip byte-identically. We do not understand it,
 *             cannot show it and cannot change it — but we never damage it.
 *             This is where the preserving architecture puts almost everything
 *             for free, and it is why the file is safe on day one.
 *   read      we can extract it and reason about it.
 *   edit      we can change it and write it back correctly.
 *   render    we can display it to a user. NOTHING is here yet, because there is
 *             no editor UI — this column is the honest measure of how far the
 *             product is from being an editor rather than a file processor.
 *
 * A feature at `preserve` is not a gap in FIDELITY. It is a gap in FUNCTIONALITY.
 * Conflating those two is how a roadmap ends up sequenced by what is visible
 * rather than by what is risky, so the report prints them separately.
 *
 * Keep this file honest. It is quoted in FORMAT-FIDELITY.md and it is the thing
 * a reader will use to decide whether to trust us with a file. Downgrading a
 * state when we discover a limitation is correct behaviour, not a regression.
 */

/** @typedef {'preserve'|'read'|'edit'|'render'} State */

const F = (area, feature, part, state, note) => ({ area, feature, part, state, note });

export const SPREADSHEET_FEATURES = [
  // --- what we actually operate on ---------------------------------------
  F('cells', 'Cell values (number, text, boolean)', 'xl/worksheets/*', 'render'),
  F('cells', 'Formulas', 'xl/worksheets/*', 'render', 'evaluated; grid shows the value, formula bar the formula'),
  F('calculation', 'Arithmetic, comparison, concatenation', '@rutba/formula', 'edit', 'spreadsheet precedence, including -2^2 = -4'),
  F('calculation', 'Cell and range references', '@rutba/formula', 'edit', 'relative, absolute, cross-sheet, whole column/row'),
  F('calculation', 'Error values and propagation', '@rutba/formula', 'edit', 'errors are values, not exceptions'),
  F('calculation', 'Function library', '@rutba/formula', 'edit', '134 functions; unknown ones degrade loudly'),
  F('calculation', 'Dependency-ordered recalculation', '@rutba/formula', 'edit', 'topological, with cycle detection'),
  F('calculation', 'Date serials (1900 epoch)', '@rutba/formula', 'edit', 'including the leap-year bug the format carries'),
  F('calculation', 'Array formulas and spilling', '@rutba/formula', 'edit', 'arrays spill into empty neighbours; a blocked spill is #SPILL!'),
  F('calculation', 'Structured references', '@rutba/formula', 'edit', 'Table1[Qty], [@Col], #All/#Data/#Headers/#Totals, column spans'),
  F('calculation', 'Iterative calculation', '@rutba/formula', 'preserve', 'cycles are reported, not iterated'),
  F('cells', 'Inline strings', 'xl/worksheets/*', 'edit', 'our write path, to avoid renumbering sharedStrings'),
  F('cells', 'Shared strings', 'xl/sharedStrings.xml', 'read', 'read and resolved; never renumbered'),
  F('structure', 'Sheets and sheet order', 'xl/workbook.xml', 'render', 'switchable tabs; cannot add or reorder yet'),
  F('interaction', 'Grid, selection and keyboard navigation', '@rutba/sheet-view', 'render', 'virtualised; ctrl+arrow, tab/enter block semantics'),
  F('interaction', 'Cell editing and the formula bar', '@rutba/sheet-view', 'render'),
  F('interaction', 'Copy, paste and clear', '@rutba/sheet-view', 'render', 'tab-separated text'),
  F('interaction', 'Selection summary (count, sum, average)', '@rutba/sheet-view', 'render'),
  F('interaction', 'Frozen panes', 'xl/worksheets/*', 'preserve', 'stored and preserved, not yet rendered'),
  F('interaction', 'Undo and redo', '@rutba/editing', 'render', 'coalesced by cell: retyping one cell is one step, moving on starts another'),
  F('structure', 'Defined names (named ranges)', 'xl/workbook.xml', 'edit', 'the business-data binding anchor'),
  F('structure', 'Calculation chain hint', 'xl/workbook.xml', 'edit', 'fullCalcOnLoad, for what we cannot evaluate'),

  // --- preserved, and that is currently enough ----------------------------
  F('format', 'Number formats', 'xl/styles.xml', 'render', 'built-in ids and custom codes, incl. sections and dates'),
  F('format', 'Fonts, fills and borders', 'xl/styles.xml', 'render', 'read and painted; never rewritten — an edited cell keeps its style index'),
  F('format', 'Theme colours and tint', 'xl/theme/theme1.xml', 'render', 'index order and luminance tint per ECMA-376, not the scheme document order'),
  F('format', 'Indexed (legacy) colours', 'xl/styles.xml', 'render', 'the 56-colour palette; 64 and 65 fall back to the reader own ink'),
  F('format', 'Cell alignment, wrap and indent', 'xl/styles.xml', 'render', 'the file alignment beats the number format default'),
  F('format', 'Conditional formatting', 'xl/worksheets/*', 'render', 'cellIs/expression/text/top-N/duplicates through their dxf, colour scales and data bars painted; icon sets and timePeriod preserved only'),
  F('format', 'Column widths and row heights', 'xl/worksheets/*', 'render'),
  F('format', 'Merged cells', 'xl/worksheets/*', 'render', 'drawn once by the top-left cell across its range; covered cells are not emitted'),
  F('data', 'Data validation', 'xl/worksheets/*', 'render', 'enforced on direct entry in the author own words; list rules offer an in-cell dropdown; custom rules preserved only'),
  F('data', 'Autofilter and sort state', 'xl/tables/*', 'edit', 'value filters read and written, with the excluded rows hidden; customFilters/top10/dynamicFilter preserved'),
  F('data', 'Tables (ListObjects)', 'xl/tables/*', 'edit', 'range, header/totals counts and column names modelled; structured references calculate; table styling not yet painted'),
  F('data', 'External data connections', 'xl/connections.xml', 'preserve'),
  F('protection', 'Sheet and workbook protection', 'xl/worksheets/*', 'edit', 'sheet protection enforced (locked cells refuse, the lock is a cell format) and toggled; a password-sealed sheet refuses both the edit and the unprotect — that gate is Excel’s to lift; workbook-level protection preserved only'),
  F('view', 'Frozen panes and sheet views', 'xl/worksheets/*', 'edit', 'the pane element read, honoured in the grid, and written'),
  F('print', 'Page setup, margins, print areas', 'xl/worksheets/*', 'preserve'),
  F('objects', 'Charts', 'xl/charts/*', 'render', 'read and re-drawn by @rutba/drawing in our palette; the part is preserved untouched'),
  F('objects', 'Pivot tables and caches', 'xl/pivotTables/*', 'edit', 'read, created and recomputed from the live source; a refresh leaves the existing cache to be rebuilt on open (refreshOnLoad) while creation authors one outright; grouped fields, calculated fields and external sources refuse with a reason'),
  F('objects', 'Drawing anchors', 'xl/drawings/*', 'render', 'placement, preset shapes and embedded pictures all drawn'),
  F('objects', 'Preset shapes', 'a:prstGeom', 'render', '18 common presets drawn properly; anything else becomes a box of the right size and colour'),
  F('objects', 'Embedded pictures', 'xl/media/*', 'render', 'web image types inline as data URIs; EMF and WMF are preserved but not shown'),
  F('objects', 'Slicers', 'xl/slicers/*', 'preserve'),
  F('comments', 'Comments and threaded comments', 'xl/comments*.xml', 'render', 'shown as a corner mark with the note and its author; writing a comment is not implemented'),
  F('macros', 'VBA project', 'xl/vbaProject.bin', 'preserve', 'never executed'),
  F('metadata', 'Document properties', 'docProps/*', 'preserve'),
  F('metadata', 'Custom XML parts', 'customXml/*', 'preserve'),
  F('security', 'Digital signature parts', '_xmlsignatures/*', 'preserve', 'preserved, but any edit invalidates the signature'),
];

export const DOCUMENT_FEATURES = [
  F('text', 'Paragraph text', 'word/document.xml', 'render', 'editable; refuses paragraphs carrying structure'),
  F('text', 'Run properties on edited text', 'word/document.xml', 'render', 'carried across unchanged'),
  F('text', 'Bold, italic and underline', 'word/document.xml', 'render', 'toggled over a selection or armed at the caret'),
  F('structure', 'Content controls (w:sdt)', 'word/document.xml', 'render', 'the business-data binding anchor; fillable while the paragraph is locked'),
  F('interaction', 'Caret and selection across paragraphs', '@rutba/doc-view', 'render', 'character offsets, not run indices'),
  F('interaction', 'Insert, delete, split and merge', '@rutba/doc-view', 'render', 'runs coalesce so typing does not fragment a paragraph'),
  F('interaction', 'Copy and paste', '@rutba/doc-view', 'render', 'newlines become paragraph splits'),
  F('interaction', 'Undo and redo', '@rutba/editing', 'render', 'coalesced by paragraph: a word is one step, a paste is one step'),
  F('structure', 'Bookmarks', 'word/document.xml', 'read'),
  F('structure', 'Section properties', 'word/document.xml', 'render', 'page size, margins, header/footer references and titlePg all honoured'),
  F('layout', 'Pagination', '@rutba/doc-view/paginate', 'render', 'content breaks across sheets; line breaking is ours, so measurement and rendering cannot disagree'),
  F('layout', 'Explicit page breaks', 'w:pageBreakBefore, w:br', 'render', 'both the paragraph property and the run-level break'),
  F('structure', 'Tables', 'word/document.xml', 'render', 'grid, spans, merges, shading and nesting drawn; cells editable (type, Enter, Tab cell to cell, Tab in the last cell adds a row), tables insertable; structure edits beyond a new row wait'),
  F('structure', 'Table borders and shading', 'w:tblBorders, w:shd', 'render', 'including insideH/insideV, which is what draws a TableGrid'),
  F('fields', 'Field codes (DOCPROPERTY, …)', 'word/document.xml', 'preserve', 'preserved with the result Word cached, not re-evaluated'),
  F('fields', 'PAGE and NUMPAGES', 'word/footer*.xml', 'render', 'the two pagination makes knowable; computed per sheet'),
  F('format', 'Paragraph styles (heading, list)', 'word/styles.xml', 'render', 'resolved from styles.xml — basedOn chains, docDefaults, CSS px — so pagination measures what paint draws'),
  F('format', 'Numbering and lists', 'word/numbering.xml', 'render', 'decimal, letters, roman, bullets, multilevel with restart; deleting item two renumbers item three'),
  F('layout', 'Headers and footers', 'word/header*.xml', 'render', 'default, first and even references; read from sectPr, not by globbing the parts'),
  F('notes', 'Footnotes and endnotes', 'word/footnotes.xml', 'preserve'),
  F('review', 'Comments', 'word/comments.xml', 'preserve'),
  F('review', 'Tracked changes', 'word/document.xml', 'preserve', 'preserved; we do not accept or reject'),
  F('objects', 'Images and drawings', 'word/media/*', 'render', 'inline and anchored pictures drawn at their stated extent; EMF/WMF preserved and named, not shown'),
  F('objects', 'Embedded objects (OLE)', 'word/embeddings/*', 'preserve'),
  F('metadata', 'Document properties', 'docProps/*', 'preserve'),
  F('metadata', 'Custom XML parts', 'customXml/*', 'preserve'),
];

export const PRESENTATION_FEATURES = [
  F('package', 'Open, preserve, save unchanged', 'ppt/*', 'preserve', 'no view; a .pptx round trips untouched'),
];

export const ALL_FEATURES = [
  ...SPREADSHEET_FEATURES.map((f) => ({ ...f, kind: 'sheet' })),
  ...DOCUMENT_FEATURES.map((f) => ({ ...f, kind: 'document' })),
  ...PRESENTATION_FEATURES.map((f) => ({ ...f, kind: 'presentation' })),
];

/**
 * Ordered weakest to strongest. Each state IMPLIES every state below it: a
 * feature we can render we can obviously also preserve. Scoring has to respect
 * that or adding a capability makes the numbers go down.
 *
 * `lost` exists and is deliberately unused. It is the state for a feature we
 * discover we damage, and having it means the fidelity number is a measurement
 * rather than a tautology — if something ever lands here, the score drops and
 * the CI gate in tools/fidelity-report.js is what should have caught it first.
 */
export const STATES = ['lost', 'preserve', 'read', 'edit', 'render'];
const RANK = Object.fromEntries(STATES.map((s, i) => [s, i]));

/** Counts per state, per kind, plus the two summary numbers that matter. */
export function summarise(features = ALL_FEATURES) {
  const byKind = {};
  for (const f of features) {
    byKind[f.kind] ??= { lost: 0, preserve: 0, read: 0, edit: 0, render: 0, total: 0 };
    byKind[f.kind][f.state] += 1;
    byKind[f.kind].total += 1;
  }
  const total = features.length;
  const atLeast = (state) => features.filter((f) => RANK[f.state] >= RANK[state]).length;
  const safe = atLeast('preserve');
  const operable = atLeast('read');
  const renderable = atLeast('render');
  return {
    total,
    byKind,
    // Two different questions, deliberately never averaged into one score:
    fidelity: safe / total, // "will we damage this file?"
    functionality: operable / total, // "can we do anything with it?"
    rendering: renderable / total, // "can a user see it?"
  };
}

/** Features whose state is below the one required — the gap list for a goal. */
export function gapsFor(targetState, features = ALL_FEATURES) {
  return features.filter((f) => RANK[f.state] < RANK[targetState]);
}
