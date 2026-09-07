/**
 * @rutba/sheet-view — the renderable model of a sheet.
 *
 * The last of the three engines FORMAT-FIDELITY.md named as missing. Calculation
 * came first because a data sheet that does not compute is not a data sheet;
 * this is the surface a person actually looks at.
 *
 * Deliberately DOM-free. The hard parts — number formats, virtualisation
 * geometry, selection semantics, the display/edit split — are pure logic and
 * tested in Node. The browser shell is a thin renderer over `view.render()`,
 * small enough to be obviously correct by reading it.
 */
export {
  formatValue, readNumberFormats, BUILTIN_FORMATS, splitSections, tokenize, isDateFormat,
} from './numfmt.js';
export {
  readStyles, readMergedCells, readDataValidations, readConditionalFormatting,
  readTheme, applyTint, readColourElement,
  INDEXED_COLOURS, THEME_SLOTS, BORDER_WIDTHS,
} from './styles.js';
export {
  SheetGeometry, charWidthToPixels, pointsToPixels,
  DEFAULT_COL_WIDTH_CHARS, DEFAULT_ROW_HEIGHT_POINTS, MAX_ROWS, MAX_COLS,
} from './geometry.js';
export { Selection, colName, ref } from './selection.js';
export { SheetView, coerceInput } from './view.js';
