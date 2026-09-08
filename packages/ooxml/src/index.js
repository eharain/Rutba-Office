/**
 * @rutba/ooxml — the owned document core.
 *
 * Decision context: the engine choice was reopened on 2026-08-20 because a
 * licensed engine cannot be modified by us (ONLYOFFICE Developer Edition's EULA
 * forbids it outright), and ownership was made a hard requirement. Microsoft
 * format fidelity was simultaneously called critical, because customer files go
 * to banks, suppliers and government.
 *
 * Those two requirements together put OOXML fidelity on the critical path, which
 * is why this package exists before any editor UI does. Its strategy:
 *
 *     Model only what we edit. Preserve everything else byte-identically.
 *
 * See FORMAT-FIDELITY.md for what that buys, what it does not, and how to
 * measure it against a real corpus.
 */
export { OoxmlPackage, OoxmlError, attrs, esc } from './package.js';
export { readZip, writeZip, ZipEntry, crc32 } from './zip.js';
export {
  Workbook, colToIndex, indexToCol, parseRef, makeRef, unesc,
} from './workbook.js';
export {
  Document, textOf, parseRuns, renderRuns, hasToggle, withToggle, WORD_NS,
} from './document.js';
export {
  parseTable, parseSection, twipsToPx, eighthPointsToPx, TWIPS_PER_INCH,
} from './table.js';
export { readHeadersAndFooters, parseBand } from './headers.js';
export { readParagraphStyles, readCharacterStyles, readNumberingDefs, readThemeFonts, readThemeColours, STANDARD_PARAGRAPH_STYLES, STANDARD_STYLES_XML } from './docstyles.js';
export {
  recalculateWorkbook, inspectCalculation, toSpreadsheet, parseDefinedNameRange,
} from './recalc.js';
export {
  fidelityOf, comparePackages, formatReport, preservedButUnsupported, xmlEquivalent,
} from './fidelity.js';
export {
  buildXlsx, buildDocx, buildLetterDocx, colName, cellRef,
  readWorkbook, readSheetValues, readSheetNames, readDefinedNames,
} from './build.js';
