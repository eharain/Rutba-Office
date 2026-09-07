/**
 * Bridge: a real `.xlsx` in, calculated values out, and back again — preserving.
 *
 * This closes the loop that FORMAT-FIDELITY.md listed as the top functional gap.
 * Until now we wrote formulas and asked Excel to recalculate on open
 * (`fullCalcOnLoad`). Now we can evaluate them ourselves, which is what lets a
 * bound sheet be correct *before* anyone opens it — the difference between a
 * report a customer can trust and one whose totals are a day old.
 *
 * The preserving discipline is unchanged. Recalculation writes only cached
 * values (`<v>`) into cells that already hold formulas; it never touches the
 * formula text, the style index, or anything outside `<sheetData>`. A workbook
 * with no formulas comes back byte-identical.
 *
 * Deliberately NOT done here: adding formulas, changing them, or "fixing" ones
 * we cannot parse. A formula we fail to understand keeps its existing cached
 * value and is reported — degrading loudly beats silently writing a wrong
 * number into a file that goes to a bank.
 */
import { Spreadsheet, isError, formatNumber } from '@rutba/formula';
import { Workbook, parseRef, makeRef } from './workbook.js';
import { OoxmlPackage } from './package.js';

/**
 * Load a workbook's cells and defined names into a calculation model.
 * @returns {{sheet: Spreadsheet, cells: Array}}
 */
export function toSpreadsheet(wb, { now } = {}) {
  const sheet = new Spreadsheet({ now });
  const loaded = [];

  for (const { name } of wb.sheets()) {
    sheet.addSheet(name);
    const { part } = wb._sheetPart(name);

    // Pass 1: the ranges covered by array formulas (`<f t="array" ref>` —
    // ours or Excel's). The cells an array covers carry cached VALUES with
    // no formula of their own; loading those as constants would block the
    // spill the anchor is about to re-produce and turn a working file into
    // #SPILL!. They are results, not inputs, so they load as empty and the
    // engine derives them again.
    const covered = [];
    for (const row of part.rows) {
      for (const cellXml of row.inner.match(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g) ?? []) {
        const f = /<f\b([^>]*)>/.exec(cellXml);
        if (!f || !/\bt="array"/.test(f[1])) continue;
        const refAttr = /\bref="([^"]+)"/.exec(f[1]);
        const cellRef = /\br="([A-Z]+\d+)"/.exec(cellXml);
        if (!refAttr || !cellRef) continue;
        const [a, b] = refAttr[1].split(':');
        if (!b) continue; // a 1x1 array covers only its own anchor
        try {
          const p1 = parseRef(a);
          const p2 = parseRef(b);
          const anchor = parseRef(cellRef[1]);
          covered.push({
            top: Math.min(p1.row, p2.row),
            left: Math.min(p1.col, p2.col),
            bottom: Math.max(p1.row, p2.row),
            right: Math.max(p1.col, p2.col),
            anchorRow: anchor.row,
            anchorCol: anchor.col,
          });
        } catch { /* a ref we cannot parse covers nothing */ }
      }
    }
    const isGhost = (r, c) => covered.some((v) => r >= v.top && r <= v.bottom
      && c >= v.left && c <= v.right && !(r === v.anchorRow && c === v.anchorCol));

    for (const row of part.rows) {
      const cells = row.inner.match(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g) ?? [];
      for (const cellXml of cells) {
        const refMatch = /\br="([A-Z]+\d+)"/.exec(cellXml);
        if (!refMatch) continue;
        const { row: r, col: c } = parseRef(refMatch[1]);
        const value = wb.getCell(name, refMatch[1]);
        if (value === null) continue;
        const isFormula = typeof value === 'string' && value.startsWith('=');
        if (!isFormula && covered.length && isGhost(r, c)) continue;
        sheet.setCell(name, r, c, value);
        if (isFormula) {
          loaded.push({ sheet: name, ref: refMatch[1], formula: value });
        }
      }
    }
  }

  for (const { name, ref } of wb.definedNames()) {
    const parsed = parseDefinedNameRange(ref);
    if (parsed) sheet.defineName(name, parsed);
  }

  for (const t of wb.tables()) {
    const [a, b] = String(t.ref).split(':');
    if (!a) continue;
    const start = parseRef(a.replace(/\$/g, ''));
    const end = b ? parseRef(b.replace(/\$/g, '')) : start;
    sheet.defineTable(t.name, {
      sheet: t.sheet,
      top: Math.min(start.row, end.row),
      left: Math.min(start.col, end.col),
      bottom: Math.max(start.row, end.row),
      right: Math.max(start.col, end.col),
      headerRows: t.headerRowCount,
      totalsRows: t.totalsRowCount,
      columns: t.columns,
    });
  }
  return { sheet, formulas: loaded };
}

/** `Stock!$A$3:$F$6` -> { sheet, start, end }. Returns null for anything else. */
export function parseDefinedNameRange(ref) {
  const m = /^(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_.]*))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(String(ref));
  if (!m) return null;
  const sheetName = m[1] ?? m[2];
  const start = parseRef(m[3] + m[4]);
  const end = m[5] ? parseRef(m[5] + m[6]) : start;
  return { sheet: sheetName, start, end };
}

/**
 * Recalculate every formula in a workbook and write the results back as cached
 * values.
 *
 * @param {Buffer} buf
 * @param {object} [opts]
 * @param {() => Date} [opts.now]
 * @returns {{output: Buffer, calculated: number, cycles: string[], errors: Array, unparsed: Array}}
 */
export function recalculateWorkbook(buf, { now } = {}) {
  const wb = Workbook.open(buf);
  const { sheet, formulas } = toSpreadsheet(wb, { now });
  const { calculated, cycles } = sheet.recalculate();

  const errors = [];
  const unparsed = [];

  for (const { sheet: sheetName, ref, formula } of formulas) {
    const { row, col } = parseRef(ref);
    const cell = sheet.cell(sheetName, row, col);
    if (!cell) continue;

    if (cell.error && cell.error.type === '#NAME?' && /unknown function/.test(cell.error.detail ?? '')) {
      // We could not evaluate it, so we leave whatever Excel last cached rather
      // than replacing a good number with our failure.
      unparsed.push({ sheet: sheetName, ref, formula, reason: cell.error.detail });
      continue;
    }
    if (isError(cell.value)) {
      errors.push({ sheet: sheetName, ref, formula, error: cell.value.type, detail: cell.value.detail });
    }
    writeCachedValue(wb, sheetName, ref, formula, cell.value);
  }

  return { output: wb.save(), calculated, cycles, errors, unparsed, sheet };
}

/**
 * Replace a formula cell's cached `<v>` while keeping `<f>`, `r`, `s` and `t`.
 *
 * Done as a surgical rewrite of the one cell rather than through `setCell`,
 * because `setCell` would replace the formula with the value — the opposite of
 * what a recalculation means.
 */
export function writeCachedValue(wb, sheetName, ref, formula, value, { arrayRef = null } = {}) {
  const { part } = wb._sheetPart(sheetName);
  const { row, col } = parseRef(ref);
  const rowRec = part.rows.find((r) => r.index === row);
  if (!rowRec) return false;

  const cellRe = new RegExp('<c\\b[^>]*?\\br="' + ref + '"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)');
  const existing = cellRe.exec(rowRec.inner);
  if (!existing) return false;

  const attrsMatch = /^<c\b([^>]*?)>/.exec(existing[0]) ?? /^<c\b([^>]*?)\/>/.exec(existing[0]);
  let attrs = attrsMatch ? attrsMatch[1] : ' r="' + ref + '"';
  const fMatch = /<f\b[^>]*>[\s\S]*?<\/f>|<f\b[^>]*\/>/.exec(existing[0]);
  // A spilling anchor writes the ARRAY form — `t="array"` with the spilled
  // range — which is what makes the file mean the same thing in an Excel
  // without dynamic arrays: a CSE array over the range rather than an
  // implicit intersection of one cell. Rebuilt fresh each time so the ref
  // FOLLOWS the spill as edits grow or shrink it; a cell that is not an
  // anchor keeps whatever `<f>` it already had, verbatim.
  const fXml = arrayRef
    ? '<f t="array" ref="' + arrayRef + '">' + escapeXml(String(formula).replace(/^=/, '')) + '</f>'
    : (fMatch ? fMatch[0] : '<f>' + String(formula).replace(/^=/, '') + '</f>');

  // The cached value's type attribute must agree with what we cached.
  attrs = attrs.replace(/\s+t="[^"]*"/, '');
  let vXml;
  if (isError(value)) {
    attrs += ' t="e"';
    vXml = '<v>' + value.type + '</v>';
  } else if (typeof value === 'boolean') {
    attrs += ' t="b"';
    vXml = '<v>' + (value ? 1 : 0) + '</v>';
  } else if (typeof value === 'number') {
    vXml = '<v>' + formatNumber(value) + '</v>';
  } else if (value === '' || value === null || value === undefined) {
    vXml = '';
  } else {
    attrs += ' t="str"'; // a formula that returns text caches it as str
    vXml = '<v>' + escapeXml(String(value)) + '</v>';
  }

  const rebuilt = '<c' + attrs + '>' + fXml + vXml + '</c>';
  // Idempotent: a cache that already agrees is left exactly as it is, so
  // re-caching every formula on save costs nothing for the ones that did not
  // move and never marks a part modified without cause.
  if (rebuilt === existing[0]) return false;

  rowRec.inner = rowRec.inner.replace(existing[0], rebuilt);
  rowRec.dirty = true;
  part.dirty = true;
  return true;
}

const escapeXml = (s) =>
  String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** Report a workbook's calculation state without writing anything. */
export function inspectCalculation(buf, { now } = {}) {
  const wb = Workbook.open(buf);
  const { sheet, formulas } = toSpreadsheet(wb, { now });
  const { calculated, cycles } = sheet.recalculate();
  const errors = [];
  for (const { sheet: sheetName, ref, formula } of formulas) {
    const { row, col } = parseRef(ref);
    const cell = sheet.cell(sheetName, row, col);
    if (cell && isError(cell.value)) {
      errors.push({ sheet: sheetName, ref, formula, error: cell.value.type, detail: cell.value.detail });
    }
  }
  return { formulaCount: formulas.length, calculated, cycles, errors };
}

export { OoxmlPackage, makeRef };
