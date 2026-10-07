// A binary workbook's model, as an .xlsx.
//
// msxls.js reads a .xls of any age into sheets of cells, the fonts, number
// formats and cell formats they use, and the workbook's names; this writes
// them as SpreadsheetML — every cell format as a cell format (its font,
// fill, borders, alignment, number format and protection), every cell with
// its value and, for a formula, the formula and the value Excel last
// calculated, each sheet's column widths, row heights, hidden rows and
// columns, merged areas, frozen panes and window, and the defined names.
// A sheet's pictures, hyperlinks and notes come too; its charts do not.

import zlib from 'node:zlib';
import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { Workbook } from '@rutba/ooxml/workbook';

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const esc = (s) => String(s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const H_ALIGN = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
const V_ALIGN = ['top', 'center', 'bottom', 'justify', 'distributed'];
const LINES = ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'];
const PATTERNS = ['none', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625'];
const UNDERLINE = { 1: 'single', 2: 'double', 0x21: 'singleAccounting', 0x22: 'doubleAccounting' };

const colName = (c) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const ref = (row, col) => colName(col) + (row + 1);
const rangeRef = ({ top, left, bottom, right }) => (top === bottom && left === right ? ref(top, left) : `${ref(top, left)}:${ref(bottom, right)}`);

/** A sheet's name as .xlsx allows one: none of []:*?/\, at most 31 characters, and its own. */
function sheetNames(sheets) {
  const used = new Set();
  return sheets.map((s, i) => {
    let base = String(s.name || '').replace(/[[\]:*?/\\]/g, '_').trim().slice(0, 31) || `Sheet${i + 1}`;
    let name = base;
    for (let n = 2; used.has(name.toLowerCase()); n++) name = base.slice(0, 31 - String(n).length - 1) + '_' + n;
    used.add(name.toLowerCase());
    return name;
  });
}

/**
 * Write the model as an .xlsx: returns the file's bytes.
 * @param {object} book what msxls.js's readXls gives
 */
export function xlsModelToXlsx(book) {
  const names = sheetNames(book.sheets);
  // The skeleton: one sheet part per sheet, each with its pictures' drawing,
  // and a styles part (a styled cell asks for one); the sheets and styles
  // are written over below.
  const pkg = OoxmlPackage.read(buildXlsx({
    sheets: names.map((name, i) => ({ name, rows: [], styles: i === 0 ? { A1: { bold: true } } : {}, drawings: pictureDrawings(book.sheets[i]) })),
  }));
  const colour = (index) => {
    if (index == null || index >= 64 || index < 0) return null;
    return book.palette[index] || null;
  };
  const styles = stylesXml(book, colour);
  pkg.write_('xl/styles.xml', styles.xml);
  book.sheets.forEach((sheet, i) => pkg.write_(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(sheet, i === book.activeSheet, pictureDrawings(sheet).length > 0)));
  pkg.write_('xl/workbook.xml', workbookXml(book, names));
  let bytes = pkg.write();
  // Links and notes, as the workbook engine writes them: a link's relationship, a note's comments part and the box Excel draws it in.
  if (book.sheets.some((s) => s.links?.length || s.notes?.length)) {
    const wb = Workbook.open(Buffer.from(bytes));
    book.sheets.forEach((s, i) => {
      for (const l of s.links || []) {
        try { wb.setHyperlink(names[i], rangeRef(l.range), { href: l.href || null, location: l.location || null, tooltip: l.tooltip || null, display: l.display || null }); } catch { /* a link this cannot write is left out */ }
      }
      for (const n of s.notes || []) {
        try { wb.setComment(names[i], ref(n.row, n.col), { author: n.author || '', text: n.text || '' }); } catch { /* so is a note */ }
      }
    });
    bytes = wb.save();
  }
  return bytes;
}

/**
 * A sheet's pictures as the workbook builder places them: from the cell its
 * top-left corner is in, so far into it, at the size its two corners give —
 * measured in the sheet's own column widths and row heights.
 */
function pictureDrawings(sheet) {
  const kinds = { png: 'png', jpeg: 'jpeg', jpg: 'jpeg', gif: 'gif', bmp: 'bmp', emf: 'emf', wmf: 'wmf' };
  const colPx = (c) => {
    const entry = sheet.cols.find((x) => c >= x.first && c <= x.last);
    if (entry) return entry.hidden ? 0 : (entry.width / 256) * 7;
    return sheet.standardWidth ? (sheet.standardWidth / 256) * 7 : 64;
  };
  const rowPx = (r) => {
    const info = sheet.rows.get(r);
    if (info?.hidden) return 0;
    return (info?.height || sheet.defaultRowHeight || 300) / 15;
  };
  const span = (size, a, b) => {
    if (a.at === b.at) return (b.f - a.f) * size(a.at);
    let total = (1 - a.f) * size(a.at) + b.f * size(b.at);
    for (let i = a.at + 1; i < b.at; i++) total += size(i);
    return total;
  };
  return (sheet.pictures || []).filter((p) => kinds[p.blip?.ext]).map((p, i) => {
    let bytes = p.blip.bytes;
    if (p.blip.deflated) { try { bytes = zlib.inflateSync(Buffer.from(bytes)); } catch { bytes = Buffer.from(bytes); } }
    return {
      kind: 'picture', name: p.name || `Picture ${i + 1}`, bytes: Buffer.from(bytes), extension: kinds[p.blip.ext],
      from: { col: p.from.col, row: p.from.row, colOff: Math.round(p.from.dx * colPx(p.from.col) * 9525), rowOff: Math.round(p.from.dy * rowPx(p.from.row) * 9525) },
      widthPx: Math.max(1, span(colPx, { at: p.from.col, f: p.from.dx }, { at: p.to.col, f: p.to.dx })),
      heightPx: Math.max(1, span(rowPx, { at: p.from.row, f: p.from.dy }, { at: p.to.row, f: p.to.dy })),
    };
  });
}

/** The styles part: number formats, fonts, fills, borders and one cell format per XF, in the XFs' order. */
function stylesXml(book, colour) {
  // Number formats: the workbook's own each given an id of the .xlsx's own; the built-in ones as they are.
  const numFmts = [];
  const fmtId = new Map();
  for (const [id, code] of book.formats) {
    fmtId.set(id, 164 + numFmts.length);
    numFmts.push(`<numFmt numFmtId="${164 + numFmts.length}" formatCode="${esc(code)}"/>`);
  }
  const numFmtOf = (id) => (fmtId.has(id) ? fmtId.get(id) : book.biff >= 5 && id <= 49 ? id : 0);

  const fonts = book.fonts.length ? book.fonts.map((f) => {
    const c = f.colour != null && f.colour !== 0x7fff ? colour(f.colour) : null;
    return '<font>' + (f.bold ? '<b/>' : '') + (f.italic ? '<i/>' : '') + (f.strike ? '<strike/>' : '')
      + (UNDERLINE[f.underline] ? `<u${f.underline === 1 ? '' : ` val="${UNDERLINE[f.underline]}"`}/>` : '')
      + (f.script === 1 ? '<vertAlign val="superscript"/>' : f.script === 2 ? '<vertAlign val="subscript"/>' : '')
      + `<sz val="${Math.max(1, f.height / 20)}"/>` + (c ? `<color rgb="FF${c}"/>` : '<color theme="1"/>')
      + `<name val="${esc(f.name || 'Arial')}"/>` + (f.family ? `<family val="${f.family}"/>` : '') + (f.charset ? `<charset val="${f.charset}"/>` : '')
      + '</font>';
  }) : ['<font><sz val="10"/><name val="Arial"/></font>'];
  // Font 4 is never written: the fourth and later are one out.
  const fontOf = (i) => Math.min(fonts.length - 1, i < 4 ? i : i === 4 ? 0 : i - 1);

  const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const fillIds = new Map();
  const fillOf = ({ pattern, fg, bg }) => {
    if (!pattern || !PATTERNS[pattern]) return 0;
    const f = colour(fg);
    const b = colour(bg);
    const xml = `<fill><patternFill patternType="${PATTERNS[pattern]}">${f ? `<fgColor rgb="FF${f}"/>` : '<fgColor indexed="64"/>'}${b ? `<bgColor rgb="FF${b}"/>` : '<bgColor indexed="65"/>'}</patternFill></fill>`;
    if (!fillIds.has(xml)) { fillIds.set(xml, fills.length); fills.push(xml); }
    return fillIds.get(xml);
  };

  const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const borderIds = new Map();
  const side = (name, l) => {
    if (!l || !LINES[l.style] || l.style === 0) return `<${name}/>`;
    const c = colour(l.colour);
    return `<${name} style="${LINES[l.style]}">${c ? `<color rgb="FF${c}"/>` : '<color indexed="64"/>'}</${name}>`;
  };
  const borderOf = (b) => {
    const xml = `<border>${side('left', b.left)}${side('right', b.right)}${side('top', b.top)}${side('bottom', b.bottom)}<diagonal/></border>`;
    if (xml === borders[0]) return 0;
    if (!borderIds.has(xml)) { borderIds.set(xml, borders.length); borders.push(xml); }
    return borderIds.get(xml);
  };

  const xfs = (book.xfs.length ? book.xfs : [{ font: 0, format: 0, h: 0, v: 2, border: {}, fill: {} }]).map((x) => {
    const numFmtId = numFmtOf(x.format);
    const fontId = fontOf(x.font);
    const fillId = fillOf(x.fill || {});
    const borderId = borderOf(x.border || {});
    const al = [];
    if (x.h) al.push(`horizontal="${H_ALIGN[x.h] || 'general'}"`);
    if (x.v !== 2 && V_ALIGN[x.v]) al.push(`vertical="${V_ALIGN[x.v]}"`);
    if (x.wrap) al.push('wrapText="1"');
    if (x.rotation) al.push(`textRotation="${x.rotation}"`);
    if (x.indent) al.push(`indent="${x.indent}"`);
    if (x.shrink) al.push('shrinkToFit="1"');
    if (x.rtl) al.push('readingOrder="2"');
    const prot = !x.locked || x.hidden ? `<protection${x.locked ? '' : ' locked="0"'}${x.hidden ? ' hidden="1"' : ''}/>` : '';
    const inner = (al.length ? `<alignment ${al.join(' ')}/>` : '') + prot;
    return `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"`
      + (numFmtId ? ' applyNumberFormat="1"' : '') + (fontId ? ' applyFont="1"' : '') + (fillId ? ' applyFill="1"' : '')
      + (borderId ? ' applyBorder="1"' : '') + (al.length ? ' applyAlignment="1"' : '') + (prot ? ' applyProtection="1"' : '')
      + (inner ? `>${inner}</xf>` : '/>');
  });

  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
    + (numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.join('')}</numFmts>` : '')
    + `<fonts count="${fonts.length}">${fonts.join('')}</fonts>`
    + `<fills count="${fills.length}">${fills.join('')}</fills>`
    + `<borders count="${borders.length}">${borders.join('')}</borders>`
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>`
    + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    + '</styleSheet>';
  return { xml };
}

/** One sheet: its window, column widths, rows of cells, merged areas. */
function sheetXml(sheet, active, drawing = false) {
  const cells = [...sheet.cells.values()].sort((a, b) => a.row - b.row || a.col - b.col);
  let maxRow = 0;
  let maxCol = 0;
  for (const c of cells) { maxRow = Math.max(maxRow, c.row); maxCol = Math.max(maxCol, c.col); }

  // The window: gridlines, headings, formulas shown, frozen panes.
  const view = [`workbookViewId="0"`];
  if (active) view.push('tabSelected="1"');
  if (sheet.grid === false) view.push('showGridLines="0"');
  if (sheet.headings === false) view.push('showRowColHeaders="0"');
  if (sheet.formulas) view.push('showFormulas="1"');
  let pane = '';
  if (sheet.frozen) {
    const { rows, cols } = sheet.frozen;
    const where = rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight';
    pane = `<pane${cols ? ` xSplit="${cols}"` : ''}${rows ? ` ySplit="${rows}"` : ''} topLeftCell="${ref(rows, cols)}" activePane="${where}" state="frozen"/><selection pane="${where}"/>`;
  }
  const views = `<sheetViews><sheetView ${view.join(' ')}>${pane}</sheetView></sheetViews>`;

  const fmt = [`defaultRowHeight="${sheet.defaultRowHeight ? sheet.defaultRowHeight / 20 : 12.75}"`];
  if (sheet.defaultColWidth) fmt.push(`baseColWidth="${sheet.defaultColWidth}"`);
  if (sheet.standardWidth) fmt.push(`defaultColWidth="${Math.round((sheet.standardWidth / 256) * 100) / 100}"`);

  const cols = sheet.cols
    .filter((c) => c.last >= c.first)
    .sort((a, b) => a.first - b.first)
    .map((c) => {
      const width = Math.round((c.width / 256) * 100) / 100;
      return `<col min="${c.first + 1}" max="${Math.min(16384, c.last + 1)}" width="${width || 0}" customWidth="1"${c.hidden || !width ? ' hidden="1"' : ''}${c.xf ? ` style="${c.xf}"` : ''}${c.level ? ` outlineLevel="${c.level}"` : ''}/>`;
    });

  // Rows, each with its cells; a row with a height or hidden but no cells still written.
  const byRow = new Map();
  for (const c of cells) {
    if (!byRow.has(c.row)) byRow.set(c.row, []);
    byRow.get(c.row).push(c);
  }
  for (const r of sheet.rows.keys()) if (!byRow.has(r)) byRow.set(r, []);
  const rowsXml = [...byRow.keys()].sort((a, b) => a - b).map((r) => {
    const info = sheet.rows.get(r) || {};
    const attrs = [`r="${r + 1}"`];
    if (info.height) attrs.push(`ht="${info.height / 20}"`, 'customHeight="1"');
    if (info.hidden) attrs.push('hidden="1"');
    if (info.level) attrs.push(`outlineLevel="${info.level}"`);
    return `<row ${attrs.join(' ')}>${byRow.get(r).map(cellXml).join('')}</row>`;
  });

  const merges = sheet.merges.length ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${rangeRef(m)}"/>`).join('')}</mergeCells>` : '';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<worksheet ${NS}>`
    + `<dimension ref="${cells.length ? `A1:${ref(maxRow, maxCol)}` : 'A1'}"/>`
    + views
    + `<sheetFormatPr ${fmt.join(' ')}/>`
    + (cols.length ? `<cols>${cols.join('')}</cols>` : '')
    + `<sheetData>${rowsXml.join('')}</sheetData>`
    + merges
    // The drawing that holds its pictures: the skeleton's first relationship.
    + (drawing ? '<drawing r:id="rId1"/>' : '')
    + '</worksheet>';
}

/** A cell: its value, or its formula with the value Excel last calculated. */
function cellXml(c) {
  const r = ref(c.row, c.col);
  const s = c.xf ? ` s="${c.xf}"` : '';
  if (c.formula) {
    const f = c.array ? `<f t="array" ref="${rangeRef(c.array)}">${esc(c.formula)}</f>` : `<f>${esc(c.formula)}</f>`;
    if (c.t === 'str') return `<c r="${r}"${s} t="str">${f}<v>${esc(c.v ?? '')}</v></c>`;
    if (c.t === 'b') return `<c r="${r}"${s} t="b">${f}<v>${c.v ? 1 : 0}</v></c>`;
    if (c.t === 'e') return `<c r="${r}"${s} t="e">${f}<v>${esc(c.v)}</v></c>`;
    return `<c r="${r}"${s}>${f}${Number.isFinite(c.v) ? `<v>${c.v}</v>` : ''}</c>`;
  }
  switch (c.t) {
    case 'n': return Number.isFinite(c.v) ? `<c r="${r}"${s}><v>${c.v}</v></c>` : `<c r="${r}"${s}/>`;
    case 's': case 'str': return `<c r="${r}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(c.v ?? '')}</t></is></c>`;
    case 'b': return `<c r="${r}"${s} t="b"><v>${c.v ? 1 : 0}</v></c>`;
    case 'e': return `<c r="${r}"${s} t="e"><v>${esc(c.v)}</v></c>`;
    default: return s ? `<c r="${r}"${s}/>` : '';
  }
}

/** The workbook part: its sheets (hidden ones hidden), the active one, the 1904 dates, the names. */
function workbookXml(book, names) {
  const sheets = book.sheets.map((s, i) => `<sheet name="${esc(names[i])}" sheetId="${i + 1}"${s.state && s.state !== 'visible' ? ` state="${s.state}"` : ''} r:id="rId${i + 1}"/>`).join('');
  const defined = book.names
    .filter((n) => n.formula && n.formula !== '#REF!')
    .map((n) => `<definedName name="${esc(n.name)}"${n.sheet != null ? ` localSheetId="${n.sheet}"` : ''}${n.hidden || n.name === '_xlnm._FilterDatabase' ? ' hidden="1"' : ''}>${esc(n.formula)}</definedName>`);
  // A name may be given once in each scope.
  const seen = new Set();
  const unique = defined.filter((x) => { const key = /name="([^"]*)"(?: localSheetId="(\d+)")?/.exec(x); const k = `${key[1].toLowerCase()}|${key[2] ?? ''}`; if (seen.has(k)) return false; seen.add(k); return true; });
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<workbook ${NS}>`
    + (book.date1904 ? '<workbookPr date1904="1"/>' : '<workbookPr/>')
    + `<bookViews><workbookView activeTab="${book.activeSheet || 0}"/></bookViews>`
    + `<sheets>${sheets}</sheets>`
    + (unique.length ? `<definedNames>${unique.join('')}</definedNames>` : '')
    + '</workbook>';
}
