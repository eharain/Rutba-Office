// A binary workbook's model, as an .xlsx.
//
// msxls.js reads a .xls of any age into sheets of cells, the fonts, number
// formats and cell formats they use, and the workbook's names; this writes
// them as SpreadsheetML — every cell format as a cell format (its font,
// fill, borders, alignment, number format and protection), every cell with
// its value and, for a formula, the formula and the value Excel last
// calculated, each sheet's column widths, row heights, hidden rows and
// columns, merged areas, frozen panes and window, and the defined names.
// A sheet's pictures, charts, shapes, tables, hyperlinks and notes come too, and the
// workbook's theme when Excel 2007 or later kept it in the file.

import zlib from 'node:zlib';
import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { Workbook, unesc } from '@rutba/ooxml/workbook';
import { readZip } from '@rutba/ooxml/zip';
import { chartXml } from './msxls-chart.js';

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const esc = (s) => String(s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** SUBTOTAL's function numbers, as a totals row names them. */
const TOTALS = { 101: 'average', 102: 'countNums', 103: 'count', 104: 'max', 105: 'min', 107: 'stdDev', 109: 'sum', 110: 'var', 1: 'average', 2: 'countNums', 3: 'count', 4: 'max', 5: 'min', 7: 'stdDev', 9: 'sum', 10: 'var' };
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
  // The skeleton: one sheet part per sheet, each with its pictures' and
  // charts' drawing, and a styles part (a styled cell asks for one); the
  // sheets and styles are written over below.
  const drawings = book.sheets.map((s) => sheetDrawings(s, book));
  const pkg = OoxmlPackage.read(buildXlsx({
    sheets: names.map((name, i) => ({ name, rows: [], styles: i === 0 ? { A1: { bold: true } } : {}, drawings: drawings[i] })),
  }));
  const colour = (index) => {
    if (index == null || index >= 64 || index < 0) return null;
    return book.palette[index] || null;
  };
  // The conditional formats' own formatting, as the styles part's dxfs, each rule pointing at its own.
  const dxfs = [];
  for (const sheet of book.sheets) for (const block of sheet.condFormats || []) for (const rule of block.rules) if (rule.dxf) { rule.dxfId = dxfs.length; dxfs.push(rule.dxf); }
  const styles = stylesXml(book, colour, dxfs);
  pkg.write_('xl/styles.xml', styles.xml);
  book.sheets.forEach((sheet, i) => pkg.write_(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(sheet, i === book.activeSheet, drawings[i].length > 0, styles.runPr)));
  pkg.write_('xl/workbook.xml', workbookXml(book, names));
  // The theme Excel 2007 and later keep, zipped, in the file: the colours a chart's theme colours are.
  const theme = themeXml(book.theme);
  if (theme) {
    pkg.addPart('xl/theme/theme1.xml', theme, 'application/vnd.openxmlformats-officedocument.theme+xml');
    pkg.addRelationshipTo('xl/workbook.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', 'theme/theme1.xml');
  }
  let bytes = pkg.write();
  // Links and notes, as the workbook engine writes them: a link's relationship, a note's comments part and the box Excel draws it in.
  if (book.sheets.some((s) => s.links?.length || s.notes?.length || s.tables?.length)) {
    const wb = Workbook.open(Buffer.from(bytes));
    book.sheets.forEach((s, i) => {
      for (const l of s.links || []) {
        try { wb.setHyperlink(names[i], rangeRef(l.range), { href: l.href || null, location: l.location || null, tooltip: l.tooltip || null, display: l.display || null }); } catch { /* a link this cannot write is left out */ }
      }
      for (const n of s.notes || []) {
        try { wb.setComment(names[i], ref(n.row, n.col), { author: n.author || '', text: n.text || '' }); } catch { /* so is a note */ }
      }
      // Its tables: their range, style and name, the header row's words their columns', and a totals row's labels and functions.
      for (const t of s.tables || []) {
        try {
          const at = (r, c) => s.cells.get(r * 0x4000 + c);
          const columns = [];
          for (let c = t.ref.left; c <= t.ref.right; c++) columns.push(t.header ? String(at(t.ref.top, c)?.v ?? '') : '');
          const totals = [];
          if (t.totals) {
            for (let c = t.ref.left; c <= t.ref.right; c++) {
              const cell = at(t.ref.bottom, c);
              const code = /^SUBTOTAL\((\d+),/i.exec(cell?.formula || '')?.[1];
              totals.push(code ? { fn: TOTALS[code] } : cell && cell.t === 's' && !cell.formula ? { label: String(cell.v) } : null);
            }
          }
          wb.addTable(names[i], rangeRef(t.ref), {
            name: t.name, style: t.style || 'TableStyleMedium2', stripes: t.stripes, headerNames: columns, totalsRow: Boolean(t.totals), totals,
            firstColumn: t.firstColumn, lastColumn: t.lastColumn, columnStripes: t.columnStripes,
          });
        } catch { /* a table this cannot write leaves its cells */ }
      }
    });
    bytes = wb.save();
  }
  return bytes;
}

/** The theme part inside the zip a THEME record holds, or null. */
function themeXml(zip) {
  if (!zip || zip[0] !== 0x50 || zip[1] !== 0x4b) return null;
  try {
    const entry = readZip(Buffer.from(zip)).entries.find((e) => /(^|\/)theme\d*\.xml$/.test(e.name) && !/_rels/.test(e.name));
    const xml = entry ? entry.data.toString('utf8') : null;
    return xml && /<a:theme\b/.test(xml) ? xml : null;
  } catch {
    return null;
  }
}

/**
 * A chart's range, as its cells' values: "Sales!$B$2:$B$13", a sheet's name
 * quoted when it has to be, several areas in brackets. Null for a range on
 * no sheet of this workbook.
 */
function rangeValues(book, text) {
  const colNumber = (letters) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const out = [];
  const areas = String(text).replace(/^\((.*)\)$/, '$1').split(/,(?=(?:[^']*'[^']*')*[^']*$)/);
  for (const area of areas) {
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/.exec(area.trim());
    if (!m) return null;
    const name = m[1] != null ? m[1].replace(/''/g, "'") : m[2];
    const sheet = book.sheets.find((x) => x.name === name);
    if (!sheet) return null;
    const c1 = colNumber(m[3]);
    const r1 = Number(m[4]) - 1;
    const c2 = m[5] ? colNumber(m[5]) : c1;
    const r2 = m[6] ? Number(m[6]) - 1 : r1;
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
      for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
        const cell = sheet.cells.get(r * 0x4000 + c);
        out.push(!cell || cell.t === 'e' || cell.v == null ? null : cell.t === 'b' ? Boolean(cell.v) : cell.v);
      }
    }
  }
  return out;
}

/**
 * A sheet's pictures and charts as the workbook builder places them: from
 * the cell a drawing's top-left corner is in, so far into it — a picture at
 * the size its two corners give, a chart to the cell its other corner is
 * in — measured in the sheet's own column widths and row heights. A chart
 * sheet's chart fills a page's worth of cells.
 */
function sheetDrawings(sheet, book) {
  const kinds = { png: 'png', jpeg: 'jpeg', jpg: 'jpeg', gif: 'gif', bmp: 'bmp', emf: 'emf', wmf: 'wmf' };
  const colPx = (c) => {
    const entry = sheet.cols.find((x) => c >= x.first && c <= x.last);
    if (entry) return entry.hidden ? 0 : (entry.width / 256) * 7;
    return sheet.standardWidth ? (sheet.standardWidth / 256) * 7 : 64;
  };
  const rowPx = (r) => {
    const info = sheet.rows.get(r);
    if (info?.hidden) return 0;
    return (info?.height || info?.fitted || sheet.defaultRowHeight || 300) / 15;
  };
  // The sheet in pixels: where a column or row starts, a cell marker's point, and the marker at a point.
  const lefts = [0];
  const tops = [0];
  const left = (c) => { for (let i = lefts.length; i <= c; i++) lefts[i] = lefts[i - 1] + colPx(i - 1); return lefts[c]; };
  const top = (r) => { for (let i = tops.length; i <= r; i++) tops[i] = tops[i - 1] + rowPx(i - 1); return tops[r]; };
  const pointOf = (m) => ({ x: left(m.col) + (m.dx || 0) * colPx(m.col), y: top(m.row) + (m.dy || 0) * rowPx(m.row) });
  const markerAt = (x, y) => {
    let col = 0;
    while (col < 16383 && left(col + 1) <= x) col += 1;
    let row = 0;
    while (row < 1048575 && top(row + 1) <= y) row += 1;
    return { col, row, colOff: Math.max(0, Math.round((x - left(col)) * 9525)), rowOff: Math.max(0, Math.round((y - top(row)) * 9525)) };
  };
  const marker = (m) => ({ col: m.col, row: m.row, colOff: Math.round((m.dx || 0) * colPx(m.col) * 9525), rowOff: Math.round((m.dy || 0) * rowPx(m.row) * 9525) });
  /** A drawing's box in pixels: from its cells, or from its group's box through the group's coordinates. */
  const rectOf = (item) => {
    if (item.from && item.to) {
      const a = pointOf(item.from);
      const b = pointOf(item.to);
      return { x: a.x, y: a.y, w: Math.max(0, b.x - a.x), h: Math.max(0, b.y - a.y) };
    }
    const place = item.place;
    if (!place) return null;
    if (place.anchor) return rectOf(place.anchor);
    const frame = rectOf(place.frame.place.anchor ? place.frame.place.anchor : { place: place.frame.place });
    if (!frame) return null;
    const c = place.frame.coords;
    const sx = frame.w / (c.w || 1);
    const sy = frame.h / (c.h || 1);
    return { x: frame.x + (place.box.x - c.x) * sx, y: frame.y + (place.box.y - c.y) * sy, w: place.box.w * sx, h: place.box.h * sy };
  };
  const pictures = (sheet.pictures || []).filter((p) => kinds[p.blip?.ext]).map((p, i) => {
    let bytes = p.blip.bytes;
    if (p.blip.deflated) { try { bytes = zlib.inflateSync(Buffer.from(bytes)); } catch { bytes = Buffer.from(bytes); } }
    const r = rectOf(p);
    if (!r) return null;
    return {
      kind: 'picture', name: p.name || `Picture ${i + 1}`, bytes: Buffer.from(bytes), extension: kinds[p.blip.ext], order: p.order ?? i,
      from: p.from ? marker(p.from) : markerAt(r.x, r.y), to: p.to ? marker(p.to) : markerAt(r.x + r.w, r.y + r.h),
      editAs: p.from ? p.edit || 'twoCell' : 'oneCell', widthPx: Math.max(1, r.w), heightPx: Math.max(1, r.h),
    };
  }).filter(Boolean);
  const page = { from: { col: 0, row: 0, dx: 0, dy: 0 }, to: { col: 14, row: 32, dx: 0, dy: 0 } };
  const charts = (sheet.charts || []).map((c, i) => {
    let xml = null;
    try { xml = c.chart ? chartXml(c.chart, (r) => rangeValues(book, r)) : null; } catch { xml = null; }
    if (!xml) return null;
    const at = c.anchor ?? page;
    return { kind: 'chart', name: c.name || `Chart ${i + 1}`, chartXml: xml, from: marker(at.from), to: marker(at.to), order: c.order ?? 1e6 + i };
  }).filter(Boolean);
  // A shape turned a quarter or so has its box kept turned with it: its own box is the other way round about the same middle.
  const shapes = (sheet.shapes || []).map((sh, i) => {
    let r = rectOf(sh);
    if (!r) return null;
    const turn = ((sh.rotation % 360) + 360) % 360;
    if ((turn >= 45 && turn < 135) || (turn >= 225 && turn < 315)) r = { x: r.x + r.w / 2 - r.h / 2, y: r.y + r.h / 2 - r.w / 2, w: r.h, h: r.w };
    // As Excel 2007 or later drew it, in its own cells, where the workbook kept that and it is still this shape.
    const kept = keptShape(sh);
    return {
      kind: 'raw', order: sh.order ?? 2e6 + i, from: kept?.from ?? markerAt(r.x, r.y), to: kept?.to ?? markerAt(r.x + r.w, r.y + r.h), editAs: sh.edit || undefined,
      contentXml: (id) => (kept ? keptShapeXml(kept, id, r) : drawnShapeXml(sh, id, r)),
    };
  }).filter(Boolean);
  // In the order they were drawn, the later over the earlier.
  return [...pictures, ...charts, ...shapes].sort((a, b) => a.order - b.order).map(({ order, ...d }) => d);
}

/**
 * A shape as Excel 2007 or later drew it, which they keep in a 97-2003
 * workbook beside its older description (Office Art's metroBlob, a package
 * holding the shape's DrawingML and the cells it was in): its element, and
 * its cells when it had its own. Only while it is still this shape — an
 * older Excel that changed it since rewrote only the older description, so
 * a shape whose cells or words are not those the package says is read from
 * that. A shape that would need the package's own relationships is too.
 */
function keptShape(sh) {
  if (!sh.drawingML || sh.drawingML[0] !== 0x50 || sh.drawingML[1] !== 0x4b) return null;
  let entries;
  try { entries = readZip(Buffer.from(sh.drawingML)).entries; } catch { return null; }
  const part = (name) => { const e = entries.find((x) => x.name === name); return e ? Buffer.from(e.data).toString('utf8') : null; };
  const xml = part('drs/shapexml.xml') ?? part('drs/connectorxml.xml');
  const el = xml && /<xdr:(sp|cxnSp)\b[\s\S]*<\/xdr:\1>/.exec(xml)?.[0];
  if (!el || /\sr:(id|embed|link|pict)=/.test(el)) return null;
  // The cells its older description had when this was kept, against those it has.
  const was = /<a:from row="(\d+)" col="(\d+)"[^>]*\/>\s*<a:to row="(\d+)" col="(\d+)"/.exec(part('drs/downrev.xml') || '');
  if (sh.from && was && was.slice(1).join() !== [sh.from.row, sh.from.col, sh.to.row, sh.to.col].join()) return null;
  // Its words, against the text object's.
  const words = (el.match(/<a:p\b[\s\S]*?<\/a:p>/g) || []).map((p) => [...p.matchAll(/<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>|<a:br\b/g)].map((m) => (m[1] == null ? '\n' : unesc(m[1]))).join('')).join('\n');
  const own = (sh.paragraphs || []).map((p) => p.runs.map((r) => r.text).join('')).join('\n');
  if (words.trimEnd() !== own.trimEnd()) return null;
  const m = /<xdr:from><xdr:col>(\d+)<\/xdr:col><xdr:colOff>(-?\d+)<\/xdr:colOff><xdr:row>(\d+)<\/xdr:row><xdr:rowOff>(-?\d+)<\/xdr:rowOff><\/xdr:from><xdr:to><xdr:col>(\d+)<\/xdr:col><xdr:colOff>(-?\d+)<\/xdr:colOff><xdr:row>(\d+)<\/xdr:row><xdr:rowOff>(-?\d+)<\/xdr:rowOff><\/xdr:to>/.exec(xml);
  const at = (i) => ({ col: Number(m[i]), colOff: Math.max(0, Number(m[i + 1])), row: Number(m[i + 2]), rowOff: Math.max(0, Number(m[i + 3])) });
  return m && sh.from ? { xml: el, own: true, from: at(1), to: at(5) } : { xml: el, own: false };
}

/** A kept shape given this drawing's id; one placed in a group, now on its own, its box where the group put it. */
function keptShapeXml(kept, id, r) {
  const emu = (px) => Math.round(px * 9525);
  let xml = kept.xml.replace(/(<xdr:cNvPr\b[^>]*?\bid=")\d+"/, (_, head) => `${head}${id}"`);
  if (!kept.own) xml = xml.replace(/(<a:xfrm\b[^>]*>)\s*<a:off [^>]*\/>\s*<a:ext [^>]*\/>/, (_, head) => `${head}<a:off x="${emu(r.x)}" y="${emu(r.y)}"/><a:ext cx="${emu(r.w)}" cy="${emu(r.h)}"/>`);
  return xml;
}

/** A drawn shape as DrawingML: a connector, or a shape with its outline, fill, line, shadow and words. */
function drawnShapeXml(sh, id, r) {
  const emu = (px) => Math.round(px * 9525);
  const hex = (c) => String(typeof c === 'object' ? c.color : c).replace('#', '').toUpperCase();
  const colourXml = (c) => `<a:srgbClr val="${hex(c)}">${typeof c === 'object' && c.alpha != null && c.alpha < 1 ? `<a:alpha val="${Math.round(c.alpha * 100000)}"/>` : ''}</a:srgbClr>`;
  const fillXml = (fill) => {
    if (fill == null) return '';
    if (fill === 'none') return '<a:noFill/>';
    if (fill.gradient) {
      const stops = fill.gradient.stops.map((st) => `<a:gs pos="${Math.round(st.pos * 100000)}">${colourXml(st)}</a:gs>`).join('');
      return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:lin ang="${Math.round(fill.gradient.angle * 60000)}" scaled="1"/></a:gradFill>`;
    }
    return `<a:solidFill>${colourXml(fill)}</a:solidFill>`;
  };
  const lineXml = (line) => (line === 'none' ? '<a:ln><a:noFill/></a:ln>' : line ? `<a:ln w="${Math.round(line.width * 12700)}"><a:solidFill>${colourXml(line.color)}</a:solidFill></a:ln>` : '');
  const turn = sh.rotation ? ` rot="${Math.round(sh.rotation * 60000)}"` : '';
  const flips = (sh.flipH ? ' flipH="1"' : '') + (sh.flipV ? ' flipV="1"' : '');
  const xfrm = `<a:xfrm${turn}${flips}><a:off x="${emu(r.x)}" y="${emu(r.y)}"/><a:ext cx="${emu(r.w)}" cy="${emu(r.h)}"/></a:xfrm>`;
  const name = esc(sh.name || (sh.textBox ? 'TextBox ' : sh.connector ? 'Connector ' : 'Shape ') + id);
  if (sh.connector) {
    return `<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>`
      + `<xdr:spPr>${xfrm}<a:prstGeom prst="${sh.preset}"><a:avLst/></a:prstGeom>${lineXml(sh.line)}</xdr:spPr></xdr:cxnSp>`;
  }
  let geom = `<a:prstGeom prst="${sh.preset || 'rect'}"><a:avLst/></a:prstGeom>`;
  if (sh.path) {
    const pt = (p) => `<a:pt x="${Math.round(p[0])}" y="${Math.round(p[1])}"/>`;
    const cmds = sh.path.commands.map((c) => (c.op === 'M' ? `<a:moveTo>${pt(c.pts[0])}</a:moveTo>` : c.op === 'L' ? `<a:lnTo>${pt(c.pts[0])}</a:lnTo>` : c.op === 'C' ? `<a:cubicBezTo>${c.pts.map(pt).join('')}</a:cubicBezTo>` : '<a:close/>')).join('');
    geom = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="100000" h="100000"${sh.path.filled ? '' : ' fill="none"'}>${cmds}</a:path></a:pathLst></a:custGeom>`;
  }
  const shadow = sh.shadow ? `<a:effectLst><a:outerShdw blurRad="${Math.round(sh.shadow.blur * 12700)}" dist="${Math.round(sh.shadow.dist * 12700)}" dir="${Math.round(sh.shadow.dir * 60000)}" algn="ctr" rotWithShape="0">${colourXml({ color: sh.shadow.color, alpha: sh.shadow.alpha })}</a:outerShdw></a:effectLst>` : '';
  const runXml = (run) => {
    const attrs = ` lang="en-US"${run.size ? ` sz="${Math.round(run.size * 100)}"` : ''}${run.bold ? ' b="1"' : ''}${run.italic ? ' i="1"' : ''}${run.underline ? ' u="sng"' : ''}${run.strike ? ' strike="sngStrike"' : ''}${run.script === 1 ? ' baseline="30000"' : run.script === 2 ? ' baseline="-25000"' : ''}`;
    const inner = (run.colour ? `<a:solidFill><a:srgbClr val="${run.colour}"/></a:solidFill>` : '') + (run.font ? `<a:latin typeface="${esc(run.font)}"/>` : '');
    return `<a:r><a:rPr${attrs}${inner ? `>${inner}</a:rPr>` : '/>'}<a:t>${esc(run.text)}</a:t></a:r>`;
  };
  const [lIns, tIns, rIns, bIns] = sh.insets || [91440, 45720, 91440, 45720];
  const paragraphs = sh.paragraphs?.length
    ? sh.paragraphs.map((p) => `<a:p><a:pPr algn="${sh.align || 'l'}"/>${p.runs.map(runXml).join('')}</a:p>`).join('')
    : '<a:p><a:endParaRPr lang="en-US"/></a:p>';
  const body = `<xdr:txBody><a:bodyPr vertOverflow="clip" wrap="square" lIns="${lIns}" tIns="${tIns}" rIns="${rIns}" bIns="${bIns}" rtlCol="0" anchor="${sh.anchor || 't'}"/><a:lstStyle/>${paragraphs}</xdr:txBody>`;
  return `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvSpPr${sh.textBox ? ' txBox="1"' : ''}/></xdr:nvSpPr>`
    + `<xdr:spPr>${xfrm}${geom}${fillXml(sh.fill)}${lineXml(sh.line)}${shadow}</xdr:spPr>${body}</xdr:sp>`;
}

/** A conditional format's formatting: its font's colour, bold, italic and strike, and its fill. */
function dxfXml(x, colour) {
  // Excel 2007's exact colours where it kept them (CFEX), else the palette's.
  const exact = x.exact || {};
  const paletteEl = (name, index) => { const c = index != null ? colour(index) : null; return c ? `<${name} rgb="FF${c}"/>` : ''; };
  const text = exact.text ? colourEl('color', exact.text) : paletteEl('color', x.colour);
  const font = x.bold != null || x.italic != null || x.strike || text ? '<font>' + (x.bold ? '<b/>' : '') + (x.italic ? '<i/>' : '') + (x.strike ? '<strike/>' : '') + text + '</font>' : '';
  const fg = exact.fillFg ? colourEl('fgColor', exact.fillFg) : paletteEl('fgColor', x.fg);
  // A rule's solid fill shows its background colour, as Excel writes a dxf's.
  const bg = exact.fillBg ? colourEl('bgColor', exact.fillBg) : x.bg != null ? paletteEl('bgColor', x.bg) : fg.replace('<fgColor', '<bgColor');
  const fill = fg || bg ? `<fill><patternFill${x.pattern && x.pattern !== 1 && PATTERNS[x.pattern] ? ` patternType="${PATTERNS[x.pattern]}"` : ''}>${fg}${bg}</patternFill></fill>` : '';
  return `<dxf>${font}${fill}</dxf>`;
}

/** A sheet's conditional formats, each block's ranges and rules in the order the schema wants. */
function condFormatsXml(sheet) {
  let priority = 0;
  const cfvo = (t) => `<cfvo type="${t.type}"${t.formula != null ? ` val="${esc(t.formula)}"` : t.value != null ? ` val="${t.value}"` : ''}${t.gte === false ? ' gte="0"' : ''}/>`;
  const colourOf = (c) => (c ? colourEl('color', c) : '<color rgb="FF638EC6"/>');
  return (sheet.condFormats || []).map((block) => {
    const rules = block.rules.map((r) => {
      const p = ` priority="${r.priority || (priority += 1)}"`;
      if (r.type === 'colorScale') return `<cfRule type="colorScale"${p}><colorScale>${r.thresholds.map(cfvo).join('')}${r.colours.map(colourOf).join('')}</colorScale></cfRule>`;
      if (r.type === 'dataBar') return `<cfRule type="dataBar"${p}><dataBar${r.minLength != null && r.minLength !== 10 ? ` minLength="${r.minLength}"` : ''}${r.maxLength != null && r.maxLength !== 90 ? ` maxLength="${r.maxLength}"` : ''}${r.showValue ? '' : ' showValue="0"'}>${r.thresholds.map(cfvo).join('')}${colourOf(r.colours[0])}</dataBar></cfRule>`;
      if (r.type === 'iconSet') return `<cfRule type="iconSet"${p}><iconSet iconSet="${r.iconSet}"${r.showValue ? '' : ' showValue="0"'}${r.reverse ? ' reverse="1"' : ''}>${r.thresholds.map(cfvo).join('')}</iconSet></cfRule>`;
      const formulas = (r.formulas || []).map((x) => `<formula>${esc(x)}</formula>`).join('');
      if (!formulas) return '';
      const dxf = r.dxfId != null ? ` dxfId="${r.dxfId}"` : '';
      return r.type === 'expression' ? `<cfRule type="expression"${dxf}${p}>${formulas}</cfRule>` : `<cfRule type="cellIs"${dxf}${p}${r.operator ? ` operator="${r.operator}"` : ''}>${formulas}</cfRule>`;
    }).join('');
    return rules ? `<conditionalFormatting sqref="${block.ranges.map(rangeRef).join(' ')}">${rules}</conditionalFormatting>` : '';
  }).join('');
}

/** A colour element for an exact colour: its RGB, its theme colour and tint, or its palette index. */
function colourEl(name, c) {
  const tint = c.tint ? ` tint="${Math.round(c.tint * 1e6) / 1e6}"` : '';
  if (c.rgb) return `<${name} rgb="FF${c.rgb}"${tint}/>`;
  if (c.theme != null) return `<${name} theme="${c.theme}"${tint}/>`;
  return `<${name} indexed="${c.indexed ?? 64}"${tint}/>`;
}

/** The styles part: number formats, fonts, fills, borders and one cell format per XF, in the XFs' order. */
function stylesXml(book, colour, dxfs = []) {
  // Number formats: the workbook's own each given an id of the .xlsx's own; the built-in ones as they are.
  const numFmts = [];
  const fmtId = new Map();
  for (const [id, code] of book.formats) {
    fmtId.set(id, 164 + numFmts.length);
    numFmts.push(`<numFmt numFmtId="${164 + numFmts.length}" formatCode="${esc(code)}"/>`);
  }
  const numFmtOf = (id) => (fmtId.has(id) ? fmtId.get(id) : book.biff >= 5 && id <= 49 ? id : 0);

  const fontXml = (f, exact = null) => {
    const c = f.colour != null && f.colour !== 0x7fff ? colour(f.colour) : null;
    return '<font>' + (f.bold ? '<b/>' : '') + (f.italic ? '<i/>' : '') + (f.strike ? '<strike/>' : '')
      + (UNDERLINE[f.underline] ? `<u${f.underline === 1 ? '' : ` val="${UNDERLINE[f.underline]}"`}/>` : '')
      + (f.script === 1 ? '<vertAlign val="superscript"/>' : f.script === 2 ? '<vertAlign val="subscript"/>' : '')
      + `<sz val="${Math.max(1, f.height / 20)}"/>` + (exact ? colourEl('color', exact) : c ? `<color rgb="FF${c}"/>` : '<color theme="1"/>')
      + `<name val="${esc(f.name || 'Arial')}"/>` + (f.family ? `<family val="${f.family}"/>` : '') + (f.charset ? `<charset val="${f.charset}"/>` : '')
      + '</font>';
  };
  // The first font is the Normal style's, in the colour Excel 2007 kept for that style: a cell format that only
  // says so again has the workbook's default font still, which a table's style may colour over.
  const fonts = book.fonts.length ? book.fonts.map((f, i) => fontXml(f, i === 0 ? book.xfExt?.get(0)?.text : null)) : ['<font><sz val="10"/><name val="Arial"/></font>'];
  // Font 4 is never written: the fourth and later are one out.
  const fontOf = (i) => Math.min(fonts.length - 1, i < 4 ? i : i === 4 ? 0 : i - 1);
  // A font in an XF's exact colour: one more font, once for each font and colour.
  const exactFonts = new Map();
  const fontWith = (i, exact) => {
    const base = book.fonts[fontOf(i)];
    if (!exact || !base) return fontOf(i);
    const xml = fontXml(base, exact);
    if (xml === fonts[fontOf(i)]) return fontOf(i);
    if (!exactFonts.has(xml)) { exactFonts.set(xml, fonts.length); fonts.push(xml); }
    return exactFonts.get(xml);
  };

  const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const fillIds = new Map();
  const fillOf = ({ pattern, fg, bg }, ext = {}) => {
    if (!pattern || !PATTERNS[pattern]) return 0;
    const f = colour(fg);
    const b = colour(bg);
    const fgXml = ext.fillFg ? colourEl('fgColor', ext.fillFg) : f ? `<fgColor rgb="FF${f}"/>` : '<fgColor indexed="64"/>';
    const bgXml = ext.fillBg ? colourEl('bgColor', ext.fillBg) : b ? `<bgColor rgb="FF${b}"/>` : '<bgColor indexed="65"/>';
    const xml = `<fill><patternFill patternType="${PATTERNS[pattern]}">${fgXml}${bgXml}</patternFill></fill>`;
    if (!fillIds.has(xml)) { fillIds.set(xml, fills.length); fills.push(xml); }
    return fillIds.get(xml);
  };

  const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const borderIds = new Map();
  const side = (name, l, exact) => {
    if (!l || !LINES[l.style] || l.style === 0) return `<${name}/>`;
    const c = colour(l.colour);
    return `<${name} style="${LINES[l.style]}">${exact ? colourEl('color', exact) : c ? `<color rgb="FF${c}"/>` : '<color indexed="64"/>'}</${name}>`;
  };
  const borderOf = (b, ext = {}) => {
    const xml = `<border>${side('left', b.left, ext.left)}${side('right', b.right, ext.right)}${side('top', b.top, ext.top)}${side('bottom', b.bottom, ext.bottom)}<diagonal/></border>`;
    if (xml === borders[0]) return 0;
    if (!borderIds.has(xml)) { borderIds.set(xml, borders.length); borders.push(xml); }
    return borderIds.get(xml);
  };

  const xfs = (book.xfs.length ? book.xfs : [{ font: 0, format: 0, h: 0, v: 2, border: {}, fill: {} }]).map((x, i) => {
    // Excel 2007's exact colours where it kept them, over the palette's nearest.
    const ext = book.xfExt?.get(i) || {};
    const numFmtId = numFmtOf(x.format);
    const fontId = fontWith(x.font, ext.text);
    const fillId = fillOf(x.fill || {}, ext);
    const borderId = borderOf(x.border || {}, ext);
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
    + (dxfs.length ? `<dxfs count="${dxfs.length}">${dxfs.map((x) => dxfXml(x, colour)).join('')}</dxfs>` : '')
    + '</styleSheet>';
  // A rich string's run in its font: rPr in the order the schema wants.
  const runPr = (i) => {
    const f = book.fonts[fontOf(i)];
    if (!f) return '';
    const c = f.colour != null && f.colour !== 0x7fff ? colour(f.colour) : null;
    return `<rPr><rFont val="${esc(f.name || 'Arial')}"/>` + (f.charset ? `<charset val="${f.charset}"/>` : '') + (f.family ? `<family val="${f.family}"/>` : '')
      + (f.bold ? '<b/>' : '') + (f.italic ? '<i/>' : '') + (f.strike ? '<strike/>' : '') + (c ? `<color rgb="FF${c}"/>` : '')
      + `<sz val="${Math.max(1, f.height / 20)}"/>` + (UNDERLINE[f.underline] ? `<u${f.underline === 1 ? '' : ` val="${UNDERLINE[f.underline]}"`}/>` : '')
      + (f.script === 1 ? '<vertAlign val="superscript"/>' : f.script === 2 ? '<vertAlign val="subscript"/>' : '') + '</rPr>';
  };
  return { xml, runPr };
}

/** One sheet: its window, column widths, rows of cells, merged areas. */
function sheetXml(sheet, active, drawing = false, runPr = () => '') {
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
  // A row of no cells for what it says of itself; Excel sizes one from its fonts again.
  for (const [r, info] of sheet.rows) if (!byRow.has(r) && Object.keys(info).some((k) => k !== 'fitted')) byRow.set(r, []);
  const rowsXml = [...byRow.keys()].sort((a, b) => a - b).map((r) => {
    const info = sheet.rows.get(r) || {};
    const attrs = [`r="${r + 1}"`];
    if (info.height) attrs.push(`ht="${info.height / 20}"`, 'customHeight="1"');
    if (info.hidden) attrs.push('hidden="1"');
    if (info.level) attrs.push(`outlineLevel="${info.level}"`);
    return `<row ${attrs.join(' ')}>${byRow.get(r).map((c) => cellXml(c, runPr)).join('')}</row>`;
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
    + condFormatsXml(sheet)
    // The drawing that holds its pictures: the skeleton's first relationship.
    + (drawing ? '<drawing r:id="rId1"/>' : '')
    + '</worksheet>';
}

/** A cell: its value, or its formula with the value Excel last calculated. */
function cellXml(c, runPr = () => '') {
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
    case 's': case 'str': {
      const text = String(c.v ?? '');
      if (!c.runs?.length) return `<c r="${r}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(text)}</t></is></c>`;
      // Rich text: the words before the first run in the cell's own font, then each run in its font.
      const cuts = [{ at: 0, font: null }, ...c.runs.filter((x) => x.at < text.length)];
      const runs = cuts.map((x, k) => ({ font: x.font, text: text.slice(x.at, cuts[k + 1]?.at ?? text.length) })).filter((x) => x.text);
      return `<c r="${r}"${s} t="inlineStr"><is>${runs.map((x) => `<r>${x.font == null ? '' : runPr(x.font)}<t xml:space="preserve">${esc(x.text)}</t></r>`).join('')}</is></c>`;
    }
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
