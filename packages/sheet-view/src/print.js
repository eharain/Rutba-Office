/**
 * A worksheet on paper.
 *
 * A spreadsheet is one endless plane and a page is a fixed rectangle, so
 * printing a sheet is not "draw what is on screen, larger": it is deciding
 * where the plane is cut. Excel's rules, which people know without being able
 * to state them, are the ones followed here:
 *
 *   - Only the used range is printed, or the print area if one is set.
 *   - Columns fill a page left to right until the next one will not fit, then
 *     the whole block of remaining columns goes to a later set of pages —
 *     never a column split down its middle.
 *   - Rows fill a page top to bottom the same way, and the run of pages goes
 *     DOWN first and then across, because a person reading a wide report wants
 *     page 2 to be the rest of the same rows.
 *   - "Fit to width" scales everything by one factor so the widest block fits,
 *     rather than shrinking columns unevenly.
 *   - Rows repeated at the top (print titles) are drawn on every page.
 *
 * The output is HTML, because the thing that turns a page into paper or into a
 * PDF is Chromium, which already knows how to break pages, embed fonts, and
 * talk to a printer. What this file owns is the layout decision; what the
 * browser owns is the ink. Both halves are testable: this one in node, against
 * the strings it produces, and the other by looking at the paper.
 */

import { ref } from './selection.js';

/** Paper sizes in millimetres, portrait. */
export const PAPER = {
  A3: { width: 297, height: 420, css: 'A3' },
  A4: { width: 210, height: 297, css: 'A4' },
  A5: { width: 148, height: 210, css: 'A5' },
  Letter: { width: 215.9, height: 279.4, css: 'Letter' },
  Legal: { width: 215.9, height: 355.6, css: 'Legal' },
  Tabloid: { width: 279.4, height: 431.8, css: 'Tabloid' },
};

/** CSS pixels per millimetre at the 96 dpi every browser lays out in. */
export const PX_PER_MM = 96 / 25.4;

/**
 * What a page setup says. Margins are millimetres, as every print dialog in
 * the world states them; everything else is a name.
 */
export const DEFAULT_PAGE_SETUP = {
  paper: 'A4',
  orientation: 'portrait',
  margins: { top: 12.7, right: 12.7, bottom: 12.7, left: 12.7 },
  /** 'none' keeps the sheet's own size; 'width' fits all columns across; 'page' fits everything on one. */
  fit: 'none',
  scale: 1,
  gridlines: false,
  headings: false,
  /** Rows repeated at the top of every page, as a count from row 1. */
  repeatRows: 0,
  /** An A1:D40 range to print instead of the used range. */
  area: null,
  /** Centre the block on the page, as Excel's "center on page" does. */
  centre: { horizontal: false, vertical: false },
  header: null,
  footer: '&P of &N',
  /** Down first, then across — Excel's default order. */
  order: 'down',
};

export function pageSetup(options = {}) {
  return {
    ...DEFAULT_PAGE_SETUP,
    ...options,
    margins: { ...DEFAULT_PAGE_SETUP.margins, ...(options.margins || {}) },
    centre: { ...DEFAULT_PAGE_SETUP.centre, ...(options.centre || {}) },
  };
}

/** The printable rectangle in CSS pixels, orientation applied. */
export function printableArea(setup) {
  const paper = PAPER[setup.paper] || PAPER.A4;
  const landscape = setup.orientation === 'landscape';
  const width = (landscape ? paper.height : paper.width) - setup.margins.left - setup.margins.right;
  const height = (landscape ? paper.width : paper.height) - setup.margins.top - setup.margins.bottom;
  return { width: Math.max(1, width * PX_PER_MM), height: Math.max(1, height * PX_PER_MM), paper, landscape };
}

/** A1:D40 -> { top, left, bottom, right }, or null. */
function parseArea(area) {
  if (!area) return null;
  const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(String(area).trim());
  if (!m) return null;
  const col = (letters) => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
  const top = Number(m[2]) - 1;
  const left = col(m[1]);
  return {
    top,
    left,
    bottom: m[4] ? Number(m[4]) - 1 : top,
    right: m[3] ? col(m[3]) : left,
  };
}

/**
 * Where the cuts fall.
 *
 * Answers the blocks of columns and rows each page holds, and the one scale
 * everything is drawn at. Nothing here reads a cell: it is arithmetic over the
 * geometry, which is what makes it cheap on a sheet of a million rows and
 * testable without a workbook.
 */
/**
 * The room a running header and footer take out of the page.
 *
 * Counted in the plan and reserved in the CSS, because the two must agree: a
 * page that is planned as full and then drawn with a footer under it overflows
 * by exactly one line, and the browser starts a new page for it. Three pages
 * were planned and five came out of the printer.
 */
export function runningHeight(setup) {
  const lines = (setup.header ? 1 : 0) + (setup.footer ? 1 : 0);
  return lines * 22; // 9pt of type with its padding, in CSS pixels
}

export function paginate({ geo, range, setup, maxPages = 2000 }) {
  const full = printableArea(setup);
  const area = { ...full, height: Math.max(1, full.height - runningHeight(setup)) };
  const cols = [];
  for (let c = range.left; c <= range.right; c++) cols.push({ index: c, size: geo.colWidth(c) });

  // Print titles: the first rows of the range, drawn again at the top of every
  // page and left out of the flow, so page one does not carry its heading
  // twice and page four carries it at all. Capped, because a title band taller
  // than the page leaves no room for the report.
  const titleRows = [];
  const repeat = Math.max(0, Math.min(Number(setup.repeatRows) || 0, 20));
  for (let r = range.top; r < range.top + repeat && r <= range.bottom; r++) {
    titleRows.push({ index: r, size: geo.rowHeight(r) });
  }
  const rows = [];
  for (let r = range.top + titleRows.length; r <= range.bottom; r++) rows.push({ index: r, size: geo.rowHeight(r) });
  const titleHeight = titleRows.reduce((n, r) => n + r.size, 0);

  const total = {
    width: cols.reduce((n, c) => n + c.size, 0),
    height: rows.reduce((n, r) => n + r.size, 0) + titleHeight,
  };

  // One scale for the whole sheet, as Excel does it: shrinking columns
  // unevenly to fit would change the shape of the report.
  let scale = Number(setup.scale) || 1;
  if (setup.fit === 'width' && total.width > 0) scale = Math.min(scale, area.width / total.width);
  else if (setup.fit === 'page' && total.width > 0 && total.height > 0) {
    scale = Math.min(scale, area.width / total.width, area.height / total.height);
  }
  scale = Math.max(0.1, Math.min(1, scale));

  const fit = (items, room) => {
    const bands = [];
    let band = [];
    let used = 0;
    // Half a pixel of slack. Fitting to width makes the scaled total exactly
    // the page width, and floating-point arithmetic then puts the last column
    // a fraction over it — which spilled one column onto a second page and
    // made "fit to one page" two.
    const limit = room + 0.5;
    for (const item of items) {
      const size = item.size * scale;
      // A single item taller or wider than the page still gets its own page:
      // an empty page helps nobody.
      if (band.length && used + size > limit) {
        bands.push(band);
        band = [];
        used = 0;
      }
      band.push(item);
      used += size;
    }
    if (band.length) bands.push(band);
    return bands.length ? bands : [[]];
  };

  const colBands = fit(cols, area.width);
  const rowBands = fit(rows, Math.max(1, area.height - titleHeight * scale));

  const pages = [];
  const down = setup.order !== 'across';
  const outer = down ? colBands : rowBands;
  const inner = down ? rowBands : colBands;
  for (let o = 0; o < outer.length; o++) {
    for (let i = 0; i < inner.length; i++) {
      if (pages.length >= maxPages) return { pages, scale, area, truncated: true, total };
      pages.push({
        cols: down ? colBands[o] : colBands[i],
        rows: down ? rowBands[i] : rowBands[o],
        titleRows,
      });
    }
  }
  return { pages, scale, area, truncated: false, total };
}

const esc = (s) =>
  String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

/** A cell's appearance as inline CSS — the same properties the grid paints. */
function cellCss(cell, setup) {
  const css = [];
  const font = cell.style?.font;
  if (font?.bold) css.push('font-weight:700');
  if (font?.italic) css.push('font-style:italic');
  if (font?.underline) css.push('text-decoration:underline');
  if (font?.strike) css.push('text-decoration:line-through');
  if (font?.name) css.push(`font-family:${JSON.stringify(font.name)},Calibri,Arial,sans-serif`);
  if (font?.size) css.push(`font-size:${(Number(font.size) * 4) / 3}px`);
  const colour = cell.colour || font?.colour;
  if (colour) css.push(`color:${colour}`);
  if (cell.style?.fill?.colour) css.push(`background:${cell.style.fill.colour}`);
  if (cell.align) css.push(`text-align:${cell.align}`);
  if (cell.valign) css.push(`vertical-align:${cell.valign === 'center' ? 'middle' : cell.valign}`);
  if (cell.wrap) css.push('white-space:pre-wrap');
  if (cell.indent) css.push(`padding-left:${3 + cell.indent * 9}px`);
  for (const [side, key] of [['top', 'top'], ['right', 'right'], ['bottom', 'bottom'], ['left', 'left']]) {
    const b = cell.style?.border?.[key];
    if (b?.style && b.style !== 'none') css.push(`border-${side}:${b.width || 1}px ${b.style === 'double' ? 'double' : b.style === 'dashed' ? 'dashed' : b.style === 'dotted' ? 'dotted' : 'solid'} ${b.colour || '#000'}`);
  }
  if (!cell.style?.border && setup.gridlines) css.push('border:1px solid #d7d7d7');
  return css.join(';');
}

/** &P, &N, &A, &D, &F — the header and footer codes people already know. */
function fields(text, { page, pages, sheet, file, date }) {
  return esc(String(text ?? ''))
    .replace(/&amp;P/g, String(page))
    .replace(/&amp;N/g, String(pages))
    .replace(/&amp;A/g, esc(sheet))
    .replace(/&amp;F/g, esc(file || ''))
    .replace(/&amp;D/g, esc(date.toLocaleDateString()))
    .replace(/&amp;T/g, esc(date.toLocaleTimeString()));
}

const COLUMN_NAME = (index) => {
  let n = index + 1;
  let out = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    out = String.fromCharCode(65 + r) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
};

/**
 * The sheet, laid out on pages, as a page of HTML.
 *
 * `sheets` chooses what is printed: the active sheet by default, or every
 * sheet, or a named list — a workbook printed whole is one document with the
 * pages in tab order.
 */
export function printHtml(view, options = {}) {
  const setup = pageSetup(options);
  const date = options.now ? new Date(options.now) : new Date();
  const file = options.file || '';
  const wanted = options.sheets === 'all' ? view.sheetNames() : Array.isArray(options.sheets) && options.sheets.length ? options.sheets : [view.activeSheet];
  const was = view.activeSheet;

  const sections = [];
  let pageNumber = 0;
  const plans = [];
  for (const name of wanted) {
    view.activeSheet = name;
    const bounds = view.calc.usedBounds(name);
    const area = parseArea(setup.area);
    const range = area || { top: 0, left: 0, bottom: bounds.maxRow, right: bounds.maxCol };
    plans.push({ name, plan: paginate({ geo: view.geometry.get(name), range, setup }), range });
  }
  const totalPages = plans.reduce((n, p) => n + p.plan.pages.length, 0);

  for (const { name, plan } of plans) {
    view.activeSheet = name;
    for (const page of plan.pages) {
      pageNumber += 1;
      const rows = [...page.titleRows, ...page.rows];
      const widths = page.cols.map((c) => `<col style="width:${Math.round(c.size * plan.scale)}px">`).join('');
      const body = rows
        .map((row) => {
          const cells = page.cols
            .map((col) => {
              const merge = view.mergeAt(row.index, col.index);
              if (merge && (merge.top !== row.index || merge.left !== col.index)) return '';
              const display = view.displayValue(row.index, col.index);
              const style = view.styleFor(row.index, col.index);
              const cell = {
                text: display.text,
                align: style?.align?.horizontal ?? display.align,
                valign: style?.align?.vertical ?? null,
                wrap: Boolean(style?.align?.wrap),
                indent: style?.align?.indent ?? 0,
                colour: display.colour,
                style: style ? { font: style.font, fill: style.fill, border: style.border } : null,
              };
              const span = merge
                ? ` colspan="${Math.min(merge.right, page.cols[page.cols.length - 1].index) - merge.left + 1}" rowspan="${Math.min(merge.bottom, rows[rows.length - 1].index) - merge.top + 1}"`
                : '';
              return `<td${span} style="${cellCss(cell, setup)}">${esc(cell.text)}</td>`;
            })
            .join('');
          const heading = setup.headings ? `<th class="rh">${row.index + 1}</th>` : '';
          return `<tr style="height:${Math.round(row.size * plan.scale)}px">${heading}${cells}</tr>`;
        })
        .join('');
      const headings = setup.headings
        ? `<tr class="ch"><th class="rh"></th>${page.cols.map((c) => `<th>${COLUMN_NAME(c.index)}</th>`).join('')}</tr>`
        : '';
      const context = { page: pageNumber, pages: totalPages, sheet: name, file, date };
      sections.push(
        `<section class="page${setup.centre.horizontal ? ' centre-h' : ''}${setup.centre.vertical ? ' centre-v' : ''}">` +
          (setup.header ? `<div class="running head">${fields(setup.header, context)}</div>` : '') +
          `<div class="grid"><table><colgroup>${setup.headings ? '<col style="width:34px">' : ''}${widths}</colgroup>${headings}${body}</table></div>` +
          (setup.footer ? `<div class="running foot">${fields(setup.footer, context)}</div>` : '') +
          '</section>'
      );
    }
  }
  view.activeSheet = was;

  const paper = PAPER[setup.paper] || PAPER.A4;
  return `<!doctype html>
<meta charset="utf-8">
<title>${esc(options.title || file || wanted.join(', '))}</title>
<style>
  @page { size: ${paper.css} ${setup.orientation}; margin: ${setup.margins.top}mm ${setup.margins.right}mm ${setup.margins.bottom}mm ${setup.margins.left}mm; }
  html, body { margin: 0; padding: 0; }
  body { font: 11pt/1.25 Calibri, Arial, sans-serif; color: #000; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  /* An exact page, not a minimum: a section that grows by one line starts a
     second sheet of paper, and the plan above said there would be one. */
  .page { break-after: page; display: flex; flex-direction: column; height: ${(printableArea(setup).height / PX_PER_MM).toFixed(2)}mm; overflow: hidden; }
  .page:last-child { break-after: auto; }
  .page.centre-h .grid { align-self: center; }
  .page.centre-v { justify-content: center; }
  .grid { overflow: hidden; }
  table { border-collapse: collapse; table-layout: fixed; }
  td, th { padding: 0 3px; overflow: hidden; text-overflow: clip; vertical-align: bottom; font-weight: inherit; }
  .ch th, th.rh { background: #f2f2f2; border: 1px solid #b7b7b7; font: 9pt Calibri, Arial, sans-serif; text-align: center; color: #333; }
  .running { font: 9pt Calibri, Arial, sans-serif; color: #444; padding: 2px 0; }
  .running.head { border-bottom: 0; }
  .running.foot { margin-top: auto; text-align: center; }
</style>
${sections.join('\n')}
`;
}

/** What the print dialog needs to say before anything is drawn. */
export function printSummary(view, options = {}) {
  const setup = pageSetup(options);
  const names = options.sheets === 'all' ? view.sheetNames() : Array.isArray(options.sheets) && options.sheets.length ? options.sheets : [view.activeSheet];
  let pages = 0;
  const per = [];
  for (const name of names) {
    const bounds = view.calc.usedBounds(name);
    const area = parseArea(setup.area);
    const range = area || { top: 0, left: 0, bottom: bounds.maxRow, right: bounds.maxCol };
    const plan = paginate({ geo: view.geometry.get(name), range, setup });
    pages += plan.pages.length;
    per.push({ sheet: name, pages: plan.pages.length, scale: plan.scale, ref: `${ref(range.top, range.left)}:${ref(range.bottom, range.right)}` });
  }
  return { pages, sheets: per, setup };
}
