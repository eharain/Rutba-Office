// The page an app lays out for the printer itself.
//
// A document, a workbook and a deck are cut into pages by their engines. A
// message, a calendar, a card and a picture have no engine that knows paper:
// each app writes what it prints as one page of HTML, and Chromium cuts it
// into sheets in the same hidden window the engines' pages go through
// (main/print.js, pageSummary, page and pagePdf).
//
// This file is what those pages share: the paper, the frame every page is
// written in, and pictures laid on sheets (Image and Pictures both print
// them). Nothing here needs a window, so each app's page is checked as the
// text it is.
//
// The frame's policy lets a page draw only what it carries, the pictures the
// main process holds for the window (rutba://blob) and files on this computer
// (rutba://file). Nothing is fetched from anywhere else unless the app says
// the person already chose to (`remote`: Mail, with a message's remote
// pictures shown). A page that holds someone else's words as markup (Mail's
// message) asks for held pictures only (`files: false`), so nothing a sender
// writes has a file on this computer drawn. The hidden window runs no script
// whatever the page holds.

/** Paper, in millimetres, portrait: the sizes the print dialog offers. */
export const PAPER_MM = {
  A4: { width: 210, height: 297 },
  Letter: { width: 215.9, height: 279.4 },
  Legal: { width: 215.9, height: 355.6 },
  A3: { width: 297, height: 420 },
  A5: { width: 148, height: 210 },
};

/** CSS pixels to a millimetre: 96 to the inch, as Chromium prints. */
export const PX_PER_MM = 96 / 25.4;

const NORMAL = { top: 12.7, right: 12.7, bottom: 12.7, left: 12.7 };

/** Text made safe to write into a page: nothing a name or a subject holds becomes markup. */
export const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

/** The print dialog's choices as a page's setup: a paper this file knows, a direction, four margins. */
export function pageSetup(options = {}) {
  const paper = PAPER_MM[options.paper] ? options.paper : 'A4';
  const orientation = options.orientation === 'landscape' ? 'landscape' : 'portrait';
  const m = options.margins || {};
  const side = (k) => (Number.isFinite(Number(m[k])) && Number(m[k]) >= 0 ? Number(m[k]) : NORMAL[k]);
  return { paper, orientation, margins: { top: side('top'), right: side('right'), bottom: side('bottom'), left: side('left') } };
}

/** Whether a way round was chosen by a person: 'portrait' or 'landscape', and nothing else. */
export const wayChosen = (orientation) => orientation === 'portrait' || orientation === 'landscape';

/**
 * The way the paper is turned for something with a shape of its own (a picture, a slide): wider
 * than it is tall goes on paper turned the same way, and anything else stays upright.
 */
export function wayOfShape(width, height) {
  return Number(width) > 0 && Number(height) > 0 && Number(width) > Number(height) ? 'landscape' : 'portrait';
}

const DIGITS = /[٠-٩۰-۹]/g;
const plainDigit = (d) => String(d.charCodeAt(0) - (d.charCodeAt(0) >= 0x06f0 ? 0x06f0 : 0x0660));

/**
 * The pages a person typed in the Pages box, of a print of `count` pages.
 *
 * Page numbers and ranges, parted by commas, semicolons or spaces: "1,2,5,6", "1-5", "2-4, 9", "7-"
 * (seven to the end), "-3" (the start to three). The numbers may be written in Arabic or Persian
 * digits and the dash may be any dash. The pages come back counted from 1, in the order typed, each
 * once: more of a page is what Copies is for. A range that runs past the last page stops at it.
 *
 * @returns {{ all: true } | { pages: number[] } | { wrong: true } | { waiting: true }}
 *   `all` when nothing is typed, `wrong` when what is typed is not pages of this print, `waiting`
 *   when the pages are not counted yet
 */
export function pagesTyped(text, count) {
  const typed = String(text ?? '')
    .replace(DIGITS, plainDigit)
    .replace(/[‐-―−]/g, '-')
    .replace(/\s*-\s*/g, '-')
    .trim();
  if (!typed) return { all: true };
  const total = Number(count);
  if (!Number.isInteger(total) || total < 1) return { waiting: true };
  const pages = [];
  const add = (n) => {
    if (!pages.includes(n)) pages.push(n);
  };
  for (const part of typed.split(/[,;،؛\s]+/).filter(Boolean)) {
    const range = /^(\d*)-(\d*)$/.exec(part);
    if (range && (range[1] || range[2])) {
      const from = range[1] ? Number(range[1]) : 1;
      const to = range[2] ? Math.min(Number(range[2]), total) : total;
      if (from < 1 || from > total || from > to) return { wrong: true };
      for (let n = from; n <= to; n++) add(n);
    } else if (/^\d+$/.test(part)) {
      const n = Number(part);
      if (n < 1 || n > total) return { wrong: true };
      add(n);
    } else {
      return { wrong: true };
    }
  }
  return pages.length ? { pages } : { wrong: true };
}

/** The room inside the margins, in millimetres, the way the paper is turned. */
export function printableArea(setup) {
  const paper = PAPER_MM[setup.paper] || PAPER_MM.A4;
  const [w, h] = setup.orientation === 'landscape' ? [paper.height, paper.width] : [paper.width, paper.height];
  return { width: w - setup.margins.left - setup.margins.right, height: h - setup.margins.top - setup.margins.bottom };
}

/** A size in bytes as a list of attachments says it: 0 B, 12.4 KB, 3.1 MB. */
export function sizeWords(n) {
  if (!Number.isFinite(n) || n <= 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.min(units.length - 1, Math.floor(Math.log(n) / Math.log(1024)));
  const v = n / 1024 ** i;
  return `${v >= 100 || i === 0 ? Math.round(v) : v.toFixed(1)} ${units[i]}`;
}

/** A name made fit to be a file's: what a PDF of a print is offered as. */
export const printFileStem = (name) => String(name || '').replace(/[\/:*?"<>|\u0000-\u001f]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80) || 'print';

const mm = (n) => `${Math.round(n * 100) / 100}mm`;

/**
 * The whole page: the paper and margins as CSS, the policy, the app's own
 * styles and its body. `remote` lets pictures come from the web, for a page
 * whose reader has already let them. `files: false` keeps the page to the
 * pictures the main process holds for the window, for a page that carries
 * markup its app did not write.
 */
export function pageHtml({ title = '', options = {}, css = '', body = '', remote = false, files = true, lang = 'en', dir = 'ltr' } = {}) {
  const setup = pageSetup(options);
  const paper = PAPER_MM[setup.paper];
  const { margins } = setup;
  const size = setup.orientation === 'landscape' ? `${mm(paper.height)} ${mm(paper.width)}` : `${mm(paper.width)} ${mm(paper.height)}`;
  const csp = `default-src 'none'; img-src data: ${files ? 'rutba:' : 'rutba://blob'}${remote ? ' https:' : ''}; style-src 'unsafe-inline'; font-src data:`;
  return `<!doctype html>
<html lang="${esc(lang)}" dir="${dir === 'rtl' ? 'rtl' : 'ltr'}">
<head>
<meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="${csp}">
<title>${esc(title)}</title>
<style>
  @page { size: ${size}; margin: ${mm(margins.top)} ${mm(margins.right)} ${mm(margins.bottom)} ${mm(margins.left)}; }
  html, body { margin: 0; padding: 0; }
  body { font: 10.5pt/1.45 "Segoe UI", system-ui, -apple-system, Roboto, Arial, sans-serif; color: #111; background: #fff; -webkit-print-color-adjust: exact; print-color-adjust: exact; overflow-wrap: anywhere; }
  h1 { font-size: 15pt; line-height: 1.25; margin: 0 0 3mm; }
  table { border-collapse: collapse; }
  img { max-width: 100%; }
${css}
</style>
</head>
<body>
${body}
</body>
</html>`;
}

/* ── pictures on sheets ─────────────────────────────────────────────────── */

/** How many across and down a sheet takes for so many pictures to a page, the way the paper is turned. */
export function gridFor(perPage, orientation = 'portrait') {
  const n = [1, 2, 4, 6, 9].includes(Number(perPage)) ? Number(perPage) : 1;
  const [short, long] = { 1: [1, 1], 2: [1, 2], 4: [2, 2], 6: [2, 3], 9: [3, 3] }[n];
  return orientation === 'landscape' ? { cols: long, rows: short, perPage: n } : { cols: short, rows: long, perPage: n };
}

const SHEET_CSS = `
  .sheet { break-after: page; break-inside: avoid; overflow: hidden; position: relative; display: grid; }
  .sheet:last-child { break-after: auto; }
  .cell { display: flex; flex-direction: column; align-items: center; justify-content: center; min-width: 0; min-height: 0; overflow: hidden; }
  .cell .pic { flex: 1 1 0; min-height: 0; width: 100%; display: flex; align-items: center; justify-content: center; }
  .cell img { width: 100%; height: 100%; max-width: none; object-fit: contain; display: block; }
  .cell .cap { flex: 0 0 auto; width: 100%; text-align: center; font-size: 8.5pt; line-height: 1.3; color: #333; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; padding-top: 1.5mm; }
  .tile img { position: absolute; max-width: none; max-height: none; }
`;

/** A sheet is a hair shorter than its page, so a rounding of the paper's size never spills one onto a blank sheet. */
const SHEET_SLACK = 0.6;

/**
 * Pictures laid on sheets, each made to fit its place: one to a sheet, or
 * several in a grid, with their names beneath when asked. A picture is
 * `{ url, name }`. Returns the page's body and how many sheets it is.
 */
export function pictureSheets({ pictures = [], perPage = 1, captions = false, options = {} } = {}) {
  const setup = pageSetup(options);
  const area = printableArea(setup);
  const { cols, rows, perPage: each } = gridFor(perPage, setup.orientation);
  const gap = each > 1 ? 4 : 0;
  const height = area.height - SHEET_SLACK;
  const sheets = [];
  for (let i = 0; i < pictures.length; i += each) {
    const cells = pictures.slice(i, i + each).map((p) =>
      `<div class="cell"><div class="pic"><img src="${esc(p.url)}" alt=""></div>${captions && p.name ? `<div class="cap">${esc(p.name)}</div>` : ''}</div>`
    );
    sheets.push(`<section class="sheet" style="width:${mm(area.width)};height:${mm(height)};grid-template-columns:repeat(${cols},minmax(0,1fr));grid-template-rows:repeat(${rows},minmax(0,1fr));gap:${mm(gap)}">${cells.join('')}</section>`);
  }
  return { css: SHEET_CSS, body: sheets.join('\n'), sheets: sheets.length };
}

/** The most sheets a picture at its own size is cut into: past this, the dialog says so and prints nothing. */
export const MOST_TILES = 200;

/** How many sheets across and down a picture of this many pixels takes at its own size (96 pixels to the inch). */
export function tilesFor({ width, height, options = {} }) {
  const area = printableArea(pageSetup(options));
  const w = Math.max(1, Number(width) || 1) / PX_PER_MM;
  const h = Math.max(1, Number(height) || 1) / PX_PER_MM;
  // A picture a breath wider than the sheet is not worth a second sheet.
  const across = Math.max(1, Math.ceil(w / area.width - 0.005));
  const down = Math.max(1, Math.ceil((h / (area.height - SHEET_SLACK)) - 0.005));
  return { across, down, sheets: across * down, width: w, height: h };
}

/**
 * One picture at its own size, 96 pixels to the inch as a window shows it at
 * 100%: on one sheet when it fits, cut across as many as it needs when it
 * does not, read left to right and then down, to be laid edge to edge.
 */
export function posterSheets({ url, width, height, options = {} } = {}) {
  const setup = pageSetup(options);
  const area = printableArea(setup);
  const sheetHeight = area.height - SHEET_SLACK;
  const { across, down, width: w, height: h } = tilesFor({ width, height, options });
  if (across * down > MOST_TILES) return { css: SHEET_CSS, body: '', sheets: across * down, tooMany: true };
  const sheets = [];
  for (let r = 0; r < down; r++) {
    for (let c = 0; c < across; c++) {
      sheets.push(`<section class="sheet tile" style="width:${mm(area.width)};height:${mm(sheetHeight)}"><img src="${esc(url)}" alt="" style="left:${mm(-c * area.width)};top:${mm(-r * sheetHeight)};width:${mm(w)};height:${mm(h)}"></section>`);
    }
  }
  return { css: SHEET_CSS, body: sheets.join('\n'), sheets: sheets.length, across, down };
}
