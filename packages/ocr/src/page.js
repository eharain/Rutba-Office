// One recognition, start to finish, on the thread that holds the engine:
// the picture as it came (pixels, or a PNG file's bytes) made grey, looked
// at by our own analysis (analyse.js), prepared, read, and the answer put
// back into the picture's own coordinates.
//
// The preparation, each step skipped when it would change nothing:
//   turned to its negative when it is light letters on a dark ground (a dark
//     theme, a sign, a price on a black label), as everything after is made
//     for dark ink on light paper;
//   turned back level when the text is turned by two degrees or more (the
//     engine reads a smaller tilt as well itself, and turning a soft scan
//     by resampling it costs more than the tilt: docs/OCR.md);
//   scaled when the text is very small (a screenshot, a price cut from one)
//     or very large (a poster, a 600 dpi scan), towards the height the
//     models read best;
//   handed over as ink and paper (Sauvola, its window wider than the
//     heaviest stroke) when the paper is lit unevenly,
//     and otherwise as grey, which the engine thresholds itself.

import { checkLanguages } from './languages.js';
import { isPng, decodePng } from './png.js';
import { toGrey } from './pixels.js';
import { binarise, estimateSkew, rotate, unrotatePoint, scale, textHeight, unevenPaper, darkPaper, invert, rulings, cropPadded, inkIn, fitToInk } from './analyse.js';
import { tableCells } from './layout.js';
import { documentFields } from './fields.js';
import { findCheckboxes } from './checkboxes.js';
import { fromEngine } from './results.js';
import { PSM } from './engine.js';

/**
 * A table read from its picture, by our own layout over the engine's words.
 * The engine's page segmentation is made for pages, not grids: it can drop
 * a narrow column of numbers whole, and its sparse mode, which keeps the
 * column, drops a figure standing alone. So the words of both readings are
 * merged (one word where two overlap, the surer kept), the grid is found
 * from the table's rules or its spacing, and a cell that came out empty
 * though there is ink in it is read again on its own as one line.
 */
function readTable(engine, { grey, prepared, how, langs, dpi, result, map }) {
  const level = how.turn ? rotate(grey, -how.turn) : grey;
  const binary = binarise(level, { method: how.uneven ? 'sauvola' : 'otsu', textHeight: how.letters });
  const rules = rulings(binary);
  const rtl = result.lines.filter((l) => l.direction === 'rtl').length > result.lines.length / 2;
  // Boxes on the level picture: the rules were found there.
  const toLevel = (b) => (how.turn ? turnBox(b, how.turn, grey.width, grey.height) : b);
  const sparse = fromEngine(engine.read(prepared, { languages: langs, dpi: dpi ? dpi * how.factor : 0, psm: PSM.sparse }).json, { width: grey.width, height: grey.height, languages: langs, map });
  const words = mergeWords(result.words.map((w) => ({ ...w, box: toLevel(w.box) })), sparse.words.map((w) => ({ ...w, box: toLevel(w.box) })));
  const grid = tableCells(words, rules, { rtl });
  // An empty cell with ink in it: read alone, as a line.
  const pad = 3;
  grid.rows.forEach((row, r) => {
    grid.columns.forEach((col, c) => {
      if (grid.cells[r][c]) return;
      const x = col.x0 + pad, y = row.y0 + pad, w = col.x1 - col.x0 - pad * 2, h = row.y1 - row.y0 - pad * 2;
      if (w < 4 || h < 4 || inkIn(binary, x, y, w, h) < 12) return;
      const crop = cropPadded(level, x, y, w, h);
      const again = fromEngine(engine.read(crop, { languages: langs, psm: PSM.line }).json, { width: crop.width, height: crop.height, languages: langs });
      const text = again.text.replace(/\s+/g, ' ').trim().replace(/^[|│¦_]+|[|│¦_]+$/g, '').trim();
      if (text) {
        grid.cells[r][c] = text;
        grid.confidence[r][c] = again.words.length ? Math.min(...again.words.map((x2) => x2.confidence)) : 0;
      }
    });
  });
  return grid;
}

/**
 * The words of two readings of one picture as one list: the first reading's
 * words in their order, each replaced by the second's where that is surer by
 * more than ten and no shorter, and the second's words that lie over none of
 * the first's added after them (with no line of their own) when they are
 * at least `least` sure: a sparse reading also reads a dashed rule or a
 * grain of paper as letters, unsure ones.
 */
export function mergeWords(words, more, { least = 0 } = {}) {
  const out = words.map((w) => ({ ...w }));
  const overlap = (a, b) => {
    const w = Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x);
    const h = Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y);
    return w > 0 && h > 0 ? (w * h) / Math.min(a.width * a.height, b.width * b.height) : 0;
  };
  const first = out.length;
  for (const s of more) {
    const same = out.findIndex((w) => overlap(w.box, s.box) > 0.4);
    if (same < 0) { if (s.confidence >= least) out.push({ ...s, line: -1 }); }
    else if (same < first && s.confidence > out[same].confidence + 10 && s.text.length >= out[same].text.length) out[same] = { ...s, line: out[same].line };
  }
  return out;
}

/** A box in the picture as it lies in the picture turned back by `turn`. */
function turnBox(b, turn, width, height) {
  const corners = [[b.x, b.y], [b.x + b.width, b.y], [b.x, b.y + b.height], [b.x + b.width, b.y + b.height]].map(([x, y]) => unrotatePoint(x, y, turn, width, height));
  const xs = corners.map((c) => c.x), ys = corners.map((c) => c.y);
  return { x: Math.min(...xs), y: Math.min(...ys), width: Math.max(...xs) - Math.min(...xs), height: Math.max(...ys) - Math.min(...ys) };
}

/** Pixels from whatever was handed over. */
export function picture(image) {
  if (image instanceof Uint8Array || image instanceof ArrayBuffer) {
    const bytes = image instanceof ArrayBuffer ? new Uint8Array(image) : image;
    if (isPng(bytes)) return decodePng(bytes);
    throw Object.assign(new Error('Text can be read from pixels or a PNG picture.'), { code: 'OCR_IMAGE' });
  }
  const { width, height, data } = image || {};
  if (!(width > 0 && height > 0 && data && data.length >= width * height)) {
    throw Object.assign(new Error('The picture to read has no pixels.'), { code: 'OCR_IMAGE' });
  }
  return { width, height, data: data instanceof Uint8Array ? data : new Uint8Array(data.buffer ?? data, data.byteOffset ?? 0, data.length), channels: image.channels };
}

/** Text smaller than this (in pixels, a lower-case letter's height) is scaled up, larger than LARGE down. */
const SMALL = 10;
const LARGE = 64;
const TARGET = 20;
/**
 * A picture of one line (a price, a label cut from a screenshot) is measured
 * by its letters' full height, a figure's or a capital's, not a page's
 * lower-case: under LINE_SMALL it is scaled up towards LINE_TARGET. Small
 * heavy figures most of all are read better larger: their counters and
 * decimal points are a pixel or two.
 */
const LINE_SMALL = 14;
const LINE_TARGET = 36;
/** A page turned by this many degrees or more is turned back before reading. */
const TURN = 2;
/** Analysis looks at a copy of at most this many pixels; the answer does not need more. */
const LOOK = 3e6;

/**
 * What to do to a grey picture before reading it: the angle to turn it back,
 * the factor to scale it by, whether to hand over ink and paper, and
 * whether to read its negative.
 */
export function plan(grey, { deskew = true, rescale = true } = {}) {
  const shrink = Math.min(1, Math.sqrt(LOOK / (grey.width * grey.height)));
  let look = shrink < 1 ? scale(grey, shrink) : grey;
  const negative = darkPaper(look);
  if (negative) look = invert(look);
  const uneven = unevenPaper(look);
  const binary = binarise(look, { method: uneven ? 'sauvola' : 'otsu' });
  const angle = deskew ? estimateSkew(binary) : 0;
  const height = textHeight(binary) / shrink;
  // A picture of one line has letters too tall for a page's: measured as a line, for its scale and Sauvola's window.
  const letters = height || textHeight(binary, { least: 2, line: true }) / shrink;
  let factor = 1;
  if (rescale && height > 0 && height < SMALL) factor = Math.min(4, TARGET / height);
  else if (rescale && height > LARGE) factor = Math.max(0.25, TARGET * 1.6 / height);
  else if (rescale && !height && letters > 0 && letters < LINE_SMALL) factor = Math.min(4, LINE_TARGET / letters);
  return { angle, turn: Math.abs(angle) >= TURN ? angle : 0, factor: Math.round(factor * 100) / 100, uneven, negative, textHeight: Math.round(height), letters: Math.round(letters) };
}

/**
 * @param engine from createEngine
 * @param {{ image, languages?, dpi?, mode?: 'auto'|'block'|'line'|'sparse', deskew?: boolean, rescale?: boolean, table?: boolean, document?: boolean }} job
 *   `table`: the picture is a table; the answer carries `table` too, its rows,
 *   columns and cells from its ruling lines or its spacing (layout.js)
 *   `document`: the picture is a bill, an invoice, a receipt, a letter or a
 *   form; the page is read twice (as a page, and as sparse text, which keeps
 *   a label or a figure standing alone that the first drops), the words of
 *   both merged, the figures beside its labels and its items' amounts read
 *   again each as one line, and the answer carries `document` too: its
 *   kind, its fields, its items, whether its sums agree (fields.js) and its
 *   tick boxes, ticked or not (checkboxes.js), with
 *   the merged words and the second readings the fields were read from
 */
export function recognizePage(engine, { image, languages, dpi = 0, mode = 'auto', deskew = true, rescale = true, table = false, document = false }) {
  const langs = checkLanguages(languages);
  const started = Date.now();
  const pixels = picture(image);
  const seen = toGrey(pixels);
  const how = plan(seen, { deskew, rescale });
  // Light letters on a dark ground are read as their negative, dark on light, as everything after expects.
  const grey = how.negative ? invert(seen) : seen;
  let prepared = how.turn ? rotate(grey, -how.turn) : grey;
  if (how.factor !== 1) prepared = scale(prepared, how.factor);
  if (how.uneven) prepared = binarise(prepared, { method: 'sauvola', textHeight: how.letters * how.factor });
  const { json } = engine.read(prepared, { languages: langs, dpi: dpi ? dpi * how.factor : 0, psm: PSM[mode] ?? PSM.auto });
  // A box in what was read, back in the picture: unscaled, then turned as the picture is turned.
  const map = (b) => {
    const x0 = b.x / how.factor, y0 = b.y / how.factor, x1 = (b.x + b.width) / how.factor, y1 = (b.y + b.height) / how.factor;
    if (!how.turn) return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
    const corners = [[x0, y0], [x1, y0], [x0, y1], [x1, y1]].map(([x, y]) => unrotatePoint(x, y, -how.turn, grey.width, grey.height));
    const xs = corners.map((c) => c.x), ys = corners.map((c) => c.y);
    return { x: Math.min(...xs), y: Math.min(...ys), width: Math.max(...xs) - Math.min(...xs), height: Math.max(...ys) - Math.min(...ys) };
  };
  let result = fromEngine(json, { width: pixels.width, height: pixels.height, languages: langs, angle: how.angle, map });
  // A picture of one line the engine found nothing in, though there are letters in it: its page analysis
  // can take a short line of heavy figures for a picture and drop it. Read again as the one line it is,
  // at its own size.
  if (!result.words.length && !how.textHeight && how.letters && !how.turn && mode === 'auto') {
    const own = how.uneven ? binarise(grey, { method: 'sauvola', textHeight: how.letters }) : grey;
    const again = fromEngine(engine.read(own, { languages: langs, dpi, psm: PSM.line }).json, { width: pixels.width, height: pixels.height, languages: langs, angle: how.angle });
    if (again.words.length) result = again;
  }
  // Each word's box drawn in to its own ink, side to side, on a page read as it lies (a box mapped back
  // from a turned reading is a turned box's bounds, wider than the word by its slant): what Find marks
  // and a selection covers is then the word on the page, in Arabic script as in Latin.
  if (!how.turn) fitToInk(binarise(grey, { textHeight: how.letters }), result.words);
  result.prepared = { turned: how.turn, factor: how.factor, binarised: how.uneven, inverted: how.negative, textHeight: how.textHeight };
  if (table) result.table = readTable(engine, { grey, prepared, how, langs, dpi, result, map });
  if (document) {
    const sparse = fromEngine(engine.read(prepared, { languages: langs, dpi: dpi ? dpi * how.factor : 0, psm: PSM.sparse }).json, { width: pixels.width, height: pixels.height, languages: langs, map });
    const words = mergeWords(result.words, sparse.words, { least: 40 });
    const read = { ...result, words };
    // The figures beside the labels and the items' amounts, read again each as one line: by English's
    // model, which reads Western digits best, and by the page's own, which reads its own digits.
    const collect = [];
    // The tick boxes, a sign of a form and part of what it says.
    const checkboxes = findCheckboxes(grey, { words, textHeight: how.textHeight });
    documentFields(read, { languages: langs, collect, checkboxes });
    const again = collect.map((box) => ({ box, readings: [] }));
    const sets = langs.includes('eng') && langs.length > 1 ? [['eng'], langs] : [langs];
    for (const set of sets) {
      for (const a of again) {
        const crop = cropPadded(grey, a.box.x, a.box.y, a.box.width, a.box.height);
        const r = fromEngine(engine.read(crop, { languages: set, psm: PSM.line }).json, { width: crop.width, height: crop.height, languages: set });
        const text = r.text.replace(/\s+/g, ' ').trim();
        if (text) a.readings.push({ text, confidence: r.words.length ? Math.min(...r.words.map((w) => w.confidence)) : 0, languages: set });
      }
    }
    result.document = { ...documentFields(read, { languages: langs, again, checkboxes }), checkboxes, words, again };
  }
  result.ms = Date.now() - started;
  return result;
}
