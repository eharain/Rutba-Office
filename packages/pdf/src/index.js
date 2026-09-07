'use strict';

/**
 * @rutba/pdf - writing PDFs, with nothing installed to do it.
 *
 * The suite could READ a PDF (`pdfjs-dist`, in the viewer) and could not
 * write one, which is why finance could email an invoice but never attach
 * it. The two usual answers are a headless browser - a Chromium download and
 * a process per document - or a drawing library that still leaves you to
 * write the layout. Neither is much use for the thing actually wanted: a
 * page of A4 that looks like an invoice.
 *
 * See `metrics.js` for why this needs no font files, `document.js` for the
 * file format, `sheet.js` for the layout, and `paper.js` for the business
 * document those three add up to.
 */

const { PdfDocument, Page, SIZES, colour, fingerprint } = require('./document');
const { Sheet, INK, DEFAULT_MARGINS } = require('./sheet');
const { drawPaper, renderPaper } = require('./paper');
const { decodePng, isPng } = require('./images');
const { appendExecutionPage, overlayPages, describePages, Unsupported: PdfUnsupportedError } = require('./incremental');
const metrics = require('./metrics');
const encoding = require('./encoding');

module.exports = {
  PdfDocument, Page, SIZES, colour, fingerprint,
  Sheet, INK, DEFAULT_MARGINS,
  drawPaper, renderPaper,
  decodePng, isPng,
  appendExecutionPage, overlayPages, describePages, PdfUnsupportedError,
  metrics, encoding,
  FONTS: metrics.FONTS,
};
