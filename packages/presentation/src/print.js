/**
 * A deck on paper.
 *
 * Three things people print a deck as, and they are not variations of one
 * layout: the slides themselves, one to a page, for a handbook; the speaker's
 * notes under a small slide, for talking from; and a handout of several slides
 * to a page, for the room. PowerPoint calls these Full Page Slides, Notes
 * Pages and Handouts, and anyone who has printed a deck expects all three.
 *
 * The drawing is the renderer the window already uses, so what comes out of a
 * printer is what was on the stage — the same SVG, at paper size. Pictures are
 * embedded as data, because the page is handed to a window that cannot reach
 * this process's blobs.
 */

import { renderSlide } from './render.js';

/** Paper in millimetres, portrait, as the sheet printer has them. */
export const PAPER = {
  A3: { width: 297, height: 420, css: 'A3' },
  A4: { width: 210, height: 297, css: 'A4' },
  Letter: { width: 215.9, height: 279.4, css: 'Letter' },
  Legal: { width: 215.9, height: 355.6, css: 'Legal' },
};

export const PX_PER_MM = 96 / 25.4;

export const DEFAULT_DECK_PRINT = {
  paper: 'A4',
  /** Slides are wider than they are tall, so a full-page print is landscape. */
  orientation: null,
  margins: { top: 10, right: 10, bottom: 10, left: 10 },
  /** 'slides' | 'notes' | 'handout' */
  layout: 'slides',
  /** Handout only: 1, 2, 3, 4, 6 or 9 to a page. */
  perPage: 6,
  /** A frame round each slide, as PowerPoint's "frame slides" does. */
  frame: true,
  /** Only these slides, by index, or null for all of them. */
  slides: null,
  header: null,
  footer: null,
};

const MIME = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  tiff: 'image/tiff',
  emf: 'image/emf',
  wmf: 'image/wmf',
};

const base64 = (bytes) => {
  if (typeof Buffer !== 'undefined') return Buffer.from(bytes).toString('base64');
  let binary = '';
  for (const b of bytes) binary += String.fromCharCode(b);
  return btoa(binary);
};

const esc = (s) =>
  String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

/** The picture parts, as data — a printing window cannot reach a blob held here. */
function imageResolver(deck) {
  const cache = new Map();
  return (shape) => {
    const part = shape.source?.part;
    if (!part) return null;
    if (cache.has(part)) return cache.get(part);
    let url = null;
    try {
      const bytes = deck.media(part);
      if (bytes) url = `data:${MIME[part.split('.').pop()?.toLowerCase()] || 'application/octet-stream'};base64,${base64(bytes)}`;
    } catch {
      url = null;
    }
    cache.set(part, url);
    return url;
  };
}

/** How many across and down a handout of N slides is, as PowerPoint has it. */
const HANDOUT_GRID = { 1: [1, 1], 2: [1, 2], 3: [1, 3], 4: [2, 2], 6: [2, 3], 9: [3, 3] };

export function deckPrintSetup(options = {}) {
  const setup = { ...DEFAULT_DECK_PRINT, ...options, margins: { ...DEFAULT_DECK_PRINT.margins, ...(options.margins || {}) } };
  // A slide is landscape and a page of notes is not: the default follows the
  // layout rather than making a person set it.
  if (!setup.orientation) setup.orientation = setup.layout === 'slides' ? 'landscape' : 'portrait';
  if (!HANDOUT_GRID[setup.perPage]) setup.perPage = 6;
  return setup;
}

export function deckPrintSummary(deck, options = {}) {
  const setup = deckPrintSetup(options);
  const indexes = (setup.slides && setup.slides.length ? setup.slides : deck.slideParts.map((_, i) => i)).filter((i) => i >= 0 && i < deck.slideCount);
  const perPage = setup.layout === 'handout' ? setup.perPage : 1;
  return { pages: Math.max(1, Math.ceil(indexes.length / perPage)), slides: indexes.length, setup };
}

/**
 * The deck as a page of HTML, one section per printed page.
 *
 * `outline` is used for the notes, so a deck of two hundred slides does not
 * build two hundred scenes to print six of them.
 */
export function deckPrintHtml(deck, options = {}) {
  const setup = deckPrintSetup(options);
  const paper = PAPER[setup.paper] || PAPER.A4;
  const landscape = setup.orientation === 'landscape';
  const pageWidth = (landscape ? paper.height : paper.width) - setup.margins.left - setup.margins.right;
  const pageHeight = (landscape ? paper.width : paper.height) - setup.margins.top - setup.margins.bottom;

  const all = deck.slideParts.map((_, i) => i);
  const indexes = (setup.slides && setup.slides.length ? setup.slides : all).filter((i) => i >= 0 && i < deck.slideCount);
  const resolveImage = imageResolver(deck);
  const size = deck.size || { width: 960, height: 540 };
  const aspect = size.height / (size.width || 1);

  const drawn = (index, widthMm) => {
    try {
      return renderSlide(deck.slide(index), { width: Math.round(widthMm * PX_PER_MM), resolveImage, standalone: true });
    } catch {
      return `<div class="broken">Slide ${index + 1} could not be drawn.</div>`;
    }
  };

  const sections = [];
  if (setup.layout === 'handout') {
    const [across, down] = HANDOUT_GRID[setup.perPage];
    const gapMm = 6;
    const cellW = (pageWidth - gapMm * (across - 1)) / across;
    for (let at = 0; at < indexes.length; at += setup.perPage) {
      const group = indexes.slice(at, at + setup.perPage);
      const tiles = group
        .map((i) => `<figure class="tile${setup.frame ? ' framed' : ''}" style="width:${cellW.toFixed(2)}mm"><div class="shot">${drawn(i, cellW)}</div><figcaption>${i + 1}</figcaption></figure>`)
        .join('');
      sections.push(`<section class="page handout" style="--across:${across};--down:${down};--gap:${gapMm}mm">${tiles}</section>`);
    }
  } else if (setup.layout === 'notes') {
    const shotW = pageWidth * 0.8;
    for (const i of indexes) {
      const notes = esc(deck.slide(i).notes || '').replace(/\n/g, '<br>');
      sections.push(
        `<section class="page notes">` +
          `<div class="shot${setup.frame ? ' framed' : ''}" style="width:${shotW.toFixed(2)}mm">${drawn(i, shotW)}</div>` +
          `<div class="note-text">${notes || '<span class="none">No notes on this slide.</span>'}</div>` +
          `<div class="stamp">${i + 1}</div>` +
          '</section>'
      );
    }
  } else {
    for (const i of indexes) {
      // A slide fills the page in whichever direction runs out first, so a
      // 4:3 deck on A4 landscape is not stretched.
      const byWidth = Math.min(pageWidth, pageHeight / aspect);
      sections.push(`<section class="page full"><div class="shot${setup.frame ? ' framed' : ''}" style="width:${byWidth.toFixed(2)}mm">${drawn(i, byWidth)}</div></section>`);
    }
  }

  return `<!doctype html>
<meta charset="utf-8">
<title>${esc(options.title || 'Presentation')}</title>
<style>
  @page { size: ${paper.css} ${setup.orientation}; margin: ${setup.margins.top}mm ${setup.margins.right}mm ${setup.margins.bottom}mm ${setup.margins.left}mm; }
  html, body { margin: 0; padding: 0; }
  body { font: 10pt/1.35 Calibri, Arial, sans-serif; color: #111; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  .page { break-after: page; height: ${pageHeight.toFixed(2)}mm; display: flex; }
  .page:last-child { break-after: auto; }
  .shot svg { display: block; width: 100%; height: auto; }
  .framed { outline: 0.3mm solid #9aa0a6; outline-offset: 0; }
  .full { align-items: center; justify-content: center; }
  .notes { flex-direction: column; align-items: center; gap: 6mm; }
  .notes .note-text { width: 80%; font-size: 10.5pt; white-space: pre-wrap; }
  .notes .none { color: #888; }
  .notes .stamp { margin-top: auto; align-self: center; color: #666; font-size: 9pt; }
  .handout { flex-wrap: wrap; align-content: flex-start; gap: var(--gap); }
  .handout .tile { margin: 0; }
  .handout figcaption { text-align: right; color: #666; font-size: 8.5pt; padding-top: 1mm; }
  .broken { border: 0.3mm dashed #b00; color: #b00; padding: 4mm; font-size: 10pt; }
</style>
${sections.join('\n')}
`;
}
