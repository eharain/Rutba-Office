// Insert → Photo Album: a new presentation made of pictures — a title slide,
// then the pictures one, two or four to a slide, each fitted to its share of
// the slide without stretching, with its name under it when asked. What
// PowerPoint's New Photo Album makes, as a deck like any other.

import { buildPptx, SLIDE_SIZE_16_9 } from './build.js';
import { Deck } from './deck.js';
import { emuToPx } from './units.js';

const PAD = 40;
const GAP = 24;
const CAPTION = 34;

/** A picture's name without its extension — the caption PowerPoint gives it. */
const captionOf = (name) => String(name || 'Picture').replace(/\.[^.]+$/, '');

/**
 * Where each picture on a slide goes: its box, fitted inside its cell, and
 * the caption's box under it. `size` is the slide's, in pixels.
 */
export function albumBoxes(pictures, perSlide, captions, size) {
  const cols = perSlide === 1 ? 1 : 2;
  const rows = perSlide === 4 ? 2 : 1;
  const cellW = (size.width - 2 * PAD - (cols - 1) * GAP) / cols;
  const cellH = (size.height - 2 * PAD - (rows - 1) * GAP) / rows;
  const areaH = cellH - (captions ? CAPTION : 0);
  return pictures.map((p, i) => {
    const cx = PAD + (i % cols) * (cellW + GAP);
    const cy = PAD + Math.floor(i / cols) * (cellH + GAP);
    const iw = p.width > 0 ? p.width : 4;
    const ih = p.height > 0 ? p.height : 3;
    const scale = Math.min(cellW / iw, areaH / ih);
    const w = Math.max(1, Math.round(iw * scale));
    const h = Math.max(1, Math.round(ih * scale));
    return {
      picture: { x: Math.round(cx + (cellW - w) / 2), y: Math.round(cy + (areaH - h) / 2), w, h },
      caption: captions ? { x: Math.round(cx), y: Math.round(cy + areaH + 4), w: Math.round(cellW), h: CAPTION - 4 } : null,
    };
  });
}

/**
 * The album, as a Deck.
 *
 *   pictures  [{ data, contentType, name, width, height }] — width and height
 *             the picture's own, for its shape; a picture without them is
 *             drawn four by three
 *   perSlide  1, 2 or 4
 *   captions  each picture's name under it
 *   title     the title slide's title; `subtitle` under it
 */
export function photoAlbum({ pictures = [], perSlide = 1, captions = false, title = 'Photo Album', subtitle = '', size = SLIDE_SIZE_16_9 } = {}) {
  if (![1, 2, 4].includes(perSlide)) throw new Error('a photo album puts 1, 2 or 4 pictures on a slide');
  if (!pictures.length) throw new Error('a photo album needs a picture');
  const px = { width: emuToPx(size.cx), height: emuToPx(size.cy) };
  const groups = [];
  for (let i = 0; i < pictures.length; i += perSlide) groups.push(pictures.slice(i, i + perSlide));
  const layout = groups.map((g) => albumBoxes(g, perSlide, captions, px));
  const slides = [
    { layout: 'title', title: title || 'Photo Album', body: subtitle || '' },
    ...groups.map((g, s) => ({
      layout: 'obj',
      textBoxes: captions ? g.map((p, i) => ({ ...layout[s][i].caption, paragraphs: [{ align: 'center', runs: [{ text: captionOf(p.name), size: 14 }] }] })) : [],
    })),
  ];
  const deck = Deck.open(buildPptx({ title: title || 'Photo Album', size, slides }));
  groups.forEach((g, s) => g.forEach((p, i) => {
    deck.addPicture(s + 1, { data: p.data, contentType: p.contentType, name: captionOf(p.name), ...layout[s][i].picture });
  }));
  return deck;
}
