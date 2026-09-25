/**
 * Where a floating drawing stands — one answer for the screen and the page.
 *
 * A `wp:anchor` places its drawing on each axis either by an alignment
 * (left/centre/right, top/middle/bottom) or by an offset, each measured from
 * something: the page, the margins, the column, a margin itself, or — down
 * the page — the paragraph that anchors it. The screen and the paginator
 * both ask this module, so a text box dragged to a place on screen prints
 * in that place, and one aligned to the page's centre is centred on both.
 *
 * Horizontal answers are px from the column's (the content box's) left
 * edge; vertical answers say what they are measured from: `paragraph` (the
 * anchor's top) or `page` (the sheet's top).
 */

/** Which of Word's frames a horizontal `relativeFrom` is, as a span from the column's left edge. */
function hFrame(rel, g) {
  const ml = g.marginLeftPx ?? 96;
  const content = g.contentWidthPx ?? 624;
  const page = g.pageWidthPx ?? content + ml + (g.marginRightPx ?? 96);
  switch (rel) {
    case 'page': return [-ml, page - ml];
    case 'leftMargin':
    case 'insideMargin': return [-ml, 0];
    case 'rightMargin':
    case 'outsideMargin': return [content, page - ml];
    case 'column': return [0, g.columnWidthPx ?? content];
    // margin, character, and anything a newer Word names: the text box.
    default: return [0, content];
  }
}

/** Which frame a vertical `relativeFrom` is, down the page, from the sheet's top — or null for the paragraph. */
function vFrame(rel, g) {
  const H = g.pageHeightPx ?? 1123;
  const mt = g.marginTopPx ?? 96;
  const mb = g.marginBottomPx ?? 96;
  switch (rel) {
    case 'page': return [0, H];
    case 'margin': return [mt, H - mb];
    case 'topMargin':
    case 'insideMargin': return [0, mt];
    case 'bottomMargin':
    case 'outsideMargin': return [H - mb, H];
    default: return null; // paragraph, line
  }
}

/**
 * A drawing's place: `x` from the column's left, `y` from what `yFrom`
 * names, and its box. `d` carries the anchor's fields as the engine reads
 * them (hRel, hAlign, hOffsetPx, vRel, vAlign, vOffsetPx, widthPx,
 * heightPx); `g` the page's geometry in px.
 */
export function floatPlace(d, g = {}) {
  const w = Number(d.widthPx) || 0;
  const h = Number(d.heightPx) || 0;
  const [a, b] = hFrame(d.hRel, g);
  let x;
  if (d.hAlign) {
    const align = d.hAlign === 'inside' ? 'left' : d.hAlign === 'outside' ? 'right' : d.hAlign;
    x = align === 'right' ? b - w : align === 'center' ? (a + b - w) / 2 : a;
  } else x = a + (Number(d.hOffsetPx) || 0);
  const frame = vFrame(d.vRel, g);
  let y;
  let yFrom;
  if (frame) {
    yFrom = 'page';
    if (d.vAlign) {
      const [top, bottom] = frame;
      y = d.vAlign === 'bottom' || d.vAlign === 'outside' ? bottom - h : d.vAlign === 'center' ? (top + bottom - h) / 2 : top;
    } else y = frame[0] + (Number(d.vOffsetPx) || 0);
  } else {
    yFrom = 'paragraph';
    y = Number(d.vOffsetPx) || 0;
  }
  return { x, y, yFrom, widthPx: w, heightPx: h };
}

/** Does the words' flow go round this drawing, and on which side of it — the side CSS floats it to. */
export function wrapsBeside(d) {
  return Boolean(d?.anchored) && ['square', 'tight', 'through'].includes(d.wrap) && !d.behind;
}

/** The side a beside-wrapped drawing stands at: the half of the column its middle is in. */
export function sideOf(x, widthPx, columnWidthPx) {
  return x + widthPx / 2 > columnWidthPx / 2 ? 'right' : 'left';
}

/** Is the drawing laid on the page over or under the words, out of their flow (wrap none)? */
export function layerOf(d) {
  if (!d?.anchored) return 'inline';
  if (d.wrap === 'none' || !d.wrap) return d.behind ? 'behind' : 'front';
  if (d.wrap === 'topAndBottom') return 'block';
  return 'beside';
}

/**
 * Where a drawing's top lands from its anchor paragraph's top, in the
 * paragraph's own terms — the offset a CSS float or a paginator needs.
 * `anchorTop` is the anchor paragraph's top on its page, from the sheet's
 * top (only read when the drawing is placed on the page).
 */
export function topFromAnchor(place, anchorTop) {
  return place.yFrom === 'page' ? place.y - anchorTop : place.y;
}
