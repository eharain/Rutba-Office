// View → Handout Master: where the slides go on a printed handout page —
// one to nine of them, in PowerPoint's own arrangements — drawn dashed over
// the master, as PowerPoint draws them. The slides are the Print dialog's
// to lay; the master holds only what is round them.

import React from 'react';

/** Columns and rows for each number of slides a page shows; three sit down the left, beside lines for notes. */
const GRID = { 1: [1, 1], 2: [1, 2], 3: [1, 3], 4: [2, 2], 6: [2, 3], 9: [3, 3] };

/**
 * The slots, in the page's own pixels: inside the header and footer bands,
 * each slot the slide's shape, as large as its cell allows, centred in it.
 */
export function handoutSlots(page, slide, per) {
  const [cols, rows] = GRID[per] || GRID[6];
  const W = page.width;
  const H = page.height;
  const left = W * 0.08;
  const top = H * 0.09;
  const width = (per === 3 ? W * 0.5 : W * 0.84);
  const height = H * 0.82;
  const gap = Math.min(W, H) * 0.03;
  const cellW = (width - gap * (cols - 1)) / cols;
  const cellH = (height - gap * (rows - 1)) / rows;
  const aspect = slide?.width && slide?.height ? slide.width / slide.height : 16 / 9;
  let w = cellW;
  let h = w / aspect;
  if (h > cellH) { h = cellH; w = h * aspect; }
  const out = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      out.push({ x: left + c * (cellW + gap) + (cellW - w) / 2, y: top + r * (cellH + gap) + (cellH - h) / 2, w, h });
    }
  }
  return out;
}

export function HandoutSlots({ size, slide, per }) {
  const slots = handoutSlots(size, slide, per);
  return (
    <div className="sl-handout-slots" aria-hidden="true">
      {slots.map((s, i) => (
        <div key={i} className="sl-handout-slot" style={{ left: s.x, top: s.y, width: s.w, height: s.h }} />
      ))}
      {per === 3 ? slots.map((s, i) => (
        <div key={`l${i}`} className="sl-handout-lines" style={{ left: size.width * 0.62, top: s.y, width: size.width * 0.3, height: s.h }} />
      )) : null}
    </div>
  );
}

export const HANDOUT_CSS = `
.sl-handout-slots { position: absolute; inset: 0; pointer-events: none; }
.sl-handout-slot { position: absolute; border: 2px dashed #8a94a6; border-radius: 2px; background: rgba(138, 148, 166, 0.06); box-sizing: border-box; }
.sl-handout-lines { position: absolute; background: repeating-linear-gradient(to bottom, transparent 0 22px, #b8bfcc 22px 23px); }
`;
