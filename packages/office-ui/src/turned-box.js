// A turned object's frame: where its corners, its eight handles and its
// rotate handle stand on the page, whether a point falls inside it, and the
// box a handle drag gives it, all in the object's own axes.
//
// Every app draws an inserted picture, shape, text box or chart as an
// upright box `{ x, y, w, h }` turned by `deg` degrees clockwise about its
// centre, the way OOXML's `a:xfrm rot` stores it. The frame turns with it,
// as Office's does: a click inside the turned outline takes hold of it, a
// click in the empty corner of the upright box does not, and a handle
// resizes along the object's own sides with the opposite handle staying put.

/** The eight handles: where on the box each sits, as fractions of its width and height. */
export const FRAME_HANDLES = [
  ['nw', 0, 0], ['n', 0.5, 0], ['ne', 1, 0], ['e', 1, 0.5],
  ['se', 1, 1], ['s', 0.5, 1], ['sw', 0, 1], ['w', 0, 0.5],
];

const RAD = Math.PI / 180;

/** (x, y) turned `deg` degrees clockwise (on screen, y down) about (cx, cy). */
export function turnPoint(x, y, cx, cy, deg) {
  if (!deg) return { x, y };
  const a = deg * RAD;
  const c = Math.cos(a);
  const s = Math.sin(a);
  const dx = x - cx;
  const dy = y - cy;
  return { x: cx + dx * c - dy * s, y: cy + dx * s + dy * c };
}

/** The point at (fx, fy) of the box, as fractions, where it stands once the box is turned. */
export function framePoint(box, deg, fx, fy) {
  return turnPoint(box.x + box.w * fx, box.y + box.h * fy, box.x + box.w / 2, box.y + box.h / 2, deg);
}

/** The four corners of the turned box, clockwise from its own top left. */
export function frameCorners(box, deg) {
  return [[0, 0], [1, 0], [1, 1], [0, 1]].map(([fx, fy]) => framePoint(box, deg, fx, fy));
}

/** Each of the eight handles with its place on the page and the cursor it wears there. */
export function frameHandles(box, deg) {
  return FRAME_HANDLES.map(([name, fx, fy]) => ({ name, ...framePoint(box, deg, fx, fy), cursor: handleCursor(name, deg) }));
}

/**
 * The rotate handle: `gap` above the middle of the turned top edge, and that
 * middle itself, for the stem that joins them.
 */
export function rotateHandle(box, deg, gap) {
  const top = framePoint(box, deg, 0.5, 0);
  const a = deg * RAD;
  return { x: top.x + gap * Math.sin(a), y: top.y - gap * Math.cos(a), stem: top };
}

/** The upright box round the turned one. */
export function frameBounds(box, deg) {
  const pts = frameCorners(box, deg);
  const xs = pts.map((p) => p.x);
  const ys = pts.map((p) => p.y);
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  return { x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y };
}

/** Whether (px, py) falls inside the turned box (`slack` pixels round it count too). */
export function insideFrame(box, deg, px, py, slack = 0) {
  const p = turnPoint(px, py, box.x + box.w / 2, box.y + box.h / 2, -(deg || 0));
  return p.x >= box.x - slack && p.x <= box.x + box.w + slack && p.y >= box.y - slack && p.y <= box.y + box.h + slack;
}

/**
 * The cursor a handle wears once the box is turned: its own direction turned
 * with the box, to the nearest of the four double arrows.
 */
export function handleCursor(name, deg = 0) {
  const base = { n: 0, ne: 45, e: 90, se: 135, s: 180, sw: 225, w: 270, nw: 315 }[name] ?? 0;
  const step = Math.round(((((base + (deg || 0)) % 180) + 180) % 180) / 45) % 4;
  return ['ns-resize', 'nesw-resize', 'ew-resize', 'nwse-resize'][step];
}

/**
 * The box a drag of `handle` by (dx, dy) page pixels gives a box turned by
 * `deg`: the travel is read along the box's own sides, the side or corner
 * opposite the handle stays where it is on the page, and the box keeps its
 * turn. `keep` holds the proportions on a corner (the larger change wins),
 * `min` is the smallest side.
 */
export function resizeTurned(box, deg, handle, dx, dy, { keep = false, min = 8 } = {}) {
  const a = (deg || 0) * RAD;
  const c = Math.cos(a);
  const s = Math.sin(a);
  // The pointer's travel turned back into the box's own axes.
  const ux = dx * c + dy * s;
  const uy = -dx * s + dy * c;
  let l = 0;
  let t = 0;
  let w = box.w;
  let h = box.h;
  if (handle.includes('e')) w += ux;
  if (handle.includes('s')) h += uy;
  if (handle.includes('w')) { l += ux; w -= ux; }
  if (handle.includes('n')) { t += uy; h -= uy; }
  if (keep && handle.length === 2 && box.w > 0 && box.h > 0) {
    const ratio = box.w / box.h;
    if (Math.abs(w / box.w) > Math.abs(h / box.h)) h = w / ratio; else w = h * ratio;
    if (handle.includes('w')) l = box.w - w;
    if (handle.includes('n')) t = box.h - h;
  }
  if (w < min) { if (handle.includes('w')) l = box.w - min; w = min; }
  if (h < min) { if (handle.includes('n')) t = box.h - min; h = min; }
  // The new centre, moved in the box's own axes, then turned onto the page:
  // the opposite handle, fixed in the box's axes, is fixed on the page too.
  const sx = l + w / 2 - box.w / 2;
  const sy = t + h / 2 - box.h / 2;
  const cx = box.x + box.w / 2 + sx * c - sy * s;
  const cy = box.y + box.h / 2 + sx * s + sy * c;
  return { x: cx - w / 2, y: cy - h / 2, w, h };
}
