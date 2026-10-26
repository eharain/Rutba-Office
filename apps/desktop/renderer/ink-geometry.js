// Draw: the arithmetic of ink, apart from the window so it can be tested —
// the pens and how each lays its ink, a stroke thinned, a stroke laid along
// the Ruler, and Ink to Shape's reading of a drawn rectangle, oval or
// triangle.

/** The pens the Draw tab starts with, as PowerPoint's gallery does. */
export const DEFAULT_PENS = [
  { id: 'pen', tool: 'pen', color: '#000000', width: 1 },
  { id: 'pencil', tool: 'pencil', color: '#404040', width: 1 },
  { id: 'highlighter', tool: 'highlighter', color: '#FFFF00', width: 8 },
];
export const PEN_COLOURS = ['#000000', '#FFFFFF', '#7F7F7F', '#E81123', '#FF8C00', '#FFE100', '#16A34A', '#0078D4', '#5C2D91', '#E3008C'];
export const PEN_WIDTHS = { pen: [0.5, 1, 2, 3.5, 5], pencil: [0.5, 1, 2, 3.5, 5], highlighter: [4, 6, 8, 12, 16] };

/** How each kind of pen lays its ink: see-through, and the shape of its tip. */
export const strokeLook = (pen) => ({
  color: pen.color,
  width: pen.width,
  alpha: pen.tool === 'highlighter' ? 0.5 : pen.tool === 'pencil' ? 0.85 : 1,
  cap: pen.tool === 'highlighter' ? 'sq' : 'rnd',
});

/** Whether a slide shape is a stroke of ink. */
export const isInk = (s) => s?.kind !== 'group' && /^Ink \d+$/.test(s?.name || '');

const dist = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);

/** A stroke thinned to the points that matter — a point every couple of pixels. */
export function thin(points, min = 1.5) {
  const out = [points[0]];
  for (const p of points.slice(1)) if (dist(p, out[out.length - 1]) >= min) out.push(p);
  if (out.length === 1 && points.length > 1) out.push(points[points.length - 1]);
  return out;
}

/** The ruler's drawing edge as a line: a point on it and its direction. */
function rulerEdge(r) {
  const a = (r.angle * Math.PI) / 180;
  const dir = [Math.cos(a), Math.sin(a)];
  // The edge is the ruler's upper side, its centre line moved up half the ruler's depth.
  const normal = [Math.sin(a), -Math.cos(a)];
  return { origin: [r.x + normal[0] * (r.depth / 2), r.y + normal[1] * (r.depth / 2)], dir };
}

/** A stroke begun along the ruler's edge, laid on the edge as a straight line. */
export function alongRuler(points, ruler, reach = 24) {
  if (!ruler || points.length < 2) return null;
  const { origin, dir } = rulerEdge(ruler);
  const offset = (p) => (p[0] - origin[0]) * -dir[1] + (p[1] - origin[1]) * dir[0];
  if (Math.abs(offset(points[0])) > reach) return null;
  const along = points.map((p) => (p[0] - origin[0]) * dir[0] + (p[1] - origin[1]) * dir[1]);
  const lo = Math.min(...along);
  const hi = Math.max(...along);
  const at = (t) => [origin[0] + dir[0] * t, origin[1] + dir[1] * t];
  return [at(lo), at(hi)];
}

/** A stroke sampled at `n` points evenly along its length. */
function resample(points, n = 64) {
  const total = points.slice(1).reduce((s, p, i) => s + dist(p, points[i]), 0);
  if (!total) return points.slice(0, 1);
  const step = total / (n - 1);
  const out = [points[0]];
  let acc = 0;
  for (let i = 1; i < points.length && out.length < n; i++) {
    let a = points[i - 1];
    const b = points[i];
    let d = dist(a, b);
    while (acc + d >= step && out.length < n) {
      const t = (step - acc) / d;
      a = [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t];
      out.push(a);
      d = dist(a, b);
      acc = 0;
    }
    acc += d;
  }
  while (out.length < n) out.push(points[points.length - 1]);
  return out;
}

/**
 * Ink to Shape: a closed stroke read as the shape it was meant to be — a
 * triangle, a rectangle or an oval, by how many corners it turns — or
 * null for a stroke that is none of them (it stays ink).
 */
export function recognise(points) {
  if (points.length < 8) return null;
  const length = points.slice(1).reduce((s, p, i) => s + dist(p, points[i]), 0);
  if (length < 60 || dist(points[0], points[points.length - 1]) > length * 0.18) return null;
  const xs = points.map((p) => p[0]);
  const ys = points.map((p) => p[1]);
  const box = { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) };
  if (box.w < 16 || box.h < 16) return null;
  // Corners: where the direction turns sharply, counted once each.
  const r = resample(points, 72);
  const heading = (i) => Math.atan2(r[(i + 3) % r.length][1] - r[i][1], r[(i + 3) % r.length][0] - r[i][0]);
  let corners = 0;
  let last = -10;
  for (let i = 0; i < r.length; i++) {
    let turn = Math.abs(heading((i + 3) % r.length) - heading(i));
    if (turn > Math.PI) turn = 2 * Math.PI - turn;
    if (turn > 0.85 && i - last > 6) { corners += 1; last = i; }
  }
  // How far the stroke keeps from an ellipse in its box.
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  const off = r.reduce((s, p) => s + Math.abs(Math.hypot((p[0] - cx) / (box.w / 2), (p[1] - cy) / (box.h / 2)) - 1), 0) / r.length;
  if (corners === 3) return { preset: 'triangle', ...box };
  if (corners >= 4 && corners <= 5 && off > 0.08) return { preset: 'rect', ...box };
  if (off < 0.16) return { preset: 'ellipse', ...box };
  return null;
}

export const polygonHas = (poly, p) => {
  let inside = false;
  for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
    const [xi, yi] = poly[i];
    const [xj, yj] = poly[j];
    if ((yi > p[1]) !== (yj > p[1]) && p[0] < ((xj - xi) * (p[1] - yi)) / (yj - yi) + xi) inside = !inside;
  }
  return inside;
};

