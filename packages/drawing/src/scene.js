/**
 * The scene graph — plain data, deliberately.
 *
 * A scene is JSON. It can be stored, diffed, sent over a wire, rendered on a
 * server or in a browser, and produced by a spreadsheet, a design tool or an
 * email composer without any of them knowing about the others. That is the whole
 * reason this is a data structure rather than a set of draw calls.
 *
 * Style lives on the node. There is no cascade, no inheritance and no stylesheet
 * to reason about — a node says what it looks like. That makes scenes verbose
 * and makes them trivially correct to render, which is the right trade for
 * something three products will consume.
 *
 * Coordinates are a plain top-left origin in abstract units; the renderer
 * decides what a unit is. OOXML uses EMUs, the browser uses pixels, and the
 * adapters convert at the edge rather than leaking either into the model.
 */

/** English Metric Units — OOXML's unit. 914400 per inch, 12700 per point. */
export const EMU_PER_PIXEL = 9525;
export const EMU_PER_POINT = 12700;
export const emuToPx = (emu) => Number(emu) / EMU_PER_PIXEL;
export const pxToEmu = (px) => Math.round(Number(px) * EMU_PER_PIXEL);

export const group = (children = [], props = {}) => ({ type: 'group', children, ...props });

export const rect = ({ x, y, width, height, rx, ry, fill, stroke, strokeWidth, opacity, ...rest }) =>
  ({ type: 'rect', x, y, width, height, rx, ry, fill, stroke, strokeWidth, opacity, ...rest });

export const ellipse = ({ cx, cy, rx, ry, fill, stroke, strokeWidth, opacity, ...rest }) =>
  ({ type: 'ellipse', cx, cy, rx, ry, fill, stroke, strokeWidth, opacity, ...rest });

export const line = ({ x1, y1, x2, y2, stroke, strokeWidth, dash, opacity, ...rest }) =>
  ({ type: 'line', x1, y1, x2, y2, stroke, strokeWidth, dash, opacity, ...rest });

export const polyline = ({ points, stroke, strokeWidth, fill = 'none', dash, ...rest }) =>
  ({ type: 'polyline', points, stroke, strokeWidth, fill, dash, ...rest });

export const polygon = ({ points, fill, stroke, strokeWidth, opacity, ...rest }) =>
  ({ type: 'polygon', points, fill, stroke, strokeWidth, opacity, ...rest });

export const path = ({ d, fill, stroke, strokeWidth, opacity, ...rest }) =>
  ({ type: 'path', d, fill, stroke, strokeWidth, opacity, ...rest });

/**
 * @param {object} spec
 * @param {'start'|'middle'|'end'} [spec.anchor='start']
 * @param {'alphabetic'|'middle'|'hanging'} [spec.baseline='alphabetic']
 */
export const text = ({
  x, y, value, fill, size, family, weight, style, anchor = 'start', baseline = 'alphabetic', ...rest
}) => ({ type: 'text', x, y, value, fill, size, family, weight, style, anchor, baseline, ...rest });

/** An image. `href` may be a URL or a data: URI, which is how email carries one. */
export const image = ({ x, y, width, height, href, opacity, ...rest }) =>
  ({ type: 'image', x, y, width, height, href, opacity, ...rest });

/**
 * A whole drawing: a size and a root group.
 * `title` and `description` become the SVG accessibility elements, which is the
 * only reason a screen reader will make anything of a chart.
 */
export const scene = ({ width, height, children = [], background, title, description, mode = 'light' }) => ({
  type: 'scene',
  width,
  height,
  background,
  title,
  description,
  mode,
  children,
});

/** A rounded-rectangle path with per-corner control — bars round only their data end. */
const rd = (v) => Math.round(Number(v) * 100) / 100;

export function roundedBarPath({ x, y, width, height, radius = 0, side = 'top' }) {
  const r = rd(Math.max(0, Math.min(radius, Math.min(Math.abs(width), Math.abs(height)) / 2)));
  x = rd(x); y = rd(y); width = rd(width); height = rd(height);
  if (r === 0) return `M${x} ${y}h${width}v${height}h${-width}Z`;
  switch (side) {
    case 'top':
      return `M${x} ${rd(y + height)}V${rd(y + r)}a${r} ${r} 0 0 1 ${r} ${-r}h${rd(width - 2 * r)}a${r} ${r} 0 0 1 ${r} ${r}V${rd(y + height)}Z`;
    case 'bottom':
      return `M${x} ${y}V${rd(y + height - r)}a${r} ${r} 0 0 0 ${r} ${r}h${rd(width - 2 * r)}a${r} ${r} 0 0 0 ${r} ${-r}V${y}Z`;
    case 'right':
      return `M${x} ${y}h${rd(width - r)}a${r} ${r} 0 0 1 ${r} ${r}v${rd(height - 2 * r)}a${r} ${r} 0 0 1 ${-r} ${r}H${x}Z`;
    case 'left':
      return `M${rd(x + width)} ${y}H${rd(x + r)}a${r} ${r} 0 0 0 ${-r} ${r}v${rd(height - 2 * r)}a${r} ${r} 0 0 0 ${r} ${r}H${rd(x + width)}Z`;
    default:
      return `M${x} ${y}h${width}v${height}h${-width}Z`;
  }
}

/** Axis-aligned bounds of a node, for hit-testing and for fitting a scene. */
export function bounds(node) {
  if (!node) return null;
  const merge = (a, b) => {
    if (!a) return b;
    if (!b) return a;
    return {
      x: Math.min(a.x, b.x),
      y: Math.min(a.y, b.y),
      right: Math.max(a.right, b.right),
      bottom: Math.max(a.bottom, b.bottom),
    };
  };
  const box = (x, y, w, h) => ({ x, y, right: x + w, bottom: y + h });

  switch (node.type) {
    case 'scene':
    case 'group': {
      let acc = null;
      for (const child of node.children ?? []) acc = merge(acc, bounds(child));
      return acc;
    }
    case 'rect':
    case 'image':
      return box(node.x, node.y, node.width, node.height);
    case 'ellipse':
      return box(node.cx - node.rx, node.cy - node.ry, node.rx * 2, node.ry * 2);
    case 'line':
      return {
        x: Math.min(node.x1, node.x2), y: Math.min(node.y1, node.y2),
        right: Math.max(node.x1, node.x2), bottom: Math.max(node.y1, node.y2),
      };
    case 'polyline':
    case 'polygon': {
      const xs = node.points.map((p) => p[0]);
      const ys = node.points.map((p) => p[1]);
      return { x: Math.min(...xs), y: Math.min(...ys), right: Math.max(...xs), bottom: Math.max(...ys) };
    }
    case 'text':
      // Text bounds need metrics; measure.js does that. This is the anchor point.
      return box(node.x, node.y, 0, 0);
    default:
      return null;
  }
}

/** Walk every node, depth-first. Useful for retargeting colours or hit-testing. */
export function* walk(node) {
  if (!node) return;
  yield node;
  for (const child of node.children ?? []) yield* walk(child);
}

/** Count drawable marks, so a caller can refuse to render something absurd. */
export function markCount(node) {
  let n = 0;
  for (const child of walk(node)) if (child.type !== 'group' && child.type !== 'scene') n += 1;
  return n;
}
