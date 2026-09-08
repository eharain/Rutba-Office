/**
 * Shapes and pictures.
 *
 * The other half of a drawing. A chart is data; a shape is a rectangle someone
 * drew round a total and a picture is the company logo on a quotation. Both are
 * common in real files and both currently render as nothing.
 *
 * The split mirrors charts: a DESCRIPTOR (plain data, parsed at the edge) and a
 * BUILDER (descriptor + a size -> scene). Studio produces descriptors directly
 * without ever touching OOXML; the adapter produces them from a customer's file.
 *
 * Preset geometry is where a shape library could sprawl forever — OOXML defines
 * nearly two hundred presets, most of which nobody uses. The ones below are the
 * ones that appear in real business documents. Anything else falls back to a
 * rectangle of the right size and colour, which reads as "a shape is here" rather
 * than as nothing at all — a wrong-shaped box is a far smaller lie than a blank.
 */
import { group, rect, ellipse, line, polygon, path, text as textNode, roundedBarPath } from './scene.js';
import { theme, seriesColour } from './palette.js';
import { measureText, wrapText, lineHeight } from './measure.js';

/** Presets we draw properly. Everything else becomes a rectangle. */
export const SUPPORTED_GEOMETRY = [
  'rect', 'roundRect', 'ellipse', 'line', 'straightConnector1', 'bentConnector2', 'bentConnector3', 'curvedConnector3',
  'triangle', 'diamond', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow',
  'pentagon', 'hexagon', 'star5', 'plus', 'chevron', 'parallelogram', 'trapezoid',
  'leftBrace', 'rightBrace', 'leftBracket', 'rightBracket',
];

const LINE_GEOMETRY = /^(line|straightConnector\d*|bentConnector\d*|curvedConnector\d*)$/;
const BRACE_GEOMETRY = /^(left|right)(Brace|Bracket)$/;

/**
 * A brace or bracket: a stroke and no fill, the curl at the middle of the
 * long side, opening away from the side it is named for.
 */
function bracePath(geometry, x, y, w, h) {
  const left = geometry.startsWith('left');
  const outer = left ? x + w : x;   // the side the arms reach
  const inner = left ? x : x + w;   // the side the point touches
  const mid = x + w / 2;
  const cy = y + h / 2;
  const f = (n) => Math.round(n * 100) / 100;
  if (geometry.endsWith('Bracket')) return `M${f(outer)} ${f(y)}L${f(inner)} ${f(y)}L${f(inner)} ${f(y + h)}L${f(outer)} ${f(y + h)}`;
  const r = Math.min(w, h / 4);
  return (
    `M${f(outer)} ${f(y)}Q${f(mid)} ${f(y)} ${f(mid)} ${f(y + r)}L${f(mid)} ${f(cy - r)}Q${f(mid)} ${f(cy)} ${f(inner)} ${f(cy)}` +
    `Q${f(mid)} ${f(cy)} ${f(mid)} ${f(cy + r)}L${f(mid)} ${f(y + h - r)}Q${f(mid)} ${f(y + h)} ${f(outer)} ${f(y + h)}`
  );
}


/**
 * A theme colour maps onto OUR palette rather than the customer's theme.
 *
 * We do not read their theme part, so `accent1` becoming our slot 1 keeps a
 * shape coherent with the rest of the product instead of guessing at a colour we
 * never loaded. The original part is untouched, so their theme survives the save.
 */
const SCHEME_SLOTS = {
  accent1: 0, accent2: 1, accent3: 2, accent4: 3, accent5: 4, accent6: 5,
};

export function resolveColour(descriptor, mode = 'light') {
  if (!descriptor) return null;
  if (descriptor.type === 'none') return 'none';
  if (descriptor.type === 'srgb') return '#' + descriptor.value.toLowerCase();
  if (descriptor.type === 'scheme') {
    const t = theme(mode);
    const slot = SCHEME_SLOTS[descriptor.value];
    if (slot !== undefined) return seriesColour(slot, mode);
    if (descriptor.value === 'bg1') return t.ink.surface;
    if (descriptor.value === 'tx1' || descriptor.value === 'dk1') return t.ink.primary;
    if (descriptor.value === 'bg2' || descriptor.value === 'lt2') return t.ink.gridline;
    return t.ink.muted;
  }
  return null;
}

/** Points of a regular-ish polygon inscribed in the box. */
function polygonPoints(geometry, x, y, w, h) {
  const cx = x + w / 2;
  const cy = y + h / 2;
  switch (geometry) {
    case 'triangle':
      return [[cx, y], [x + w, y + h], [x, y + h]];
    case 'diamond':
      return [[cx, y], [x + w, cy], [cx, y + h], [x, cy]];
    case 'parallelogram':
      return [[x + w * 0.25, y], [x + w, y], [x + w * 0.75, y + h], [x, y + h]];
    case 'trapezoid':
      return [[x + w * 0.25, y], [x + w * 0.75, y], [x + w, y + h], [x, y + h]];
    case 'pentagon': {
      const pts = [];
      for (let i = 0; i < 5; i++) {
        const a = -Math.PI / 2 + (i * 2 * Math.PI) / 5;
        pts.push([cx + Math.cos(a) * (w / 2), cy + Math.sin(a) * (h / 2)]);
      }
      return pts;
    }
    case 'hexagon':
      return [
        [x + w * 0.25, y], [x + w * 0.75, y], [x + w, cy],
        [x + w * 0.75, y + h], [x + w * 0.25, y + h], [x, cy],
      ];
    case 'star5': {
      const pts = [];
      for (let i = 0; i < 10; i++) {
        const a = -Math.PI / 2 + (i * Math.PI) / 5;
        const r = i % 2 === 0 ? 0.5 : 0.21;
        pts.push([cx + Math.cos(a) * w * r, cy + Math.sin(a) * h * r]);
      }
      return pts;
    }
    case 'plus': {
      const t = 0.32;
      return [
        [x + w * t, y], [x + w * (1 - t), y], [x + w * (1 - t), y + h * t],
        [x + w, y + h * t], [x + w, y + h * (1 - t)], [x + w * (1 - t), y + h * (1 - t)],
        [x + w * (1 - t), y + h], [x + w * t, y + h], [x + w * t, y + h * (1 - t)],
        [x, y + h * (1 - t)], [x, y + h * t], [x + w * t, y + h * t],
      ];
    }
    case 'chevron':
      return [
        [x, y], [x + w * 0.72, y], [x + w, cy], [x + w * 0.72, y + h],
        [x, y + h], [x + w * 0.28, cy],
      ];
    case 'rightArrow':
      return [
        [x, y + h * 0.3], [x + w * 0.6, y + h * 0.3], [x + w * 0.6, y],
        [x + w, cy], [x + w * 0.6, y + h], [x + w * 0.6, y + h * 0.7], [x, y + h * 0.7],
      ];
    case 'leftArrow':
      return [
        [x + w, y + h * 0.3], [x + w * 0.4, y + h * 0.3], [x + w * 0.4, y],
        [x, cy], [x + w * 0.4, y + h], [x + w * 0.4, y + h * 0.7], [x + w, y + h * 0.7],
      ];
    case 'upArrow':
      return [
        [x + w * 0.3, y + h], [x + w * 0.3, y + h * 0.4], [x, y + h * 0.4],
        [cx, y], [x + w, y + h * 0.4], [x + w * 0.7, y + h * 0.4], [x + w * 0.7, y + h],
      ];
    case 'downArrow':
      return [
        [x + w * 0.3, y], [x + w * 0.3, y + h * 0.6], [x, y + h * 0.6],
        [cx, y + h], [x + w, y + h * 0.6], [x + w * 0.7, y + h * 0.6], [x + w * 0.7, y],
      ];
    default:
      return null;
  }
}

/**
 * Build a shape scene fragment.
 *
 * @param {object} descriptor  from parseShapeXml, or written directly by Studio
 * @param {object} box         { x, y, width, height }
 */
export function buildShape(descriptor, box, { mode = 'light' } = {}) {
  const t = theme(mode);
  const { x = 0, y = 0, width = 100, height = 60 } = box ?? {};
  const geometry = descriptor.geometry ?? 'rect';

  const fill = resolveColour(descriptor.fill, mode) ?? seriesColour(0, mode);
  const stroke = resolveColour(descriptor.stroke, mode);
  const strokeWidth = descriptor.strokeWidth ?? (stroke && stroke !== 'none' ? 1 : 0);
  const paint = { fill: fill === 'none' ? 'none' : fill, stroke: stroke === 'none' ? null : stroke, strokeWidth: strokeWidth || null };

  const children = [];
  if (geometry === 'ellipse') {
    children.push(ellipse({ cx: x + width / 2, cy: y + height / 2, rx: width / 2, ry: height / 2, ...paint }));
  } else if (LINE_GEOMETRY.test(geometry)) {
    // Corner to corner of the box; a flip says which corners. A bent
    // connector goes half way across, turns, and finishes level.
    const x1 = descriptor.flipH ? x + width : x;
    const x2 = descriptor.flipH ? x : x + width;
    const y1 = descriptor.flipV ? y + height : y;
    const y2 = descriptor.flipV ? y : y + height;
    const strokeColour = (stroke && stroke !== 'none') ? stroke : fill;
    if (/^bentConnector/.test(geometry)) {
      const mx = (x1 + x2) / 2;
      children.push(path({ d: `M${x1} ${y1}L${mx} ${y1}L${mx} ${y2}L${x2} ${y2}`, fill: 'none', stroke: strokeColour, strokeWidth: strokeWidth || 1 }));
    } else {
      children.push(line({ x1, y1, x2, y2, stroke: strokeColour, strokeWidth: strokeWidth || 1, linecap: 'round' }));
    }
  } else if (BRACE_GEOMETRY.test(geometry)) {
    children.push(path({
      d: bracePath(geometry, x, y, width, height), fill: 'none',
      stroke: (stroke && stroke !== 'none') ? stroke : fill, strokeWidth: strokeWidth || 1,
    }));
  } else if (geometry === 'roundRect') {

    const radius = Math.min(width, height) * 0.14;
    children.push(path({ d: roundedRectPath(x, y, width, height, radius), ...paint }));
  } else {
    const points = polygonPoints(geometry, x, y, width, height);
    if (points) children.push(polygon({ points, ...paint }));
    else children.push(rect({ x, y, width, height, ...paint }));
  }

  // Shape text is centred in the box and wrapped, the way a text box behaves.
  if (descriptor.text) {
    const size = descriptor.textSize ?? t.font.label + 1;
    const colour = resolveColour(descriptor.textColour, mode) ?? t.ink.surface;
    const lines = wrapText(descriptor.text, Math.max(16, width - 12), { size });
    const lh = lineHeight(size);
    const startY = y + height / 2 - ((lines.length - 1) * lh) / 2;
    lines.forEach((lineText, i) => {
      children.push(textNode({
        x: x + width / 2, y: startY + i * lh, value: lineText,
        size, fill: colour, anchor: 'middle', baseline: 'middle',
        weight: descriptor.textBold ? '600' : null,
      }));
    });
  }

  // A rotated shape turns about its own centre — a brace laid on its side
  // under a box is a vertical brace with rot="5400000".
  const rotation = Number(descriptor.rotation) || 0;
  return group(children, {
    class: 'shape',
    ...(rotation ? { transform: `rotate(${Math.round(rotation * 100) / 100} ${x + width / 2} ${y + height / 2})` } : {}),
  });
}


/** Path data is a string, so the renderer's number formatting cannot reach it.
 *  Round here or a rounded rectangle ships with 17 digits per corner. */
const r2 = (v) => String(Math.round(Number(v) * 100) / 100);

function roundedRectPath(x, y, w, h, r) {
  const radius = Math.max(0, Math.min(r, Math.min(w, h) / 2));
  return `M${r2(x + radius)} ${r2(y)}h${r2(w - 2 * radius)}a${r2(radius)} ${r2(radius)} 0 0 1 ${r2(radius)} ${r2(radius)}`
    + `v${r2(h - 2 * radius)}a${r2(radius)} ${r2(radius)} 0 0 1 ${r2(-radius)} ${r2(radius)}`
    + `h${r2(-(w - 2 * radius))}a${r2(radius)} ${r2(radius)} 0 0 1 ${r2(-radius)} ${r2(-radius)}`
    + `v${r2(-(h - 2 * radius))}a${r2(radius)} ${r2(radius)} 0 0 1 ${r2(radius)} ${r2(-radius)}Z`;
}

/** A picture is one image node; the descriptor already carries the href. */
export function buildPicture(descriptor, box) {
  const { x = 0, y = 0, width = 100, height = 60 } = box ?? {};
  if (!descriptor?.href) return group([], { class: 'picture' });
  return group([{
    type: 'image', x, y, width, height, href: descriptor.href,
    opacity: descriptor.opacity,
  }], { class: 'picture' });
}

/** Common image types, so a media part can become a data: URI. */
export const MEDIA_TYPES = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
  bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tiff: 'image/tiff',
};

/** Bytes -> data URI. `bytes` may be a Buffer or a Uint8Array. */
export function toDataUri(bytes, partName) {
  if (!bytes) return null;
  const ext = String(partName ?? '').split('.').pop().toLowerCase();
  const type = MEDIA_TYPES[ext];
  if (!type) return null;
  const base64 = typeof Buffer !== 'undefined' && Buffer.isBuffer(bytes)
    ? bytes.toString('base64')
    : btoa(String.fromCharCode(...new Uint8Array(bytes)));
  return 'data:' + type + ';base64,' + base64;
}

export { polygonPoints, roundedRectPath, roundedBarPath, measureText };
