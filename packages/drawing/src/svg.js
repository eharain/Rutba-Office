/**
 * Scene -> SVG, as a string.
 *
 * A string rather than DOM nodes, on purpose: the same renderer then works in a
 * browser, on a server, inside an email pipeline and in a test. SVG is the right
 * first target because it is vector, themeable, accessible and text-searchable.
 *
 * EMAIL, HONESTLY. Several major mail clients — Outlook on Windows above all —
 * do not render SVG at all. So for Mail this output is the SOURCE, not the
 * delivery format: it needs a raster step to PNG, which needs a rendering engine
 * we do not have in-process. `svgDataUri()` is here for the clients that do cope,
 * and `RASTER_REQUIRED_NOTE` exists so nobody ships an email chart assuming SVG
 * is enough. That gap is real and unsolved; see FORMAT-FIDELITY.md.
 */
import { INK } from './palette.js';

export const RASTER_REQUIRED_NOTE =
  'Outlook on Windows and several other mail clients do not render SVG. ' +
  'An emailed chart must be rasterised to PNG before sending; this renderer produces the source, not the delivery format.';

const esc = (s) =>
  String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');

/** Compact, deterministic numbers — a scene must render byte-identically twice. */
const n = (v) => {
  if (v === null || v === undefined) return '0';
  const x = Number(v);
  if (!Number.isFinite(x)) return '0';
  const rounded = Math.round(x * 100) / 100;
  return String(Object.is(rounded, -0) ? 0 : rounded);
};

const attr = (name, value) =>
  value === null || value === undefined || value === '' ? '' : ' ' + name + '="' + esc(value) + '"';
const numAttr = (name, value) =>
  value === null || value === undefined ? '' : ' ' + name + '="' + n(value) + '"';

function paintAttrs(node) {
  return attr('fill', node.fill)
    + attr('stroke', node.stroke)
    + numAttr('stroke-width', node.strokeWidth)
    + attr('stroke-dasharray', node.dash)
    + attr('stroke-linecap', node.linecap)
    + attr('stroke-linejoin', node.linejoin)
    + (node.opacity === undefined ? '' : numAttr('opacity', node.opacity))
    + attr('class', node.class);
}

const BASELINE = { alphabetic: 'alphabetic', middle: 'central', hanging: 'hanging' };

function renderNode(node, theme) {
  if (!node) return '';
  switch (node.type) {
    case 'group': {
      const inner = (node.children ?? []).map((c) => renderNode(c, theme)).join('');
      if (!inner) return '';
      const open = '<g' + attr('transform', node.transform) + attr('class', node.class)
        + (node.opacity === undefined ? '' : numAttr('opacity', node.opacity)) + '>';
      return open + inner + '</g>';
    }
    case 'rect':
      return '<rect' + numAttr('x', node.x) + numAttr('y', node.y)
        + numAttr('width', Math.max(0, node.width)) + numAttr('height', Math.max(0, node.height))
        + numAttr('rx', node.rx) + numAttr('ry', node.ry) + paintAttrs(node) + '/>';
    case 'ellipse':
      return '<ellipse' + numAttr('cx', node.cx) + numAttr('cy', node.cy)
        + numAttr('rx', node.rx) + numAttr('ry', node.ry) + paintAttrs(node) + '/>';
    case 'line':
      return '<line' + numAttr('x1', node.x1) + numAttr('y1', node.y1)
        + numAttr('x2', node.x2) + numAttr('y2', node.y2) + paintAttrs(node) + '/>';
    case 'polyline':
      return '<polyline' + attr('points', node.points.map((p) => n(p[0]) + ',' + n(p[1])).join(' '))
        + paintAttrs(node) + '/>';
    case 'polygon':
      return '<polygon' + attr('points', node.points.map((p) => n(p[0]) + ',' + n(p[1])).join(' '))
        + paintAttrs(node) + '/>';
    case 'path':
      return '<path' + attr('d', node.d) + paintAttrs(node) + '/>';
    case 'text': {
      if (node.value === null || node.value === undefined || node.value === '') return '';
      return '<text' + numAttr('x', node.x) + numAttr('y', node.y)
        + attr('text-anchor', node.anchor === 'start' ? null : node.anchor)
        + attr('dominant-baseline', node.baseline === 'alphabetic' ? null : BASELINE[node.baseline])
        + attr('font-family', node.family ?? theme.fontFamily)
        + numAttr('font-size', node.size)
        + attr('font-weight', node.weight)
        + attr('font-style', node.style)
        + attr('fill', node.fill)
        + attr('class', node.class)
        + '>' + esc(node.value) + '</text>';
    }
    case 'image':
      return '<image' + numAttr('x', node.x) + numAttr('y', node.y)
        + numAttr('width', node.width) + numAttr('height', node.height)
        + attr('href', node.href)
        + (node.opacity === undefined ? '' : numAttr('opacity', node.opacity)) + '/>';
    default:
      return '';
  }
}

/**
 * @param {object} sceneNode
 * @param {object} [opts]
 * @param {boolean} [opts.standalone=true]  include xmlns, so it works as a file or data: URI
 * @param {string}  [opts.fontFamily]
 */
export function renderSvg(sceneNode, { standalone = true, fontFamily } = {}) {
  if (!sceneNode || sceneNode.type !== 'scene') throw new Error('renderSvg expects a scene');
  const ink = INK[sceneNode.mode] ?? INK.light;
  const theme = { fontFamily: fontFamily ?? 'system-ui, -apple-system, "Segoe UI", sans-serif' };

  const parts = [];
  parts.push('<svg'
    + (standalone ? ' xmlns="http://www.w3.org/2000/svg"' : '')
    + numAttr('width', sceneNode.width)
    + numAttr('height', sceneNode.height)
    + attr('viewBox', '0 0 ' + n(sceneNode.width) + ' ' + n(sceneNode.height))
    + attr('role', 'img')
    + attr('aria-label', sceneNode.title)
    + '>');

  // Accessibility first, so a screen reader meets it before the marks.
  if (sceneNode.title) parts.push('<title>' + esc(sceneNode.title) + '</title>');
  if (sceneNode.description) parts.push('<desc>' + esc(sceneNode.description) + '</desc>');

  const background = sceneNode.background === undefined ? ink.surface : sceneNode.background;
  if (background && background !== 'none') {
    parts.push('<rect x="0" y="0"' + numAttr('width', sceneNode.width)
      + numAttr('height', sceneNode.height) + attr('fill', background) + '/>');
  }
  for (const child of sceneNode.children ?? []) parts.push(renderNode(child, theme));
  parts.push('</svg>');
  return parts.join('');
}

/** For an `<img src>` or a CSS url(). Base64 rather than percent-encoding: it
 *  survives more email pipelines intact, at a small size cost. */
export function svgDataUri(sceneNode, opts) {
  const svg = renderSvg(sceneNode, opts);
  const base64 = typeof Buffer !== 'undefined'
    ? Buffer.from(svg, 'utf8').toString('base64')
    : btoa(unescape(encodeURIComponent(svg)));
  return 'data:image/svg+xml;base64,' + base64;
}

export { esc as escapeXml, n as formatNumber };
